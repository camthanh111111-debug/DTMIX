from __future__ import annotations

import copy
import html
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, Callable

import docx
from docx.oxml.ns import qn

import headless_compat
from dtmix_legacy import DTMIXApp


class ValueBox:
    """Tk-like value wrapper used by the legacy algorithm."""
    def __init__(self, value: Any = None):
        self.value = value
        self._original_val = str(value) if value is not None else ""
    def get(self):
        return self.value
    def set(self, value):
        self.value = value
    def insert(self, index, value):
        self.value = str(value)
    def delete(self, *args, **kwargs):
        self.value = ""
    def select(self):
        self.value = True
    def deselect(self):
        self.value = False
    def configure(self, *args, **kwargs):
        if "text" in kwargs:
            self.value = kwargs["text"]
        return self
    config = configure
    def pack(self, *args, **kwargs): return self
    def pack_forget(self, *args, **kwargs): return self
    def grid(self, *args, **kwargs): return self
    def bind(self, *args, **kwargs): return self
    def winfo_exists(self): return True
    def index(self, *args, **kwargs): return 0
    def icursor(self, *args, **kwargs): return None


class DummyWidget(ValueBox):
    pass


@dataclass
class MixResult:
    zip_bytes: bytes
    zip_name: str
    files: list[str]
    output_dir: str


class DTMIXWebEngine:
    """
    Headless bridge around the original DTMIX 1.3 processing algorithms.
    It intentionally does not call DTMIXApp.__init__(), because that constructor builds a desktop GUI.
    """

    def __init__(
        self,
        file_bytes: bytes,
        filename: str,
        youngmix: bool = False,
        header: dict[str, str] | None = None,
        progress_callback: Callable[[int, int, str], None] | None = None,
    ):
        self.file_bytes = bytes(file_bytes)
        self.filename = Path(filename).name or "de_goc.docx"
        if not self.filename.lower().endswith(".docx"):
            raise ValueError("DTMIX chỉ hỗ trợ tệp Word .docx")
        self.youngmix = bool(youngmix)
        self.progress_callback = progress_callback
        self.work_dir = Path(tempfile.mkdtemp(prefix="dtmix_streamlit_"))
        self.source_path = self.work_dir / self.filename
        self.source_path.write_bytes(self.file_bytes)
        self.output_base = self.work_dir / "outputs"
        self.output_base.mkdir(parents=True, exist_ok=True)
        self.last_output_dir: str | None = None
        self.header = header or {}
        self.app = self._make_headless_app()
        self.analyze()

    def close(self):
        shutil.rmtree(self.work_dir, ignore_errors=True)

    def _make_headless_app(self):
        # Bypass the desktop constructor and initialize only state required by core algorithms.
        a = object.__new__(DTMIXApp)
        a.filepath = str(self.source_path)
        a.output_dir = str(self.output_base)
        a.parsed_data = {"header": [], "parts": []}
        a.muc_entries = {}
        a.muc_shuffle_vars = {}
        a.muc_fix_entries = {}
        a.muc_order_shuffle_vars = {}
        a.errors = []
        a.dynamic_wrap_labels = []
        a.youngmix_var = ValueBox(self.youngmix)
        a.youngmix_data = []
        a.youngmix_total_q = 0
        a.ym_ui_controls = []
        a.master_fix_var = ValueBox(False)
        a.continuous_numbering_var = ValueBox(False)
        a.keep_muc_var = ValueBox(False)
        a.auto_code_var = ValueBox(True)
        a.auto_code_var_ym = ValueBox(True)
        a.thu_cong_var = ValueBox(True)
        a.thu_cong_var_ym = ValueBox(True)
        a.entry_made = ValueBox("111, 222, 333, 444")
        a.entry_made_ym = ValueBox("111, 222, 333, 444")
        a.combo_somade = ValueBox("4")
        a.combo_somade_ym = ValueBox("4")
        a.btn_tron = DummyWidget()
        a.btn_tron_ym = DummyWidget()
        a.progress_bar = DummyWidget(0)
        a.progress_bar_ym = DummyWidget(0)
        a.lbl_status = DummyWidget("")
        a.lbl_status_ym = DummyWidget("")
        a.file_bytes_in_ram = self.file_bytes
        a.global_doc = None

        defaults = {
            "so": "SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH",
            "truong": "TRƯỜNG THPT DƯƠNG BẠCH MAI",
            "kythi": "KIỂM TRA HỌC KỲ II",
            "namhoc": "NĂM HỌC 2025 - 2026",
            "monthi": "Môn: HÓA HỌC",
            "thoigian": "Thời gian làm bài: 45 phút",
        }
        defaults.update({k: v for k, v in self.header.items() if v is not None})
        a.val_so = ValueBox(defaults["so"])
        a.val_truong = ValueBox(defaults["truong"])
        a.val_kythi = ValueBox(defaults["kythi"])
        a.val_namhoc = ValueBox(defaults["namhoc"])
        a.val_mon = ValueBox(defaults["monthi"])
        a.val_thoigian = ValueBox(defaults["thoigian"])

        # Headless event/UI hooks.
        a.after = lambda delay, callback=None, *args: callback(*args) if callback else None
        a.update = lambda *args, **kwargs: None
        a.update_idletasks = lambda *args, **kwargs: None
        a.show_success_dialog = self._capture_success
        a.update_progress = self._progress
        return a

    def _capture_success(self, out_dir):
        self.last_output_dir = str(out_dir)

    def _progress(self, current, total, text):
        if self.progress_callback:
            self.progress_callback(int(current), int(total), str(text))

    def analyze(self):
        a = self.app
        doc = docx.Document(str(self.source_path))
        # Same preprocessing sequence as the desktop browse_file() method.
        a.convert_auto_numbering(doc)
        a.convert_floating_to_inline(doc)
        a.split_soft_returns(doc)
        a.preprocess_textboxes(doc)
        a.preprocess_layout_tables(doc)
        a.global_doc = doc
        a.parsed_data, a.errors = a.parse_document_structure(doc)
        if self.youngmix:
            a.parse_youngmix_structure(doc)
            self._detect_youngmix_group_types()
            self._setup_youngmix_defaults()
        else:
            self._setup_standard_defaults()
        return self.summary()

    def _setup_standard_defaults(self):
        a = self.app
        a.muc_entries = {}
        a.muc_shuffle_vars = {}
        a.muc_fix_entries = {}
        a.muc_order_shuffle_vars = {}
        for p_idx, part in enumerate(a.parsed_data.get("parts", [])):
            is_part4 = part.get("type") == 4
            multi = len(part.get("mucs", [])) > 1
            for m_idx, muc in enumerate(part.get("mucs", [])):
                q_count = len(muc.get("questions", []))
                fixed = [str(i + 1) for i, q in enumerate(muc.get("questions", [])) if q.get("is_fixed")]
                a.muc_entries[(p_idx, m_idx)] = ValueBox(str(q_count))
                a.muc_shuffle_vars[(p_idx, m_idx)] = ValueBox(not is_part4)
                a.muc_fix_entries[(p_idx, m_idx)] = ValueBox(", ".join(fixed))
                if multi:
                    a.muc_order_shuffle_vars[(p_idx, m_idx)] = ValueBox(False if is_part4 else not muc.get("is_fixed", False))

    def _detect_youngmix_group_types(self):
        a = self.app
        all_mucs = [m for p in a.parsed_data.get("parts", []) for m in p.get("mucs", [])]
        for m_idx, g in enumerate(a.youngmix_data):
            if m_idx >= len(all_mucs):
                continue
            muc = all_mucs[m_idx]
            q_type1 = q_type2 = q_essay = q_short = 0
            for q in muc.get("questions", []):
                q_type = q.get("ym_type", 4)
                num_ans = len(q.get("answers", []))
                if num_ans == 0:
                    if q_type == 3: q_short += 1
                    else: q_essay += 1
                elif num_ans >= 2:
                    if q_type == 2: q_type2 += 1
                    else: q_type1 += 1
            if q_essay > 0 and q_type1 == q_type2 == q_short == 0:
                g["q_type"] = "Tự luận"
                if g.get("mix_type") == "Câu hỏi & đáp án": g["mix_type"] = "Chỉ câu hỏi"
            elif q_type2 > 0 and q_type1 == q_essay == q_short == 0:
                g["q_type"] = "TN 2025 - Phần 2"
            elif q_short > 0 and q_type1 == q_type2 == q_essay == 0:
                g["q_type"] = "TN 2025 - Phần 3"
                if g.get("mix_type") == "Câu hỏi & đáp án": g["mix_type"] = "Chỉ câu hỏi"
            elif q_type1 > 0:
                g["q_type"] = "TN 2025 - Phần 1"

    def _setup_youngmix_defaults(self):
        a = self.app
        a.ym_ui_controls = []
        for i, group in enumerate(a.youngmix_data):
            a.ym_ui_controls.append({
                "type": ValueBox(group.get("q_type", "Trắc nghiệm")),
                "mix": ValueBox(group.get("mix_type", "Câu hỏi & đáp án")),
                "fix": ValueBox(bool(group.get("is_fixed", False))),
                "start_q1": ValueBox(True),
                "pick": ValueBox(str(group.get("q_count", 0))),
            })

    @staticmethod
    def _part_label(part: dict) -> str:
        title = (part.get("title") or "").strip()
        if title:
            return title
        p_type = part.get("type", 1)
        names = {1: "PHẦN I", 2: "PHẦN II", 3: "PHẦN III", 4: "PHẦN IV"}
        return names.get(p_type, f"PHẦN {p_type}")

    def summary(self) -> dict[str, Any]:
        a = self.app
        parts = []
        total_q = total_valid = total_missing = 0
        for p_idx, part in enumerate(a.parsed_data.get("parts", [])):
            p_type = part.get("type", 1)
            groups = []
            p_count = p_valid = p_missing = 0
            for m_idx, muc in enumerate(part.get("mucs", [])):
                q_items = []
                for q_idx, q in enumerate(muc.get("questions", [])):
                    if q.get("is_virtual"):
                        continue
                    p_count += 1; total_q += 1
                    num_ans = len(q.get("answers", []))
                    q_type = q.get("ym_type", p_type)
                    valid = self._question_has_valid_answer(q, q_type if self.youngmix else p_type)
                    if valid:
                        p_valid += 1; total_valid += 1
                    else:
                        p_missing += 1; total_missing += 1
                    q_items.append({
                        "index": q_idx + 1,
                        "raw_text": q.get("raw_text", ""),
                        "answer_count": num_ans,
                        "correct_count": sum(1 for x in q.get("answers", []) if x.get("is_true")),
                        "fixed": bool(q.get("is_fixed")),
                        "q_type": q_type,
                        "valid_answer": valid,
                    })
                groups.append({
                    "p_idx": p_idx, "m_idx": m_idx,
                    "title": muc.get("title") or f"Nhóm {m_idx + 1}",
                    "question_count": len([q for q in muc.get("questions", []) if not q.get("is_virtual")]),
                    "is_fixed": bool(muc.get("is_fixed")),
                    "questions": q_items,
                })
            parts.append({
                "p_idx": p_idx,
                "title": self._part_label(part),
                "type": p_type,
                "question_count": p_count,
                "valid_count": p_valid,
                "missing_count": p_missing,
                "groups": groups,
            })
        return {
            "filename": self.filename,
            "youngmix": self.youngmix,
            "total_questions": total_q,
            "valid_answers": total_valid,
            "missing_answers": total_missing,
            "parts": parts,
            "youngmix_groups": self.youngmix_groups(),
        }

    def _question_has_valid_answer(self, q: dict, q_type: int) -> bool:
        answers = q.get("answers", [])
        if q_type == 1:
            return sum(1 for x in answers if x.get("is_true")) == 1
        if q_type == 2:
            # In DTMIX, unmarked choices are interpreted as Sai, so a parsed 4-option D/S question is valid.
            return len(answers) >= 2
        if q_type == 3:
            for elm in q.get("q_elements", []):
                if elm.tag == qn("w:p"):
                    _, txt = self.app._build_mapping(elm)
                    m = re.search(r"(?i)(đáp\s+án\s*[:\.]|(?:^|\n|\uFFFC)\s*A\.)\s*(.*)", txt)
                    if m and re.search(r"\d", m.group(2)):
                        return True
            return False
        return True  # essay

    def youngmix_groups(self) -> list[dict[str, Any]]:
        if not self.youngmix:
            return []
        out = []
        for i, g in enumerate(self.app.youngmix_data):
            preview = "".join(chunk[0] for chunk in g.get("preview_data", []))
            out.append({
                "index": i,
                "name": g.get("name", f"NHÓM {i+1}"),
                "tag": g.get("tag", "<g>"),
                "q_type": g.get("q_type", "Trắc nghiệm"),
                "mix_type": g.get("mix_type", "Câu hỏi & đáp án"),
                "is_fixed": bool(g.get("is_fixed", False)),
                "question_count": int(g.get("q_count", 0)),
                "preview": preview.strip(),
                "preview_chunks": g.get("preview_data", []),
            })
        return out

    def preview_html(self, max_chars: int = 25000) -> str:
        """Readable preview preserving correct-answer red and underline hints in YoungMix mode."""
        if self.youngmix:
            chunks = []
            used = 0
            for g in self.app.youngmix_data:
                chunks.append(f"<h4>{html.escape(g.get('name','Nhóm'))}</h4>")
                for text, is_red, is_under in g.get("preview_data", []):
                    if used >= max_chars: break
                    safe = html.escape(text[: max_chars-used]).replace("\n", "<br>")
                    used += len(text)
                    style = []
                    if is_red: style.append("color:#dc2626;font-weight:600")
                    elif is_under: style.append("text-decoration:underline")
                    chunks.append(f"<span style=\"{';'.join(style)}\">{safe}</span>")
            return "".join(chunks)
        # Standard preview: reconstruct text from parsed XML mappings.
        chunks = []
        used = 0
        for part in self.app.parsed_data.get("parts", []):
            chunks.append(f"<h4>{html.escape(self._part_label(part))}</h4>")
            for muc in part.get("mucs", []):
                if muc.get("title"):
                    chunks.append(f"<b>{html.escape(str(muc['title']))}</b><br>")
                for q in muc.get("questions", []):
                    if q.get("is_virtual"): continue
                    qt = html.escape(q.get("raw_text", ""))
                    chunks.append(qt + "<br>")
                    used += len(qt)
                    for ans in q.get("answers", []):
                        txt = ""
                        for elm in ans.get("elements", []):
                            if elm.tag == qn("w:p"):
                                _, t = self.app._build_mapping(elm)
                                txt += t + "\n"
                        style = "color:#dc2626;font-weight:600" if ans.get("is_true") else ""
                        chunks.append(f"<span style=\"{style}\">{html.escape(txt).replace(chr(10),'<br>')}</span>")
                        used += len(txt)
                    chunks.append("<br>")
                    if used >= max_chars:
                        chunks.append("<i>…Đã rút gọn phần xem trước…</i>")
                        return "".join(chunks)
        return "".join(chunks)

    def set_header(self, header: dict[str, str]):
        mapping = {
            "so": "val_so", "truong": "val_truong", "kythi": "val_kythi",
            "namhoc": "val_namhoc", "monthi": "val_mon", "thoigian": "val_thoigian",
        }
        for key, attr in mapping.items():
            if key in header:
                getattr(self.app, attr).set(header[key])

    def configure_standard(self, config: dict[str, Any]):
        a = self.app
        a.keep_muc_var.set(bool(config.get("keep_group_titles", False)))
        groups_cfg = config.get("groups", {})
        for p_idx, part in enumerate(a.parsed_data.get("parts", [])):
            multi = len(part.get("mucs", [])) > 1
            for m_idx, muc in enumerate(part.get("mucs", [])):
                key = f"{p_idx}:{m_idx}"
                c = groups_cfg.get(key, {})
                total = len(muc.get("questions", []))
                pick = max(0, min(int(c.get("pick", total)), total))
                a.muc_entries[(p_idx, m_idx)] = ValueBox(str(pick))
                a.muc_shuffle_vars[(p_idx, m_idx)] = ValueBox(bool(c.get("shuffle_questions", part.get("type") != 4)))
                fixed_default = ", ".join(str(i+1) for i,q in enumerate(muc.get("questions", [])) if q.get("is_fixed"))
                a.muc_fix_entries[(p_idx, m_idx)] = ValueBox(str(c.get("fixed_questions", fixed_default)))
                if multi:
                    a.muc_order_shuffle_vars[(p_idx, m_idx)] = ValueBox(bool(c.get("shuffle_group_order", not muc.get("is_fixed", False) and part.get("type") != 4)))

    def configure_youngmix(self, config: dict[str, Any]):
        a = self.app
        continuous = bool(config.get("continuous_numbering", False))
        master_fix = bool(config.get("master_fix", False))
        a.continuous_numbering_var.set(continuous)
        a.master_fix_var.set(master_fix)
        groups_cfg = config.get("groups", [])
        all_mucs = [m for p in a.parsed_data.get("parts", []) for m in p.get("mucs", [])]
        controls = []
        type_map = {
            "Trắc nghiệm": 1, "TN 2025 - Phần 1": 1,
            "TN 2025 - Phần 2": 2, "TN 2025 - Phần 3": 3, "Tự luận": 4,
        }
        for i, group in enumerate(a.youngmix_data):
            c = groups_cfg[i] if i < len(groups_cfg) else {}
            qtype = c.get("q_type", group.get("q_type", "Trắc nghiệm"))
            mix_type = c.get("mix_type", group.get("mix_type", "Câu hỏi & đáp án"))
            fix = master_fix or bool(c.get("fix", group.get("is_fixed", False)))
            start_q1 = bool(c.get("start_q1", (i == 0 if continuous else True)))
            total = int(group.get("q_count", 0))
            pick = max(0, min(int(c.get("pick", total)), total))
            controls.append({
                "type": ValueBox(qtype), "mix": ValueBox(mix_type),
                "fix": ValueBox(fix), "start_q1": ValueBox(start_q1), "pick": ValueBox(str(pick)),
            })
            group["q_type"] = qtype
            # Make Streamlit's type selector effective in the output, not just in the review panel.
            if i < len(all_mucs):
                forced_type = type_map.get(qtype)
                if forced_type:
                    for q in all_mucs[i].get("questions", []):
                        if not q.get("is_virtual"):
                            q["ym_type"] = forced_type
        a.ym_ui_controls = controls

    @staticmethod
    def normalize_codes(codes: list[str] | str) -> list[str]:
        if isinstance(codes, str):
            items = re.split(r"[,;\n]+", codes)
        else:
            items = list(codes)
        cleaned = []
        for x in items:
            s = str(x).strip()
            if s and s not in cleaned:
                cleaned.append(s)
        if not cleaned:
            raise ValueError("Vui lòng nhập ít nhất một mã đề.")
        if len(cleaned) > 24:
            raise ValueError("DTMIX hỗ trợ tối đa 24 mã đề trong một lần trộn.")
        return cleaned

    def mix(
        self,
        codes: list[str] | str,
        header: dict[str, str] | None = None,
        standard_config: dict[str, Any] | None = None,
        youngmix_config: dict[str, Any] | None = None,
    ) -> MixResult:
        codes = self.normalize_codes(codes)
        if header:
            self.set_header(header)
        if self.youngmix:
            self.configure_youngmix(youngmix_config or {})
            self.app.entry_made_ym.set(", ".join(codes))
            self.app.combo_somade_ym.set(str(len(codes)))
            self.app.auto_code_var_ym.set(True)
            self.app.thu_cong_var_ym.set(True)
        else:
            self.configure_standard(standard_config or {})
            self.app.entry_made.set(", ".join(codes))
            self.app.combo_somade.set(str(len(codes)))
            self.app.auto_code_var.set(True)
            self.app.thu_cong_var.set(True)

        # Use a fresh output base for each run.
        run_root = self.work_dir / ("run_" + next(tempfile._get_candidate_names()))
        run_root.mkdir(parents=True, exist_ok=True)
        self.app.output_dir = str(run_root)
        self.last_output_dir = None
        headless_compat.reset_messages()

        self.app.tron_de()
        if headless_compat.last_error:
            raise RuntimeError(headless_compat.last_error)

        out_dir = Path(self.last_output_dir) if self.last_output_dir else None
        if not out_dir or not out_dir.exists():
            dirs = [p for p in run_root.iterdir() if p.is_dir()]
            if dirs:
                out_dir = max(dirs, key=lambda p: p.stat().st_mtime)
        if not out_dir or not out_dir.exists():
            raise RuntimeError("DTMIX không tạo được thư mục kết quả.")

        files = sorted([p for p in out_dir.rglob("*") if p.is_file()])
        if not files:
            raise RuntimeError("DTMIX không tạo được file kết quả.")

        bio = BytesIO()
        with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
            for p in files:
                z.write(p, p.relative_to(out_dir))
        bio.seek(0)
        zip_name = f"DTMIX_{Path(self.filename).stem}_KET_QUA.zip"
        return MixResult(
            zip_bytes=bio.getvalue(),
            zip_name=zip_name,
            files=[p.name for p in files],
            output_dir=str(out_dir),
        )
