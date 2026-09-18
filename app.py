from __future__ import annotations

import hashlib
import html
import re
from pathlib import Path

import streamlit as st
from docx.oxml.ns import qn

from engine import DTMIXWebEngine

# ============================================================
# DTMIX ONLINE 1.3 — ONE PAGE WORKSPACE
# ============================================================
st.set_page_config(
    page_title="DTMIX Online 1.3",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# -------------------- THEME --------------------
st.markdown(
    r"""
<style>
:root{
  --navy:#0B2447;
  --navy2:#123B72;
  --blue:#176BCE;
  --blue2:#2D8CFF;
  --sky:#EAF4FF;
  --green:#14845C;
  --green-bg:#EAF8F2;
  --amber:#B86B00;
  --amber-bg:#FFF7E8;
  --red:#C8394A;
  --red-bg:#FFF0F2;
  --ink:#172033;
  --muted:#66758A;
  --line:#DCE5F0;
  --soft:#F5F8FC;
  --white:#FFFFFF;
}
html,body,[class*="css"]{font-family:Inter,"Segoe UI",Arial,sans-serif}
.stApp{
  background:
    radial-gradient(circle at 0 0,rgba(45,140,255,.10),transparent 28rem),
    radial-gradient(circle at 100% 5%,rgba(20,132,92,.06),transparent 25rem),
    #F5F8FC;
}
.block-container{max-width:1480px;padding-top:.8rem;padding-bottom:4rem}
#MainMenu,footer,header{visibility:hidden}

/* HERO */
.hero{
  background:linear-gradient(120deg,#08244B 0%,#0C3972 48%,#176BCE 100%);
  border-radius:23px;padding:25px 29px;color:#fff;
  box-shadow:0 18px 48px rgba(11,36,71,.18);position:relative;overflow:hidden;
  margin-bottom:14px;
}
.hero:after{
  content:"";position:absolute;width:310px;height:310px;border-radius:50%;
  right:-100px;top:-135px;background:rgba(255,255,255,.08);
  box-shadow:-160px 200px 0 rgba(255,255,255,.035)
}
.hero-row{display:flex;align-items:center;gap:17px;position:relative;z-index:2}
.hero-icon{
  width:58px;height:58px;border-radius:17px;display:flex;align-items:center;justify-content:center;
  background:rgba(255,255,255,.13);border:1px solid rgba(255,255,255,.2);font-size:29px
}
.hero-title{font-size:31px;font-weight:850;letter-spacing:-.4px;line-height:1}
.hero-sub{font-size:14px;color:#D6E8FF;margin-top:8px}
.hero-tags{display:flex;gap:7px;flex-wrap:wrap;margin-top:12px}
.hero-tag{font-size:11.5px;font-weight:750;padding:5px 9px;border-radius:999px;background:rgba(255,255,255,.11);border:1px solid rgba(255,255,255,.14)}
.hero-side{margin-left:auto;text-align:right}
.hero-side b{font-size:14px}.hero-side span{display:block;color:#D6E8FF;font-size:12px;margin-top:5px}

/* FLOW */
.flow{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin:0 0 18px}
.flow-item{background:#fff;border:1px solid var(--line);border-radius:13px;padding:10px 12px;display:flex;gap:9px;align-items:center;box-shadow:0 4px 14px rgba(11,36,71,.035)}
.flow-no{min-width:27px;height:27px;border-radius:9px;display:flex;align-items:center;justify-content:center;background:#EAF3FF;color:#176BCE;font-size:12px;font-weight:850}
.flow-item b{font-size:12.5px;color:#24364D}.flow-item span{display:block;font-size:10.5px;color:#7A899D;margin-top:1px}

/* SECTION */
.sec{display:flex;gap:11px;align-items:center;margin:21px 0 10px}
.sec-no{width:34px;height:34px;border-radius:11px;background:linear-gradient(135deg,#176BCE,#2D8CFF);color:white;display:flex;align-items:center;justify-content:center;font-weight:850;box-shadow:0 6px 14px rgba(23,107,206,.2)}
.sec-title{font-size:20px;font-weight:820;color:#172033}.sec-sub{font-size:12.5px;color:#748297;margin-top:1px}

/* STREAMLIT CARDS */
[data-testid="stVerticalBlockBorderWrapper"]{
  border:1px solid var(--line)!important;border-radius:16px!important;background:rgba(255,255,255,.97);
  box-shadow:0 7px 22px rgba(11,36,71,.045);padding:3px
}
[data-testid="stMetric"]{background:linear-gradient(180deg,#fff,#FBFDFF);border:1px solid var(--line);border-radius:15px;padding:13px 15px;box-shadow:0 4px 14px rgba(11,36,71,.035)}
[data-testid="stMetricValue"]{color:#0B376B}
[data-testid="stMetricLabel"]{color:#66758A}
[data-testid="stFileUploaderDropzone"]{background:#F7FBFF;border:1.5px dashed #8DB9E9;border-radius:13px}
.stButton>button{border-radius:10px;font-weight:760;min-height:42px}
.stButton>button[kind="primary"]{background:linear-gradient(90deg,#176BCE,#2D8CFF);border:0;box-shadow:0 7px 17px rgba(23,107,206,.2)}
.stDownloadButton>button{border-radius:10px;font-weight:780}
.stTextInput input,.stNumberInput input{border-radius:9px}
.stSelectbox div[data-baseweb="select"]>div{border-radius:9px}
.stExpander{background:#fff;border:1px solid #E3EAF3!important;border-radius:12px!important}
hr{border-color:#E8EEF5}

/* STATUS */
.status-good{background:var(--green-bg);border:1px solid #BFE9D7;color:#0B6D4B;border-radius:12px;padding:10px 13px;font-weight:760}
.status-warn{background:var(--amber-bg);border:1px solid #F1D8A7;color:#955600;border-radius:12px;padding:10px 13px;font-weight:760}
.status-bad{background:var(--red-bg);border:1px solid #F0C3C9;color:#AA2E3D;border-radius:12px;padding:10px 13px;font-weight:760}
.file-pill{display:inline-block;background:#EEF5FF;color:#145FAA;border:1px solid #D5E6FA;border-radius:999px;padding:5px 10px;font-size:12px;font-weight:700}
.mini-help{font-size:12px;color:#68788E;background:#F7FAFE;border:1px solid #E1E9F3;padding:9px 11px;border-radius:10px;line-height:1.55}
.audit-row{display:flex;align-items:center;justify-content:space-between;padding:8px 0;border-bottom:1px solid #EEF2F7;font-size:13px}
.audit-row:last-child{border-bottom:0}
.audit-ok{color:#14845C;font-weight:800}.audit-warn{color:#B86B00;font-weight:800}.audit-bad{color:#C8394A;font-weight:800}

/* PREVIEW */
.preview-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:9px}
.preview-title{font-weight:820;color:#17365D;font-size:15px}
.qcard{background:#fff;border:1px solid #E1E8F1;border-radius:12px;padding:13px 14px;margin:0 0 10px;box-shadow:0 3px 10px rgba(11,36,71,.025)}
.qcard.issue{border-color:#F0C8A5;background:#FFFCF7}
.qnum{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:850;color:#176BCE;background:#EDF5FF;border-radius:999px;padding:4px 8px;margin-bottom:7px}
.qtext{font-family:"Times New Roman",serif;font-size:17px;line-height:1.43;color:#161B22}
.opt{font-family:"Times New Roman",serif;font-size:16.2px;line-height:1.38;padding:4px 8px;margin-top:3px;border-radius:7px}
.opt.correct{background:#EAF8F2;color:#0B6D4B;font-weight:700;border-left:3px solid #28A979}
.opt.normal{background:#FAFCFF;color:#283342}
.short-answer{margin-top:7px;background:#EAF8F2;border-left:3px solid #28A979;border-radius:7px;padding:6px 9px;color:#0B6D4B;font-weight:700;font-family:"Times New Roman",serif}
.media-chip{display:inline-block;margin:6px 5px 0 0;background:#F3F0FF;color:#5B42A8;border:1px solid #E2DAFF;border-radius:999px;padding:3px 7px;font-size:11.5px;font-family:Inter,"Segoe UI",sans-serif}
.issue-chip{display:inline-block;margin:6px 5px 0 0;background:#FFF2E8;color:#A95B00;border:1px solid #F1D3B1;border-radius:999px;padding:3px 7px;font-size:11.5px;font-family:Inter,"Segoe UI",sans-serif}

/* YOUNGMIX */
.g-badge{display:inline-block;border-radius:7px;padding:3px 7px;font-size:11px;font-weight:850;background:#EAF3FF;color:#176BCE;border:1px solid #D1E3FA}
.g1{background:#EDF8FF;color:#126A9A}.g2{background:#F2EEFF;color:#6743B1}.g3{background:#EAF8F2;color:#0B7A53}.g0{background:#F2F4F7;color:#536174}.g4{background:#FFF3E7;color:#A45B00}

/* OUTPUT */
.action-card{background:linear-gradient(115deg,#0C315F,#176BCE);border-radius:17px;color:#fff;padding:19px 20px;margin-top:8px}
.action-card b{font-size:16px}.action-card span{display:block;color:#D8E9FF;font-size:12px;margin-top:4px}
.result-ok{background:#EAF8F2;border:1px solid #BFE9D7;border-radius:13px;color:#0B6D4B;padding:12px 14px;font-weight:760}
.footer{text-align:center;color:#97A4B4;font-size:11.5px;margin-top:26px}

@media(max-width:1000px){.flow{grid-template-columns:1fr 1fr}.hero-side{display:none}}
@media(max-width:620px){.flow{grid-template-columns:1fr}.block-container{padding-left:.75rem;padding-right:.75rem}.hero{padding:20px}.hero-title{font-size:25px}}
</style>

<div class="hero">
 <div class="hero-row">
   <div class="hero-icon">🧪</div>
   <div>
     <div class="hero-title">DTMIX Online <span style="font-size:17px;color:#BFD9FF;font-weight:650">V1.3</span></div>
     <div class="hero-sub">Trộn đề Word trực tuyến • rà soát cấu trúc • kiểm tra đáp án • YoungMix g1/g2/g3</div>
     <div class="hero-tags">
       <span class="hero-tag">DOCX</span><span class="hero-tag">PHẦN I–IV</span>
       <span class="hero-tag">g1 • g2 • g3</span><span class="hero-tag">Xem trước online</span>
       <span class="hero-tag">ZIP • Word • SmartTest</span>
     </div>
   </div>
   <div class="hero-side"><b>Một trang làm việc duy nhất</b><span>Tải đề → Rà soát → Cấu hình → Trộn → Tải về</span></div>
 </div>
</div>
<div class="flow">
 <div class="flow-item"><div class="flow-no">1</div><div><b>Tải đề</b><span>DOCX + chế độ</span></div></div>
 <div class="flow-item"><div class="flow-no">2</div><div><b>Phân tích</b><span>Câu, phần, đáp án</span></div></div>
 <div class="flow-item"><div class="flow-no">3</div><div><b>Xem & rà soát</b><span>Preview online</span></div></div>
 <div class="flow-item"><div class="flow-no">4</div><div><b>Cấu hình trộn</b><span>Tự động / g1 g2 g3</span></div></div>
 <div class="flow-item"><div class="flow-no">5</div><div><b>Xuất đề</b><span>Mã đề + tải kết quả</span></div></div>
</div>
""",
    unsafe_allow_html=True,
)

# ============================================================
# STATE / HELPERS
# ============================================================
DEFAULT_HEADER = {
    "so": "SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH",
    "truong": "TRƯỜNG THPT DƯƠNG BẠCH MAI",
    "kythi": "KIỂM TRA HỌC KỲ II",
    "namhoc": "NĂM HỌC 2025 - 2026",
    "monthi": "Môn: HÓA HỌC",
    "thoigian": "Thời gian làm bài: 45 phút",
}
for _k, _v in DEFAULT_HEADER.items():
    st.session_state.setdefault(f"hdr_{_k}", _v)


def sec(n: int, title: str, subtitle: str) -> None:
    st.markdown(
        f'<div class="sec"><div class="sec-no">{n}</div><div>'
        f'<div class="sec-title">{title}</div><div class="sec-sub">{subtitle}</div></div></div>',
        unsafe_allow_html=True,
    )


def header_values() -> dict[str, str]:
    return {k: st.session_state.get(f"hdr_{k}", v) for k, v in DEFAULT_HEADER.items()}


def clear_engine() -> None:
    old = st.session_state.pop("dtmix_engine", None)
    if old:
        try:
            old.close()
        except Exception:
            pass
    st.session_state.pop("dtmix_signature", None)
    st.session_state.pop("mix_result", None)


def esc(s: object) -> str:
    return html.escape(str(s or ""))


def element_text(engine: DTMIXWebEngine, elm) -> str:
    try:
        if elm.tag == qn("w:p"):
            return engine.app._build_mapping(elm)[1]
    except Exception:
        pass
    texts = []
    try:
        for t in elm.iter(qn("w:t")):
            if t.text:
                texts.append(t.text)
    except Exception:
        pass
    return "".join(texts)


def media_counts(elements) -> dict[str, int]:
    counts = {"image": 0, "math": 0, "table": 0}
    for elm in elements or []:
        tag0 = str(getattr(elm, "tag", ""))
        if tag0.endswith("}tbl"):
            counts["table"] += 1
        try:
            for node in elm.iter():
                tag = str(node.tag)
                if tag.endswith("}drawing") or tag.endswith("}pict") or tag.endswith("}imagedata"):
                    counts["image"] += 1
                elif tag.endswith("}oMath") or tag.endswith("}oMathPara") or tag.endswith("}object") or tag.endswith("}OLEObject"):
                    counts["math"] += 1
        except Exception:
            pass
    return counts


def extract_short_answer(engine: DTMIXWebEngine, q: dict) -> str:
    for elm in q.get("q_elements", []):
        txt = element_text(engine, elm)
        m = re.search(r"(?i)(?:đáp\s*án\s*[:\.]|(?:^|\n|\uFFFC)\s*A\.)\s*(.+)", txt)
        if m and m.group(1).strip():
            return m.group(1).strip()
    return ""


def answer_texts(engine: DTMIXWebEngine, q: dict) -> list[dict]:
    out = []
    for i, a in enumerate(q.get("answers", [])):
        txts = []
        media = {"image": 0, "math": 0, "table": 0}
        for elm in a.get("elements", []):
            txts.append(element_text(engine, elm))
            mc = media_counts([elm])
            for k in media:
                media[k] += mc[k]
        text = "\n".join(x for x in txts if x).strip()
        out.append(
            {
                "index": i,
                "text": text,
                "correct": bool(a.get("is_true")),
                "fixed": bool(a.get("is_fixed")),
                "media": media,
            }
        )
    return out


def duplicate_answer_pairs(engine: DTMIXWebEngine, q: dict) -> list[str]:
    answers = answer_texts(engine, q)
    normalized = []
    for a in answers:
        t = re.sub(r"^\s*#?[A-Ga-g][\.\)]\s*", "", a["text"]).strip()
        t = re.sub(r"\s+", "", t).rstrip(".").lower()
        if any(a["media"].values()):
            t += f"|media:{a['index']}"
        normalized.append(t)
    labels = "ABCDEFG"
    dup = []
    for i in range(len(normalized)):
        for j in range(i + 1, len(normalized)):
            if normalized[i] and normalized[i] == normalized[j]:
                li = labels[i] if i < len(labels) else str(i + 1)
                lj = labels[j] if j < len(labels) else str(j + 1)
                dup.append(f"{li}–{lj}")
    return dup


def standard_preview_data(engine: DTMIXWebEngine) -> list[dict]:
    rows = []
    for p_idx, part in enumerate(engine.app.parsed_data.get("parts", [])):
        title = engine._part_label(part)
        p_type = part.get("type", 1)
        q_display = 0
        for m_idx, muc in enumerate(part.get("mucs", [])):
            gtitle = muc.get("title") or (f"Nhóm {m_idx+1}" if len(part.get("mucs", [])) > 1 else "")
            for q in muc.get("questions", []):
                if q.get("is_virtual"):
                    continue
                q_display += 1
                answers = answer_texts(engine, q)
                valid = engine._question_has_valid_answer(q, p_type)
                media = media_counts(q.get("q_elements", []))
                for a in answers:
                    for k in media:
                        media[k] += a["media"][k]
                rows.append(
                    {
                        "p_idx": p_idx,
                        "part": title,
                        "type": p_type,
                        "group": gtitle,
                        "number": q_display,
                        "text": q.get("raw_text", "").strip(),
                        "answers": answers,
                        "valid": valid,
                        "short_answer": extract_short_answer(engine, q) if p_type == 3 else "",
                        "duplicates": duplicate_answer_pairs(engine, q) if p_type in (1, 2) else [],
                        "fixed": bool(q.get("is_fixed")),
                        "media": media,
                    }
                )
    return rows


def render_media(media: dict) -> str:
    chips = []
    if media.get("image"):
        chips.append(f'<span class="media-chip">🖼️ Hình ảnh ×{media["image"]}</span>')
    if media.get("math"):
        chips.append(f'<span class="media-chip">∑ Công thức ×{media["math"]}</span>')
    if media.get("table"):
        chips.append(f'<span class="media-chip">▦ Bảng ×{media["table"]}</span>')
    return "".join(chips)


def render_standard_question(q: dict, show_key: bool) -> None:
    issue = (not q["valid"]) or bool(q["duplicates"])
    cls = "qcard issue" if issue else "qcard"
    group = f" · {esc(q['group'])}" if q["group"] else ""
    fixed = " · 🔒 cố định" if q["fixed"] else ""
    body = [
        f'<div class="{cls}">',
        f'<div class="qnum">Câu {q["number"]}{fixed}</div>',
        f'<div class="qtext">{esc(q["text"])}</div>',
        render_media(q["media"]),
    ]
    for a in q["answers"]:
        optcls = "opt correct" if (show_key and a["correct"]) else "opt normal"
        tick = " ✓" if (show_key and a["correct"]) else ""
        lock = " 🔒" if a["fixed"] else ""
        body.append(f'<div class="{optcls}">{esc(a["text"])}{tick}{lock}</div>')
    if q["short_answer"] and show_key:
        body.append(f'<div class="short-answer">Đáp án: {esc(q["short_answer"])}</div>')
    if not q["valid"]:
        body.append('<span class="issue-chip">⚠️ Cần kiểm tra đáp án</span>')
    if q["duplicates"]:
        body.append(f'<span class="issue-chip">⚠️ Phương án trùng: {esc(", ".join(q["duplicates"]))}</span>')
    body.append("</div>")
    st.markdown("".join(body), unsafe_allow_html=True)


def audit_standard(engine: DTMIXWebEngine, preview_rows: list[dict]) -> list[dict]:
    issues = []
    for q in preview_rows:
        if not q["valid"]:
            issues.append(
                {
                    "level": "error",
                    "part": q["part"],
                    "question": q["number"],
                    "text": "Chưa nhận diện đáp án hợp lệ",
                }
            )
        if q["duplicates"]:
            issues.append(
                {
                    "level": "warning",
                    "part": q["part"],
                    "question": q["number"],
                    "text": "Phương án trùng: " + ", ".join(q["duplicates"]),
                }
            )
        if q["type"] == 1 and len(q["answers"]) not in (4,):
            issues.append(
                {
                    "level": "warning",
                    "part": q["part"],
                    "question": q["number"],
                    "text": f"Nhận diện {len(q['answers'])} phương án (thường là 4)",
                }
            )
    return issues


def youngmix_mode_label(mix_type: str) -> str:
    return {
        "Không hoán vị": "g0 · Không hoán vị",
        "Chỉ câu hỏi": "g1 · Chỉ trộn câu hỏi",
        "Chỉ đáp án": "g2 · Chỉ trộn đáp án",
        "Câu hỏi & đáp án": "g3 · Trộn câu hỏi + đáp án",
    }.get(mix_type, "g3 · Trộn câu hỏi + đáp án")


def mix_type_from_label(label: str) -> str:
    if label.startswith("g0"): return "Không hoán vị"
    if label.startswith("g1"): return "Chỉ câu hỏi"
    if label.startswith("g2"): return "Chỉ đáp án"
    return "Câu hỏi & đáp án"


def codes_ui(prefix: str) -> list[str]:
    c1, c2, c3 = st.columns([.75, 1.25, 2.2])
    n = int(c1.number_input("Số đề", 1, 24, 4, step=1, key=f"{prefix}_n"))
    mode = c2.selectbox(
        "Kiểu mã đề",
        ["Chuẩn 111, 222...", "Liên tiếp", "Thủ công"],
        key=f"{prefix}_code_mode",
    )
    if mode == "Chuẩn 111, 222...":
        codes = [str(111 * i) for i in range(1, n + 1)]
        c3.text_input("Mã đề sẽ tạo", ", ".join(codes), disabled=True, key=f"{prefix}_codes_std")
    elif mode == "Liên tiếp":
        start = c3.text_input("Mã bắt đầu", "101", key=f"{prefix}_code_start")
        if start.strip().isdigit():
            width = len(start.strip())
            first = int(start.strip())
            codes = [str(first + i).zfill(width) for i in range(n)]
        else:
            codes = []
            c3.error("Mã bắt đầu phải là số.")
    else:
        raw = c3.text_input(
            "Nhập mã đề (cách nhau bằng dấu phẩy)",
            "111, 222, 333, 444",
            key=f"{prefix}_codes_manual",
        )
        codes = [x.strip() for x in re.split(r"[,;\n]+", raw) if x.strip()]
        if len(codes) != n:
            c3.caption(f"Đã nhập {len(codes)} mã / số đề đang chọn {n}. DTMIX sẽ dùng danh sách mã đã nhập.")
    return codes


def run_mix(engine, codes, std_cfg=None, ym_cfg=None) -> None:
    if not codes:
        st.error("Chưa có mã đề hợp lệ.")
        return
    progress = st.progress(0, text="Đang chuẩn bị trộn đề...")

    def cb(cur, total, text):
        pct = int(100 * cur / total) if total else 0
        progress.progress(min(max(pct, 0), 100), text=text)

    engine.progress_callback = cb
    try:
        result = engine.mix(
            codes,
            header=header_values(),
            standard_config=std_cfg,
            youngmix_config=ym_cfg,
        )
        progress.progress(100, text="Hoàn tất")
        st.session_state.mix_result = result
        st.rerun()
    except Exception as exc:
        st.error(f"Không thể trộn đề: {exc}")
        with st.expander("Chi tiết lỗi kỹ thuật"):
            st.exception(exc)


def download_results() -> None:
    result = st.session_state.get("mix_result")
    if not result:
        return
    st.markdown(
        f'<div class="result-ok">✅ Đã tạo {len(result.files)} file kết quả. Chọn tải toàn bộ hoặc tải từng file.</div>',
        unsafe_allow_html=True,
    )
    st.download_button(
        "⬇️ TẢI TOÀN BỘ KẾT QUẢ (.ZIP)",
        data=result.zip_bytes,
        file_name=result.zip_name,
        mime="application/zip",
        type="primary",
        use_container_width=True,
        key="download_all_zip",
    )
    outdir = Path(result.output_dir)
    with st.expander("Tải từng file", expanded=False):
        files = [p for p in outdir.rglob("*") if p.is_file()] if outdir.exists() else []
        for i in range(0, len(files), 3):
            cols = st.columns(3)
            for j, p in enumerate(files[i:i+3]):
                cols[j].download_button(
                    f"⬇ {p.name}",
                    data=p.read_bytes(),
                    file_name=p.name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        if p.suffix.lower() == ".docx"
                        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        if p.suffix.lower() == ".xlsx"
                        else "application/octet-stream"
                    ),
                    use_container_width=True,
                    key=f"file_dl_{i}_{j}_{p.name}",
                )


# ============================================================
# 1 — UPLOAD & SETTINGS
# ============================================================
sec(1, "Tải đề & chọn chế độ trộn", "Tải file Word, chọn nhận diện tự động hoặc nhóm g1/g2/g3. Thông tin đầu đề được đặt gọn trong phần mở rộng.")

with st.container(border=True):
    top_l, top_r = st.columns([1.55, 1], gap="large")
    with top_l:
        mode = st.radio(
            "Chế độ xử lý",
            [
                "Tự động theo PHẦN I / II / III / IV",
                "Theo nhóm g1 / g2 / g3 / g4 (YoungMix)",
            ],
            horizontal=True,
            key="dtmix_mode",
        )
        is_youngmix = mode.startswith("Theo nhóm")
        uploaded = st.file_uploader(
            "Kéo thả hoặc chọn file đề gốc (.docx)",
            type=["docx"],
            accept_multiple_files=False,
            key="source_docx",
        )
    with top_r:
        st.markdown("**DTMIX sẽ kiểm tra**")
        st.markdown(
            """
<div class="mini-help">
<b>• Cấu trúc:</b> PHẦN I–IV hoặc các nhóm g1/g2/g3<br>
<b>• Số câu:</b> theo từng phần/nhóm và toàn đề<br>
<b>• Đáp án:</b> đã nhận diện hay còn thiếu<br>
<b>• Lỗi:</b> phương án trùng, số phương án bất thường<br>
<b>• Xem trước:</b> nội dung câu hỏi và đáp án ngay trên web
</div>
""",
            unsafe_allow_html=True,
        )
        with st.expander("📝 Thông tin đầu trang đề", expanded=False):
            st.text_input("Sở GD&ĐT / Phòng", key="hdr_so")
            st.text_input("Tên trường", key="hdr_truong")
            c1, c2 = st.columns(2)
            c1.text_input("Tên kỳ thi", key="hdr_kythi")
            c2.text_input("Năm học", key="hdr_namhoc")
            c3, c4 = st.columns(2)
            c3.text_input("Môn thi", key="hdr_monthi")
            c4.text_input("Thời gian làm bài", key="hdr_thoigian")

    current_sig = None
    if uploaded:
        raw = uploaded.getvalue()
        current_sig = hashlib.sha256(raw + str(is_youngmix).encode()).hexdigest()
        info1, info2, info3 = st.columns([2, 1, 1])
        info1.markdown(f'<span class="file-pill">📄 {esc(uploaded.name)}</span>', unsafe_allow_html=True)
        info2.caption(f"{len(raw)/1024:.1f} KB")
        info3.caption("YoungMix" if is_youngmix else "Tự động")

        same = st.session_state.get("dtmix_signature") == current_sig
        b1, b2 = st.columns([3, 1])
        if b1.button(
            "✅ ĐÃ PHÂN TÍCH — DÙNG KẾT QUẢ BÊN DƯỚI" if same else "🔎 PHÂN TÍCH & RÀ SOÁT ĐỀ",
            type="primary",
            use_container_width=True,
            disabled=same,
        ):
            clear_engine()
            with st.spinner("DTMIX đang đọc cấu trúc Word, câu hỏi, đáp án, bảng, hình ảnh và công thức..."):
                try:
                    eng = DTMIXWebEngine(
                        raw,
                        uploaded.name,
                        youngmix=is_youngmix,
                        header=header_values(),
                    )
                    st.session_state.dtmix_engine = eng
                    st.session_state.dtmix_signature = current_sig
                    st.session_state.mix_result = None
                    st.rerun()
                except Exception as exc:
                    st.error(f"Không phân tích được đề: {exc}")
                    with st.expander("Chi tiết lỗi"):
                        st.exception(exc)
        if b2.button("↻ Đề khác", use_container_width=True, disabled=not same):
            clear_engine()
            st.rerun()
    else:
        st.info("Chọn một file .docx để bắt đầu.")

engine = st.session_state.get("dtmix_engine")
if current_sig is not None and st.session_state.get("dtmix_signature") != current_sig:
    engine = None

# ============================================================
# 2 — OVERVIEW
# ============================================================
sec(2, "Tổng quan & tình trạng đề", "Ngay sau khi phân tích, DTMIX cho biết số phần, số câu và mức độ hoàn chỉnh của đáp án.")

if not engine:
    with st.container(border=True):
        st.caption("Chưa có dữ liệu. Sau khi phân tích, thống kê đề sẽ xuất hiện tại đây.")
else:
    summary = engine.summary()
    parts = summary.get("parts", [])
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Phần / nhóm", len(summary.get("youngmix_groups", [])) if engine.youngmix else len(parts))
    k2.metric("Tổng số câu", summary["total_questions"])
    k3.metric("Có đáp án / hợp lệ", summary["valid_answers"])
    k4.metric("Cần kiểm tra", summary["missing_answers"])
    readiness = "SẴN SÀNG" if summary["missing_answers"] == 0 else "CẦN RÀ SOÁT"
    k5.metric("Trạng thái", readiness)

    if summary["missing_answers"] == 0:
        st.markdown('<div class="status-good">✅ Đề đã nhận diện đầy đủ các đáp án bắt buộc. Có thể tiếp tục cấu hình và trộn đề.</div>', unsafe_allow_html=True)
    else:
        st.markdown(
            f'<div class="status-warn">⚠️ Có {summary["missing_answers"]} câu chưa nhận diện đáp án hợp lệ hoặc thuộc dạng cần kiểm tra. Xem danh sách ở khu vực Rà soát.</div>',
            unsafe_allow_html=True,
        )

    if not engine.youngmix:
        cols = st.columns(max(1, min(4, len(parts))))
        for i, part in enumerate(parts):
            with cols[i % len(cols)]:
                ok = part["missing_count"] == 0
                st.markdown(
                    f'<div class="audit-row"><span><b>{esc(part["title"])}</b><br>'
                    f'<span style="color:#7A899D">{part["question_count"]} câu</span></span>'
                    f'<span class="{"audit-ok" if ok else "audit-warn"}">{part["valid_count"]}/{part["question_count"]} ĐA</span></div>',
                    unsafe_allow_html=True,
                )
    else:
        groups = summary.get("youngmix_groups", [])
        cols = st.columns(max(1, min(4, len(groups))))
        for i, g in enumerate(groups):
            tag = re.sub(r"[<>#]", "", g.get("tag", "g")).lower()
            st_cls = tag if tag in ("g0","g1","g2","g3","g4") else "g3"
            with cols[i % len(cols)]:
                st.markdown(
                    f'<div class="audit-row"><span><b>{esc(g["name"])}</b> '
                    f'<span class="g-badge {st_cls}">{esc(g["tag"])}</span><br>'
                    f'<span style="color:#7A899D">{g["question_count"]} câu</span></span>'
                    f'<span style="font-size:12px;color:#536174">{esc(g["mix_type"])}</span></div>',
                    unsafe_allow_html=True,
                )

# ============================================================
# 3 — WORKSPACE: PREVIEW + REVIEW/CONFIG
# ============================================================
sec(3, "Xem trước online • Rà soát • Cấu hình", "Vùng làm việc chính: xem đề bên trái, rà soát và thiết lập cách trộn bên phải.")

std_config = None
ym_config = None

if not engine:
    with st.container(border=True):
        st.caption("Phân tích đề để mở khu vực xem trước và cấu hình.")
else:
    if not engine.youngmix:
        preview_rows = standard_preview_data(engine)
        issues = audit_standard(engine, preview_rows)

        left, right = st.columns([1.85, 1], gap="large")
        with left:
            with st.container(border=True):
                st.markdown('<div class="preview-title">👁️ XEM TRƯỚC ĐỀ ONLINE</div>', unsafe_allow_html=True)
                part_names = list(dict.fromkeys(q["part"] for q in preview_rows))
                f1, f2, f3 = st.columns([1.5, 1.1, .9])
                selected_part = f1.selectbox("Phần đang xem", ["Tất cả"] + part_names, key="preview_part")
                search = f2.text_input("Tìm câu / từ khóa", "", placeholder="VD: Câu 12, glucose", key="preview_search")
                show_key = f3.toggle("Hiện đáp án", value=True, key="preview_show_key")

                selected = preview_rows
                if selected_part != "Tất cả":
                    selected = [q for q in selected if q["part"] == selected_part]
                if search.strip():
                    s = search.strip().lower()
                    selected = [q for q in selected if s in q["text"].lower() or s in f"câu {q['number']}".lower()]

                st.caption(f"Đang hiển thị {len(selected)} / {len(preview_rows)} câu")
                with st.container(height=650, border=False):
                    last_part = None
                    for q in selected:
                        if q["part"] != last_part:
                            st.markdown(f"#### {q['part']}")
                            last_part = q["part"]
                        render_standard_question(q, show_key)

        with right:
            with st.container(border=True):
                st.markdown("#### 🩺 Rà soát đề")
                if not issues:
                    st.success("Không phát hiện lỗi đáp án/phương án nổi bật.")
                else:
                    err = sum(1 for x in issues if x["level"] == "error")
                    warn = len(issues) - err
                    c1, c2 = st.columns(2)
                    c1.metric("Lỗi đáp án", err)
                    c2.metric("Cảnh báo", warn)
                    with st.container(height=210, border=False):
                        for item in issues:
                            icon = "❌" if item["level"] == "error" else "⚠️"
                            st.markdown(
                                f'<div class="audit-row"><span>{icon} <b>{esc(item["part"])}</b> · Câu {item["question"]}</span>'
                                f'<span class="{"audit-bad" if item["level"]=="error" else "audit-warn"}">{esc(item["text"])}</span></div>',
                                unsafe_allow_html=True,
                            )

                st.divider()
                st.markdown("#### ⚙️ Cấu hình trộn tự động")
                keep_titles = st.checkbox("Giữ tiêu đề nhóm/mục", value=False, key="std_keep_titles_v2")
                std_groups = {}
                with st.container(height=440, border=False):
                    for part in summary["parts"]:
                        icon = "✅" if part["missing_count"] == 0 else "⚠️"
                        with st.expander(f"{icon} {part['title']} · {part['question_count']} câu", expanded=False):
                            multi = len(part["groups"]) > 1
                            for g in part["groups"]:
                                key = f"s_{g['p_idx']}_{g['m_idx']}"
                                if multi:
                                    st.markdown(f"**{g['title']}**")
                                c1, c2 = st.columns(2)
                                shuffle_q = c1.toggle("Trộn thứ tự câu", value=(part["type"] != 4), key=key+"_sq")
                                pick = c2.number_input("Số câu lấy", 0, g["question_count"], g["question_count"], key=key+"_pick")
                                c3, c4 = st.columns(2)
                                fixed = c3.text_input(
                                    "Cố định câu",
                                    value=", ".join(str(q["index"]) for q in g["questions"] if q["fixed"]),
                                    placeholder="VD: 1, 5",
                                    key=key+"_fix",
                                )
                                shuffle_group = False
                                if multi:
                                    shuffle_group = c4.toggle(
                                        "Trộn vị trí nhóm",
                                        value=(not g["is_fixed"] and part["type"] != 4),
                                        key=key+"_sg",
                                    )
                                else:
                                    c4.caption("Không có nhiều nhóm trong phần này.")
                                std_groups[f"{g['p_idx']}:{g['m_idx']}"] = {
                                    "shuffle_questions": shuffle_q,
                                    "pick": int(pick),
                                    "fixed_questions": fixed,
                                    "shuffle_group_order": shuffle_group,
                                }
                std_config = {"keep_group_titles": keep_titles, "groups": std_groups}

    else:
        groups = summary.get("youngmix_groups", [])
        left, right = st.columns([1.85, 1], gap="large")

        with left:
            with st.container(border=True):
                st.markdown('<div class="preview-title">👁️ XEM TRƯỚC NHÓM g1 / g2 / g3 ONLINE</div>', unsafe_allow_html=True)
                choices = [f"{g['name']}  {g['tag']}  · {g['question_count']} câu" for g in groups]
                idx = st.selectbox(
                    "Nhóm đang xem",
                    range(len(groups)),
                    format_func=lambda i: choices[i],
                    key="ym_preview_group",
                ) if groups else None
                show_mark = st.toggle("Hiện đánh dấu đáp án từ file gốc", value=True, key="ym_show_mark")
                with st.container(height=650, border=False):
                    if idx is not None:
                        g = groups[idx]
                        tag = re.sub(r"[<>#]", "", g.get("tag", "g3")).lower()
                        st.markdown(
                            f'### {esc(g["name"])} &nbsp; <span class="g-badge {tag if tag in ("g0","g1","g2","g3","g4") else "g3"}">{esc(g["tag"])}</span>',
                            unsafe_allow_html=True,
                        )
                        st.caption(f"{g['q_type']} · {g['mix_type']} · {g['question_count']} câu")
                        if show_mark:
                            chunks = []
                            for text, is_red, is_under in g.get("preview_chunks", []):
                                style = ""
                                if is_red:
                                    style = "background:#EAF8F2;color:#0B6D4B;font-weight:700;border-radius:4px;padding:0 2px"
                                elif is_under:
                                    style = "text-decoration:underline"
                                chunks.append(f'<span style="{style}">{esc(text).replace(chr(10),"<br>")}</span>')
                            st.markdown(
                                '<div class="qcard"><div class="qtext">' + "".join(chunks) + "</div></div>",
                                unsafe_allow_html=True,
                            )
                        else:
                            st.markdown(
                                f'<div class="qcard"><div class="qtext">{esc(g["preview"]).replace(chr(10),"<br>")}</div></div>',
                                unsafe_allow_html=True,
                            )

        with right:
            with st.container(border=True):
                st.markdown("#### 🩺 Rà soát YoungMix")
                if summary["missing_answers"] == 0:
                    st.success("Các câu cần đáp án đã được nhận diện.")
                else:
                    st.warning(f"Có {summary['missing_answers']} câu cần kiểm tra hoặc thuộc dạng tự luận.")

                c1, c2 = st.columns(2)
                continuous = c1.toggle("Đánh số liên tục", value=False, key="ym_cont_v2")
                master_fix = c2.toggle("Cố định tất cả nhóm", value=False, key="ym_master_v2")

                st.divider()
                st.markdown("#### 🧩 Cấu hình g1 / g2 / g3")
                ym_groups_cfg = []
                with st.container(height=505, border=False):
                    for i, g in enumerate(groups):
                        with st.expander(f"{g['name']} {g['tag']} · {g['question_count']} câu", expanded=(i == 0)):
                            tag_class = re.sub(r"[<>#]", "", g.get("tag", "g3")).lower()
                            st.markdown(
                                f'<span class="g-badge {tag_class if tag_class in ("g0","g1","g2","g3","g4") else "g3"}">{esc(g["tag"])}</span> '
                                f'<span style="font-size:12px;color:#66758A">{esc(g["q_type"])}</span>',
                                unsafe_allow_html=True,
                            )
                            mode_options = [
                                "g0 · Không hoán vị",
                                "g1 · Chỉ trộn câu hỏi",
                                "g2 · Chỉ trộn đáp án",
                                "g3 · Trộn câu hỏi + đáp án",
                            ]
                            default_mode = youngmix_mode_label(g["mix_type"])
                            if default_mode not in mode_options:
                                default_mode = mode_options[-1]
                            mix_label = st.selectbox(
                                "Cách trộn nhóm",
                                mode_options,
                                index=mode_options.index(default_mode),
                                key=f"ym_mode_v2_{i}",
                            )
                            type_opts = ["TN 2025 - Phần 1", "TN 2025 - Phần 2", "TN 2025 - Phần 3", "Trắc nghiệm", "Tự luận"]
                            default_type = g["q_type"] if g["q_type"] in type_opts else "Trắc nghiệm"
                            q_type = st.selectbox(
                                "Loại câu",
                                type_opts,
                                index=type_opts.index(default_type),
                                key=f"ym_type_v2_{i}",
                            )
                            c3, c4 = st.columns(2)
                            pick = c3.number_input("Số câu lấy", 0, g["question_count"], g["question_count"], key=f"ym_pick_v2_{i}")
                            fix = c4.checkbox(
                                "Cố định nhóm",
                                value=(master_fix or g["is_fixed"]),
                                disabled=master_fix,
                                key=f"ym_fix_v2_{i}",
                            )
                            start_q1 = st.checkbox(
                                "Đánh lại từ Câu 1 ở nhóm này",
                                value=((i == 0) if continuous else True),
                                disabled=continuous,
                                key=f"ym_start_v2_{i}",
                            )
                            ym_groups_cfg.append(
                                {
                                    "q_type": q_type,
                                    "mix_type": mix_type_from_label(mix_label),
                                    "pick": int(pick),
                                    "fix": fix,
                                    "start_q1": start_q1,
                                }
                            )
                ym_config = {
                    "continuous_numbering": continuous,
                    "master_fix": master_fix,
                    "groups": ym_groups_cfg,
                }

# ============================================================
# 4 — CODES + MIX
# ============================================================
sec(4, "Mã đề & trộn đề", "Chọn số lượng đề và cách sinh mã. DTMIX sử dụng thuật toán trộn gốc để tạo file Word kết quả.")

if not engine:
    with st.container(border=True):
        st.caption("Khu vực mã đề sẽ hoạt động sau khi đề được phân tích.")
else:
    with st.container(border=True):
        prefix = "ym2" if engine.youngmix else "std2"
        codes = codes_ui(prefix)
        st.caption("Tối đa 24 mã đề trong một lần trộn.")

        ready = summary["missing_answers"] == 0
        status_text = (
            f"✅ {summary['total_questions']} câu · {len(codes)} mã đề · sẵn sàng trộn"
            if ready
            else f"⚠️ {summary['total_questions']} câu · {summary['missing_answers']} câu cần kiểm tra · {len(codes)} mã đề"
        )
        st.markdown(
            f'<div class="{"status-good" if ready else "status-warn"}">{status_text}</div>',
            unsafe_allow_html=True,
        )
        st.write("")
        if st.button("🚀 TRỘN ĐỀ VÀ XUẤT FILE", type="primary", use_container_width=True, key="mix_now_v2"):
            if engine.youngmix:
                run_mix(engine, codes, ym_cfg=ym_config)
            else:
                run_mix(engine, codes, std_cfg=std_config)

# ============================================================
# 5 — RESULTS
# ============================================================
sec(5, "Kết quả", "Tải toàn bộ bằng ZIP hoặc tải riêng từng mã đề / file đáp án.")

if not st.session_state.get("mix_result"):
    with st.container(border=True):
        st.caption("Sau khi trộn thành công, các nút tải file sẽ xuất hiện tại đây.")
else:
    with st.container(border=True):
        download_results()

# ============================================================
# GUIDE
# ============================================================
st.markdown("---")
with st.expander("📖 Hướng dẫn nhanh & quy ước g1/g2/g3", expanded=False):
    st.markdown(
        """
**Chế độ tự động:** dùng đề được chia theo `PHẦN I`, `PHẦN II`, `PHẦN III`, `PHẦN IV`.

**YoungMix:**
- `g0`: không hoán vị.
- `g1`: chỉ hoán vị câu hỏi.
- `g2`: chỉ hoán vị phương án/đáp án.
- `g3`: hoán vị cả câu hỏi và phương án.
- `g4`: nhóm tự luận.
- `<#g1>`, `<#g3>`...: cố định vị trí nhóm.

**Đáp án:** Phần I/II nên gạch chân hoặc tô đỏ đáp án đúng trong Word. Phần III dùng `Đáp án:` hoặc `A. giá trị`.

**Quy trình:** tải file → phân tích/rà soát → xem trước → cấu hình → chọn mã đề → trộn → tải kết quả.
"""
    )

st.markdown(
    '<div class="footer">DTMIX Online 1.3 • Giao diện một trang • Thuật toán xử lý Word kế thừa từ DTMIX 1.3</div>',
    unsafe_allow_html=True,
)
