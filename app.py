from __future__ import annotations

import base64
import copy
import hashlib
import html
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
from docx.oxml.ns import qn
from lxml import etree

try:
    import fitz  # PyMuPDF - dùng để dựng trang PDF thành ảnh preview
except Exception:
    fitz = None

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
# Người dùng có thể đổi bảng màu; toàn bộ đều là nền sáng, không dùng nền đen.
_THEME_PRESETS = {
    "Xanh biển sáng": {
        "bg": "#F3F8FE", "card": "#FFFFFF", "primary": "#2469B5", "primary2": "#4B9AE8",
        "hero1": "#E9F3FF", "hero2": "#D8EBFF", "hero_text": "#173A63",
        "line": "#D7E4F2", "soft": "#F7FAFE", "muted": "#63758A",
    },
    "Xanh ngọc": {
        "bg": "#F2FAF7", "card": "#FFFFFF", "primary": "#247A65", "primary2": "#48A98E",
        "hero1": "#E6F6F0", "hero2": "#D6EFE7", "hero_text": "#195646",
        "line": "#D5E9E1", "soft": "#F7FCFA", "muted": "#637970",
    },
    "Tím dịu": {
        "bg": "#F7F4FC", "card": "#FFFFFF", "primary": "#7056AE", "primary2": "#9A82D3",
        "hero1": "#EFEAFA", "hero2": "#E3DCF5", "hero_text": "#4E3A7D",
        "line": "#E1DAEE", "soft": "#FBF9FE", "muted": "#726B80",
    },
    "Kem sáng": {
        "bg": "#FBF8F2", "card": "#FFFFFF", "primary": "#8D6738", "primary2": "#B98A50",
        "hero1": "#F7EEDD", "hero2": "#F0DFC5", "hero_text": "#62451F",
        "line": "#E9DDCB", "soft": "#FEFCF8", "muted": "#7D7161",
    },
}
_theme_name = st.session_state.get("ui_theme", "Xanh biển sáng")
_theme = _THEME_PRESETS.get(_theme_name, _THEME_PRESETS["Xanh biển sáng"])

st.markdown(
    f"""
<style>
:root{{
  --bg:{_theme["bg"]};
  --card:{_theme["card"]};
  --primary:{_theme["primary"]};
  --primary2:{_theme["primary2"]};
  --hero1:{_theme["hero1"]};
  --hero2:{_theme["hero2"]};
  --hero-text:{_theme["hero_text"]};
  --line:{_theme["line"]};
  --soft:{_theme["soft"]};
  --muted:{_theme["muted"]};
  --ink:#203044;
  --green:#17785B;
  --green-bg:#EAF7F1;
  --amber:#A96A0B;
  --amber-bg:#FFF6E7;
  --red:#B93E4B;
  --red-bg:#FFF0F2;
}}
html,body,[class*="css"]{{font-family:Inter,"Segoe UI",Arial,sans-serif}}
.stApp{{
  background:
    radial-gradient(circle at 0 0, color-mix(in srgb, var(--primary2) 12%, transparent), transparent 28rem),
    radial-gradient(circle at 100% 0, color-mix(in srgb, var(--primary) 8%, transparent), transparent 25rem),
    var(--bg);
  color:var(--ink);
}}
.block-container{{max-width:none;width:100%;padding:.18rem 1.0rem 1.4rem}}
#MainMenu,footer,header{{visibility:hidden}}

/* Hero sáng, không nền đen */
.hero{{
  background:linear-gradient(120deg,var(--hero1),var(--hero2));
  border:1px solid color-mix(in srgb,var(--primary) 18%, white);
  border-radius:14px;padding:8px 15px;color:var(--hero-text);
  box-shadow:0 12px 30px rgba(58,91,124,.10);position:relative;overflow:hidden;margin-bottom:6px;
}}
.hero:after{{content:"";position:absolute;width:270px;height:270px;border-radius:50%;right:-100px;top:-120px;background:rgba(255,255,255,.45)}}
.hero-row{{display:flex;align-items:center;gap:14px;position:relative;z-index:2}}
.hero-icon{{width:46px;height:46px;border-radius:13px;display:flex;align-items:center;justify-content:center;background:rgba(255,255,255,.65);border:1px solid rgba(255,255,255,.9);font-size:26.33px}}
.hero-title{{font-size:26.33px;font-weight:850;letter-spacing:-.3px;line-height:1}}
.hero-sub{{font-size:14.83px;color:color-mix(in srgb,var(--hero-text) 76%, white);margin-top:5px}}
.hero-side{{margin-left:auto;text-align:right;max-width:330px}}
.hero-side b{{font-size:14.33px}} .hero-side span{{display:block;color:color-mix(in srgb,var(--hero-text) 72%, white);font-size:13.33px;margin-top:4px}}

.sec{{display:flex;gap:10px;align-items:center;margin:14px 0 7px}}
.sec-no{{width:31px;height:31px;border-radius:10px;background:linear-gradient(135deg,var(--primary),var(--primary2));color:white;display:flex;align-items:center;justify-content:center;font-weight:850;box-shadow:0 5px 12px color-mix(in srgb,var(--primary) 24%, transparent)}}
.sec-title{{font-size:20.33px;font-weight:820;color:#203044}} .sec-sub{{font-size:14.33px;color:var(--muted);margin-top:1px}}

[data-testid="stVerticalBlockBorderWrapper"]{{
  border:1px solid var(--line)!important;border-radius:15px!important;background:rgba(255,255,255,.98);
  box-shadow:0 6px 18px rgba(55,79,104,.045);padding:3px
}}
[data-testid="stMetric"]{{background:linear-gradient(180deg,#fff,var(--soft));border:1px solid var(--line);border-radius:13px;padding:10px 13px;box-shadow:0 3px 11px rgba(55,79,104,.035)}}
[data-testid="stMetricValue"]{{color:var(--primary)}} [data-testid="stMetricLabel"]{{color:var(--muted)}}
[data-testid="stFileUploaderDropzone"]{{background:var(--soft);border:1.5px dashed color-mix(in srgb,var(--primary2) 65%, white);border-radius:12px;min-height:78px;padding:.45rem}}
[data-testid="stFileUploaderDropzone"] section{{padding:.2rem .5rem}}
.stButton>button{{border-radius:9px;font-weight:760;min-height:40px}}
.stButton>button[kind="primary"]{{background:linear-gradient(90deg,var(--primary),var(--primary2));border:0;box-shadow:0 6px 15px color-mix(in srgb,var(--primary) 22%, transparent)}}
.stDownloadButton>button{{border-radius:9px;font-weight:760}}
.stTextInput input,.stNumberInput input{{border-radius:8px;background:#fff!important;color:#23364A!important}}
.stSelectbox div[data-baseweb="select"]>div{{border-radius:8px;background:#fff!important;color:#23364A!important}}
.stRadio label,.stCheckbox label{{color:#2B3B4E!important}}

/* Chữ rõ hơn, không cắt label */
[data-testid="stWidgetLabel"] p{{font-weight:700!important;color:#30455E!important;font-size:15.83px!important}}
[data-testid="stMarkdownContainer"] p{{line-height:1.48}}
[data-testid="stToggle"] label{{width:100%!important}}
[data-testid="stToggle"] label p{{white-space:normal!important;overflow:visible!important;text-overflow:clip!important;font-weight:650!important}}
[data-baseweb="select"] *{{font-size:15.83px!important}}
.stCaptionContainer{{color:#73859A!important}}
div[data-testid="stHorizontalBlock"]{{align-items:flex-start}}

/* Thanh công cụ trên cùng */
.tool-card-title{{font-size:14.83px;font-weight:850;color:#173B65;margin-bottom:6px}}
.top-hint{{font-size:13.33px;color:#6D8096;line-height:1.35;margin-top:5px}}
.top-status{{font-size:13.83px;font-weight:750;color:#176BCE;background:#EEF6FF;border:1px solid #D4E8FB;border-radius:8px;padding:6px 8px;text-align:center}}

/* Thông tin đề dạng ngang */
.exam-info{{display:grid;grid-template-columns:1.15fr 1.45fr 1.05fr;gap:0;border:1px solid #DCE7F3;border-radius:12px;overflow:hidden;background:#fff;margin:6px 0 11px}}
.exam-info>div{{padding:9px 12px;border-right:1px solid #E6EDF5;border-bottom:1px solid #E6EDF5;font-size:13.53px;color:#30455E}}
.exam-info>div:nth-child(3n){{border-right:0}}
.exam-info>div:nth-last-child(-n+3){{border-bottom:0}}
.exam-info b{{color:#173B65;margin-right:5px}}

/* Workspace */
.preview-shell{{background:#F5F9FE;border:1px solid #DDE8F4;border-radius:14px;padding:8px}}
.preview-heading{{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:7px}}
.preview-heading b{{font-size:17.33px;color:#143A66}}
.preview-badge{{font-size:13.33px;font-weight:800;color:#0F7357;background:#E8F8F1;border:1px solid #C5EBDD;border-radius:999px;padding:5px 9px}}

/* Right panel clarity */
.control-title{{font-size:17.33px;font-weight:850;color:#173B65;margin:2px 0 8px}}
.control-note{{font-size:13.83px;color:#657A91;background:#F7FAFE;border:1px solid #DFE8F2;border-radius:9px;padding:7px 9px;margin-bottom:8px}}
.stExpander{{background:#fff;border:1px solid var(--line)!important;border-radius:11px!important}}
hr{{border-color:#E7EDF4}}

.status-good{{background:var(--green-bg);border:1px solid #BFE5D6;color:#116B50;border-radius:11px;padding:9px 12px;font-weight:750}}
.status-warn{{background:var(--amber-bg);border:1px solid #EED9AE;color:#955B00;border-radius:11px;padding:9px 12px;font-weight:750}}
.status-bad{{background:var(--red-bg);border:1px solid #EFC5CB;color:#A53240;border-radius:11px;padding:9px 12px;font-weight:750}}
.file-pill{{display:inline-block;background:color-mix(in srgb,var(--primary2) 12%, white);color:var(--primary);border:1px solid color-mix(in srgb,var(--primary2) 28%, white);border-radius:999px;padding:4px 9px;font-size:13.83px;font-weight:720}}
.mini-help{{font-size:13.83px;color:var(--muted);background:var(--soft);border:1px solid var(--line);padding:6px 9px;border-radius:9px;line-height:1.5}}
.audit-row{{display:flex;align-items:center;justify-content:space-between;padding:7px 0;border-bottom:1px solid #EEF2F6;font-size:14.83px;gap:10px}}
.audit-row:last-child{{border-bottom:0}} .audit-ok{{color:#17785B;font-weight:800}} .audit-warn{{color:#A96A0B;font-weight:800}} .audit-bad{{color:#B93E4B;font-weight:800}}
.g-badge{{display:inline-block;border-radius:7px;padding:3px 7px;font-size:12.13px;font-weight:850;background:#EAF3FF;color:#176BCE;border:1px solid #D1E3FA}}
.g1{{background:#EDF8FF;color:#126A9A}} .g2{{background:#F2EEFF;color:#6743B1}} .g3{{background:#EAF8F2;color:#0B7A53}} .g0{{background:#F2F4F7;color:#536174}} .g4{{background:#FFF3E7;color:#A45B00}}
.word-preview-note{{font-size:13.83px;color:#536A82;background:#F7FAFE;border:1px solid #DFE8F2;border-radius:9px;padding:7px 9px;margin:5px 0 9px}}
.preview-toolbar{{font-size:13.83px;color:#536A82;background:var(--soft);border:1px solid var(--line);border-radius:9px;padding:7px 9px}}
.result-ok{{background:var(--green-bg);border:1px solid #BFE5D6;border-radius:11px;color:#116B50;padding:10px 12px;font-weight:750}}
.footer{{text-align:center;color:#93A0AF;font-size:13.33px;margin-top:22px}}

/* Tổng quan gọn - tối ưu màn hình máy tính ở 100% */
.summary-compact{{
  display:grid;grid-template-columns:repeat(4,minmax(130px,1fr));gap:7px;margin:5px 0 5px;
}}
.summary-chip{{
  background:#FFFFFF;border:1px solid var(--line);border-radius:11px;padding:6px 10px;
  box-shadow:0 2px 8px rgba(45,75,105,.035);min-height:50px;
}}
.summary-chip .v{{font-size:23.33px;font-weight:850;color:var(--primary);line-height:1.05}}
.summary-chip .l{{font-size:13.83px;font-weight:700;color:var(--muted);margin-top:4px}}
.part-status-grid{{
  display:grid;grid-template-columns:repeat(4,minmax(190px,1fr));gap:7px;margin:5px 0 7px;
}}
.part-status-card{{
  background:#fff;border:1px solid var(--line);border-radius:10px;padding:6px 9px;
  display:flex;align-items:center;justify-content:space-between;gap:8px;min-height:47px;
}}
.part-status-card .pn{{font-size:14.83px;font-weight:850;color:#203C5D}}
.part-status-card .qc{{font-size:13.83px;color:#6D7F93;margin-top:2px}}
.part-status-card .ans-ok{{font-size:13.83px;font-weight:850;color:#137A5A;text-align:right}}
.part-status-card .ans-bad{{font-size:13.83px;font-weight:850;color:#B34A39;text-align:right}}
.workspace-title{{
  font-size:20.33px;font-weight:850;color:#193B61;margin:7px 0 5px;
  display:flex;align-items:center;gap:8px;
}}
.workspace-title:before{{
  content:"";width:5px;height:22px;border-radius:5px;background:linear-gradient(var(--primary),var(--primary2));
}}
@media(max-width:1200px){{
  .summary-compact{{grid-template-columns:repeat(2,1fr)}}
  .part-status-grid{{grid-template-columns:repeat(2,1fr)}}
}}


@media(max-width:1000px){{.hero-side{{display:none}}}}

/* Dồn các khối Streamlit lên gần nhau hơn */
[data-testid="stVerticalBlock"]{{gap:.48rem!important}}
[data-testid="stHorizontalBlock"]{{gap:.65rem!important}}
[data-testid="stFileUploader"]{{margin-bottom:.1rem!important}}
[data-testid="stFileUploaderDropzone"]{{padding:.45rem .7rem!important;min-height:72px!important}}
[data-testid="stMetric"]{{padding:7px 10px!important}}
.stElementContainer{{margin-bottom:.05rem!important}}
hr{{margin:.45rem 0!important}}


/* Cấu hình trộn tự động luôn mở, hiển thị theo hàng ngang */
.auto-config-wrap{{
  margin:6px 0 10px;padding:10px 11px;border:1px solid #DCE7F3;
  border-radius:13px;background:linear-gradient(180deg,#FFFFFF,#F8FBFF);
}}
.auto-config-head{{
  display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:7px;
}}
.auto-config-title{{font-size:18.33px;font-weight:850;color:#173B65}}
.auto-config-sub{{font-size:13.33px;color:#6D8096}}
.config-part-label{{
  font-size:14.83px;font-weight:850;color:#173B65;line-height:1.35;padding-top:3px;
}}
.config-part-meta{{font-size:13.33px;color:#70839A;margin-top:3px}}
.config-divider{{height:1px;background:#E8EEF5;margin:7px 0}}
/* Rà soát từng phần */
.review-part{{
  border:1px solid #DCE7F3;border-radius:11px;background:#FFFFFF;padding:9px 10px;margin:0 0 8px;
}}
.review-part-title{{font-size:14.83px;font-weight:850;color:#173B65}}
.review-part-line{{font-size:13.83px;color:#40566F;margin-top:4px;line-height:1.45}}
.review-ok{{color:#137A5A;font-weight:850}}
.review-bad{{color:#B33E35;font-weight:850}}
.review-missing{{
  margin-top:6px;padding:6px 8px;border-radius:8px;background:#FFF4F1;
  border:1px solid #F2D0C9;color:#A03B33;font-size:13.33px;font-weight:750;
}}
.review-allok{{
  margin-top:6px;padding:6px 8px;border-radius:8px;background:#ECF8F2;
  border:1px solid #C8ECDD;color:#137A5A;font-size:13.33px;font-weight:750;
}}


/* Phân tích chi tiết đề / nhóm đặt phía trên preview */
.analysis-detail-grid{{display:grid;grid-template-columns:repeat(3,minmax(250px,1fr));gap:7px;margin:5px 0 9px}}
.analysis-card{{background:#fff;border:1px solid #DCE7F3;border-radius:11px;padding:8px 10px;min-height:74px}}
.analysis-card-title{{font-size:14.83px;font-weight:850;color:#173B65;line-height:1.3}}
.analysis-card-meta{{font-size:13.83px;color:#40566F;margin-top:4px;line-height:1.45}}
.analysis-card-ok{{font-size:13.33px;color:#137A5A;font-weight:800;margin-top:5px}}
.analysis-card-bad{{font-size:13.33px;color:#B33E35;font-weight:800;margin-top:5px}}
/* YoungMix cấu hình luôn mở theo hàng ngang */
.ym-config-head{{font-size:18.33px;font-weight:850;color:#173B65;margin:6px 0 5px}}
.ym-config-note{{font-size:13.33px;color:#6C8097;margin-bottom:6px}}
.preview-side-title{{font-size:16.33px;font-weight:850;color:#173B65;margin-bottom:6px}}
.preview-side-box{{background:#F8FBFF;border:1px solid #DCE7F3;border-radius:10px;padding:8px 9px;font-size:13.83px;color:#526A83;line-height:1.5}}
@media(max-width:1200px){{.analysis-detail-grid{{grid-template-columns:repeat(2,1fr)}}}}


.codes-live{{
  margin-top:4px;padding:7px 9px;border-radius:9px;
  background:#F2F7FD;border:1px solid #DCE8F4;color:#274761;
  font-size:13.83px;line-height:1.35;word-break:break-word;
}}


/* Giá trị trong các ô nhập căn giữa, dễ quan sát trên một hàng */
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input{{
  text-align:center!important;
}}


/* Căn giữa nhãn và giá trị của các ô thông tin đầu đề */
[data-testid="stTextInput"] [data-testid="stWidgetLabel"]{{
  width:100%!important;
  display:flex!important;
  justify-content:center!important;
  text-align:center!important;
}}
[data-testid="stTextInput"] [data-testid="stWidgetLabel"] p{{
  width:100%!important;
  text-align:center!important;
}}
[data-testid="stTextInput"] input{{
  text-align:center!important;
}}



/* Tăng thêm khoảng 1pt cho chữ điều khiển của Streamlit */
.stButton > button,
.stDownloadButton > button,
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-baseweb="select"] *,
[data-testid="stRadio"] label p,
[data-testid="stCheckbox"] label p,
[data-testid="stToggle"] label p,
.stCaptionContainer p{{
  font-size:15.8px!important;
}}
</style>

<div class="hero">
 <div class="hero-row">
   <div class="hero-icon">🧪</div>
   <div>
     <div class="hero-title">DTMIX Online <span style="font-size:16px;color:var(--primary);font-weight:700">V1.3</span></div>
     <div class="hero-sub">Trộn đề Word trực tuyến • rà soát đáp án • YoungMix g1/g2/g3 • xem trước giống Word</div>
   </div>
 </div>
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
    "namhoc": "NĂM HỌC 2026 - 2027",
    "monthi": "Môn: HÓA HỌC",
    "thoigian": "Thời gian làm bài: 45 phút",
}
for _k, _v in DEFAULT_HEADER.items():
    st.session_state.setdefault(f"hdr_{_k}", _v)

# Nâng năm học mặc định của các phiên cũ lên 2026 - 2027.
if st.session_state.get("hdr_namhoc") in {
    "NĂM HỌC 2025 - 2026",
    "NĂM HỌC 2025-2026",
    "2025 - 2026",
    "2025-2026",
}:
    st.session_state["hdr_namhoc"] = "NĂM HỌC 2026 - 2027"


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




def _question_number_from_text(raw_text: str, fallback: int) -> int:
    m = re.search(r"(?i)(?:#\s*)?(?:Câu|Question)\s*(\d+)", raw_text or "")
    return int(m.group(1)) if m else int(fallback)


def _normalize_question_anchor(text_value: str) -> str:
    """Chuẩn hóa phần đầu câu hỏi để ghép đúng câu trong DOCX."""
    s = (text_value or "").replace("\uFFFC", " ")
    s = re.sub(r"(?i)^\s*#?\s*(?:Câu|Question)\s*\d+\s*[\.\:\)]*\s*", "", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    # 120 ký tự đủ phân biệt các câu nhưng không quá nhạy với công thức/ảnh.
    return s[:120]


def _answer_format_evidence(answer: dict) -> int:
    """
    Điểm bằng chứng định dạng trong CHÍNH phương án:
    - màu đỏ có trọng số cao;
    - gạch chân tính theo số ký tự.
    Chỉ dùng làm fallback khi parser trả 0 hoặc >1 đáp án ở Phần I.
    """
    red_score = 0
    underline_score = 0
    for elm in answer.get("elements", []):
        try:
            runs = elm.iter(qn("w:r"))
        except Exception:
            continue
        for run in runs:
            txt = "".join((t.text or "") for t in run.iter(qn("w:t")))
            weight = max(1, len(txt.strip()))
            rpr = run.find(qn("w:rPr"))
            if rpr is None:
                continue

            u = rpr.find(qn("w:u"))
            if u is not None:
                uval = (u.get(qn("w:val")) or "single").lower()
                if uval not in ("none", "0", "false"):
                    underline_score += weight

            color = rpr.find(qn("w:color"))
            if color is not None:
                cval = (color.get(qn("w:val")) or "").replace("#", "").upper()
                if len(cval) == 6:
                    try:
                        rr = int(cval[0:2], 16)
                        gg = int(cval[2:4], 16)
                        bb = int(cval[4:6], 16)
                        if rr > 150 and gg < 130 and bb < 130:
                            red_score += weight
                    except Exception:
                        pass

    return red_score * 4 + underline_score


def _correct_letters_for_preview(answers: list, q_type: int) -> list[str]:
    """
    Chọn đáp án cho PREVIEW.
    - Phần I: bắt buộc tối đa 1 phương án đỏ.
    - Phần II: có thể nhiều mệnh đề đúng.
    - Không tự suy diễn nếu dữ liệu mơ hồ; thà không tô hơn tô sai.
    """
    letters = "ABCDEFG"
    parser_true = [
        letters[i]
        for i, ans in enumerate(answers[:len(letters)])
        if ans.get("is_true")
    ]

    if q_type != 1:
        return parser_true

    # Trắc nghiệm nhiều lựa chọn: nếu parser đã xác định duy nhất 1 đáp án, dùng luôn.
    if len(parser_true) == 1:
        return parser_true

    # Parser trả 0 hoặc nhiều đáp án: dùng bằng chứng định dạng trong từng phương án.
    scores = [_answer_format_evidence(ans) for ans in answers[:len(letters)]]
    if not scores:
        return []

    best = max(scores)
    if best <= 0:
        return []

    winners = [i for i, score in enumerate(scores) if score == best]

    # Chỉ tô khi có một phương án nổi bật duy nhất.
    if len(winners) == 1:
        return [letters[winners[0]]]

    # Nếu đồng điểm nhưng parser có đúng 1 phương án nằm trong nhóm đồng điểm, dùng nó.
    if len(parser_true) == 1:
        idx = letters.index(parser_true[0])
        if idx in winners:
            return parser_true

    # Mơ hồ => không tô đỏ để tránh tô sai.
    return []


def build_answer_annotation_spec(engine: DTMIXWebEngine) -> dict:
    """
    Bản đồ câu/đáp án dùng riêng cho xem trước.
    Mỗi câu có anchor nội dung để DOCX preview ghép đúng câu, tránh nhầm khi
    số câu lặp lại giữa PHẦN I / II / III.
    """
    spec = {"youngmix": bool(engine.youngmix), "parts": [], "groups": []}

    if not engine.youngmix:
        for p_idx, part in enumerate(engine.app.parsed_data.get("parts", [])):
            p_type = int(part.get("type", 1))
            q_seq = 0
            questions = []
            for muc in part.get("mucs", []):
                for q in muc.get("questions", []):
                    if q.get("is_virtual"):
                        continue
                    q_seq += 1
                    raw = q.get("raw_text", "")
                    answers = q.get("answers", [])
                    questions.append({
                        "number": _question_number_from_text(raw, q_seq),
                        "anchor": _normalize_question_anchor(raw),
                        "correct_letters": _correct_letters_for_preview(answers, p_type),
                        "q_type": p_type,
                        "short_answer": extract_short_answer(engine, q) if p_type == 3 else "",
                    })
            spec["parts"].append({
                "p_idx": p_idx,
                "type": p_type,
                "questions": questions,
            })
    else:
        all_mucs = [
            m
            for p in engine.app.parsed_data.get("parts", [])
            for m in p.get("mucs", [])
        ]
        for g_idx, muc in enumerate(all_mucs):
            q_seq = 0
            questions = []
            for q in muc.get("questions", []):
                if q.get("is_virtual"):
                    continue
                q_seq += 1
                q_type = int(q.get("ym_type", 1) or 1)
                raw = q.get("raw_text", "")
                answers = q.get("answers", [])
                questions.append({
                    "number": _question_number_from_text(raw, q_seq),
                    "anchor": _normalize_question_anchor(raw),
                    "correct_letters": _correct_letters_for_preview(answers, q_type),
                    "q_type": q_type,
                    "short_answer": extract_short_answer(engine, q) if q_type == 3 else "",
                })
            spec["groups"].append({
                "g_idx": g_idx,
                "questions": questions,
            })
    return spec


@st.cache_data(show_spinner=False, ttl=3600, max_entries=8)
def _prepare_browser_preview_docx(file_bytes: bytes, signature: str, annotation_json: str = "") -> tuple[bytes, dict]:
    """
    Tạo DOCX CHỈ phục vụ preview.
    Điểm quan trọng: đáp án đỏ lấy từ parsed_data của DTMIX, không tô theo heuristic
    "đoạn nào có chữ đỏ/gạch chân thì cả phương án đỏ" như các bản trước.
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    XML = "http://www.w3.org/XML/1998/namespace"
    NS = {"w": W}
    stats = {"converted_wmf_emf": 0, "wmf_emf_total": 0, "answer_marks": 0}
    try:
        spec = json.loads(annotation_json) if annotation_json else {"youngmix": False, "parts": [], "groups": []}
    except Exception:
        spec = {"youngmix": False, "parts": [], "groups": []}

    def _run_text(run):
        return "".join((t.text or "") for t in run.xpath(".//w:t", namespaces=NS))

    def _ensure_rpr(run):
        rpr = run.find(f"{{{W}}}rPr")
        if rpr is None:
            rpr = etree.Element(f"{{{W}}}rPr")
            run.insert(0, rpr)
        return rpr

    def _set_color(run, value):
        rpr = _ensure_rpr(run)
        color = rpr.find(f"{{{W}}}color")
        if color is None:
            color = etree.SubElement(rpr, f"{{{W}}}color")
        color.set(f"{{{W}}}val", value)
        for attr in ("themeColor", "themeTint", "themeShade"):
            color.attrib.pop(f"{{{W}}}{attr}", None)

    def _neutralize_existing_red(root):
        # Preview theo quy ước mới: chỉ đáp án DTMIX nhận diện được phép màu đỏ.
        for color in root.xpath('.//w:rPr/w:color', namespaces=NS):
            val = (color.get(f"{{{W}}}val") or "").replace("#", "").upper()
            if len(val) == 6:
                try:
                    rr, gg, bb = int(val[:2],16), int(val[2:4],16), int(val[4:],16)
                    if rr > 150 and gg < 130 and bb < 130:
                        color.set(f"{{{W}}}val", "000000")
                        for attr in ("themeColor", "themeTint", "themeShade"):
                            color.attrib.pop(f"{{{W}}}{attr}", None)
                except Exception:
                    pass

    def _simple_text_run(run):
        # Chỉ split run thuần text; run có drawing/object/math thì giữ nguyên để tránh hỏng công thức/ảnh.
        allowed = {f"{{{W}}}rPr", f"{{{W}}}t"}
        return all(ch.tag in allowed for ch in list(run)) and bool(run.xpath('./w:t', namespaces=NS))

    def _clone_text_run(run, txt, color=None):
        nr = copy.deepcopy(run)
        for t in nr.xpath('.//w:t', namespaces=NS):
            parent = t.getparent()
            parent.remove(t)
        nt = etree.Element(f"{{{W}}}t")
        if txt.startswith(' ') or txt.endswith(' '):
            nt.set(f"{{{XML}}}space", "preserve")
        nt.text = txt
        nr.append(nt)
        if color:
            _set_color(nr, color)
        return nr

    def _apply_option_intervals(p, infos, intervals):
        # intervals: [(start,end,'FF0000'/'000000')]. Xử lý một lần để không tô cả run
        # khi nhiều phương án A/B/C/D cùng nằm chung một run Word.
        for run, rs, re_, txt in list(infos):
            hits = []
            for a,b,color in intervals:
                if b > rs and a < re_:
                    hits.append((max(a,rs)-rs, min(b,re_)-rs, color))
            if not hits:
                continue
            if not _simple_text_run(run):
                # Nếu run phức tạp nhưng nằm trọn trong một phương án, đổi màu được an toàn.
                full = [h for h in hits if h[0] <= 0 and h[1] >= len(txt)]
                if full:
                    _set_color(run, full[-1][2])
                continue
            bounds = {0, len(txt)}
            for a,b,_ in hits:
                bounds.add(max(0,a)); bounds.add(min(len(txt),b))
            bounds = sorted(bounds)
            parent = run.getparent()
            if parent is None:
                continue
            pos = parent.index(run)
            pieces = []
            for i in range(len(bounds)-1):
                a,b = bounds[i], bounds[i+1]
                if b <= a: continue
                piece = txt[a:b]
                color = None
                mid = (a+b)/2
                for ia,ib,ic in hits:
                    if ia <= mid < ib:
                        color = ic
                pieces.append(_clone_text_run(run, piece, color))
            for off,nr in enumerate(pieces):
                parent.insert(pos+off, nr)
            parent.remove(run)

    def _normalize_preview_anchor(text_value):
        s = (text_value or "").replace("\uFFFC", " ")
        s = re.sub(r"(?i)^\s*#?\s*(?:Câu|Question)\s*\d+\s*[\.\:\)]*\s*", "", s)
        s = re.sub(r"\s+", " ", s).strip().lower()
        return s[:120]

    def _all_question_specs(mode):
        units = spec.get("groups" if mode == "ym" else "parts", [])
        flat = []
        for unit_idx, unit in enumerate(units):
            for seq_idx, q in enumerate(unit.get("questions", [])):
                item = dict(q)
                item["_unit_idx"] = unit_idx
                item["_seq_idx"] = seq_idx
                flat.append(item)
        return flat

    def _lookup_question_by_anchor(mode, unit_idx, seq_idx, paragraph_text, used_keys):
        """
        Ưu tiên ghép bằng chính nội dung câu hỏi, sau đó mới fallback theo thứ tự.
        Điều này loại lỗi lấy nhầm đáp án của câu khác/phần khác.
        """
        p_anchor = _normalize_preview_anchor(paragraph_text)
        candidates = _all_question_specs(mode)

        # Ưu tiên đúng unit trước.
        ordered = [q for q in candidates if q["_unit_idx"] == unit_idx]
        ordered += [q for q in candidates if q["_unit_idx"] != unit_idx]

        best = None
        best_score = -1

        for q in ordered:
            key = (q["_unit_idx"], q["_seq_idx"])
            if key in used_keys:
                continue
            anchor = q.get("anchor", "")
            if not anchor or not p_anchor:
                continue

            # exact prefix / containment rất đáng tin vì raw_text được lấy từ cùng DOCX.
            if p_anchor.startswith(anchor[:60]) or anchor.startswith(p_anchor[:60]):
                score = 1000 + min(len(anchor), len(p_anchor))
            elif anchor[:45] and anchor[:45] in p_anchor:
                score = 900 + len(anchor[:45])
            else:
                # điểm giao nhau theo token, chỉ dùng để hỗ trợ.
                at = set(re.findall(r"\w+", anchor))
                pt = set(re.findall(r"\w+", p_anchor))
                score = int(100 * len(at & pt) / max(1, len(at | pt)))

            if score > best_score:
                best_score = score
                best = q

        # Chỉ nhận fuzzy nếu khá chắc.
        if best is not None and best_score >= 55:
            return best

        # Fallback đúng unit + đúng thứ tự.
        units = spec.get("groups" if mode == "ym" else "parts", [])
        if units:
            unit_idx = max(0, min(unit_idx, len(units) - 1))
            qs = units[unit_idx].get("questions", [])
            if 0 <= seq_idx < len(qs):
                item = dict(qs[seq_idx])
                item["_unit_idx"] = unit_idx
                item["_seq_idx"] = seq_idx
                return item
        return None

    def _annotate_document_xml(xml_bytes):
        """
        Quy tắc tô đỏ preview:
        1) Xóa màu đỏ cũ khỏi bản preview.
        2) Đi qua câu hỏi theo đúng THỨ TỰ xuất hiện trong DOCX.
        3) Với mỗi câu, chỉ lấy đáp án từ parsed_data của DTMIX.
        4) Chỉ phương án đúng được tô đỏ; các phương án còn lại ép về đen.
        """
        try:
            root = etree.fromstring(xml_bytes)
        except Exception:
            return xml_bytes

        _neutralize_existing_red(root)

        mode = "ym" if spec.get("youngmix") else "std"
        units = spec.get("groups" if mode == "ym" else "parts", [])
        current_unit = 0
        question_seq = -1
        current_question = None
        seen_ym_tag = False
        used_question_keys = set()
        roman_map = {"I": 0, "II": 1, "III": 2, "IV": 3}

        # Chỉ coi là câu hỏi khi "Câu n" nằm ở ĐẦU đoạn.
        # Nhờ vậy câu mô tả kiểu "trả lời từ câu 1 đến câu 12" không bị hiểu nhầm.
        question_heading_re = re.compile(
            r"(?i)^\s*#?\s*(?:Câu|Question)\s*\d+\b"
        )

        for p in root.xpath('.//w:p', namespaces=NS):
            runs = p.xpath('./w:r', namespaces=NS)
            infos = []
            pos = 0
            for run in runs:
                txt = _run_text(run)
                if txt:
                    infos.append((run, pos, pos + len(txt), txt))
                    pos += len(txt)

            full = ''.join(x[3] for x in infos)
            if not full.strip():
                continue

            # ----- Xác định phần/nhóm -----
            if mode == "std":
                pm = re.search(r'(?i)\bPHẦN\s+(IV|III|II|I)\b', full)
                if pm:
                    candidate = roman_map.get(pm.group(1).upper(), current_unit)
                    if candidate < len(units):
                        current_unit = candidate
                        question_seq = -1
                        current_question = None
            else:
                # Với YoungMix có tag thật: mỗi <g0>...<g4> mở một group mới.
                # Với YoungMix "Tự động" không có tag: giữ nguyên 1 group và đi tuần tự toàn đề.
                if re.search(r'(?i)<\s*#?g[0-4]\s*>', full):
                    if seen_ym_tag:
                        current_unit = min(current_unit + 1, max(0, len(units) - 1))
                    else:
                        seen_ym_tag = True
                        current_unit = 0
                    question_seq = -1
                    current_question = None

            # ----- Bắt đầu câu mới theo thứ tự xuất hiện -----
            if question_heading_re.search(full):
                question_seq += 1
                current_question = _lookup_question_by_anchor(
                    mode,
                    current_unit,
                    question_seq,
                    full,
                    used_question_keys,
                )
                if current_question:
                    used_question_keys.add((
                        current_question.get("_unit_idx", current_unit),
                        current_question.get("_seq_idx", question_seq),
                    ))

            if not current_question:
                continue

            correct_list = list(current_question.get("correct_letters", []))
            q_type_now = int(current_question.get("q_type", 0) or 0)

            # PHẦN I là trắc nghiệm một lựa chọn: preview không bao giờ được tô >1 đáp án.
            if q_type_now == 1 and len(correct_list) > 1:
                correct_list = []
            correct = set(correct_list)

            # ----- Trắc nghiệm / Đúng-Sai -----
            # Nhận cả A. B. C. D. và a) b) c) d), nhưng chỉ tô đỏ nếu parsed_data nói là đúng.
            marks = list(
                re.finditer(
                    r'(?i)(?:^|[\s\u00A0\u200B\uFFFC])#?([A-G])[\.\)]',
                    full,
                )
            )
            if marks:
                intervals = []
                for i, m in enumerate(marks):
                    seg_start = m.start(1)
                    seg_end = marks[i + 1].start(1) if i + 1 < len(marks) else len(full)
                    letter = m.group(1).upper()
                    is_correct = letter in correct
                    intervals.append(
                        (seg_start, seg_end, 'FF0000' if is_correct else '000000')
                    )
                    if is_correct:
                        stats['answer_marks'] += 1

                _apply_option_intervals(p, infos, intervals)
                continue

            # ----- Trả lời ngắn -----
            if int(current_question.get('q_type', 0) or 0) == 3:
                # Chỉ tô chính đoạn đáp án, không tô các dòng khác của câu.
                if re.search(r'(?i)^\s*Đáp\s*án\s*[:\.]', full):
                    for run, *_ in infos:
                        _set_color(run, 'FF0000')
                    stats['answer_marks'] += 1
                elif re.search(r'(?i)^\s*A\.\s*[-+]?\d', full):
                    for run, *_ in infos:
                        _set_color(run, 'FF0000')
                    stats['answer_marks'] += 1

        return etree.tostring(
            root,
            xml_declaration=True,
            encoding='UTF-8',
            standalone='yes',
        )

    try:
        zin=zipfile.ZipFile(io.BytesIO(file_bytes),'r')
        infos_zip=zin.infolist()
        entries={info.filename: zin.read(info.filename) for info in infos_zip}
        zin.close()
        if 'word/document.xml' in entries:
            entries['word/document.xml']=_annotate_document_xml(entries['word/document.xml'])

        legacy=[n for n in entries if n.lower().startswith('word/media/') and n.lower().endswith(('.wmf','.emf'))]
        stats['wmf_emf_total']=len(legacy)
        converter=shutil.which('magick') or shutil.which('convert')
        if legacy and converter:
            temp_root=Path(tempfile.mkdtemp(prefix='dtmix_vec_'))
            try:
                for idx,name in enumerate(legacy):
                    ext=Path(name).suffix.lower(); srcf=temp_root/f'source_{idx}{ext}'; dstf=temp_root/f'result_{idx}.png'
                    srcf.write_bytes(entries[name])
                    try:
                        subprocess.run([converter,str(srcf),str(dstf)],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,timeout=20)
                        if dstf.exists() and dstf.stat().st_size>32:
                            entries[name]=dstf.read_bytes(); stats['converted_wmf_emf']+=1
                    except Exception:
                        pass
            finally:
                shutil.rmtree(temp_root,ignore_errors=True)
            if stats['converted_wmf_emf'] and '[Content_Types].xml' in entries:
                try:
                    ct=etree.fromstring(entries['[Content_Types].xml'])
                    for node in ct:
                        if (node.get('Extension') or '').lower() in ('wmf','emf'):
                            node.set('ContentType','image/png')
                    entries['[Content_Types].xml']=etree.tostring(ct,xml_declaration=True,encoding='UTF-8',standalone='yes')
                except Exception: pass

        out_io=io.BytesIO()
        with zipfile.ZipFile(out_io,'w',zipfile.ZIP_DEFLATED) as zout:
            seen=set()
            for info in infos_zip:
                if info.filename in entries:
                    zout.writestr(info,entries[info.filename]); seen.add(info.filename)
            for name,data in entries.items():
                if name not in seen: zout.writestr(name,data)
        return out_io.getvalue(), stats
    except Exception:
        return file_bytes, stats


def browser_docx_preview(engine: DTMIXWebEngine, key_prefix: str, height: int = 920) -> None:
    """
    Trình xem DOCX chạy trực tiếp trong trình duyệt bằng docx-preview 0.4.0.
    Ưu điểm:
      • Không phụ thuộc LibreOffice trên máy chủ.
      • Hỗ trợ ảnh, bảng, header/footer và Office Math (OMML).
      • Dùng toàn bộ chiều rộng vùng preview ~70%.
    Với MathType/OLE legacy dạng WMF/EMF, DTMIX thử chuyển sang PNG nếu máy chủ có ImageMagick.
    """
    # Xem trực tiếp CHÍNH FILE ĐỀ GỐC người dùng đã tải lên.
    # Không chỉnh màu đáp án, không chèn đánh dấu và không thay nội dung DOCX.
    preview_bytes = engine.file_bytes
    b64 = base64.b64encode(preview_bytes).decode("ascii")
    js_data = json.dumps(b64)
    conversion_note = " • đang hiển thị nguyên bản DOCX đã tải lên"

    html_doc = f"""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/docx-preview@0.4.0/dist/docx-preview.min.js"></script>
<style>
  *{{box-sizing:border-box}}
  html,body{{margin:0;background:#F3F7FC;color:#1F3147;font-family:Inter,Segoe UI,Arial,sans-serif}}
  .topbar{{
    position:sticky;top:0;z-index:10;display:flex;align-items:center;justify-content:space-between;
    gap:10px;background:rgba(255,255,255,.96);border-bottom:1px solid #DCE7F3;padding:9px 12px;
    box-shadow:0 3px 12px rgba(38,75,114,.06)
  }}
  .status{{font-size:13px;font-weight:750;color:#176BCE}}
  .tools{{display:flex;gap:7px;align-items:center}}
  button{{border:1px solid #CFE0F1;background:#fff;color:#244B76;border-radius:8px;padding:6px 9px;font-weight:750;cursor:pointer}}
  button:hover{{background:#EEF6FF}}
  #wrap{{padding:14px;overflow:auto;min-height:{height-46}px}}
  #doc{{transform-origin:top center;transition:transform .12s ease}}
  .docx-wrapper{{background:#EDF3FA!important;padding:16px!important}}
  .docx-wrapper>section.docx{{
    margin:0 auto 18px!important;box-shadow:0 6px 22px rgba(26,60,96,.13)!important;
    border:1px solid #D9E3EE!important;background:white!important
  }}
  math{{font-family:"Cambria Math","STIX Two Math","Times New Roman",serif!important}}
  img{{max-width:100%!important}}
  .err{{margin:20px;padding:14px;border:1px solid #F0C9CE;background:#FFF2F3;color:#A43B48;border-radius:10px}}
</style>
</head>
<body>
<div class="topbar">
  <div class="status" id="status">Đang dựng đề Word trong trình duyệt…{conversion_note}</div>
  <div class="tools">
    <button onclick="zoomOut()">−</button>
    <button id="zoomLabel" onclick="resetZoom()">100%</button>
    <button onclick="zoomIn()">+</button>
    <button onclick="fitWidth()">Vừa khung</button>
  </div>
</div>
<div id="wrap"><div id="doc"></div></div>
<script>
let scale=1;
const data64={js_data};
function bytesFromBase64(b64){{
  const bin=atob(b64); const arr=new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
  return arr;
}}
function applyZoom(){{
  document.getElementById('doc').style.zoom=scale;
  document.getElementById('zoomLabel').textContent=Math.round(scale*100)+'%';
}}
function zoomIn(){{scale=Math.min(1.4,scale+.1);applyZoom()}}
function zoomOut(){{scale=Math.max(.6,scale-.1);applyZoom()}}
function resetZoom(){{scale=1;applyZoom()}}
function fitWidth(){{
  const wrap=document.getElementById('wrap');
  const page=document.querySelector('section.docx');
  if(!page) return;
  const w=page.getBoundingClientRect().width/scale;
  scale=Math.min(1,(wrap.clientWidth-38)/w);
  applyZoom();
}}
(async()=>{{
 try{{
   if(!window.docx || !window.docx.renderAsync) throw new Error('Không tải được thư viện docx-preview.');
   const data=bytesFromBase64(data64);
   await window.docx.renderAsync(data, document.getElementById('doc'), null, {{
      inWrapper:true,
      ignoreWidth:false,
      ignoreHeight:false,
      ignoreFonts:false,
      breakPages:true,
      debug:false,
      experimental:true,
      renderHeaders:true,
      renderFooters:true,
      renderFootnotes:true,
      renderEndnotes:true,
      useBase64URL:true,
      ignoreLastRenderedPageBreak:false
   }});
   document.getElementById('status').textContent='✓ Đề gốc DOCX đã tải lên{conversion_note}';
   setTimeout(fitWidth,250);
 }}catch(e){{
   document.getElementById('status').textContent='Không dựng được DOCX trực tiếp.';
   document.getElementById('doc').innerHTML='<div class="err"><b>Lỗi xem trước:</b> '+String(e)+'</div>';
 }}
}})();
</script>
</body>
</html>
"""
    components.html(html_doc, height=height, scrolling=True)


@st.cache_data(show_spinner=False, ttl=3600, max_entries=8)
def _docx_to_pdf_bytes(file_bytes: bytes, signature: str) -> tuple[bytes | None, str | None]:
    """Chuyển DOCX -> PDF bằng LibreOffice để giữ công thức, WMF/OLE, ảnh và bố cục."""
    office = shutil.which("libreoffice") or shutil.which("soffice")
    if not office:
        return None, "Máy chủ chưa có LibreOffice."
    tmp = Path(tempfile.mkdtemp(prefix="dtmix_preview_"))
    try:
        src = tmp / "preview.docx"
        src.write_bytes(file_bytes)
        cmd = [office, "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(src)]
        p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=75)
        pdf = tmp / "preview.pdf"
        if not pdf.exists():
            return None, (p.stdout or "Không tạo được PDF xem trước.").strip()
        return pdf.read_bytes(), None
    except subprocess.TimeoutExpired:
        return None, "LibreOffice xử lý quá thời gian cho phép."
    except Exception as exc:
        return None, str(exc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@st.cache_data(show_spinner=False, ttl=3600, max_entries=24)
def _pdf_page_png(pdf_bytes: bytes, page_index: int, zoom: float = 1.65) -> tuple[bytes | None, int, str | None]:
    if fitz is None:
        return None, 0, "Thiếu thư viện PyMuPDF."
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        total = doc.page_count
        if total <= 0:
            return None, 0, "PDF không có trang."
        page_index = max(0, min(page_index, total - 1))
        page = doc.load_page(page_index)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        return pix.tobytes("png"), total, None
    except Exception as exc:
        return None, 0, str(exc)


def exact_word_preview(engine: DTMIXWebEngine, key_prefix: str) -> None:
    """Preview ưu tiên độ trung thực: DOCX -> PDF -> ảnh trang."""
    signature = hashlib.sha256(engine.file_bytes).hexdigest()
    with st.spinner("Đang dựng bản xem trước Word chính xác..."):
        # Bản PDF cũng được dựng trực tiếp từ file đề gốc, không qua bước tô/đánh dấu đáp án.
        pdf_bytes, err = _docx_to_pdf_bytes(engine.file_bytes, signature + "_original")

    if pdf_bytes:
        # Lấy số trang nhẹ nhàng
        _, total, err2 = _pdf_page_png(pdf_bytes, 0, 1.0)
        if total:
            st.session_state.setdefault(f"{key_prefix}_page", 1)
            current = int(st.session_state.get(f"{key_prefix}_page", 1))
            current = max(1, min(current, total))
            st.session_state[f"{key_prefix}_page"] = current

            a, b, c, d = st.columns([.72, 1.15, 1.15, 2.2])
            if a.button("◀", key=f"{key_prefix}_prev", use_container_width=True, disabled=current <= 1):
                st.session_state[f"{key_prefix}_page"] = current - 1
                st.rerun()
            page_no = b.selectbox(
                "Trang",
                list(range(1, total + 1)),
                index=current - 1,
                key=f"{key_prefix}_page_select",
                label_visibility="collapsed",
            )
            if int(page_no) != current:
                st.session_state[f"{key_prefix}_page"] = int(page_no)
                st.rerun()
            if c.button("▶", key=f"{key_prefix}_next", use_container_width=True, disabled=current >= total):
                st.session_state[f"{key_prefix}_page"] = current + 1
                st.rerun()
            d.markdown(
                f'<div class="preview-toolbar"><b>Trang {current}/{total}</b> · Dựng từ Word bằng LibreOffice nên hiển thị tốt công thức Toán/Hóa, ảnh, bảng, WMF/OLE.</div>',
                unsafe_allow_html=True,
            )
            png, _, perr = _pdf_page_png(pdf_bytes, current - 1, 1.7)
            if png:
                st.image(png, use_container_width=True)
                return
            st.warning(f"Không dựng được ảnh trang: {perr}")

    # Fallback khi môi trường chưa có LibreOffice/PyMuPDF
    st.info("Chế độ PDF chưa sẵn sàng trên máy chủ. Hãy chọn “Xem trực tiếp DOCX” ở phía trên; chế độ đó không cần LibreOffice.")
    if err:
        st.caption(err)
    show_rich_preview(rich_standard_part_html(engine, None, True), height=780)


def _default_manual_codes(n: int) -> list[str]:
    """Sinh đúng n mã mặc định, ổn định và dễ nhìn."""
    return [str(111 * i) for i in range(1, n + 1)]


def _normalize_manual_codes(raw: str, n: int) -> list[str]:
    """
    Luôn trả về ĐÚNG n mã.
    - Giữ các mã người dùng đã nhập theo thứ tự.
    - Nếu thiếu thì tự bổ sung mã chưa trùng.
    - Nếu thừa thì cắt còn n.
    """
    codes = [x.strip() for x in re.split(r"[,;\n]+", raw or "") if x.strip()]
    # loại trùng nhưng giữ thứ tự
    unique = []
    for c in codes:
        if c not in unique:
            unique.append(c)
    codes = unique[:n]

    # bổ sung nếu còn thiếu
    candidate_pool = _default_manual_codes(max(n, 24))
    next_seq = 101
    while len(codes) < n:
        candidate = candidate_pool[len(codes)] if len(codes) < len(candidate_pool) else str(next_seq)
        while candidate in codes:
            next_seq += 1
            candidate = str(next_seq)
        codes.append(candidate)
        next_seq += 1
    return codes


def _sync_manual_code_count(prefix: str) -> None:
    """
    Callback chạy NGAY khi ô 'Số đề' thay đổi.
    Nếu đang dùng mã thủ công, tự co/giãn danh sách mã về đúng số đề.
    """
    n = int(st.session_state.get(f"{prefix}_top_n", 4))
    manual_key = f"{prefix}_top_manual"
    current = st.session_state.get(manual_key, "")
    st.session_state[manual_key] = ", ".join(_normalize_manual_codes(current, n))
    st.session_state[f"{prefix}_last_n"] = n


def compact_codes_ui(prefix: str) -> list[str]:
    c1, c2 = st.columns([.8, 1.2], gap="small")

    n = int(
        c1.number_input(
            "Số đề",
            min_value=1,
            max_value=24,
            value=4,
            step=1,
            key=f"{prefix}_top_n",
            on_change=_sync_manual_code_count,
            args=(prefix,),
        )
    )

    kind = c2.selectbox(
        "Kiểu mã",
        ["111, 222...", "Liên tiếp", "Thủ công"],
        key=f"{prefix}_top_kind",
    )

    # Ghi nhớ số đề hiện tại để khi vừa chuyển sang "Thủ công" cũng đồng bộ ngay.
    last_n_key = f"{prefix}_last_n"
    manual_key = f"{prefix}_top_manual"
    if st.session_state.get(last_n_key) != n:
        # Chỉ sửa state trước khi widget manual được tạo.
        st.session_state[manual_key] = ", ".join(
            _normalize_manual_codes(st.session_state.get(manual_key, ""), n)
        )
        st.session_state[last_n_key] = n

    if kind == "111, 222...":
        codes = [str(111 * i) for i in range(1, n + 1)]
        # Không dùng text_input disabled vì widget state có thể giữ giá trị cũ.
        st.markdown(
            f'<div class="codes-live"><b>{n} mã đề:</b> {esc(", ".join(codes))}</div>',
            unsafe_allow_html=True,
        )

    elif kind == "Liên tiếp":
        start_code = st.text_input("Mã đầu", "101", key=f"{prefix}_top_start")
        if start_code.strip().isdigit():
            first = int(start_code.strip())
            width = len(start_code.strip())
            codes = [str(first + i).zfill(width) for i in range(n)]
            st.markdown(
                f'<div class="codes-live"><b>{n} mã đề:</b> {esc(", ".join(codes))}</div>',
                unsafe_allow_html=True,
            )
        else:
            codes = []
            st.caption("⚠ Mã đầu phải là số.")

    else:
        # Bảo đảm state tồn tại và có đúng n mã trước khi tạo widget.
        if manual_key not in st.session_state:
            st.session_state[manual_key] = ", ".join(_default_manual_codes(n))
        else:
            normalized = _normalize_manual_codes(st.session_state[manual_key], n)
            if len([x for x in re.split(r"[,;\n]+", st.session_state[manual_key]) if x.strip()]) != n:
                st.session_state[manual_key] = ", ".join(normalized)

        raw_codes = st.text_input(
            "Mã đề",
            key=manual_key,
            help="Danh sách luôn được đồng bộ theo số đề. Có thể sửa từng mã nếu muốn.",
        )
        codes = _normalize_manual_codes(raw_codes, n)

        # Nếu người dùng vừa sửa danh sách không đủ n mã, vẫn dùng đúng n mã và báo rõ.
        entered = [x.strip() for x in re.split(r"[,;\n]+", raw_codes or "") if x.strip()]
        if len(entered) != n:
            st.caption(f"DTMIX sẽ dùng đúng {n} mã: " + ", ".join(codes))
        else:
            st.caption(f"✓ {n} mã đề")

    return codes


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



M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _local(el) -> str:
    return str(getattr(el, "tag", "")).split("}")[-1]


def _attr_local(el, local: str, default=""):
    for k, v in getattr(el, "attrib", {}).items():
        if str(k).split("}")[-1] == local:
            return v
    return default


def _math_token_html(text: str) -> str:
    if not text:
        return ""
    out = []
    buf = ""
    kind = None
    def flush():
        nonlocal buf, kind
        if not buf:
            return
        safe = esc(buf)
        if kind == "num": out.append(f"<mn>{safe}</mn>")
        elif kind == "id": out.append(f"<mi>{safe}</mi>")
        elif kind == "op": out.append(f"<mo>{safe}</mo>")
        else: out.append(f"<mtext>{safe}</mtext>")
        buf = ""; kind = None
    ops = set("=+-−×÷·*/<>≤≥≈≠±∑∫∞→←↔()[]{}|,:;^")
    for ch in text:
        k = "num" if ch.isdigit() or (ch in ".," and kind == "num") else "id" if ch.isalpha() else "op" if ch in ops else "text"
        if kind is None or k == kind:
            buf += ch; kind = k
        else:
            flush(); buf = ch; kind = k
    flush()
    return "".join(out)


def _omml_children(el) -> str:
    return "".join(_omml_node(c) for c in el)


def _first_child(el, name: str):
    for c in el:
        if _local(c) == name:
            return c
    return None


def _omml_node(el) -> str:
    name = _local(el)
    if name in ("oMath", "oMathPara"):
        return f'<span class="doc-math"><math xmlns="http://www.w3.org/1998/Math/MathML"><mrow>{_omml_children(el)}</mrow></math></span>'
    if name == "t":
        return _math_token_html(el.text or "")
    if name == "r":
        text = "".join((x.text or "") for x in el.iter() if _local(x) == "t")
        return _math_token_html(text)
    if name in ("e", "num", "den", "sup", "sub", "deg", "fName", "lim"):
        return f"<mrow>{_omml_children(el)}</mrow>"
    if name == "f":
        num = _first_child(el, "num"); den = _first_child(el, "den")
        return f"<mfrac>{_omml_node(num) if num is not None else '<mrow/>'}{_omml_node(den) if den is not None else '<mrow/>'}</mfrac>"
    if name == "sSup":
        e = _first_child(el, "e"); sup = _first_child(el, "sup")
        return f"<msup>{_omml_node(e) if e is not None else '<mrow/>'}{_omml_node(sup) if sup is not None else '<mrow/>'}</msup>"
    if name == "sSub":
        e = _first_child(el, "e"); sub = _first_child(el, "sub")
        return f"<msub>{_omml_node(e) if e is not None else '<mrow/>'}{_omml_node(sub) if sub is not None else '<mrow/>'}</msub>"
    if name == "sSubSup":
        e = _first_child(el, "e"); sub = _first_child(el, "sub"); sup = _first_child(el, "sup")
        return f"<msubsup>{_omml_node(e) if e is not None else '<mrow/>'}{_omml_node(sub) if sub is not None else '<mrow/>'}{_omml_node(sup) if sup is not None else '<mrow/>'}</msubsup>"
    if name == "rad":
        e = _first_child(el, "e"); deg = _first_child(el, "deg")
        ehtml = _omml_node(e) if e is not None else "<mrow/>"
        dhtml = _omml_node(deg) if deg is not None else ""
        plain_deg = "".join((x.text or "") for x in deg.iter() if _local(x) == "t").strip() if deg is not None else ""
        return f"<mroot>{ehtml}{dhtml}</mroot>" if plain_deg else f"<msqrt>{ehtml}</msqrt>"
    if name == "d":
        e = _first_child(el, "e")
        beg, end = "(", ")"
        pr = _first_child(el, "dPr")
        if pr is not None:
            for x in pr:
                if _local(x) == "begChr": beg = _attr_local(x, "val", beg)
                if _local(x) == "endChr": end = _attr_local(x, "val", end)
        return f'<mfenced open="{esc(beg)}" close="{esc(end)}">{_omml_node(e) if e is not None else "<mrow/>"}</mfenced>'
    if name == "nary":
        e = _first_child(el, "e"); sub = _first_child(el, "sub"); sup = _first_child(el, "sup")
        op = "∫"
        pr = _first_child(el, "naryPr")
        if pr is not None:
            for x in pr:
                if _local(x) == "chr": op = _attr_local(x, "val", op)
        base = f"<mo>{esc(op)}</mo>"
        if sub is not None and sup is not None: ophtml = f"<munderover>{base}{_omml_node(sub)}{_omml_node(sup)}</munderover>"
        elif sub is not None: ophtml = f"<munder>{base}{_omml_node(sub)}</munder>"
        elif sup is not None: ophtml = f"<mover>{base}{_omml_node(sup)}</mover>"
        else: ophtml = base
        return f"<mrow>{ophtml}{_omml_node(e) if e is not None else ''}</mrow>"
    if name == "m":
        rows = []
        for mr in el:
            if _local(mr) != "mr": continue
            cells = [f"<mtd>{_omml_node(c)}</mtd>" for c in mr if _local(c) == "e"]
            rows.append("<mtr>" + "".join(cells) + "</mtr>")
        return "<mtable>" + "".join(rows) + "</mtable>"
    if name == "eqArr":
        rows = [f"<mtr><mtd>{_omml_node(c)}</mtd></mtr>" for c in el if _local(c) == "e"]
        return "<mtable>" + "".join(rows) + "</mtable>"
    if name in ("limLow", "limUpp"):
        e = _first_child(el, "e"); lim = _first_child(el, "lim")
        tag = "munder" if name == "limLow" else "mover"
        return f"<{tag}>{_omml_node(e) if e is not None else '<mrow/>'}{_omml_node(lim) if lim is not None else '<mrow/>'}</{tag}>"
    if name == "func":
        fn = _first_child(el, "fName"); e = _first_child(el, "e")
        return f"<mrow>{_omml_node(fn) if fn is not None else ''}{_omml_node(e) if e is not None else ''}</mrow>"
    if name == "acc":
        e = _first_child(el, "e"); ch = "^"
        pr = _first_child(el, "accPr")
        if pr is not None:
            for x in pr:
                if _local(x) == "chr": ch = _attr_local(x, "val", ch)
        return f"<mover accent='true'>{_omml_node(e) if e is not None else '<mrow/>'}<mo>{esc(ch)}</mo></mover>"
    if name == "bar":
        e = _first_child(el, "e")
        return f"<mover accent='true'>{_omml_node(e) if e is not None else '<mrow/>'}<mo>¯</mo></mover>"
    if name.endswith("Pr") or name in ("ctrlPr",):
        return ""
    return _omml_children(el)


def _image_html(engine: DTMIXWebEngine, node) -> str:
    rid = None
    try:
        for x in node.iter():
            lname = _local(x)
            if lname in ("blip", "imagedata"):
                for k, v in x.attrib.items():
                    an = str(k).split("}")[-1]
                    if an in ("embed", "id"):
                        rid = v; break
            if rid: break
        if not rid:
            return ""
        part = engine.app.global_doc.part.related_parts.get(rid)
        if part is None or not hasattr(part, "blob"):
            return ""
        mime = getattr(part, "content_type", "image/png") or "image/png"
        b64 = base64.b64encode(part.blob).decode("ascii")
        return f'<div class="doc-img-wrap"><img class="doc-img" src="data:{esc(mime)};base64,{b64}"></div>'
    except Exception:
        return ""


def _run_html(engine: DTMIXWebEngine, run_el) -> str:
    style = []
    rpr = _first_child(run_el, "rPr")
    vert = None
    if rpr is not None:
        for x in rpr:
            n = _local(x)
            if n == "b": style.append("font-weight:700")
            elif n == "i": style.append("font-style:italic")
            elif n == "u": style.append("text-decoration:underline")
            elif n == "color":
                val = _attr_local(x, "val", "")
                if val and val.lower() != "auto" and re.fullmatch(r"[0-9A-Fa-f]{6}", val): style.append(f"color:#{val}")
            elif n == "vertAlign": vert = _attr_local(x, "val", "")
    chunks = []
    for c in run_el:
        n = _local(c)
        if n == "t": chunks.append(esc(c.text or ""))
        elif n == "tab": chunks.append("&emsp;")
        elif n in ("br", "cr"): chunks.append("<br>")
        elif n in ("drawing", "pict", "object"):
            chunks.append(_image_html(engine, c) or '<span class="media-chip">Đối tượng Word</span>')
        elif n in ("oMath", "oMathPara"):
            chunks.append(_omml_node(c))
        elif n in ("AlternateContent", "Choice", "Fallback", "sdt", "smartTag"):
            img = _image_html(engine, c)
            if img:
                chunks.append(img)
            for sub in c.iter():
                if sub is c: continue
                if _local(sub) in ("oMath", "oMathPara"):
                    chunks.append(_omml_node(sub))
    inner = "".join(chunks)
    if not inner: return ""
    if vert == "subscript": inner = f"<sub>{inner}</sub>"
    elif vert == "superscript": inner = f"<sup>{inner}</sup>"
    return f'<span style="{";".join(style)}">{inner}</span>' if style else inner


def _paragraph_html(engine: DTMIXWebEngine, p_el) -> str:
    align = ""
    ppr = _first_child(p_el, "pPr")
    if ppr is not None:
        jc = _first_child(ppr, "jc")
        if jc is not None:
            val = _attr_local(jc, "val", "")
            if val in ("center", "right", "both", "justify"):
                align = "justify" if val in ("both", "justify") else val
    out = []
    def walk(parent):
        for c in parent:
            n = _local(c)
            if n == "r": out.append(_run_html(engine, c))
            elif n in ("oMath", "oMathPara"): out.append(_omml_node(c))
            elif n in ("hyperlink", "smartTag", "sdt", "ins"):
                walk(c)
    walk(p_el)
    inner = "".join(out).strip()
    if not inner: inner = "&nbsp;"
    return f'<div class="doc-p {align}">{inner}</div>'


def _table_html(engine: DTMIXWebEngine, tbl_el) -> str:
    rows = []
    for tr in tbl_el:
        if _local(tr) != "tr": continue
        cells = []
        for tc in tr:
            if _local(tc) != "tc": continue
            cell_parts = []
            for c in tc:
                n = _local(c)
                if n == "p": cell_parts.append(_paragraph_html(engine, c))
                elif n == "tbl": cell_parts.append(_table_html(engine, c))
            cells.append("<td>" + "".join(cell_parts) + "</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return '<table class="doc-table">' + "".join(rows) + "</table>"


def _elements_html(engine: DTMIXWebEngine, elements) -> str:
    out = []
    for el in elements or []:
        n = _local(el)
        if n == "p": out.append(_paragraph_html(engine, el))
        elif n == "tbl": out.append(_table_html(engine, el))
    return "".join(out)


def _answer_elements_html(engine: DTMIXWebEngine, answer: dict, show_key: bool) -> str:
    cls = "doc-answer correct" if show_key and answer.get("is_true") else "doc-answer normal"
    return f'<div class="{cls}">{_elements_html(engine, answer.get("elements", []))}</div>'


def rich_standard_part_html(engine: DTMIXWebEngine, part_idx: int | None, show_key: bool) -> str:
    chunks = ['<div class="doc-page">']
    for p_i, part in enumerate(engine.app.parsed_data.get("parts", [])):
        if part_idx is not None and p_i != part_idx: continue
        chunks.append(f'<div class="doc-part-title">{esc(engine._part_label(part))}</div>')
        for muc in part.get("mucs", []):
            if muc.get("title"):
                chunks.append(f'<div class="doc-group-title">{esc(muc.get("title"))}</div>')
            chunks.append(_elements_html(engine, muc.get("description_elements", [])))
            for q in muc.get("questions", []):
                if q.get("is_virtual"): continue
                chunks.append(_elements_html(engine, q.get("q_elements", [])))
                for ans in q.get("answers", []):
                    chunks.append(_answer_elements_html(engine, ans, show_key))
        chunks.append('<div style="height:8px"></div>')
    chunks.append('</div>')
    return "".join(chunks)


def rich_youngmix_group_html(engine: DTMIXWebEngine, group_idx: int, show_key: bool) -> str:
    all_mucs = [m for p in engine.app.parsed_data.get("parts", []) for m in p.get("mucs", [])]
    if group_idx < 0 or group_idx >= len(all_mucs):
        return '<div class="doc-page"><div class="doc-empty">Không tìm thấy nội dung nhóm.</div></div>'
    muc = all_mucs[group_idx]
    chunks = ['<div class="doc-page">']
    if muc.get("title"):
        chunks.append(f'<div class="doc-part-title">{esc(muc.get("title"))}</div>')
    chunks.append(_elements_html(engine, muc.get("description_elements", [])))
    for q in muc.get("questions", []):
        if q.get("is_virtual"): continue
        chunks.append(_elements_html(engine, q.get("q_elements", [])))
        for ans in q.get("answers", []):
            chunks.append(_answer_elements_html(engine, ans, show_key))
    chunks.append('</div>')
    return "".join(chunks)


def show_rich_preview(body_html: str, height: int = 760) -> None:
    css = """
    <style>
    *{box-sizing:border-box} body{margin:0;background:#EEF3F8;font-family:Inter,'Segoe UI',Arial,sans-serif}
    .doc-page{background:#fff;color:#111827;margin:0 auto;padding:26px 32px;min-height:100%;font-family:'Times New Roman',serif;font-size:18px;line-height:1.42;box-shadow:0 0 0 1px #DCE5F0 inset}
    .doc-part-title{font-size:20px;font-weight:800;margin:8px 0}.doc-group-title{font-size:17px;font-weight:700;margin:7px 0;color:#24364D}
    .doc-p{margin:3px 0}.doc-p.center{text-align:center}.doc-p.right{text-align:right}.doc-p.justify{text-align:justify}
    .doc-answer{margin:2px 0 2px 14px;padding:2px 7px;border-radius:6px}.doc-answer.correct{background:#EAF8F2;color:#0B6D4B;border-left:3px solid #28A979;font-weight:700}
    .doc-img-wrap{text-align:center;margin:9px 0}.doc-img{max-width:96%;height:auto}.doc-table{width:100%;border-collapse:collapse;margin:8px 0;font-size:16px}.doc-table td{border:1px solid #BBC7D5;padding:5px;vertical-align:top}
    .doc-math{display:inline-block;vertical-align:middle;margin:0 2px}.doc-math math{font-size:1.08em}.media-chip{display:inline-block;background:#F3F0FF;color:#5B42A8;border-radius:999px;padding:2px 6px;font-size:12px}
    sub,sup{line-height:0}
    </style>
    """
    components.html(css + body_html, height=height, scrolling=True)


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



def _safe_zip_name(source_filename: str) -> str:
    stem = Path(source_filename or "de_goc").stem.strip() or "de_goc"
    stem = re.sub(r'[\\/:*?"<>|]+', "_", stem).strip(" .")
    return f"dtmix_{stem}.zip"


def auto_download_zip(zip_bytes: bytes, filename: str) -> None:
    """Tự tải ZIP sau khi trộn xong."""
    payload = base64.b64encode(zip_bytes).decode("ascii")
    html_download = f"""
<!doctype html>
<html><body style="margin:0;background:transparent">
<script>
(function(){{
  const b64 = {json.dumps(payload)};
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);
  const blob = new Blob([bytes], {{type:"application/zip"}});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = {json.dumps(filename)};
  document.body.appendChild(a);
  setTimeout(() => {{
    a.click();
    setTimeout(() => {{
      URL.revokeObjectURL(url);
      a.remove();
    }}, 1500);
  }}, 200);
}})();
</script>
</body></html>
"""
    components.html(html_download, height=0, scrolling=False)


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
        progress.progress(100, text="Hoàn tất — đang tải ZIP...")
        st.session_state.mix_result = result
        zip_name = _safe_zip_name(engine.filename)
        auto_download_zip(result.zip_bytes, zip_name)
        st.toast(f"Đã trộn xong. Đang tải {zip_name}", icon="✅")
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
# 1 — HEADER + TOOLBAR
# ============================================================
with st.container(border=True):
    title_row, theme_row = st.columns([4.6, 1], gap="small")
    title_row.markdown("**📝 Thông tin đầu trang đề**")
    theme_row.selectbox(
        "Màu giao diện",
        list(_THEME_PRESETS.keys()),
        key="ui_theme",
        label_visibility="collapsed",
    )

    h1, h2 = st.columns([1.15, 1.35], gap="small")
    h1.text_input("Sở GD&ĐT / Phòng", key="hdr_so")
    h2.text_input("Tên trường", key="hdr_truong")
    h3, h4, h5, h6 = st.columns([1.18, .82, 1.0, 1.18], gap="small")
    h3.text_input("Tên kỳ thi", key="hdr_kythi")
    h4.text_input("Năm học", key="hdr_namhoc")
    h5.text_input("Môn thi", key="hdr_monthi")
    h6.text_input("Thời gian làm bài", key="hdr_thoigian")

    st.divider()

    # Bố cục người dùng yêu cầu: Đề gốc | Chế độ | Mã đề | Trộn & xuất
    file_col, mode_col, code_col, action_col = st.columns([2.15, 1.35, 1.55, 1.25], gap="medium")

    current_sig = None
    raw = None
    with file_col:
        st.markdown('<div class="tool-card-title">1. 📄 Đề gốc (.docx)</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Đề gốc",
            type=["docx"],
            accept_multiple_files=False,
            key="source_docx",
            label_visibility="collapsed",
        )
        if uploaded:
            raw = uploaded.getvalue()
            st.markdown(f'<span class="file-pill">📄 {esc(uploaded.name)}</span>', unsafe_allow_html=True)
            st.caption(f"{len(raw)/1024:.1f} KB")

    with mode_col:
        st.markdown('<div class="tool-card-title">2. ⚙️ Chế độ xử lý</div>', unsafe_allow_html=True)
        mode = st.radio(
            "Chế độ",
            ["Tự động PHẦN I–IV", "YoungMix g1/g2/g3/g4"],
            key="dtmix_mode",
            label_visibility="collapsed",
        )
        is_youngmix = mode.startswith("YoungMix")
        st.caption("g1: câu • g2: đáp án • g3: cả hai")

    with code_col:
        st.markdown('<div class="tool-card-title">3. 🏷️ Mã đề / số lượng</div>', unsafe_allow_html=True)
        top_codes = compact_codes_ui("ym_top" if is_youngmix else "std_top")

    # Xác định engine hiện tại có đúng file/chế độ không
    if raw is not None:
        current_sig = hashlib.sha256(raw + str(is_youngmix).encode()).hexdigest()

    existing_engine = st.session_state.get("dtmix_engine")
    engine_ready = bool(
        existing_engine
        and current_sig
        and st.session_state.get("dtmix_signature") == current_sig
    )

    # TỰ ĐỘNG PHÂN TÍCH ngay khi người dùng chọn file hoặc đổi chế độ xử lý.
    # Không cần bấm nút "Phân tích đề".
    auto_analysis_error = None
    if raw is not None and not engine_ready:
        clear_engine()
        with st.spinner("DTMIX đang tự động phân tích câu hỏi, đáp án, hình ảnh, bảng và công thức..."):
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
                existing_engine = eng
                engine_ready = True
            except Exception as exc:
                auto_analysis_error = exc
                engine_ready = False

    with action_col:
        st.markdown('<div class="tool-card-title">4. 🚀 Trộn & xuất</div>', unsafe_allow_html=True)
        mix_clicked = st.button(
            "🚀 TRỘN & TẢI ZIP",
            type="primary",
            use_container_width=True,
            disabled=not engine_ready,
            key="top_mix",
        )
        if engine_ready:
            st.caption("✅ Đã tự động phân tích")
        elif auto_analysis_error is not None:
            st.caption("❌ Phân tích chưa thành công")
        else:
            st.caption("Chọn file DOCX")

    if auto_analysis_error is not None:
        st.error(f"Không phân tích được đề: {auto_analysis_error}")
        with st.expander("Chi tiết lỗi"):
            st.exception(auto_analysis_error)

    if mix_clicked:
        st.session_state["pending_mix"] = True
        st.session_state["pending_codes"] = top_codes

engine = st.session_state.get("dtmix_engine")
if current_sig is not None and st.session_state.get("dtmix_signature") != current_sig:
    engine = None

# ============================================================
# PHÂN TÍCH CHI TIẾT ĐỀ + CẤU HÌNH (ĐẶT TRÊN PREVIEW)
# ============================================================
std_config = None
ym_config = None


def _summary_question_label(q: dict, fallback: int) -> str:
    raw = q.get("raw_text", "") or ""
    m = re.search(r"(?i)(?:#\s*)?(?:Câu|Question)\s*(\d+)", raw)
    return m.group(1) if m else str(fallback)


if engine:
    summary = engine.summary()
    parts = summary.get("parts", [])
    unit_count = len(summary.get("youngmix_groups", [])) if engine.youngmix else len(parts)
    st.markdown(
        f"""
<div class="summary-compact">
  <div class="summary-chip"><div class="v">{unit_count}</div><div class="l">Phần / nhóm nhận diện</div></div>
  <div class="summary-chip"><div class="v">{summary["total_questions"]}</div><div class="l">Tổng số câu</div></div>
  <div class="summary-chip"><div class="v">{summary["valid_answers"]}/{summary["total_questions"]}</div><div class="l">Câu đã có đáp án / hợp lệ</div></div>
  <div class="summary-chip"><div class="v">{summary["missing_answers"]}</div><div class="l">Câu còn thiếu đáp án</div></div>
</div>
""",
        unsafe_allow_html=True,
    )

    # -------- PHÂN TÍCH CHI TIẾT: luôn nằm trên xem trước --------
    detail_cards = []
    if not engine.youngmix:
        roman = {1:"PHẦN I",2:"PHẦN II",3:"PHẦN III",4:"PHẦN IV"}
        for part in parts:
            total = part["question_count"]
            valid = part["valid_count"]
            missing_labels = []
            seq = 0
            for g in part.get("groups", []):
                for q in g.get("questions", []):
                    seq += 1
                    if not q.get("valid_answer"):
                        missing_labels.append(_summary_question_label(q, seq))
            title = roman.get(part.get("type"), part.get("title", "PHẦN"))
            if int(part.get("type", 1)) == 4:
                status = '<div class="analysis-card-ok">Tự luận: không bắt buộc đánh dấu đáp án.</div>'
            elif missing_labels:
                status = f'<div class="analysis-card-bad">Thiếu đáp án: Câu {esc(", ".join(missing_labels))}</div>'
            else:
                status = '<div class="analysis-card-ok">✓ Đủ đáp án cho tất cả câu.</div>'
            detail_cards.append(
                f'<div class="analysis-card"><div class="analysis-card-title">{esc(title)}</div>'
                f'<div class="analysis-card-meta">{total} câu · Có đáp án/hợp lệ: <b>{valid}/{total}</b> · Thiếu: <b>{len(missing_labels)}</b></div>'
                f'{status}</div>'
            )
    else:
        ym_groups = summary.get("youngmix_groups", [])
        flat = [g for p in parts for g in p.get("groups", [])]
        for i, yg in enumerate(ym_groups):
            sg = flat[i] if i < len(flat) else {"questions": [], "question_count": yg.get("question_count",0)}
            qs = sg.get("questions", [])
            total = int(yg.get("question_count", len(qs)))
            valid = sum(1 for q in qs if q.get("valid_answer"))
            missing = [_summary_question_label(q, j+1) for j,q in enumerate(qs) if not q.get("valid_answer")]
            fixed_text = "Cố định" if (yg.get("is_fixed") or True) else "Có thể đổi vị trí"
            if missing:
                status = f'<div class="analysis-card-bad">Thiếu đáp án: Câu {esc(", ".join(missing))}</div>'
            else:
                status = '<div class="analysis-card-ok">✓ Đủ đáp án cho các câu cần đáp án.</div>'
            detail_cards.append(
                f'<div class="analysis-card"><div class="analysis-card-title">{esc(yg.get("name",f"NHÓM {i+1}"))} {esc(yg.get("tag","<g>"))}</div>'
                f'<div class="analysis-card-meta">{total} câu · Có đáp án/hợp lệ: <b>{valid}/{total}</b> · Thiếu: <b>{len(missing)}</b><br>'
                f'{esc(yg.get("q_type",""))} · {esc(yg.get("mix_type",""))} · {fixed_text}</div>{status}</div>'
            )
    st.markdown('<div class="analysis-detail-grid">'+''.join(detail_cards)+'</div>', unsafe_allow_html=True)

    # -------- CẤU HÌNH TRỘN: cũng đặt trên preview --------
    if not engine.youngmix:
        st.markdown('<div class="auto-config-title">⚙️ Cấu hình trộn tự động</div>', unsafe_allow_html=True)
        keep_titles = st.checkbox("Giữ tiêu đề nhóm/mục", value=False, key="std_keep_titles_v64")
        std_groups = {}
        h0,h1,h2,h3,h4 = st.columns([2.4,1.0,1.0,1.35,1.0], gap="small")
        h0.markdown("**Phần / nhóm**"); h1.markdown("**Trộn câu**"); h2.markdown("**Số câu lấy**"); h3.markdown("**Giữ vị trí câu**"); h4.markdown("**Trộn nhóm**")
        for part in parts:
            multi = len(part.get("groups", [])) > 1
            for g in part.get("groups", []):
                key=f"s64_{g['p_idx']}_{g['m_idx']}"
                with st.container(border=True):
                    c0,c1,c2,c3,c4=st.columns([2.4,1.0,1.0,1.35,1.0], gap="small")
                    with c0:
                        title=part["title"] + (f' · {g["title"]}' if multi and g.get("title") else '')
                        st.markdown(f'<div class="config-part-label">{esc(title)}</div><div class="config-part-meta">{g["question_count"]} câu</div>',unsafe_allow_html=True)
                    with c1:
                        shuffle_q=st.toggle("Trộn câu",value=(part["type"]!=4),key=key+"_sq",label_visibility="collapsed")
                    with c2:
                        pick=st.number_input("Số câu lấy",0,g["question_count"],g["question_count"],key=key+"_pick",label_visibility="collapsed")
                    with c3:
                        fixed=st.text_input("Giữ vị trí câu",value=", ".join(str(q["index"]) for q in g["questions"] if q["fixed"]),placeholder="VD: 1, 5",key=key+"_fix",label_visibility="collapsed")
                    with c4:
                        shuffle_group=st.toggle("Trộn nhóm",value=(multi and not g["is_fixed"] and part["type"]!=4),disabled=not multi,key=key+"_sg",label_visibility="collapsed")
                    std_groups[f"{g['p_idx']}:{g['m_idx']}"]={"shuffle_questions":shuffle_q,"pick":int(pick),"fixed_questions":fixed,"shuffle_group_order":shuffle_group}
        std_config={"keep_group_titles":keep_titles,"groups":std_groups}
    else:
        groups=summary.get("youngmix_groups",[])
        st.markdown('<div class="ym-config-head">🩺 Rà soát & cấu hình YoungMix</div>', unsafe_allow_html=True)
        st.markdown('<div class="ym-config-note">Cấu trúc, số câu, tình trạng đáp án và toàn bộ tùy chọn YoungMix được đặt ở đây trước phần xem trước.</div>', unsafe_allow_html=True)
        y0,y1,y2=st.columns([1.55,1.65,2.8],gap="small")
        with y0:
            continuous=st.toggle("Đánh số câu liên tục giữa các nhóm",value=False,key="ym_cont_v64")
        with y1:
            master_fix=st.toggle("Cố định vị trí tất cả các nhóm",value=True,key="ym_master_v64",help="Mặc định bật theo yêu cầu: các nhóm giữ nguyên vị trí.")
        with y2:
            st.caption("Mặc định: cố định vị trí nhóm. Chỉ nội dung bên trong nhóm được trộn theo g1/g2/g3.")

        ym_groups_cfg=[]
        hh0,hh1,hh2,hh3,hh4,hh5=st.columns([1.8,1.45,2.0,.9,1.05,1.1],gap="small")
        hh0.markdown("**Nhóm**"); hh1.markdown("**Loại câu**"); hh2.markdown("**Cách trộn**"); hh3.markdown("**Số câu lấy**"); hh4.markdown("**Cố định nhóm**"); hh5.markdown("**Đánh lại Câu 1**")
        for i,g in enumerate(groups):
            with st.container(border=True):
                c0,c1,c2,c3,c4,c5=st.columns([1.8,1.45,2.0,.9,1.05,1.1],gap="small")
                with c0:
                    st.markdown(f'<div class="config-part-label">{esc(g["name"])} {esc(g["tag"])}</div><div class="config-part-meta">{g["question_count"]} câu</div>',unsafe_allow_html=True)
                type_opts=["TN 2025 - Phần 1","TN 2025 - Phần 2","TN 2025 - Phần 3","Trắc nghiệm","Tự luận"]
                default_type=g["q_type"] if g["q_type"] in type_opts else "Trắc nghiệm"
                with c1:
                    q_type=st.selectbox("Loại câu",type_opts,index=type_opts.index(default_type),key=f"ym_type_v64_{i}",label_visibility="collapsed")
                modes=["g0 · Không hoán vị","g1 · Chỉ trộn câu hỏi","g2 · Chỉ trộn đáp án","g3 · Trộn câu hỏi + đáp án"]
                dm=youngmix_mode_label(g["mix_type"]); dm=dm if dm in modes else modes[-1]
                with c2:
                    mix_label=st.selectbox("Cách trộn",modes,index=modes.index(dm),key=f"ym_mode_v64_{i}",label_visibility="collapsed")
                with c3:
                    pick=st.number_input("Số câu lấy",0,g["question_count"],g["question_count"],key=f"ym_pick_v64_{i}",label_visibility="collapsed")
                with c4:
                    fix=st.checkbox("Cố định nhóm",value=(master_fix or g["is_fixed"]),disabled=master_fix,key=f"ym_fix_v64_{i}",label_visibility="collapsed")
                    st.caption("Có" if fix else "Không")
                with c5:
                    start_q1=st.checkbox("Đánh lại Câu 1",value=((i==0) if continuous else True),disabled=continuous,key=f"ym_start_v64_{i}",label_visibility="collapsed")
                    st.caption("Có" if start_q1 else "Không")
                ym_groups_cfg.append({"q_type":q_type,"mix_type":mix_type_from_label(mix_label),"pick":int(pick),"fix":fix,"start_q1":start_q1})
        ym_config={"continuous_numbering":continuous,"master_fix":master_fix,"groups":ym_groups_cfg}

    # -------- XEM TRƯỚC: phía dưới toàn bộ phân tích/cấu hình --------
    st.markdown('<div class="workspace-title">Xem trước đề online</div>', unsafe_allow_html=True)
    pv,side=st.columns([2.35,1],gap="large")
    with pv:
        with st.container(border=True):
            st.markdown('<div class="preview-heading"><b>👁️ Đề gốc đã nhận diện</b><span class="preview-badge">~70% không gian</span></div>', unsafe_allow_html=True)
            st.markdown('<div class="word-preview-note"><b>Quy ước màu:</b> DTMIX đưa toàn bộ chữ đỏ gốc về màu thường trong bản preview và <b style="color:#D71920">chỉ tô đỏ phương án mà bộ phân tích xác định là đáp án đúng</b>. File Word gốc không bị sửa.</div>',unsafe_allow_html=True)
            if st.session_state.get("preview_mode_v64","Xem trực tiếp DOCX").startswith("Xem trực tiếp"):
                browser_docx_preview(engine,"preview_browser_v64",height=900)
            else:
                exact_word_preview(engine,"preview_exact_v64")
    with side:
        with st.container(border=True):
            st.markdown('<div class="preview-side-title">Hiển thị xem trước</div>',unsafe_allow_html=True)
            preview_mode=st.radio("Kiểu xem",["Xem trực tiếp DOCX","Bản in PDF (nếu máy chủ có LibreOffice)"],key="preview_mode_v64")
            st.markdown('<div class="preview-side-box"><b>Đỏ:</b> đáp án đúng DTMIX đã nhận diện.<br><b>Đen:</b> nội dung/đáp án còn lại.<br><br>Nếu đáp án có công thức hoặc hình ảnh, ít nhất ký hiệu phương án A/B/C/D của đáp án đúng sẽ được tô đỏ.</div>',unsafe_allow_html=True)
else:
    summary=None; parts=[]
    with st.container(border=True):
        st.caption("Tải và phân tích đề để hiển thị cấu trúc, cấu hình và xem trước.")

# ============================================================
# MIX REQUEST FROM TOP TOOLBAR
# ============================================================
if engine and st.session_state.get("pending_mix"):
    st.session_state["pending_mix"] = False
    _codes = st.session_state.get("pending_codes", [])
    if engine.youngmix:
        run_mix(engine, _codes, ym_cfg=ym_config)
    else:
        run_mix(engine, _codes, std_cfg=std_config)

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

**Quy trình:** tải file → phân tích/rà soát → xem trước → cấu hình → chọn mã đề → bấm **TRỘN & TẢI ZIP**. Khi trộn xong, ZIP được tải tự động.
"""
    )

st.markdown(
    '<div class="footer">DTMIX Online 1.3 • Tối ưu màn hình máy tính ở 100% • Preview tô đỏ đáp án theo kết quả phân tích DTMIX</div>',
    unsafe_allow_html=True,
)
