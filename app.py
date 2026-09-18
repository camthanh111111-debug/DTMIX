from __future__ import annotations

import base64
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
.block-container{{max-width:1580px;padding-top:.55rem;padding-bottom:3.2rem}}
#MainMenu,footer,header{{visibility:hidden}}

/* Hero sáng, không nền đen */
.hero{{
  background:linear-gradient(120deg,var(--hero1),var(--hero2));
  border:1px solid color-mix(in srgb,var(--primary) 18%, white);
  border-radius:18px;padding:15px 20px;color:var(--hero-text);
  box-shadow:0 12px 30px rgba(58,91,124,.10);position:relative;overflow:hidden;margin-bottom:10px;
}}
.hero:after{{content:"";position:absolute;width:270px;height:270px;border-radius:50%;right:-100px;top:-120px;background:rgba(255,255,255,.45)}}
.hero-row{{display:flex;align-items:center;gap:14px;position:relative;z-index:2}}
.hero-icon{{width:46px;height:46px;border-radius:13px;display:flex;align-items:center;justify-content:center;background:rgba(255,255,255,.65);border:1px solid rgba(255,255,255,.9);font-size:25px}}
.hero-title{{font-size:26px;font-weight:850;letter-spacing:-.3px;line-height:1}}
.hero-sub{{font-size:12.5px;color:color-mix(in srgb,var(--hero-text) 76%, white);margin-top:5px}}
.hero-tags{{display:flex;gap:6px;flex-wrap:wrap;margin-top:7px}}
.hero-tag{{font-size:10.8px;font-weight:760;padding:4px 8px;border-radius:999px;background:rgba(255,255,255,.62);border:1px solid rgba(255,255,255,.9);color:var(--hero-text)}}
.hero-side{{margin-left:auto;text-align:right;max-width:330px}}
.hero-side b{{font-size:13px}} .hero-side span{{display:block;color:color-mix(in srgb,var(--hero-text) 72%, white);font-size:11px;margin-top:4px}}

.sec{{display:flex;gap:10px;align-items:center;margin:14px 0 7px}}
.sec-no{{width:31px;height:31px;border-radius:10px;background:linear-gradient(135deg,var(--primary),var(--primary2));color:white;display:flex;align-items:center;justify-content:center;font-weight:850;box-shadow:0 5px 12px color-mix(in srgb,var(--primary) 24%, transparent)}}
.sec-title{{font-size:19px;font-weight:820;color:#203044}} .sec-sub{{font-size:12px;color:var(--muted);margin-top:1px}}

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
[data-testid="stWidgetLabel"] p{font-weight:700!important;color:#30455E!important;font-size:13.5px!important}
[data-testid="stMarkdownContainer"] p{line-height:1.48}
[data-testid="stToggle"] label{width:100%!important}
[data-testid="stToggle"] label p{white-space:normal!important;overflow:visible!important;text-overflow:clip!important;font-weight:650!important}
[data-baseweb="select"] *{font-size:13.5px!important}
.stCaptionContainer{color:#73859A!important}
div[data-testid="stHorizontalBlock"]{align-items:flex-start}

/* Thanh công cụ trên cùng */
.tool-card-title{font-size:13.5px;font-weight:850;color:#173B65;margin-bottom:6px}
.top-hint{font-size:11px;color:#6D8096;line-height:1.35;margin-top:5px}
.top-status{font-size:11.5px;font-weight:750;color:#176BCE;background:#EEF6FF;border:1px solid #D4E8FB;border-radius:8px;padding:6px 8px;text-align:center}

/* Thông tin đề dạng ngang */
.exam-info{display:grid;grid-template-columns:1.15fr 1.45fr 1.05fr;gap:0;border:1px solid #DCE7F3;border-radius:12px;overflow:hidden;background:#fff;margin:6px 0 11px}
.exam-info>div{padding:9px 12px;border-right:1px solid #E6EDF5;border-bottom:1px solid #E6EDF5;font-size:12.2px;color:#30455E}
.exam-info>div:nth-child(3n){border-right:0}
.exam-info>div:nth-last-child(-n+3){border-bottom:0}
.exam-info b{color:#173B65;margin-right:5px}

/* Workspace */
.preview-shell{background:#F5F9FE;border:1px solid #DDE8F4;border-radius:14px;padding:8px}
.preview-heading{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:7px}
.preview-heading b{font-size:16px;color:#143A66}
.preview-badge{font-size:11px;font-weight:800;color:#0F7357;background:#E8F8F1;border:1px solid #C5EBDD;border-radius:999px;padding:5px 9px}

/* Right panel clarity */
.control-title{font-size:16px;font-weight:850;color:#173B65;margin:2px 0 8px}
.control-note{font-size:11.5px;color:#657A91;background:#F7FAFE;border:1px solid #DFE8F2;border-radius:9px;padding:7px 9px;margin-bottom:8px}
.stExpander{{background:#fff;border:1px solid var(--line)!important;border-radius:11px!important}}
hr{{border-color:#E7EDF4}}

.status-good{{background:var(--green-bg);border:1px solid #BFE5D6;color:#116B50;border-radius:11px;padding:9px 12px;font-weight:750}}
.status-warn{{background:var(--amber-bg);border:1px solid #EED9AE;color:#955B00;border-radius:11px;padding:9px 12px;font-weight:750}}
.status-bad{{background:var(--red-bg);border:1px solid #EFC5CB;color:#A53240;border-radius:11px;padding:9px 12px;font-weight:750}}
.file-pill{{display:inline-block;background:color-mix(in srgb,var(--primary2) 12%, white);color:var(--primary);border:1px solid color-mix(in srgb,var(--primary2) 28%, white);border-radius:999px;padding:4px 9px;font-size:11.5px;font-weight:720}}
.mini-help{{font-size:11.5px;color:var(--muted);background:var(--soft);border:1px solid var(--line);padding:8px 10px;border-radius:9px;line-height:1.5}}
.audit-row{{display:flex;align-items:center;justify-content:space-between;padding:7px 0;border-bottom:1px solid #EEF2F6;font-size:12.5px;gap:10px}}
.audit-row:last-child{{border-bottom:0}} .audit-ok{{color:#17785B;font-weight:800}} .audit-warn{{color:#A96A0B;font-weight:800}} .audit-bad{{color:#B93E4B;font-weight:800}}
.g-badge{{display:inline-block;border-radius:7px;padding:3px 7px;font-size:10.8px;font-weight:850;background:#EAF3FF;color:#176BCE;border:1px solid #D1E3FA}}
.g1{{background:#EDF8FF;color:#126A9A}} .g2{{background:#F2EEFF;color:#6743B1}} .g3{{background:#EAF8F2;color:#0B7A53}} .g0{{background:#F2F4F7;color:#536174}} .g4{{background:#FFF3E7;color:#A45B00}}
.word-preview-note{{font-size:11.5px;color:#536A82;background:#F7FAFE;border:1px solid #DFE8F2;border-radius:9px;padding:7px 9px;margin:5px 0 9px}}
.preview-toolbar{{font-size:11.5px;color:#536A82;background:var(--soft);border:1px solid var(--line);border-radius:9px;padding:7px 9px}}
.result-ok{{background:var(--green-bg);border:1px solid #BFE5D6;border-radius:11px;color:#116B50;padding:10px 12px;font-weight:750}}
.footer{{text-align:center;color:#93A0AF;font-size:11px;margin-top:22px}}

@media(max-width:1000px){{.hero-side{{display:none}}}}
</style>

<div class="hero">
 <div class="hero-row">
   <div class="hero-icon">🧪</div>
   <div>
     <div class="hero-title">DTMIX Online <span style="font-size:16px;color:var(--primary);font-weight:700">V1.3</span></div>
     <div class="hero-sub">Trộn đề Word trực tuyến • rà soát đáp án • YoungMix g1/g2/g3 • xem trước giống Word</div>
     <div class="hero-tags">
       <span class="hero-tag">DOCX</span><span class="hero-tag">PHẦN I–IV</span>
       <span class="hero-tag">g1 • g2 • g3</span><span class="hero-tag">Công thức • Hình ảnh • Bảng</span>
     </div>
   </div>
   <div class="hero-side"><b>Một trang làm việc duy nhất</b><span>Tải đề → Chọn chế độ → Mã đề → Trộn & xuất</span></div>
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




@st.cache_data(show_spinner=False, ttl=3600, max_entries=8)
def _prepare_browser_preview_docx(file_bytes: bytes, signature: str) -> tuple[bytes, dict]:
    """
    Chuẩn bị DOCX cho trình xem trong trình duyệt.
    - Giữ nguyên DOCX gốc cho engine trộn đề.
    - Chỉ tạo một bản tạm phục vụ preview.
    - Nếu máy chủ có công cụ chuyển WMF/EMF -> PNG, thay ảnh vector legacy
      bằng PNG nhưng giữ đường dẫn quan hệ, giúp trình duyệt hiển thị MathType/OLE cũ tốt hơn.
    """
    stats = {"converted_wmf_emf": 0, "wmf_emf_total": 0}
    try:
        zin = zipfile.ZipFile(io.BytesIO(file_bytes), "r")
        names = zin.namelist()
        legacy = [n for n in names if n.lower().startswith("word/media/") and n.lower().endswith((".wmf", ".emf"))]
        stats["wmf_emf_total"] = len(legacy)
        if not legacy:
            zin.close()
            return file_bytes, stats

        converter = shutil.which("magick") or shutil.which("convert")
        if not converter:
            zin.close()
            return file_bytes, stats

        replacement = {}
        temp_root = Path(tempfile.mkdtemp(prefix="dtmix_vec_"))
        try:
            for name in legacy:
                raw = zin.read(name)
                ext = Path(name).suffix.lower()
                srcf = temp_root / ("source" + ext)
                dstf = temp_root / "result.png"
                srcf.write_bytes(raw)
                if dstf.exists():
                    dstf.unlink()
                cmd = [converter, str(srcf), str(dstf)]
                try:
                    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=20)
                    if dstf.exists() and dstf.stat().st_size > 32:
                        replacement[name] = dstf.read_bytes()
                        stats["converted_wmf_emf"] += 1
                except Exception:
                    pass

            if not replacement:
                zin.close()
                return file_bytes, stats

            out_io = io.BytesIO()
            with zipfile.ZipFile(out_io, "w", zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    data = replacement.get(info.filename, zin.read(info.filename))
                    if info.filename == "[Content_Types].xml":
                        txt = data.decode("utf-8", errors="ignore")
                        if any(k.lower().endswith(".wmf") for k in replacement):
                            txt = re.sub(
                                r'(<Default[^>]*Extension=["\']wmf["\'][^>]*ContentType=["\'])[^"\']+(["\'][^>]*/>)',
                                r'\1image/png\2', txt, flags=re.I
                            )
                        if any(k.lower().endswith(".emf") for k in replacement):
                            txt = re.sub(
                                r'(<Default[^>]*Extension=["\']emf["\'][^>]*ContentType=["\'])[^"\']+(["\'][^>]*/>)',
                                r'\1image/png\2', txt, flags=re.I
                            )
                        data = txt.encode("utf-8")
                    zout.writestr(info, data)
            zin.close()
            return out_io.getvalue(), stats
        finally:
            shutil.rmtree(temp_root, ignore_errors=True)
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
    sig = hashlib.sha256(engine.file_bytes).hexdigest()
    preview_bytes, stats = _prepare_browser_preview_docx(engine.file_bytes, sig)
    b64 = base64.b64encode(preview_bytes).decode("ascii")
    js_data = json.dumps(b64)

    conversion_note = ""
    if stats.get("wmf_emf_total", 0):
        if stats.get("converted_wmf_emf", 0):
            conversion_note = f" • đã chuyển {stats['converted_wmf_emf']}/{stats['wmf_emf_total']} ảnh WMF/EMF sang PNG"
        else:
            conversion_note = f" • phát hiện {stats['wmf_emf_total']} đối tượng WMF/EMF legacy"

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
  .status{{font-size:12px;font-weight:750;color:#176BCE}}
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
   document.getElementById('status').textContent='✓ Đã dựng DOCX: chữ, bảng, ảnh và Office Math/OMML{conversion_note}';
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
        pdf_bytes, err = _docx_to_pdf_bytes(engine.file_bytes, signature)

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


def compact_codes_ui(prefix: str) -> list[str]:
    c1, c2 = st.columns([.8, 1.2], gap="small")
    n = int(c1.number_input("Số đề", 1, 24, 4, step=1, key=f"{prefix}_top_n"))
    kind = c2.selectbox("Kiểu mã", ["111, 222...", "Liên tiếp", "Thủ công"], key=f"{prefix}_top_kind")

    if kind == "111, 222...":
        codes = [str(111 * i) for i in range(1, n + 1)]
        st.text_input("Mã đề", ", ".join(codes), disabled=True, key=f"{prefix}_top_std")
    elif kind == "Liên tiếp":
        start = st.text_input("Mã đầu", "101", key=f"{prefix}_top_start")
        if start.strip().isdigit():
            first = int(start.strip()); width = len(start.strip())
            codes = [str(first + i).zfill(width) for i in range(n)]
            st.caption("→ " + ", ".join(codes))
        else:
            codes = []
            st.caption("⚠ Mã đầu phải là số.")
    else:
        raw_codes = st.text_input("Mã đề", "111, 222, 333, 444", key=f"{prefix}_top_manual")
        codes = [x.strip() for x in re.split(r"[,;\n]+", raw_codes) if x.strip()]
        st.caption(f"{len(codes)} mã")
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
    .doc-math{display:inline-block;vertical-align:middle;margin:0 2px}.doc-math math{font-size:1.08em}.media-chip{display:inline-block;background:#F3F0FF;color:#5B42A8;border-radius:999px;padding:2px 6px;font-size:11px}
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
# 1 — HEADER + TOOLBAR
# ============================================================
sec(1, "Chuẩn bị đề & trộn nhanh", "Mọi thao tác chính nằm trên một hàng: Đề gốc → Chế độ xử lý → Mã đề → Trộn & xuất file.")

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

    with action_col:
        st.markdown('<div class="tool-card-title">4. 🚀 Trộn & xuất</div>', unsafe_allow_html=True)
        analyze_clicked = st.button(
            "🔎 PHÂN TÍCH ĐỀ",
            type="primary",
            use_container_width=True,
            disabled=(raw is None or engine_ready),
            key="top_analyze",
        )
        mix_clicked = st.button(
            "🚀 TRỘN & XUẤT",
            type="primary",
            use_container_width=True,
            disabled=not engine_ready,
            key="top_mix",
        )
        if engine_ready:
            st.caption("✅ Đã phân tích")
        elif raw is not None:
            st.caption("Chưa phân tích")
        else:
            st.caption("Chọn file DOCX")

        if st.session_state.get("mix_result"):
            result = st.session_state["mix_result"]
            st.download_button(
                "⬇ ZIP KẾT QUẢ",
                data=result.zip_bytes,
                file_name=result.zip_name,
                mime="application/zip",
                use_container_width=True,
                key="top_download_zip",
            )

    if analyze_clicked and raw is not None:
        clear_engine()
        with st.spinner("DTMIX đang đọc câu hỏi, đáp án, hình ảnh, bảng và công thức..."):
            try:
                eng = DTMIXWebEngine(raw, uploaded.name, youngmix=is_youngmix, header=header_values())
                st.session_state.dtmix_engine = eng
                st.session_state.dtmix_signature = current_sig
                st.session_state.mix_result = None
                st.rerun()
            except Exception as exc:
                st.error(f"Không phân tích được đề: {exc}")
                with st.expander("Chi tiết lỗi"):
                    st.exception(exc)

    if mix_clicked:
        st.session_state["pending_mix"] = True
        st.session_state["pending_codes"] = top_codes

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
sec(3, "Xem trước đề online • Rà soát • Cấu hình", "Khu vực xem đề chiếm khoảng 70% chiều rộng. DTMIX ưu tiên dựng DOCX trực tiếp trong trình duyệt để hiển thị chữ, ảnh, bảng và công thức Office Math.")

std_config = None
ym_config = None

if not engine:
    with st.container(border=True):
        st.caption("Phân tích đề để mở khu vực xem trước và cấu hình.")
else:
    if not engine.youngmix:
        preview_rows = standard_preview_data(engine)
        issues = audit_standard(engine, preview_rows)
        left, right = st.columns([2.35, 1], gap="large")

        with left:
            with st.container(border=True):
                st.markdown('<div class="preview-heading"><b>👁️ Xem trước đề online</b><span class="preview-badge">70% không gian làm việc</span></div>', unsafe_allow_html=True)
                st.markdown(
                    '<div class="word-preview-note"><b>Trình xem mới:</b> DOCX được dựng trực tiếp trong trình duyệt bằng docx-preview 0.4.0. '
                    'Office Math/OMML, ảnh, bảng, header/footer được ưu tiên hiển thị mà không phụ thuộc LibreOffice.</div>',
                    unsafe_allow_html=True,
                )
                preview_mode = st.radio(
                    "Kiểu xem",
                    ["Xem trực tiếp DOCX", "Bản in PDF (nếu máy chủ có LibreOffice)"],
                    horizontal=True,
                    key="std_preview_mode_v5",
                    label_visibility="collapsed",
                )
                if preview_mode.startswith("Xem trực tiếp"):
                    browser_docx_preview(engine, "std_browser", height=930)
                else:
                    exact_word_preview(engine, "std_exact")

        with right:
            with st.container(border=True):
                st.markdown('<div class="control-title">🩺 Rà soát nhanh</div>', unsafe_allow_html=True)
                if not issues:
                    st.success("Không phát hiện lỗi đáp án/phương án nổi bật.")
                else:
                    err = sum(1 for x in issues if x["level"] == "error")
                    warn = len(issues) - err
                    c1, c2 = st.columns(2)
                    c1.metric("Thiếu/lỗi đáp án", err)
                    c2.metric("Cảnh báo", warn)
                    with st.container(height=175, border=False):
                        for item in issues:
                            icon = "❌" if item["level"] == "error" else "⚠️"
                            st.markdown(
                                f'<div class="audit-row"><span>{icon} <b>{esc(item["part"])}</b> · Câu {item["question"]}</span>'
                                f'<span class="{"audit-bad" if item["level"]=="error" else "audit-warn"}">{esc(item["text"])}</span></div>',
                                unsafe_allow_html=True,
                            )

                st.divider()
                st.markdown('<div class="control-title">⚙️ Cấu hình trộn tự động</div>', unsafe_allow_html=True)
                keep_titles = st.checkbox("Giữ tiêu đề nhóm/mục", value=False, key="std_keep_titles_v3")
                std_groups = {}
                with st.container(height=500, border=False):
                    for part in summary["parts"]:
                        icon = "✅" if part["missing_count"] == 0 else "⚠️"
                        with st.expander(f"{icon} {part['title']} · {part['question_count']} câu", expanded=False):
                            multi = len(part["groups"]) > 1
                            for g in part["groups"]:
                                key = f"s3_{g['p_idx']}_{g['m_idx']}"
                                if multi: st.markdown(f"**{g['title']}**")
                                c1, c2 = st.columns(2)
                                shuffle_q = c1.toggle("Trộn câu", value=(part["type"] != 4), key=key+"_sq")
                                pick = c2.number_input("Số câu lấy", 0, g["question_count"], g["question_count"], key=key+"_pick")
                                c3, c4 = st.columns(2)
                                fixed = c3.text_input(
                                    "Giữ vị trí câu",
                                    value=", ".join(str(q["index"]) for q in g["questions"] if q["fixed"]),
                                    placeholder="VD: 1, 5",
                                    key=key+"_fix",
                                )
                                shuffle_group = c4.toggle(
                                    "Trộn nhóm",
                                    value=(multi and not g["is_fixed"] and part["type"] != 4),
                                    disabled=not multi,
                                    key=key+"_sg",
                                )
                                std_groups[f"{g['p_idx']}:{g['m_idx']}"] = {
                                    "shuffle_questions": shuffle_q,
                                    "pick": int(pick),
                                    "fixed_questions": fixed,
                                    "shuffle_group_order": shuffle_group,
                                }
                std_config = {"keep_group_titles": keep_titles, "groups": std_groups}

    else:
        groups = summary.get("youngmix_groups", [])
        left, right = st.columns([2.35, 1], gap="large")

        with left:
            with st.container(border=True):
                st.markdown('<div class="preview-heading"><b>👁️ Xem trước đề online</b><span class="preview-badge">70% không gian làm việc</span></div>', unsafe_allow_html=True)
                st.markdown(
                    '<div class="word-preview-note"><b>Trình xem mới:</b> DOCX được dựng trực tiếp trong trình duyệt. '
                    'Công thức Office Math/OMML, ảnh, bảng và định dạng Word được ưu tiên hiển thị; g1/g2/g3 vẫn cấu hình ở cột bên phải.</div>',
                    unsafe_allow_html=True,
                )
                preview_mode = st.radio(
                    "Kiểu xem",
                    ["Xem trực tiếp DOCX", "Bản in PDF (nếu máy chủ có LibreOffice)"],
                    horizontal=True,
                    key="ym_preview_mode_v5",
                    label_visibility="collapsed",
                )
                if preview_mode.startswith("Xem trực tiếp"):
                    browser_docx_preview(engine, "ym_browser", height=930)
                else:
                    exact_word_preview(engine, "ym_exact")

        with right:
            with st.container(border=True):
                st.markdown('<div class="control-title">🩺 Rà soát & cấu hình YoungMix</div>', unsafe_allow_html=True)
                if summary["missing_answers"] == 0:
                    st.success("Các câu cần đáp án đã được nhận diện.")
                else:
                    st.warning(f"Có {summary['missing_answers']} câu cần kiểm tra hoặc thuộc dạng tự luận.")

                st.markdown('<div class="control-note">Các tùy chọn bên dưới được viết đầy đủ để dễ hiểu khi sử dụng trên màn hình nhỏ.</div>', unsafe_allow_html=True)
                continuous = st.toggle(
                    "Đánh số câu liên tục giữa các nhóm",
                    value=False,
                    key="ym_cont_v3",
                    help="Bật để số câu chạy liên tục từ nhóm trước sang nhóm sau."
                )
                master_fix = st.toggle(
                    "Cố định vị trí tất cả các nhóm",
                    value=False,
                    key="ym_master_v3",
                    help="Bật để các nhóm giữ nguyên vị trí, chỉ nội dung bên trong nhóm được trộn theo cấu hình."
                )
                show_source_answers = st.toggle(
                    "Hiển thị đáp án đã nhận diện trong phần rà soát",
                    value=True,
                    key="ym_show_answers_v5",
                )
                st.divider()
                st.markdown('<div class="control-title">🧩 Cấu hình g1 / g2 / g3</div>', unsafe_allow_html=True)
                ym_groups_cfg = []
                with st.container(height=540, border=False):
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
                            if default_mode not in mode_options: default_mode = mode_options[-1]
                            mix_label = st.selectbox("Cách trộn", mode_options, index=mode_options.index(default_mode), key=f"ym_mode_v3_{i}")
                            type_opts = ["TN 2025 - Phần 1", "TN 2025 - Phần 2", "TN 2025 - Phần 3", "Trắc nghiệm", "Tự luận"]
                            default_type = g["q_type"] if g["q_type"] in type_opts else "Trắc nghiệm"
                            q_type = st.selectbox("Loại câu", type_opts, index=type_opts.index(default_type), key=f"ym_type_v3_{i}")
                            c3, c4 = st.columns(2)
                            pick = c3.number_input("Số câu lấy", 0, g["question_count"], g["question_count"], key=f"ym_pick_v3_{i}")
                            fix = c4.checkbox("Cố định nhóm", value=(master_fix or g["is_fixed"]), disabled=master_fix, key=f"ym_fix_v3_{i}")
                            start_q1 = st.checkbox("Đánh lại từ Câu 1", value=((i == 0) if continuous else True), disabled=continuous, key=f"ym_start_v3_{i}")
                            ym_groups_cfg.append({
                                "q_type": q_type,
                                "mix_type": mix_type_from_label(mix_label),
                                "pick": int(pick),
                                "fix": fix,
                                "start_q1": start_q1,
                            })
                ym_config = {"continuous_numbering": continuous, "master_fix": master_fix, "groups": ym_groups_cfg}

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
# 4 — RESULTS
# ============================================================
sec(4, "Kết quả", "Sau khi trộn, tải nhanh ZIP ở thanh trên hoặc tải riêng từng file tại đây.")

if not st.session_state.get("mix_result"):
    with st.container(border=True):
        st.caption("Chưa có kết quả. Sau khi rà soát/cấu hình, dùng nút **TRỘN & XUẤT** ở thanh trên cùng.")
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
    '<div class="footer">DTMIX Online 1.3 • Giao diện sáng • Preview Word chính xác • Thuật toán DTMIX 1.3</div>',
    unsafe_allow_html=True,
)
