from __future__ import annotations

import hashlib
import re
from pathlib import Path

import streamlit as st

from engine import DTMIXWebEngine

st.set_page_config(
    page_title="DTMIX 1.3 - Trộn đề trực tuyến",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.block-container{max-width:1450px;padding-top:1rem;padding-bottom:3rem}
.dtmix-top{background:linear-gradient(90deg,#1e3a8a,#2563eb);color:white;padding:14px 22px;border-radius:14px;margin-bottom:12px;display:flex;align-items:center;gap:14px}
.dtmix-title{font-size:30px;font-weight:800}.dtmix-ver{font-size:18px;color:#dbeafe}.dtmix-right{margin-left:auto;color:#dbeafe}
.card{background:#fff;border:1px solid #bfdbfe;border-radius:14px;padding:16px 18px;margin:8px 0 14px 0}
.card-blue{background:#eff6ff}.card-green{background:#ecfdf5}.card-purple{background:#faf5ff}.card-rose{background:#fff1f2}
.small-muted{color:#64748b;font-size:.93rem}.good{color:#059669;font-weight:700}.bad{color:#dc2626;font-weight:700}
.preview{background:white;border:1px solid #e2e8f0;border-radius:10px;padding:14px;max-height:600px;overflow:auto;font-family:"Times New Roman",serif;font-size:17px;line-height:1.45}
[data-testid="stMetric"]{background:white;border:1px solid #e2e8f0;padding:10px;border-radius:12px}
</style>
<div class="dtmix-top"><span class="dtmix-title">DTMIX</span><span class="dtmix-ver">| V.1.3 STREAMLIT</span><span class="dtmix-right">Công cụ trộn đề Word trực tuyến</span></div>
""", unsafe_allow_html=True)

DEFAULT_HEADER = {
    "so": "SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH",
    "truong": "TRƯỜNG THPT DƯƠNG BẠCH MAI",
    "kythi": "KIỂM TRA HỌC KỲ II",
    "namhoc": "NĂM HỌC 2025 - 2026",
    "monthi": "Môn: HÓA HỌC",
    "thoigian": "Thời gian làm bài: 45 phút",
}

for k, v in DEFAULT_HEADER.items():
    st.session_state.setdefault(f"hdr_{k}", v)


def current_header():
    return {k: st.session_state.get(f"hdr_{k}", v) for k, v in DEFAULT_HEADER.items()}


def clear_engine():
    old = st.session_state.pop("dtmix_engine", None)
    if old:
        try: old.close()
        except Exception: pass
    st.session_state.pop("analysis_signature", None)
    st.session_state.pop("mix_result", None)


def codes_ui(prefix: str):
    st.markdown("#### 🔢 Mã đề")
    c1, c2, c3 = st.columns([1.2, 1.3, 2.5])
    n = c1.number_input("Số đề", 1, 24, 4, step=1, key=f"{prefix}_num")
    mode = c2.selectbox("Cách sinh mã", ["111, 222, 333...", "Liên tiếp", "Thủ công"], key=f"{prefix}_code_mode")
    if mode == "111, 222, 333...":
        codes = [str(111 * i) for i in range(1, int(n) + 1)]
        c3.text_input("Danh sách mã đề", ", ".join(codes), disabled=True, key=f"{prefix}_show_default")
    elif mode == "Liên tiếp":
        start = c3.text_input("Mã bắt đầu", "101", key=f"{prefix}_start")
        if start.strip().isdigit():
            width = len(start.strip())
            first = int(start.strip())
            codes = [str(first + i).zfill(width) for i in range(int(n))]
        else:
            codes = []
            c3.error("Mã bắt đầu phải là số.")
    else:
        raw = c3.text_input("Nhập mã đề, cách nhau bằng dấu phẩy", "111, 222, 333, 444", key=f"{prefix}_manual")
        codes = [x.strip() for x in re.split(r"[,;\n]+", raw) if x.strip()]
    st.caption("Mã đề sẽ được ghi vào đầu đề và chân trang giống bản desktop.")
    return codes


def run_mix(engine, codes, header, standard_config=None, youngmix_config=None):
    bar = st.progress(0, text="Đang chuẩn bị...")
    def progress(cur, total, text):
        pct = int(100 * cur / total) if total else 0
        bar.progress(min(max(pct, 0), 100), text=text)
    engine.progress_callback = progress
    try:
        result = engine.mix(codes, header=header, standard_config=standard_config, youngmix_config=youngmix_config)
        bar.progress(100, text="Hoàn tất!")
        st.session_state.mix_result = result
        st.success(f"✅ Đã tạo {len(result.files)} file kết quả.")
    except Exception as exc:
        st.error(f"❌ Không thể trộn đề: {exc}")
        st.exception(exc)


def download_result():
    result = st.session_state.get("mix_result")
    if not result:
        return
    st.download_button(
        "⬇️ TẢI TOÀN BỘ KẾT QUẢ (.ZIP)",
        data=result.zip_bytes,
        file_name=result.zip_name,
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )
    with st.expander("📦 Danh sách file trong gói kết quả", expanded=False):
        for f in result.files:
            st.write("•", f)


tab1, tab2, tab3, tab4 = st.tabs([
    "1. Chọn Đề & Cài Đặt",
    "2. Cấu Trúc & Trộn tự động",
    "3. Trộn nhóm dạng <g3>...",
    "4. Hướng dẫn sử dụng",
])

with tab1:
    left, right = st.columns([1, 1], gap="large")
    with left:
        st.markdown('<div class="card card-blue"><h3>📝 THÔNG TIN ĐẦU TRANG ĐỀ</h3></div>', unsafe_allow_html=True)
        st.text_input("Sở GD&ĐT / Phòng", key="hdr_so")
        st.text_input("Tên trường", key="hdr_truong")
        c1, c2 = st.columns(2)
        c1.text_input("Tên kỳ thi", key="hdr_kythi")
        c2.text_input("Năm học", key="hdr_namhoc")
        c3, c4 = st.columns(2)
        c3.text_input("Môn thi", key="hdr_monthi")
        c4.text_input("Thời gian làm bài", key="hdr_thoigian")

    with right:
        st.markdown('<div class="card card-green"><h3>⚙️ TẢI LÊN TỆP ĐỀ GỐC</h3></div>', unsafe_allow_html=True)
        youngmix = st.toggle("Phần/nhóm dạng <g3>, <g2>... và đề Tiếng Anh", value=False, key="youngmix_mode")
        uploaded = st.file_uploader("📄 Tải file đề gốc (.docx)", type=["docx"], accept_multiple_files=False)
        st.info("Online không cần chọn thư mục lưu. DTMIX sẽ đóng gói toàn bộ kết quả thành ZIP để tải về.")

        if uploaded:
            file_bytes = uploaded.getvalue()
            sig = hashlib.sha256(file_bytes + str(youngmix).encode()).hexdigest()
            st.success(f"Đã chọn: {uploaded.name} — {len(file_bytes)/1024:.1f} KB")
            if st.button("🔎 PHÂN TÍCH ĐỀ", type="primary", use_container_width=True):
                if st.session_state.get("analysis_signature") != sig:
                    clear_engine()
                    with st.spinner("Đang nhận diện cấu trúc, bảng, hình ảnh, công thức và đáp án..."):
                        try:
                            engine = DTMIXWebEngine(file_bytes, uploaded.name, youngmix=youngmix, header=current_header())
                            st.session_state.dtmix_engine = engine
                            st.session_state.analysis_signature = sig
                            st.session_state.mix_result = None
                            st.success("✅ Phân tích đề thành công. Chuyển sang Tab 2 hoặc Tab 3 để cấu hình và trộn.")
                        except Exception as exc:
                            st.error(f"Không đọc được file: {exc}")
                            st.exception(exc)
                else:
                    st.success("Đề này đã được phân tích.")
        else:
            st.warning("Chưa có file nào được chọn.")

        st.markdown("""
        <div class="card">
        <b>💡 DTMIX - CÔNG CỤ TRỘN ĐỀ TRẮC NGHIỆM</b><br>
        • Nhận diện Phần, Câu/Question, phương án, bảng, hình ảnh và công thức.<br>
        • Hoán vị câu hỏi và đáp án; hỗ trợ câu cố định bằng dấu #.<br>
        • Xuất đề gốc, các mã đề, file tổng hợp đáp án và Excel SmartTest 2025.<br>
        • Chế độ YoungMix hỗ trợ &lt;g0&gt;, &lt;g1&gt;, &lt;g2&gt;, &lt;g3&gt;, &lt;g4&gt;.
        </div>
        """, unsafe_allow_html=True)

engine = st.session_state.get("dtmix_engine")

with tab2:
    if not engine:
        st.warning("Hãy tải file và bấm **PHÂN TÍCH ĐỀ** ở Tab 1 trước.")
    elif engine.youngmix:
        st.info("Đề đang được phân tích ở chế độ `<g...>`. Hãy dùng **Tab 3** để cấu hình nhóm và trộn.")
    else:
        summary = engine.summary()
        st.markdown("### 📋 CẤU TRÚC CHI TIẾT CỦA ĐỀ ĐÃ NHẬN DIỆN")
        m1, m2, m3 = st.columns(3)
        m1.metric("Tổng số câu", summary["total_questions"])
        m2.metric("Câu hợp lệ/có đáp án", summary["valid_answers"])
        m3.metric("Cần kiểm tra", summary["missing_answers"])

        left, right = st.columns([1.65, 1], gap="large")
        std_groups = {}
        with left:
            keep_titles = st.checkbox("Giữ tiêu đề Nhóm/Mục khi xuất đề", value=False, key="std_keep_titles")
            for part in summary["parts"]:
                with st.expander(f"{part['title']} — {part['question_count']} câu", expanded=True):
                    if part["missing_count"]:
                        st.warning(f"Có {part['missing_count']} câu cần kiểm tra đáp án/nhận diện.")
                    multi = len(part["groups"]) > 1
                    for g in part["groups"]:
                        key = f"std_{g['p_idx']}_{g['m_idx']}"
                        st.markdown(f"**{g['title']}** — {g['question_count']} câu")
                        cols = st.columns([1, 1, 1.1, 1]) if multi else st.columns([1, 1.4, 1])
                        pos = 0
                        shuffle_group = False
                        if multi:
                            shuffle_group = cols[pos].checkbox("Đảo nhóm", value=(not g['is_fixed'] and part['type'] != 4), key=key+"_gm"); pos += 1
                        shuffle_q = cols[pos].checkbox("Đảo câu", value=(part['type'] != 4), key=key+"_q"); pos += 1
                        fixed = cols[pos].text_input("Cố định câu", value=", ".join(str(q['index']) for q in g['questions'] if q['fixed']), placeholder="VD: 1,3", key=key+"_fix"); pos += 1
                        pick = cols[pos].number_input("Số câu lấy", 0, max(g['question_count'], 0), g['question_count'], key=key+"_pick")
                        std_groups[f"{g['p_idx']}:{g['m_idx']}"] = {
                            "shuffle_group_order": shuffle_group,
                            "shuffle_questions": shuffle_q,
                            "fixed_questions": fixed,
                            "pick": int(pick),
                        }
            with st.expander("📖 Xem trước nội dung đã nhận diện", expanded=False):
                st.markdown(f'<div class="preview">{engine.preview_html()}</div>', unsafe_allow_html=True)

        with right:
            st.markdown("### 📊 RÀ SOÁT ĐỀ GỐC")
            for part in summary["parts"]:
                status = "✅" if part["missing_count"] == 0 else "⚠️"
                st.write(f"{status} **{part['title']}**: {part['question_count']} câu")
                for g in part["groups"]:
                    for q in g["questions"]:
                        if not q["valid_answer"]:
                            st.caption(f"⚠️ {q['raw_text'][:100]} — {q['answer_count']} phương án")
            codes = codes_ui("std")
            st.divider()
            if st.button("🚀 TRỘN ĐỀ VÀ XUẤT FILE", type="primary", use_container_width=True, key="mix_std"):
                run_mix(
                    engine, codes, current_header(),
                    standard_config={"keep_group_titles": keep_titles, "groups": std_groups},
                )
            download_result()

with tab3:
    if not engine:
        st.warning("Hãy tải file và bấm **PHÂN TÍCH ĐỀ** ở Tab 1 trước.")
    elif not engine.youngmix:
        st.info("Đề hiện được phân tích theo PHẦN I/II/III/IV. Nếu muốn dùng `<g...>`, hãy bật chế độ đó ở Tab 1 rồi phân tích lại.")
    else:
        summary = engine.summary()
        groups = summary["youngmix_groups"]
        st.markdown("### 📋 CẤU TRÚC YOUNGMIX / TIẾNG ANH")
        a, b, c = st.columns(3)
        a.metric("Số nhóm", len(groups))
        b.metric("Tổng số câu", summary["total_questions"])
        c.metric("Câu hợp lệ/có đáp án", summary["valid_answers"])

        master1, master2 = st.columns(2)
        continuous = master1.toggle("Đánh số câu liên tiếp", value=False, key="ym_continuous")
        master_fix = master2.toggle("Cố định tất cả các nhóm", value=False, key="ym_master_fix")

        ym_cfg = []
        left, right = st.columns([1.65, 1], gap="large")
        with left:
            for i, g in enumerate(groups):
                with st.expander(f"{g['name']} {g['tag']} — {g['question_count']} câu", expanded=True):
                    c1, c2, c3 = st.columns([1.4, 1.4, 1])
                    type_opts = ["Trắc nghiệm", "Tự luận", "TN 2025 - Phần 1", "TN 2025 - Phần 2", "TN 2025 - Phần 3"]
                    default_type = g["q_type"] if g["q_type"] in type_opts else "Trắc nghiệm"
                    q_type = c1.selectbox("Nhóm câu hỏi", type_opts, index=type_opts.index(default_type), key=f"ym_type_{i}")
                    mix_opts = ["Không hoán vị", "Chỉ câu hỏi", "Chỉ đáp án", "Câu hỏi & đáp án"]
                    default_mix = g["mix_type"] if g["mix_type"] in mix_opts else "Câu hỏi & đáp án"
                    mix_type = c2.selectbox("Hoán vị", mix_opts, index=mix_opts.index(default_mix), key=f"ym_mix_{i}")
                    pick = c3.number_input("Số câu cần lấy", 0, max(g["question_count"], 0), g["question_count"], key=f"ym_pick_{i}")
                    c4, c5 = st.columns(2)
                    fix = c4.checkbox("Cố định nhóm", value=(master_fix or g["is_fixed"]), disabled=master_fix, key=f"ym_fix_{i}")
                    start_default = (i == 0) if continuous else True
                    start_q1 = c5.checkbox("Bắt đầu câu 1", value=start_default, disabled=continuous, key=f"ym_start_{i}")
                    ym_cfg.append({"q_type": q_type, "mix_type": mix_type, "pick": int(pick), "fix": fix, "start_q1": start_q1})
                    if g["preview"]:
                        st.text_area("Xem nhanh nội dung nhóm", g["preview"][:6000], height=170, disabled=True, key=f"ym_prev_{i}")
            with st.expander("📖 Xem trước toàn bộ (đáp án tô đỏ)", expanded=False):
                st.markdown(f'<div class="preview">{engine.preview_html()}</div>', unsafe_allow_html=True)

        with right:
            st.markdown("### 📊 RÀ SOÁT ĐỀ GỐC")
            if summary["missing_answers"]:
                st.warning(f"Có {summary['missing_answers']} câu chưa nhận diện đáp án hoặc là tự luận.")
            else:
                st.success("Đã nhận diện đáp án cho các câu cần đáp án.")
            st.markdown("**Quy ước nhóm:**")
            st.write("`<g0>` Không hoán vị")
            st.write("`<g1>` Chỉ hoán vị câu hỏi")
            st.write("`<g2>` Chỉ hoán vị phương án")
            st.write("`<g3>` Hoán vị câu hỏi và phương án")
            st.write("`<g4>` Tự luận, chỉ hoán vị câu hỏi")
            st.write("`<#g...>` Cố định vị trí nhóm")
            codes = codes_ui("ym")
            if st.button("🚀 TRỘN ĐỀ VÀ XUẤT FILE", type="primary", use_container_width=True, key="mix_ym"):
                run_mix(
                    engine, codes, current_header(),
                    youngmix_config={"continuous_numbering": continuous, "master_fix": master_fix, "groups": ym_cfg},
                )
            download_result()

with tab4:
    st.markdown("## 📖 HƯỚNG DẪN SỬ DỤNG DTMIX")
    st.markdown("""
### 1. Chuẩn bị file gốc (.docx)
- Soạn bằng MS Word và lưu định dạng **.docx**.
- Không dùng đánh số tự động (Numbering) cho câu hỏi/phương án; nên gõ thủ công `Câu 1.`, `Question 1:`, `A.`, `B.`, `C.`, `D.`.
- Chế độ tự động: phân chia bằng các tiêu đề **PHẦN I, PHẦN II, PHẦN III, PHẦN IV**.
- Phần I và II: đáp án đúng được **gạch chân** (khuyến nghị) hoặc **tô đỏ** trong file gốc.
- Phần III: cuối câu có `Đáp án: ...` hoặc `A. ...`, ví dụ `Đáp án: 7`.
- Cố định câu/phương án: thêm dấu `#` liền trước đối tượng.

### 2. Chế độ nhóm `<g...>` / YoungMix / đề Tiếng Anh
- `<g0>`: không trộn.
- `<g1>`: chỉ hoán vị câu hỏi.
- `<g2>`: chỉ hoán vị phương án.
- `<g3>`: hoán vị cả câu hỏi và phương án.
- `<g4>`: nhóm tự luận.
- `<#g1>`, `<#g3>`...: cố định vị trí nhóm.
- Có thể dùng một thẻ nhóm rồi tùy chỉnh lại **loại câu**, **kiểu hoán vị**, **số câu lấy**, **cố định nhóm** ngay trên Tab 3.

### 3. Các bước thực hiện online
1. Điền thông tin đầu trang đề ở Tab 1.
2. Nếu đề dùng `<g...>` hoặc là đề Tiếng Anh cần chia nhóm, bật công tắc YoungMix.
3. Tải file `.docx` và bấm **PHÂN TÍCH ĐỀ**.
4. Rà soát số câu, số phương án và đáp án nhận diện.
5. Chọn số đề/mã đề và cấu hình hoán vị.
6. Bấm **TRỘN ĐỀ VÀ XUẤT FILE**.
7. Tải ZIP; trong đó có các mã đề Word, đề gốc, file tổng hợp đề + đáp án và file Excel/CSV đáp án SmartTest.

### 4. Lưu ý khi chạy trên Streamlit
- Máy chủ Streamlit dùng bộ nhớ tạm. Hãy tải ZIP về máy sau khi trộn xong.
- Không cần Tkinter/CustomTkinter trên máy chủ; giao diện online dùng Streamlit, còn thuật toán Word vẫn tái sử dụng từ DTMIX 1.3.
- Với đề lớn nhiều hình/công thức, nên trộn số lượng mã đề vừa phải trong mỗi lần chạy để giảm RAM.
""")

st.caption("DTMIX Streamlit 1.3 — giao diện online sử dụng thuật toán xử lý Word từ mã nguồn DTMIX 1.3.")
