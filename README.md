# DTMIX Online 1.3 V5

Giao diện Streamlit một trang, sáng và hiện đại.

## Điểm mới V5
- Thanh thao tác trên cùng: Đề gốc → Chế độ xử lý → Mã đề → Trộn & xuất.
- Khu vực xem trước chiếm ~70%, cấu hình/rà soát ~30%.
- Preview mặc định bằng `docx-preview 0.4.0` chạy trực tiếp trong trình duyệt:
  - ảnh
  - bảng
  - header/footer
  - Office Math / OMML
  - định dạng DOCX
- Không còn phụ thuộc LibreOffice để xem trước mặc định.
- Chế độ PDF/LibreOffice là lựa chọn phụ.
- Nếu máy chủ có ImageMagick + libwmf, DTMIX thử đổi WMF/EMF legacy sang PNG cho preview.
- Nhãn YoungMix được viết đầy đủ, không còn bị cắt chữ.

## Chạy local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud
Đặt `app.py`, `requirements.txt` và `packages.txt` ở root repo.
Main file: `app.py`.
