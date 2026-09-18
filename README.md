# DTMIX 1.3 - Streamlit

Bản online của DTMIX, dùng **Streamlit** cho giao diện và tái sử dụng thuật toán xử lý Word từ DTMIX 1.3.

## Cấu trúc

- `app.py`: giao diện Streamlit.
- `engine.py`: cầu nối headless giữa Streamlit và thuật toán DTMIX.
- `dtmix_legacy.py`: mã nguồn DTMIX 1.3 đã bỏ phụ thuộc giao diện khi chạy server.
- `customtkinter.py`: shim headless để legacy core import được trên Linux.
- `headless_compat.py`: thay thế dialog Tkinter.
- `requirements.txt`: thư viện Python.
- `.streamlit/config.toml`: cấu hình upload và theme.

## Chạy trên máy

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Sau đó mở địa chỉ Streamlit hiển thị trong Terminal, thường là `http://localhost:8501`.

## Đưa lên Streamlit Community Cloud

1. Tạo/ dùng repository GitHub `DTMIX`.
2. Đưa toàn bộ các file trong thư mục này lên **root** của repository.
3. Trong Streamlit Community Cloud, tạo app mới từ GitHub repository.
4. Chọn branch `main` và **Main file path = `app.py`**.
5. Deploy.

## Lưu ý quan trọng

- Không thêm `customtkinter` vào `requirements.txt`; file `customtkinter.py` trong project là shim headless.
- Không cần cài Tkinter trên Streamlit server.
- File được người dùng tải lên và kết quả được lưu tạm trên server, sau đó tải về dưới dạng ZIP.
- Nếu đổi thuật toán trong DTMIX desktop, cần đồng bộ thay đổi tương ứng vào `dtmix_legacy.py`.
