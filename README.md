# DTMIX Web 1.3

Bản chuyển DTMIX 1.3 từ CustomTkinter sang FastAPI + HTML/CSS/JS. Thuật toán Word gốc nằm trong `dtmix_legacy.py`; `engine.py` là lớp tương thích web.

## Chạy local
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```
Mở http://127.0.0.1:8000

## Triển khai
Start command: `uvicorn app:app --host 0.0.0.0 --port $PORT`
Health check: `/health`

## Lưu ý
File upload/output hiện lưu tạm trên filesystem của server. Với nhiều người dùng hoặc hosting filesystem tạm, nên chuyển sang object storage ở giai đoạn production.
