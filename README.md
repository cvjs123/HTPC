# HTPC Excel Processing Web App

Ứng dụng xử lý dữ liệu Excel cho công tác phổ cập giáo dục, bao gồm:

- Tổng hợp phiếu điều tra
- Báo cáo M1 / M2
- Danh sách SPC1
- Rà soát dữ liệu thiếu
- Xuất phụ lục
- Tổng hợp 3 biểu Tiểu học

## Yêu cầu

- Python 3.10+
- Pip / virtual environment

## Cài đặt

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# Linux/macOS
source .venv/bin/activate

pip install -r requirements.txt
```

## Chạy ứng dụng

```bash
python web_app.py
```

Mở trình duyệt tại:

```text
http://127.0.0.1:5000
```

## Deploy miễn phí lên Render

1. Tạo repository trên GitHub.
2. Đẩy code lên GitHub.
3. Vào Render và chọn "New" → "Web Service".
4. Chọn GitHub repo.
5. Cấu hình máy chủ:

```bash
pip install -r requirements.txt
```

```bash
gunicorn web_app:app
```

Hoặc dùng file `Procfile` với nội dung:

```text
web: gunicorn web_app:app --bind 0.0.0.0:$PORT
```

## Lưu ý quan trọng

GitHub chỉ lưu trữ mã nguồn. Nếu bạn muốn ứng dụng chạy trực tuyến trên internet, bạn cần deploy lên một nền tảng như:

- Render
- Railway
- Fly.io
- Azure App Service

## Cấu trúc thư mục chính

```text
.
├── app.py
├── aggregate.py
├── web_app.py
├── templates/
├── static/
├── web_uploads/
├── web_outputs/
├── requirements.txt
├── .gitignore
├── README.md
└── ...
```

## Push lên GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin <URL_REPO_GITHUB>
git push -u origin main
```

## Xuất dữ liệu

Các file kết quả sẽ được lưu trong thư mục `web_outputs` khi chạy app trên máy cục bộ.
