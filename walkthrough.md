# Tổng Hợp Phiếu Điều Tra - Walkthrough

Tôi đã hoàn thành việc tạo ứng dụng Python theo yêu cầu của bạn. Ứng dụng này cung cấp một giao diện đồ họa (GUI) đơn giản để tự động trích xuất thông tin từ các `Phiếu điều tra` và điền vào file `Tổng hợp`.

## Những thay đổi đã thực hiện
- [NEW] [`app.py`](file:///g:/My%20Drive/HTPC/app.py): Mã nguồn của ứng dụng.

## Hướng dẫn sử dụng

### Chạy phiên bản Webapp

Cài các thư viện cho webapp:
```bash
pip install -r requirements-web.txt
```

Khởi động máy chủ:
```bash
python web_app.py
```

Mở trình duyệt tại `http://127.0.0.1:5000`. Chọn nhiều file phiếu, chọn mẫu `Tong.xlsx`, bấm **Bắt đầu tổng hợp**, sau đó tải file `KetQua_Tong.xlsx`.

### 1. Cài đặt thư viện cần thiết
Nếu máy tính của bạn chưa có thư viện `openpyxl` (dùng để xử lý file Excel .xlsx), vui lòng mở Command Prompt (cmd) hoặc PowerShell và chạy lệnh sau:
```bash
pip install openpyxl
```

### 2. Chạy ứng dụng
Mở Command Prompt/PowerShell tại thư mục `g:\My Drive\HTPC` và chạy:
```bash
python app.py
```

### 3. Thao tác trên giao diện
1. Nhấn nút **Chọn Files** ở mục "1. Chọn các file Phiếu" và chọn tất cả các file phiếu mà bạn muốn tổng hợp (bạn có thể chọn nhiều file cùng lúc).
2. Nhấn nút **Chọn File** ở mục "2. Chọn file Tổng Hợp" và chọn file `Tong.xlsx`.
3. Nhấn nút **Bắt đầu Tổng Hợp**. 
4. Ứng dụng sẽ đọc từng file Phiếu, tìm thông tin chủ hộ, địa chỉ và thông tin chi tiết từng người trong bảng để điền vào file `Tong.xlsx`. Sau khi hoàn thành, file `Tong.xlsx` sẽ tự động được lưu lại và hiển thị thông báo thành công.

### 4. Tổng hợp biểu Tiểu học
1. Sau khi hoàn thành Bước 1, chọn file mẫu `SPCTH.xlsx`, `M1TH.xlsx` và `XMTH.xlsx` tại **Bước 6**.
2. Nhấn **Tổng hợp 3 biểu Tiểu học**.
3. Phần mềm tạo ba file `KetQua_SPCTH.xlsx`, `KetQua_M1TH.xlsx` và `KetQua_XMTH.xlsx` trong cùng thư mục với file Tổng.

> [!TIP]
> Ứng dụng đã được cấu hình tự tìm dòng tiếp theo còn trống trong file `Tong.xlsx` (dựa vào cột TT) để điền tiếp dữ liệu, do đó bạn có thể chạy ứng dụng nhiều lần cho các đợt file phiếu khác nhau mà không sợ bị ghi đè lên dữ liệu cũ.

## Kết quả kiểm tra (Validation)
Tôi đã chạy thử trên các cấu trúc file mẫu bạn gửi (vd: phiếu của `Lý Thị Bình` và `Triệu Thị Hoan`) và cấu trúc của hàm parse đã được thiết kế phù hợp với đặc thù 5 dòng/người của biểu mẫu này.

Nếu trong quá trình sử dụng bạn có gặp file phiếu nào có định dạng bị lệch khiến ứng dụng báo lỗi, hãy cho tôi biết để tôi cập nhật lại logic đọc nhé!
