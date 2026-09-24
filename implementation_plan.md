# Kế hoạch chi tiết đếm số liệu cho Biểu M2

Biểu mẫu M2 yêu cầu thống kê rất chi tiết theo từng độ tuổi và cấp học cho từng thôn/bản. Để phần mềm chạy chính xác, chúng ta cần thống nhất các quy tắc đếm từ file `Tong.xlsx`. Dưới đây là logic dự kiến (giả sử năm điều tra là 2025):

## Đề xuất logic tính toán (Cần bạn xác nhận)

> [!IMPORTANT]
> Vui lòng xem kỹ các điều kiện đếm dưới đây xem đã đúng với nghiệp vụ của bạn chưa:

### 1. Huy động trẻ 6 tuổi
- **Tổng số (Cột 3):** Là những trẻ sinh năm 2019 (2025 - 6 = 2019).
- **Vào lớp 1 (Cột 4):** Là trẻ sinh năm 2019 **VÀ** có cột "Khối học" là `1`.
- **Tỷ lệ % (Cột 5):** Phần mềm sẽ tự động chia (Cột 4 / Cột 3) * 100.

### 2. Trẻ độ tuổi 11 - 14 (Sinh năm 2011 - 2014)
- **Tổng số (Cột 10):** Tổng số trẻ sinh trong khoảng 2011 - 2014.
- **Hoàn thành CT Tiểu học (Cột 11):** Trẻ 11-14 tuổi **VÀ** (cột Khối học >= 6, HOẶC đã điền hoàn thành cấp Tiểu học ở cột "Bậc tốt nghiệp").
- **Tỷ lệ % (Cột 12):** (Cột 11 / Cột 10) * 100.

### 3. Đối tượng 15 - 18 tuổi (Sinh năm 2007 - 2010)
- **Tổng số (Cột 17):** Tổng số người sinh trong khoảng 2007 - 2010.
- **Có bằng tốt nghiệp THCS (Cột 18/19/20):** Những người 15-18 tuổi **VÀ** có điền "THCS" ở cột Bậc tốt nghiệp. (Hiện tại phần mềm sẽ gộp chung vào Phổ thông - Cột 18, trừ khi có cách phân biệt GDTX).
- **Tỷ lệ % (Cột 21):** (Cột 20 / Cột 17) * 100.

### 4. Học sinh lớp 9 và HS Tốt nghiệp Tiểu học năm qua
- **HS Lớp 9 năm qua (Cột 13, 14, 15):** Cần đếm những người vừa tốt nghiệp THCS. (Lọc theo Bậc Tốt nghiệp = THCS và Năm tốt nghiệp = 2025?). 
- **HS Tốt nghiệp Tiểu học năm qua vào lớp 6 (Cột 6, 7):** Lọc theo Bậc Tốt nghiệp = TH và vào lớp 6.

## Open Questions

> [!WARNING]
> 1. Trong file `Tong.xlsx`, bạn ghi nhận học sinh Tốt nghiệp THCS và Tiểu học như thế nào? Có phải ghi vào cột **"Bậc TN"** (Cột 20) và cột **"Năm tốt nghiệp"** không?
> 2. Có cần phân biệt Phổ thông (PT) và Giáo dục thường xuyên (GDTX) không? Nếu có thì căn cứ vào cột nào?
> 3. Năm điều tra mặc định là năm 2025, có đúng không? (Nếu không phải, trẻ 6 tuổi sẽ không phải sinh năm 2019).

Nếu bạn đồng ý với các logic đếm này, hãy nhấn **Proceed / Phê duyệt**, tôi sẽ code ngay lập tức các cột của M2 theo đúng công thức này!
