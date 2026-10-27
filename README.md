# K26CNTT2_IT108_Session10_Ex09
Bài tập số 9
# Bài tập 9: Tự Động Tra Cứu Phòng Và Thù Lao Báo Cáo Viên Bằng VLOOKUP

## Phần 1. Phân Tích Bài Toán (I/O) & Chiến Lược Xử Lý Lỗi

### 1. Bảng xác định tham số hàm VLOOKUP

| Cột cần điền | Lookup_value (Giá trị tìm) | Table_array (Vùng tra cứu) | Col_index_num | Range_lookup |
| :--- | :--- | :--- | :--- | :--- |
| **Cột D: Tên Tác Giả** | `TRIM(C2)` | `DanhMucTacGia!A$2:E$81` | `2` | `FALSE` |
| **Cột E: Phòng Trình Bày** | `TRIM(C2)` | `DanhMucTacGia!A$2:E$81` | `4` | `FALSE` |
| **Cột F: Thù Lao (VNĐ)** | `TRIM(C2)` | `DanhMucTacGia!A$2:E$81` | `5` | `FALSE` |

---

### 2. Đề xuất giải pháp xử lý 3 bẫy dữ liệu

- **Xử lý Bẫy 1 (Khoảng trắng thừa):** Sử dụng hàm `TRIM(C2)` để bọc lấy `Lookup_value`. Hàm này tự động cắt bỏ khoảng trắng thừa (ví dụ chuyển `'TG001 '` thành `'TG001'`), giúp `VLOOKUP` khớp chính xác dữ liệu.
- **Xử lý Bẫy 2 (Mã tác giả ngoài trường chưa có trong danh mục):** Bọc toàn bộ hàm `VLOOKUP` bên trong hàm `IFERROR(..., "Chưa cập nhật")`. Khi không tìm thấy mã tác giả, thay vì trả về lỗi `#N/A`, hệ thống sẽ hiển thị chuỗi chữ `"Chưa cập nhật"`.
- **Xử lý Bẫy 3 (Trôi vùng tra cứu khi kéo công thức):** Sử dụng địa chỉ tuyệt đối (có dấu `$`) cho `Table_array`: `DanhMucTacGia!A$2:E$81`. Khi sao chép công thức từ ô D2 xuống D81, vùng tra cứu giữ cố định tuyệt đối không bị dịch chuyển.

---

### 3. Thiết kế các bước thực hiện trong Excel

1. **Bước 1:** Đảm bảo cả 2 sheet `DanhSachBaoCao` và `DanhMucTacGia` nằm trong cùng một tệp workbook Excel.
2. **Bước 2:** Nhập công thức kết hợp `IFERROR` + `VLOOKUP` + `TRIM` tại ô D2 của sheet `DanhSachBaoCao`.
3. **Bước 3:** Sao chép công thức sang ô E2 và F2, sau đó chỉ cần sửa chỉ số cột `Col_index_num` tương ứng (từ 2 thành 4 cho Cột E, và thành 5 cho Cột F).
4. **Bước 4:** Bôi đen vùng D2:F2 và nhấp đôi chuột vào góc dưới bên phải (Fill Handle) để Fill toàn bộ công thức xuống dòng 81.

---

## Phần 2. Triển Khai Công Thức & Giải Thích Kỹ Thuật

### 1. Công thức hoàn chỉnh chống lỗi cho dòng 2

- **Công thức tại ô D2 (Tên Tác Giả):**
  =IFERROR(VLOOKUP(TRIM(C2), DanhMucTacGia!A$2:E$81, 2, FALSE), "Chưa cập nhật")
- **Công thức tại ô E2 (Phòng trình bày):**
  =IFERROR(VLOOKUP(TRIM(C2), DanhMucTacGia!A$2:E$81, 4, FALSE), "Chưa cập nhật")
- **Công thức tại ô F2 (Thù lao):**
  =IFERROR(VLOOKUP(TRIM(C2), DanhMucTacGia!A$2:E$81, 5, FALSE), "Chưa cập nhật")

2. Giải thích hiện tượng trôi vùng tra cứu khi bỏ dấu $
Nếu bỏ dấu $ trong Table_array (viết dạng địa chỉ tương đối DanhMucTacGia!A2:E81):

Tại ô D2: Vùng tra cứu là DanhMucTacGia!A2:E81 (đúng).

Khi kéo xuống ô D3: Vùng tra cứu tự động trôi thành DanhMucTacGia!A3:E82 (bỏ sót dòng A2).

Khi kéo xuống ô D5: Vùng tra cứu bị trôi thành DanhMucTacGia!A5:E84.

Hậu quả: Nếu mã tác giả ở dòng D5 rơi vào dòng A2, A3 hoặc A4 của sheet DanhMucTacGia, hàm VLOOKUP sẽ không tìm thấy và báo lỗi #N/A (hoặc "Chưa cập nhật") dù mã đó thực tế có tồn tại trong bảng.

Phần 3. Danh Sách Kiểm Tra Hoàn Thiện (Task List)
[x] Phân tích I/O và thiết kế công thức VLOOKUP xử lý 3 bẫy dữ liệu trong README.md

[x] Tạo file Excel Danh_Sach_Tra_Cuu_Bao_Cao_Vien.xlsx tích hợp sẵn công thức chống lỗi

[ ] Push file README.md và file Excel lên GitHub Repository K26CNTT2_IT108_Session10_Ex09

[ ] Dán link Repository GitHub lên hệ thống LMS Portal
  
