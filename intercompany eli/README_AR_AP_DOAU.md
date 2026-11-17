# Workflow của AR_AP_DOAU.bas

File `AR_AP_DOAU.bas` chứa mã VBA với chức năng gộp dữ liệu từ 2 sheet báo cáo khoản phải thu (AR) và khoản phải trả (AP) thành một bảng duy nhất (Merged_AR_AP). Phiên bản này được viết có tính module và nâng cao (Advanced) hơn bằng cách sử dụng các hằng số (Constants) và hàm phụ (Subroutine) riêng biệt.

## Các bước thực hiện chính (Workflow)

1. **Thiết lập Cấu hình (Configuration)**
   - Khai báo hàng loạt các hằng số (Constants) ở đầu macro `MergeTables_Advanced()` để định nghĩa linh hoạt: tên sheet (`AR`, `AP`), dòng tiêu đề, và vị trí các cột quan trọng (Kind, Entity, Name, Current/Ending Balance) cho từng sheet.
   - Điều này giúp người dùng dễ dàng bảo trì và chỉnh sửa khi cấu trúc file Excel gốc thay đổi mà không cần can thiệp sâu vào logic code.

2. **Kiểm tra và Khởi tạo Bảng Kết Quả**
   - Kiểm tra sự tồn tại của sheet `AR` và `AP`.
   - Tạo một sheet mới tên `Merged_AR_AP` và ghi các tiêu đề chuẩn hóa: `Kind`, `Entity`, `Name`, `Ending Balance`.

3. **Gộp Dữ Liệu Thông Qua Hàm Phụ (Helper Subroutine)**
   - Thay vì lặp code xử lý 2 lần cho 2 sheet, macro gọi hàm `CopyDataFromSheet()` truyền vào các tham số cấu hình tương ứng.
   - **Bên trong `CopyDataFromSheet`:**
     - Tìm dòng dữ liệu cuối cùng của sheet nguồn.
     - Vòng lặp duyệt qua các dòng (bỏ qua dòng trống).
     - Trích xuất dữ liệu dựa trên các chỉ số cột được cấu hình và ghi trực tiếp vào sheet đích.
     - Tự động cộng dồn con trỏ dòng đích (`destRow`) để sẵn sàng cho lần chèn dữ liệu tiếp theo.

4. **Hoàn Tất**
   - Macro gọi `CopyDataFromSheet` lần lượt cho dữ liệu AR rồi tới AP.
   - Căn chỉnh tự động (AutoFit) các cột và hiện thông báo hoàn thành (MsgBox).
