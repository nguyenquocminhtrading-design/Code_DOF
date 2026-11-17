# Workflow của Create_VDR_Folders (2).bas

File `Create_VDR_Folders (2).bas` chứa script VBA dùng để tự động tạo cấu trúc cây thư mục Virtual Data Room (VDR) trên hệ điều hành Windows dựa trên danh sách cấu trúc có trong Excel.

## Các bước thực hiện chính (Workflow)

1. **Khởi tạo và Kiểm tra**
   - Macro `Create_VDR_Folder_Structure()` tìm kiếm một sheet có tên là `VDR Tracker` trong Workbook hiện tại.
   - Nếu không tìm thấy, script sẽ dừng lại và báo lỗi.

2. **Chọn Thư Mục Gốc (Root Directory)**
   - Hiển thị hộp thoại chọn thư mục (Folder Picker Dialog) để người dùng chọn đường dẫn gốc lưu trữ cây thư mục VDR. Nếu hủy, script sẽ thoát.

3. **Phân Loại Cấp Độ Thư Mục (Level Classification)**
   - Quét qua từng dòng dữ liệu từ dòng 7 cho đến dòng cuối cùng.
   - Sử dụng hàm `ClassifyRow()` để phân loại cấp độ thư mục dựa vào sự tồn tại của dữ liệu trong 3 cột A (Section), B (Subsection) và C (Item):
     - **Cấp 1 (Section):** Có giá trị ở cột A, cột B và C rỗng.
     - **Cấp 2 (Subsection):** Có giá trị ở cột A và B, cột C rỗng.
     - **Cấp 3 (Item):** Có giá trị ở cả cột A, B và C.
     - **Không xác định:** Không theo quy luật trên.

4. **Tạo Tên và Định Dạng Mã Thư Mục**
   - Mã số thư mục (đánh số thứ tự) được định dạng thành 2 chữ số (VD: 01, 01.01, 01.01.01) thông qua hàm `Format()` để đảm bảo File Explorer của Windows sắp xếp đúng thứ tự.
   - Tên thư mục được lấy từ cột I (Topic) cho cấp 1 & 2, và cột J (Key Documents) cho cấp 3.
   - Hàm `SanitizeName()` được gọi để làm sạch tên thư mục: loại bỏ các ký tự đặc biệt không được phép trong Windows (`\/:*?"<>|`), loại bỏ dấu chấm/khoảng trắng ở cuối, và giới hạn độ dài tên không quá 100 ký tự.

5. **Tạo Cây Thư Mục**
   - Dựa vào cấp độ, đường dẫn thư mục được nối dài (path1, path2, path3).
   - Hàm `EnsureFolder()` kiểm tra sự tồn tại của thư mục bằng hàm `Dir`. Nếu chưa có, lệnh `MkDir` sẽ được thực thi để tạo thư mục mới.

6. **Hoàn Tất và Báo Cáo**
   - Các dòng dữ liệu không hợp lệ (nhưng không phải dòng trống hoàn toàn) sẽ bị bỏ qua và ghi nhận số dòng.
   - Cuối cùng, hiển thị thông báo (MsgBox) thống kê số lượng thư mục đã tạo thành công và cảnh báo các dòng bị bỏ qua (nếu có).
