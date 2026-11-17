# Workflow của Module3.bas

File `Module3.bas` là một phiên bản script VBA khác thực hiện tính toán báo cáo quy mô chung (Common Size) cho vùng dữ liệu được chọn.

## Các bước thực hiện chính (Workflow)

1. **Nhận diện Vùng Dữ Liệu và Dòng Cơ Sở**
   - Người dùng chọn một vùng dữ liệu chứa các chỉ tiêu tài chính.
   - Script yêu cầu người dùng chọn kiểu báo cáo: 
     - **1**: Báo cáo Kết quả hoạt động kinh doanh (Dùng Doanh thu - Revenue làm cơ sở).
     - **2**: Bảng Cân đối kế toán (Dùng Tổng tài sản - Total Assets làm cơ sở).
   - Tương tự như `autocommonsize.bas`, script tìm kiếm dòng cơ sở thông qua mảng từ khóa chính (keywords) và từ khóa loại trừ (exclusions).

2. **Xác nhận hoặc Nhập thủ công**
   - Nếu không tìm thấy dòng cơ sở, người dùng được yêu cầu nhập tay một từ khóa để script tìm lại.
   - Khi tìm thấy, hiển thị thông báo xác nhận dòng cơ sở đã chọn.

3. **Tính toán Tỷ trọng (Common Size Calculation)**
   - Lặp qua từng cột dữ liệu, lấy giá trị của dòng cơ sở.
   - Nếu giá trị cơ sở hợp lệ (khác 0), tất cả các ô chứa dữ liệu số ở các dòng khác trong cùng cột sẽ được chia cho giá trị cơ sở đó.
   - Các giá trị không phải số sẽ được giữ nguyên.

4. **Trình Bày Dữ Liệu Đầu Ra**
   - Yêu cầu người dùng chọn ô đích để dán kết quả.
   - Dán dữ liệu tính toán và định dạng các vùng số dưới dạng phần trăm (`0.00%`).
   - Hiển thị hộp thoại hoàn thành thông báo kết quả.
