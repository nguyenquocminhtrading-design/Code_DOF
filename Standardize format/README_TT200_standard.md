# Workflow của TT200_standard.bas

File `TT200_standard.bas` chứa mã VBA để tạo các báo cáo tài chính chuẩn hóa (Bảng Cân đối kế toán và Báo cáo Kết quả hoạt động kinh doanh) theo một mẫu (template) tiếng Anh thống nhất. File này rất giống và dường như là một phiên bản cùng họ với `Module2.bas`.

## Các bước thực hiện chính (Workflow)

1. **Cấu Trúc và Cấu Hình (Configuration)**
   - Khai báo kiểu dữ liệu `EntityConfig` và `TplRow` để lưu thông tin cấu hình cho từng pháp nhân và thông tin cho từng dòng của báo cáo chuẩn.
   - Hàm `GetEntityConfigs()` định nghĩa thông số chính xác cho 8 pháp nhân (001-DOF, 002-ADF, ...), chỉ định tên sheet gốc chứa dữ liệu, tọa độ cột mã, cột dữ liệu tháng đầu tiên, tỷ giá, và giới hạn các dòng BS, PL.

2. **Template Báo Cáo Chuẩn (Standard Template)**
   - Hàm `GetTemplateLines()` chứa nội dung khai báo các dòng báo cáo (Bảng cân đối kế toán & Báo cáo kết quả hoạt động kinh doanh). Mỗi dòng được quy định: số thứ tự, tên chỉ tiêu tiếng Anh, mã số, loại dòng (D - chi tiết lấy dữ liệu, R - tính tổng, E - công thức), và vị trí thuộc phần BS hay PL.
   - Dữ liệu này được parse thành các đối tượng `TplRow` để dễ dàng duyệt qua.

3. **Xử Lý Từng Pháp Nhân (Entity Processing)**
   - Hàm `BuildAllEntitySheets()` lặp qua toàn bộ các pháp nhân được cấu hình.
   - Đối với mỗi pháp nhân, nếu sheet dữ liệu của nó tồn tại, macro gọi `BuildEntitySheet()`.
   - **Bên trong BuildEntitySheet:**
     - Xóa sheet chuẩn hóa cũ (nếu có) và tạo sheet mới với tên là mã của Entity.
     - Lấy dòng ngày tháng (header) từ sheet dữ liệu gốc để đưa vào báo cáo chuẩn.
     - In template (tên chỉ tiêu, mã số) ra sheet.
     - Dùng vòng lặp 12 tháng, kết hợp với hàm `BuildFormula()` để tự động sinh ra các công thức `SUMIFS` tương ứng, giúp lấy chính xác dữ liệu từ sheet gốc sang sheet chuẩn theo đúng mã số, hoặc sinh các công thức tính tổng (`SUM`) đối với các dòng phân nhóm.

4. **Trình Bày**
   - Định dạng in đậm các dòng tổng, tô màu nền và chữ trắng cho các dòng tiêu đề (tháng).
   - Có cột Check ở dòng 130 để tự động kiểm tra tính cân đối của báo cáo tài chính (Tổng tài sản - Tổng nguồn vốn = 0).
