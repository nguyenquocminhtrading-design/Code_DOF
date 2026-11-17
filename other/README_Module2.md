# Workflow của Module2.bas

File `Module2.bas` chứa mã VBA với mục đích tự động hóa việc trích xuất, chuẩn hóa và hợp nhất báo cáo tài chính (Bảng cân đối kế toán - Balance Sheet và Báo cáo kết quả hoạt động kinh doanh - P&L) cho nhiều công ty thành viên (entities).

## Các bước thực hiện chính (Workflow)

1. **Định nghĩa Cấu trúc Dữ liệu (Data Structures)**
   - `EntityConfig`: Cấu trúc lưu trữ cấu hình riêng cho từng entity (như mã entity, tên sheet, vị trí các cột, dòng dữ liệu BS và PL, tỷ giá...).
   - `TplRow`: Cấu trúc lưu trữ thông tin của từng dòng trong template chuẩn hóa (số dòng, tên chỉ tiêu, mã chỉ tiêu, loại công thức, etc.).

2. **Khởi tạo Template Chuẩn hóa**
   - Hàm `GetTemplateLines()` và `ParseTemplate()` tạo ra một khung chuẩn (template) với các chỉ tiêu tài chính bằng tiếng Anh (dựa theo chuẩn mực). Template này bao gồm cả phần Balance Sheet và Profit & Loss, được đánh mã số và loại công thức cụ thể (`D` cho dữ liệu chi tiết, `R` cho tính tổng một mảng, `E` cho biểu thức).

3. **Cấu hình các Entities**
   - Hàm `GetEntityConfigs()` định nghĩa cấu hình cụ thể cho 8 entities khác nhau (VD: `001-DOF`, `002-ADF`, `003-CAN`, v.v.). Mỗi cấu hình xác định chính xác dữ liệu của entity nằm ở đâu trên bảng tính gốc.

4. **Tạo Báo cáo Chuẩn hóa cho Từng Entity**
   - Sub `BuildEntitySheet(cfg)` thực hiện:
     - Tạo một sheet mới cho entity.
     - Viết các tiêu đề (Balance Sheet, P&L) và dòng tiêu đề cột chứa ngày tháng (được tự động lấy từ sheet dữ liệu gốc nhờ hàm `FindHeaderRow()`).
     - Ghi template chuẩn vào sheet mới, và ở các cột dữ liệu (tháng 1 đến tháng 12), chèn công thức (thông qua `BuildFormula()`) để tự động `SUMIFS` dữ liệu từ sheet gốc vào sheet chuẩn, hoặc tính toán tổng/công thức tự động nếu là dòng tính toán.
   - Sub `BuildAllEntitySheets()` lặp qua toàn bộ 8 entities và gọi `BuildEntitySheet` để tạo ra tất cả các sheet chuẩn.

5. **Hợp nhất (Consolidation)**
   - Sub `BuildConsolidatedSheets()` và `BuildConsolSheet()` tổng hợp tất cả các sheet entity đã được chuẩn hóa thành 2 sheet hợp nhất: `Consol - BS` và `Consol - PL`.
   - Với các dòng dữ liệu con (Kind=`D`), công thức sẽ cộng trực tiếp từ các ô tương ứng trên tất cả các sheet entity.
   - Với các dòng tổng/công thức (Kind=`R`, `E`, `L`), hệ thống sẽ tính lại công thức tự động ngay trong sheet Consol để đảm bảo tính chính xác của báo cáo hợp nhất.
