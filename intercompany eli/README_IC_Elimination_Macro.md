# Workflow của IC_Elimination_Macro.bas

File `IC_Elimination_Macro.bas` chứa mã VBA tự động hóa quá trình Loại trừ giao dịch nội bộ (Intercompany Elimination) giữa các pháp nhân (entities) trong tập đoàn DDF Group.

## Các bước thực hiện chính (Workflow)

1. **Khởi tạo và Định nghĩa Cấu trúc**
   - Macro `EliminateIntercompany()` sử dụng các hàm phụ trợ để lấy các mảng cấu hình:
     - `GetSheetKeywords()`: Xác định các sheet cần xử lý (VD: "DOF", "ADF", "CAN"...)
     - `GetRelatedParties()`: Danh sách tên đầy đủ của các công ty con/chi nhánh nội bộ trong tập đoàn.
     - Các ứng viên tên cột (Header Candidates) cho cột Tên Khách Hàng, Doanh Thu, Giá Vốn, và Lợi Nhuận Gộp.

2. **Quét và Nhận diện Tự động**
   - Duyệt qua toàn bộ các sheet trong workbook. Nếu tên sheet khớp với danh sách `GetSheetKeywords()`, macro sẽ đưa sheet đó vào quá trình xử lý (gọi Sub `ProcessSheet`).
   - Hàm `FindHeaderCell()` và `FindColInRow()` tự động tìm kiếm vị trí các cột dữ liệu dựa trên danh sách từ khóa ứng viên, giúp code linh hoạt hoạt động ngay cả khi người dùng đổi tên cột (miễn là vẫn nằm trong nhóm từ khóa cho phép).
   - Nếu không tìm thấy cột Giá Vốn (COGS) nhưng tìm thấy Lợi nhuận gộp (GP$), macro sẽ tự động tính Giá Vốn = Doanh Thu - GP$.

3. **Phân tích và Loại trừ (Elimination Logic)**
   - Vòng lặp duyệt qua từng dòng giao dịch của pháp nhân.
   - Cộng dồn Doanh thu và Giá vốn vào Tổng ban đầu.
   - Hàm `IsRelatedParty()` kiểm tra xem tên khách hàng có nằm trong danh sách các pháp nhân nội bộ hay không (có loại bỏ khoảng trắng và không phân biệt hoa/thường).
   - Nếu phát hiện giao dịch nội bộ, Doanh thu và Giá vốn của giao dịch đó sẽ được cộng dồn vào phần "Bị loại trừ" (Eliminated) theo từng đối tác.

4. **Trình bày Báo Cáo Kết Quả**
   - Macro tạo ra một sheet mới tên `IC_Elimination`.
   - **Bảng Tổng Hợp (Summary Table):** Hiển thị tổng quan mỗi pháp nhân có:
     - Tổng doanh thu / giá vốn ban đầu.
     - Số lượng doanh thu / giá vốn bị loại trừ.
     - Doanh thu / giá vốn thuần sau khi loại trừ.
   - **Bảng Chi Tiết (Detailed Breakdown):** Nằm dưới bảng tổng hợp, liệt kê chi tiết từng pháp nhân nội bộ (đối tác) đã đóng góp bao nhiêu vào con số bị loại trừ trong từng sheet, kèm số dòng giao dịch cụ thể.
   - Định dạng in đậm, tô màu nền chuyên nghiệp và tự động tính dòng "TỔNG CỘNG".
