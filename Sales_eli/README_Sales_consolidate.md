# Workflow của Sales_consolidate.bas

File `Sales_consolidate.bas` chứa mã VBA với mục đích tự động thu thập và gộp (Consolidate) dữ liệu bán hàng từ nhiều sheet nguồn khác nhau (như DOF, ADF, DOAU, CAN) vào chung một sheet kết quả chuẩn hóa. Điểm mạnh của macro này là thuật toán ánh xạ cột thông minh và tối ưu hóa tốc độ đọc/ghi dữ liệu.

## Các bước thực hiện chính (Workflow)

1. **Khởi tạo và Tối ưu Hóa**
   - Macro `RunConsolidation()` bắt đầu bằng việc tắt các tính năng làm chậm Excel (ScreenUpdating, EnableEvents, Automatic Calculation) để tăng tốc độ xử lý khi làm việc với lượng dữ liệu lớn.
   - Tạo mới (hoặc làm trống) 2 sheet: `Consolidated` (chứa dữ liệu gộp) và `Mapping Log` (chứa nhật ký ánh xạ cột).

2. **Định Nghĩa Các Trường Chuẩn (Canonical Fields)**
   - Hàm `CanonicalFields()` định nghĩa cấu trúc chuẩn cho bảng kết quả (VD: Entity, Name, Item No, Net Sales...).
   - Hàm `SynonymMap()` liệt kê các tên gọi đồng nghĩa thường gặp ở các sheet nguồn để dễ dàng nhận diện (VD: "Net Sales" cũng có thể là "Amount", "Item No" có thể là "Item No.").

3. **Ánh Xạ Cột Thông Minh (Smart Column Mapping)**
   - Macro gọi hàm `ProcessSheet()` để xử lý từng sheet nguồn.
   - Tìm dòng tiêu đề (header) thông qua hàm `FindHeaderRow()`.
   - Hàm `MapColumn()` lướt qua tiêu đề cột để ghép nối với các trường chuẩn:
     - **Bước 1:** Tìm kiếm chính xác (Exact Match) hoặc dựa trên danh sách từ đồng nghĩa.
     - **Bước 2 (Dự phòng):** Nếu không tìm thấy, sử dụng thuật toán **Fuzzy Matching** (đo khoảng cách chuỗi Levenshtein) để tìm cột có tên gần giống nhất.
   - Ghi lại quá trình ánh xạ, loại khớp (Exact/Synonym/Fuzzy/None) và mức độ tự tin (Confidence %) vào sheet `Mapping Log`.

4. **Xử Lý Dữ Liệu Tốc Độ Cao (Bulk Read/Write)**
   - Thay vì lặp qua từng ô (cell) trên bảng tính rất chậm, macro đọc TOÀN BỘ vùng dữ liệu vào bộ nhớ máy tính thông qua mảng (`srcArr`).
   - Xử lý, trích xuất dữ liệu của các cột đã được ánh xạ, và bỏ qua các dòng trống hoàn toàn. Đưa dữ liệu hợp lệ vào mảng kết quả (`outArr`).
   - Đổ toàn bộ mảng kết quả (`outArr`) xuống sheet `Consolidated` trong **MỘT LẦN GHI DUY NHẤT (Bulk Write)**. 

5. **Hoàn Tất**
   - Bật lại các tính năng của Excel.
   - Căn chỉnh độ rộng cột và hiển thị thông báo tổng kết (MsgBox) bao gồm số dòng đã gộp và thời gian thực thi.
