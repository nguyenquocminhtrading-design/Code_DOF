# Workflow của autogrwoth.bas

File `autogrwoth.bas` chứa mã VBA với chức năng tự động tính toán Tốc độ tăng trưởng (Growth rate) giữa các kỳ (ví dụ: các tháng, các quý, các năm) cho một vùng dữ liệu được chọn trên Excel.

## Các bước thực hiện chính (Workflow)

1. **Xác thực và Đọc dữ liệu (Validate & read data)**
   - Macro `CalculateGrowth_SelectedRange()` yêu cầu người dùng chọn một vùng dữ liệu cần phân tích.
   - Vùng dữ liệu phải chứa dòng tiêu đề (header), cột nhãn (chỉ tiêu) và ít nhất 2 cột dữ liệu để có thể so sánh. Dữ liệu được đưa vào mảng `dataArray` để xử lý.

2. **Xác định Cột Dữ Liệu Đầu Tiên**
   - Phân tích dòng tiêu đề để tìm cột bắt đầu chứa dữ liệu thời gian (ví dụ tìm dấu `/` cho ngày tháng hoặc "Quý").
   - Bỏ qua cột mã số (nếu có) thông qua việc tìm các từ khóa như "Mã", "Code", "STT" ở dòng tiêu đề cột thứ 2.

3. **Lọc Dòng Dữ Liệu Hợp Lệ**
   - Macro quét qua các dòng (bỏ qua header) và chỉ ghi nhận các dòng mà ô ở cột dữ liệu đầu tiên chứa số (`IsNumeric`). Các dòng này được đưa vào tập hợp `validRows` để tiến hành tính toán.

4. **Tính Toán Tốc Độ Tăng Trưởng (Growth Calculation)**
   - Vòng lặp tính toán quét qua từng cột dữ liệu của các dòng hợp lệ.
   - Cột dữ liệu đầu tiên mặc định có tốc độ tăng trưởng là `0` (vì không có kỳ trước để so sánh).
   - Từ cột thứ 2 trở đi, tính toán theo công thức: `(Giá trị kỳ này - Giá trị kỳ trước) / Giá trị kỳ trước`.
   - Nếu giá trị kỳ trước bằng 0 hoặc không phải là số, tăng trưởng được gán là `0`.

5. **Xuất Dữ Liệu (Output)**
   - Hiển thị hộp thoại yêu cầu người dùng chọn ô đích (top-left cell) để dán kết quả tính toán.
   - Kết quả được dán xuống vùng đích và các cột dữ liệu được tự động định dạng hiển thị dưới dạng phần trăm (`0.00%`).
   - Thông báo hoàn tất (MsgBox) sẽ hiện ra cùng với số dòng dữ liệu đã được xử lý.
