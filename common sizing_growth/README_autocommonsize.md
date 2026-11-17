# Workflow của autocommonsize.bas

File `autocommonsize.bas` chứa mã VBA (macro) nhằm mục đích tự động tính toán báo cáo Common Size (Báo cáo quy mô chung) cho một vùng dữ liệu được chọn trong Excel.

## Các bước thực hiện chính (Workflow)

1. **Xác thực và Đọc dữ liệu (Validate & read data)**
   - Macro `AutoCommonSize_SelectedRange()` trước tiên kiểm tra xem người dùng đã chọn một vùng dữ liệu hợp lệ chưa (yêu cầu có dòng header, cột nhãn, và ít nhất 2 cột dữ liệu).
   - Vùng dữ liệu được lưu vào một mảng `dataArray` để xử lý trên bộ nhớ, giúp thao tác nhanh hơn.

2. **Xác định Cột Dữ Liệu Đầu Tiên**
   - Macro phân tích dòng header để tìm cột bắt đầu chứa dữ liệu bằng cách bỏ qua cột mã số (code column). Nó nhận diện các cột mã số dựa trên từ khóa như "Mã", "Code", "STT" hoặc định dạng số liệu, đảm bảo phân biệt giữa các cột mô tả và cột dữ liệu tài chính.

3. **Chọn Loại Báo Cáo (Report Type)**
   - Hiển thị bảng hỏi (InputBox) yêu cầu người dùng chọn cơ sở tính toán Common Size:
     - Báo cáo kết quả hoạt động kinh doanh (Dựa trên Doanh thu).
     - Bảng cân đối kế toán (Dựa trên Tổng tài sản).
     - Chế độ tự động (Tự động tách phần Balance Sheet dùng Tổng tài sản, phần P&L dùng Doanh thu).

4. **Tìm Dòng Cơ Sở (Base Row)**
   - Macro tìm kiếm dòng chứa các từ khóa cơ sở để chia:
     - **Tổng tài sản (Total Assets):** Tìm trong mảng bằng hàm `FindBaseRowInRange()` với các từ khóa "total assets", "tổng tài sản" (có loại trừ "ngắn hạn", "dài hạn"...).
     - **Doanh thu (Revenue):** Tìm các từ khóa "revenue", "doanh thu thuần"...
   - Nếu ở chế độ "Tự động", nó sẽ xác định vị trí bắt đầu của Báo cáo KQHĐKD để chia vùng dữ liệu thành 2 phần và tìm dòng cơ sở tương ứng cho từng phần.
   - Nếu không tìm thấy tự động, Macro cho phép người dùng nhập tay từ khóa hoặc số dòng.

5. **Tính Toán Common Size**
   - Macro lặp qua từng ô chứa số trong vùng dữ liệu (của mảng `outputArray`) và chia giá trị đó cho giá trị cơ sở (base value) tương ứng ở cùng một cột.
   - Các dòng chữ (text) hoặc rỗng sẽ được giữ nguyên.

6. **Xuất Dữ Liệu (Output)**
   - Người dùng được yêu cầu chọn một ô đích (Destination) để dán kết quả.
   - Macro sẽ ghi mảng dữ liệu đã tính toán ra vùng đích tương ứng và tự động định dạng các ô là định dạng phần trăm (`0.00%`).
   - Hiển thị thông báo (MsgBox) hoàn thành, báo cáo rõ dòng cơ sở nào đã được sử dụng để người dùng kiểm chứng.

## Hàm Tiện ích (Helpers)
- `FindBaseRowInRange`: Hàm linh hoạt để tìm số dòng cơ sở theo từ khóa (có bao gồm mảng từ khóa loại trừ).
- `StripAccents`: Hàm loại bỏ dấu tiếng Việt để so sánh chuỗi chính xác hơn khi tìm từ khóa.
