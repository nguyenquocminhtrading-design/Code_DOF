# Workflow của CalculateRatios.bas

File `CalculateRatios.bas` chứa mã VBA tự động hóa quá trình tính toán các chỉ số tài chính (Financial Ratios) từ một vùng dữ liệu báo cáo tài chính được chọn trên Excel.

## Các bước thực hiện chính (Workflow)

1. **Đọc và Xác thực Dữ liệu**
   - Macro `CalculateRatios_SelectedRange()` kiểm tra vùng chọn của người dùng, đảm bảo vùng chọn có đủ dữ liệu (có tiêu đề, cột nhãn, và ít nhất 2 cột dữ liệu) rồi đưa vào mảng `dataArray`.
   - Tương tự như các module khác, nó tìm cột dữ liệu đầu tiên bằng cách bỏ qua các cột mã số (Code column) hoặc các cột không chứa dữ liệu.

2. **Khai báo Bộ Chỉ Tiêu Đầu Vào (Line Items)**
   - Một Dictionary được tạo ra chứa các chỉ tiêu tài chính cần thiết (NetSales, COGS, GrossProfit, TotalAssets, v.v.).
   - Với mỗi chỉ tiêu, khai báo sẵn một mảng các từ khóa tìm kiếm (Ví dụ: "Net revenue", "Doanh thu thuần") và mảng từ khóa loại trừ (Ví dụ: "Deferred", "Chờ phân bổ") để nhận diện chính xác dòng chứa dữ liệu.

3. **Tìm Kiếm và Ánh Xạ Dòng (Mapping)**
   - Macro dùng hàm `FindRowByKeywords()` để lướt qua cột nhãn (cột 1 của vùng dữ liệu) và tìm số dòng tương ứng cho từng chỉ tiêu đã khai báo.
   - Xử lý các logic đặc biệt như xác định Total Debt (nếu có Short-term Debt và Long-term Debt thì cộng lại, nếu không thì dùng Total Liabilities).

4. **Xác Nhận từ Người Dùng**
   - Macro hiển thị danh sách các chỉ tiêu đã tìm thấy (hoặc không tìm thấy) kèm số dòng.
   - Người dùng có thể bấm Yes để tiếp tục hoặc No để chỉnh sửa/nhập thủ công số dòng cho những chỉ tiêu mà macro nhận diện chưa chính xác.

5. **Khai Báo Bộ Chỉ Số (Ratios)**
   - Khai báo một loạt các chỉ số tài chính cần tính chia theo các nhóm:
     - **Khả năng sinh lời (Profitability):** Biên lợi nhuận gộp, ROA, ROE, ROCE...
     - **Hiệu quả hoạt động (Operation):** Vòng quay hàng tồn kho, Kỳ thu tiền bình quân...
     - **Thanh khoản (Liquidity):** Tỷ số thanh toán hiện hành, Tỷ số thanh toán nhanh...
     - **Đòn bẩy tài chính (Solvency):** Debt/Equity...

6. **Tính Toán Chỉ Số**
   - Macro duyệt qua từng cột dữ liệu, dùng hàm `GetValue` lấy giá trị của các chỉ tiêu đã ánh xạ.
   - Hàm `SafeDiv` được sử dụng để chia, tránh lỗi chia cho 0 (`#DIV/0!`).
   - Tính toán toàn bộ các công thức cho từng chỉ số theo từng kỳ (cột).

7. **Xuất Kết Quả (Output)**
   - Yêu cầu người dùng chọn một ô đích (Destination) để dán bảng kết quả tỷ số.
   - Macro dán kết quả và tự động định dạng số liệu: các chỉ số về số ngày (days) hoặc tỷ lệ (ratio, Current, Acid) sẽ được định dạng số thập phân `0.00`, còn lại định dạng phần trăm `0.00%`.
