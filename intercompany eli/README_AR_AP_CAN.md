# Workflow của AR_AP_CAN.bas

File `AR_AP_CAN.bas` chứa mã VBA với chức năng gộp dữ liệu từ bảng Khách hàng (Customer - Khoản phải thu AR) và bảng Nhà cung cấp (Vendor - Khoản phải trả AP) vào chung một bảng (Merged) để thuận tiện cho việc đối chiếu, loại trừ nội bộ hoặc phân tích.

## Các bước thực hiện chính (Workflow)

1. **Khai báo và Kiểm tra Môi trường**
   - Macro `MergeTables_Flexible()` bắt đầu bằng việc thiết lập hằng số (Constants) cho dòng tiêu đề của bảng Customer và Vendor (ví dụ: dòng 3). Điều này giúp dễ dàng điều chỉnh nếu định dạng file thay đổi.
   - Script cố gắng gắn biến đối tượng Worksheet cho hai sheet `Customer` và `Vendor`. Nếu thiếu một trong hai sheet, script sẽ thông báo lỗi và thoát.

2. **Tạo Bảng Kết Quả (Merged Sheet)**
   - Một sheet mới được tạo ra ở cuối workbook và được đặt tên là `Merged`.
   - Script tự động viết các tiêu đề cột cho sheet mới bao gồm: `Kind`, `Entity`, `Name`, và `Ending Balance`.

3. **Trích Xuất Dữ Liệu Khách Hàng (Accounts Receivable - AR)**
   - Tính toán dòng cuối cùng chứa dữ liệu của sheet `Customer` (dựa trên cột A).
   - Dùng vòng lặp quét từ dòng ngay dưới dòng tiêu đề (CUST_HEADER_ROW + 1) cho đến dòng cuối.
   - Nếu dòng không trống (cột A có dữ liệu), script sẽ trích xuất 4 trường thông tin: 
     - Kind (từ cột A)
     - Entity (từ cột B)
     - Name (từ cột D)
     - Current / Ending Balance (từ cột G của sheet Customer)
   - Dữ liệu này được ghi trực tiếp vào sheet `Merged`.

4. **Trích Xuất Dữ Liệu Nhà Cung Cấp (Accounts Payable - AP)**
   - Tương tự, tính toán dòng cuối cùng của sheet `Vendor`.
   - Quét từ dòng dưới tiêu đề cho đến dòng cuối.
   - Trích xuất 4 trường thông tin, với một sự khác biệt nhỏ về cấu trúc cột:
     - Kind (cột A)
     - Entity (cột B)
     - Name (cột D)
     - Current / Ending Balance (từ cột F của sheet Vendor - chú ý cột Ending Balance khác vị trí so với sheet Customer).
   - Dữ liệu được ghi nối tiếp xuống dưới phần dữ liệu AR trong sheet `Merged`.

5. **Hoàn Tất và Định Dạng**
   - Căn chỉnh tự động (AutoFit) độ rộng của các cột từ A đến D trên sheet `Merged` để hiển thị đẹp mắt.
   - Hiển thị hộp thoại (MsgBox) thông báo hoàn tất quá trình gộp dữ liệu.
