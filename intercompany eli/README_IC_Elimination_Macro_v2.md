# Workflow của IC_Elimination_Macro_v2.bas

File `IC_Elimination_Macro_v2.bas` là phiên bản nâng cấp của macro Loại trừ giao dịch nội bộ (Intercompany Elimination). Phiên bản này được cải tiến đặc biệt để khắc phục triệt để lỗi mã hóa (codepage) thường gặp khi so sánh chuỗi có dấu tiếng Việt trong môi trường VBA.

## Các bước thực hiện chính (Workflow)

1. **Thay Đổi Cốt Lõi (Core Improvements)**
   - Toàn bộ danh sách `GetRelatedParties()` được lưu trữ hoàn toàn bằng mã ASCII không dấu (VD: "CONG TY CO PHAN THUC PHAM DAN ON") thay vì giữ các ký tự có dấu, tránh việc dấu tiếng Việt bị hỏng (corrupt thành ký tự "?") trong mã nguồn VBA.
   - Bổ sung hàm `NormalizeName()`: Khi đọc tên khách hàng từ bảng tính Excel tại thời điểm chạy (runtime), hàm này sẽ biến đổi toàn bộ chữ cái thành in hoa (uppercase), loại bỏ tất cả dấu tiếng Việt bằng các mã Unicode chuẩn (ChrW), và gộp các khoảng trắng/dấu câu lại với nhau. Điều này giúp việc so sánh với danh sách ASCII trở nên chính xác 100%.

2. **Quét và Nhận Diện Dữ Liệu**
   - Macro tiếp tục quét qua các sheet được định nghĩa trước (VD: "DOF", "ADF", "DOAU"...).
   - Tự động nhận diện cột Tên, Doanh thu (hỗ trợ cộng dồn nếu có cả cột Net Sales và Amount), Giá vốn (hoặc tính dựa trên Lợi nhuận gộp GP$). Các từ khóa tìm kiếm cột cũng được viết không dấu hoặc tiếng Anh.

3. **Phân Tích và Loại Trừ Kép (Dual-pass Elimination Logic)**
   - **Lọc chính xác (Exact Match):** Nếu tên khách hàng (đã qua NormalizeName) khớp hoàn toàn với danh sách nội bộ, nó sẽ tự động bị loại trừ (cộng dồn vào phần Eliminated).
   - **Lọc cảnh báo (Keyword Review):** Macro có thêm danh sách `GetReviewKeywords()` chứa các cụm từ nghi ngờ (như "DAN ON", "DAN D FOODS", "DOF"). Nếu tên khách hàng không khớp hoàn toàn nhưng có chứa các cụm từ này, nó sẽ KHÔNG tự động loại trừ nhưng được ghi nhận lại vào một danh sách đặc biệt để chờ con người kiểm tra (Review).

4. **Xuất Báo Cáo Cải Tiến**
   - Macro tạo sheet `IC_Elimination` bao gồm 2 phần chính:
     - **Phần 1:** Bảng tổng hợp trước/sau loại trừ và bảng chi tiết theo từng đối tác đã được loại trừ (giống phiên bản v1).
     - **Phần 2 (Mới - REVIEW):** Nằm dưới cùng, liệt kê danh sách các khách hàng nghi ngờ là đối tác nội bộ nhưng chưa khớp danh sách chính thức. Cung cấp chi tiết tên sheet, tên khách hàng trong data, và số tiền giao dịch để người dùng tự xem xét và cập nhật lại mảng dữ liệu cấu hình nếu cần.
