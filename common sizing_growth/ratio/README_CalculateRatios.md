# HƯỚNG DẪN SETUP BẢNG DỮ LIỆU & DANH MỤC CÔNG THỨC TÍNH RATIOS
**File Macro:** `CalculateRatios_Updated (3).bas`  
**Thư mục:** `Code\common sizing_growth\ratio`

---

## PHẦN I: NGUYÊN TẮC SETUP LAYOUT BẢNG DỮ LIỆU ĐẦU VÀO

Để macro tự động nhận diện chính xác 100% cột dữ liệu và dòng chỉ tiêu tài chính mà không bao giờ bị lệch hay báo lỗi, bạn cần tuân theo các nguyên tắc setup layout như sau:

### 1. Quy chuẩn Header & Cột dữ liệu (Columns Setup)
Macro hỗ trợ tự động cả 2 dạng layout bảng phổ biến nhất:

* **Format 1 (Format Đơn nhãn):**
  - **Cột 1:** Tên chỉ tiêu (`Line item (VAS item code)`)
  - **Cột 2 trở đi:** Cột các kỳ dữ liệu (`FY2019`, `FY2020`, ..., `30 Jun 2026`)
* **Format 2 (Format Song ngữ VAS đa cột):**
  - **Cột 1:** `CHỈ TIÊU (Tiếng Việt)`
  - **Cột 2:** `CHỈ TIÊU (Tiếng Anh)`
  - **Cột 3:** `Mã số` (01, 02, 10, 100...)
  - **Cột 4:** `Thuyết minh` (VII.01, VII.02...)
  - **Cột 5 trở đi:** Cột các kỳ dữ liệu ngày tháng (`01/01/2026`, `01/02/2026`...)

> **Lưu ý quan trọng về Header:** Dòng đầu tiên (Row 1) của vùng bạn bôi chọn phải chứa các từ khóa thời gian như: `FY202x`, `202x`, `Q1..Q4`, `Quy`, `Nam`, `dd/mm/yyyy`, `Jun`, `Dec`... Macro sẽ dựa vào đó để tự động bỏ qua các cột Mã số và Thuyết minh.

---

### 2. Quy chuẩn Hàng cho Báo Cáo Hỗn Hợp (Combined Mode)
Khi chạy chế độ **Combined Mode (Lựa chọn 3: Tính cả Bảng cân đối kế toán & Báo cáo kết quả kinh doanh)**:

```
[ DÒNG HEADER: Tiêu đề các kỳ (FY2019, FY2020... hoặc 01/01/2026...) ]
------------------------------------------------------------------------
▲
│   PHẦN NỬA TRÊN: BẢNG CÂN ĐỐI KẾ TOÁN (Balance Sheet)
│   - Bắt đầu: Current Assets (Tài sản ngắn hạn)
│   - Kết thúc: Total Equity (Vốn chủ sở hữu / Tổng nguồn vốn)
▼
------------------------------------------------------------------------
▲   RANH GIỚI PHÂN CÁCH (DÒNG BẮT ĐẦU P&L):
│   Dòng này BẮT BUỘC phải là: "Gross Operating Revenue", "Revenue" hoặc "Net Sales"
▼
------------------------------------------------------------------------
▲
│   PHẦN NỬA DƯỚI: BÁO CÁO KẾT QUẢ KINH DOANH (Income Statement)
│   - COGS, Gross Profit, Chi phí bán hàng, Chi phí QLDN...
│   - Kết thúc: Net Profit After Tax Of Parent Company
▼
```

> **Nguyên tắc vàng:** Tuyệt đối không để xen kẽ các dòng P&L vào giữa phần Tài sản/Nợ. Khi đặt P&L từ dòng `Revenue` trở xuống, macro sẽ tự động kích hoạt tính năng **Phân vùng tìm kiếm độc lập (Search Partitioning)** để:
> - Chỉ tìm các khoản mục Tài sản, Nợ, Vốn CSH ở nửa trên.
> - Chỉ tìm Doanh thu, Chi phí, Lợi nhuận ở nửa dưới.
> - Chấm dứt 100% tình trạng bắt nhầm dòng (ví dụ nhầm Khấu hao lũy kế âm của tài sản cố định vào Chi phí khấu hao của P&L).

---

## PHẦN II: TỪ ĐIỂN MAPPING 19 CHỈ TIÊU CỐT LÕI

Macro sẽ tự động quét đồng thời cả Cột tiếng Việt và Cột tiếng Anh của từng dòng theo danh mục chuẩn dưới đây:

| Mục trong Code | Tên chuẩn | Từ khóa Tiếng Anh nhận diện | Từ khóa Tiếng Việt nhận diện | Luật loại trừ (Exclusions) |
| :--- | :--- | :--- | :--- | :--- |
| **`NetSales`** | Doanh thu | `gross operating revenue`, `revenue`, `net sales`, `net turnover` | `doanh thu ban hang`, `doanh thu thuan`, `tong doanh thu` | `financial`, `other income`, `revenue deductions`, `chua thuc hien` |
| **`COGS`** | Giá vốn hàng bán | `cost of goods sold`, `cost of sales`, `cogs` | `gia von hang ban`, `gia von` | `deferred` |
| **`GrossProfit`** | Lợi nhuận gộp | `gross profit` | `loi nhuan gop` | `margin`, `deferred` |
| **`Depreciation`**| Khấu hao trong kỳ | `depreciation and amortization` *(chỉ tìm ở phần P&L)* | `chi phi khau hao` | `deferred` |
| **`SellingExp`** | Chi phí bán hàng | `selling expenses`, `selling expense` | `chi phi ban hang` | `cost of goods sold` |
| **`GAExp`** | Chi phí QLDN | `general and admin expenses`, `admin expenses` | `chi phi quan ly doanh nghiep` | `deferred` |
| **`FinancialExp`**| Chi phí tài chính | `financial expenses`, `financial expense` | `chi phi tai chinh` | `interest expenses`, `lai vay` |
| **`InterestExp`** | Chi phí lãi vay | `of which: interest expenses`, `interest expenses` | `chi phi lai vay`, `lai vay` | `deferred` |
| **`OperatingProfit`**| LN từ HĐKD | `net profit from operating activities` | `loi nhuan thuan tu hdkd` | `financial`, `deferred` |
| **`ProfitBeforeTax`**| LN trước thuế | `pretax profit`, `profit before tax` | `tong loi nhuan truoc thue` | `deferred` |
| **`NetProfit`** | LN sau thuế | `net profit after tax of parent company`, `net profit after tax` | `loi nhuan sau thue` | `minority`, `operating`, `before tax` |
| **`Cash`** | Tiền & Tương đương | `cash and cash equivalents`, `cash equivalents`, `cash` | `tien va cac khoan tuong duong tien` | `provision`, `du phong` |
| **`ShortTermInvest`**| Đầu tư ngắn hạn | `short-term investments`, `held to maturity investment` | `dau tu tai chinh ngan han` | `provision`, `loss` |
| **`Receivables`** | Phải thu ngắn hạn KH| `short term account receivables`, `account receivable` | `phai thu ngan han cua khach hang` | `other`, `loan`, `doubtful` |
| **`Inventory`** | Hàng tồn kho | `inventories`, `inventory` | `hang ton kho` | `provision`, `du phong` |
| **`CurrentAssets`**| Tài sản ngắn hạn | `current assets`, `short-term assets` | `tai san ngan han` | `other current assets`, `cash`, `khac` |
| **`FixedAssets`** | Tài sản cố định | `tangible fixed assets`, `fixed assets` | `tai san co dinh huu hinh` | `depreciation`, `incomplete` |
| **`TotalAssets`** | Tổng tài sản | `total assets` | `tong cong tai san`, `tong tai san` | `current`, `short-term`, `long-term` |
| **`Payables`** | Phải trả người bán | `payment for suppliers`, `trade payables` | `phai tra nguoi ban ngan han` | `prepaid`, `tax`, `employees` |
| **`CurrentLiabilities`**| Nợ ngắn hạn | `short term liabilities`, `current liabilities` | `no ngan han` | `loan`, `provision`, `bonus`, `tax` |
| **`ShortTermDebt`**| Vay nợ ngắn hạn | `short term loan and liabilities`, `short term loans` | `vay va no thue tai chinh ngan han` | `long-term`, `dai han` |
| **`LongTermDebt`** | Vay nợ dài hạn | `loans and long term liabilities`, `long term loans` | `vay va no thue tai chinh dai han` | `short-term`, `ngan han` |
| **`TotalLiabilities`**| Tổng nợ phải trả | `total liabilities`, `liabilities` *(Mã 300)* | `no phai tra` | `short term`, `long term`, `current`, `loan` |
| **`TotalEquity`** | Vốn chủ sở hữu | `owners equity`, `capital and reserves`, `total equity` | `von chu so huu` | `paid-in`, `shares`, `premium`, `treasury` |

---

## PHẦN III: TOÀN BỘ DANH MỤC CÔNG THỨC TÀI CHÍNH ĐƯỢC CHUẨN HÓA

### 1. Nhóm Khả Năng Sinh Lời (Profitability)
1. **Cost of goods sold (% sales)**:
   $$=\frac{|\text{COGS}|}{\text{NetSales}}$$
   *Công thức Excel:* `=IFERROR(ABS(COGS)/NetSales, "")` *(Bọc ABS tránh số âm)*
2. **Gross profit margin (% sales)**:
   $$=\frac{\text{GrossProfit}}{\text{NetSales}}$$
   *Công thức Excel:* `=IFERROR(GrossProfit/NetSales, "")`
3. **Operating profit margin (%)**:
   $$=\frac{\text{OperatingProfit}}{\text{NetSales}}$$
   *Công thức Excel:* `=IFERROR(OperatingProfit/NetSales, "")`
4. **Depreciation rate (% sales)**:
   $$=\frac{\text{Depreciation}}{\text{NetSales}}$$
   *Công thức Excel:* `=IFERROR(Depreciation/NetSales, "")` *(Để trống nếu P&L không có dòng Khấu hao riêng)*
5. **Selling expense/sales (%)**:
   $$=\frac{\text{SellingExp}}{\text{NetSales}}$$
   *Công thức Excel:* `=IFERROR(SellingExp/NetSales, "")`
6. **Administrative expense/sales (%)**:
   $$=\frac{\text{GAExp}}{\text{NetSales}}$$
   *Công thức Excel:* `=IFERROR(GAExp/NetSales, "")`
7. **Financial expense/sales (%)**:
   $$=\frac{\text{FinancialExp}}{\text{NetSales}}$$
   *Công thức Excel:* `=IFERROR(FinancialExp/NetSales, "")`
8. **Net Profit margin (%)**:
   $$=\frac{\text{NetProfit}}{\text{NetSales}}$$
   *Công thức Excel:* `=IFERROR(NetProfit/NetSales, "")`
9. **Return on total assets - ROA (%)**:
   $$=\frac{\text{NetProfit}}{\text{TotalAssets}}$$
   *Công thức Excel:* `=IFERROR(NetProfit/TotalAssets, "")`
10. **Return on equity - ROE (%)**:
    $$=\frac{\text{NetProfit}}{\text{TotalEquity}}$$
    *Công thức Excel:* `=IFERROR(NetProfit/TotalEquity, "")` *(Bắt đúng dòng Owners Equity)*
11. **Return on capital employed - ROCE (%)**:
    $$=\frac{\text{EBIT}}{\text{TotalAssets} - \text{CurrentLiabilities}}$$
    *Công thức Excel:* `=IFERROR(EBIT/(TotalAssets - CurrentLiabilities), "")`  
    *(EBIT được tính bằng $\text{PretaxProfit} + \text{InterestExpenses}$ theo chuẩn IFRS)*

---

### 2. Nhóm Vòng Quay Ngày (Activity & Days)
12. **Inventory days (Số ngày tồn kho)**:
    $$=\frac{\text{Inventory} \times 365}{|\text{COGS}|}$$
    *Công thức Excel:* `=IFERROR((Inventory*365)/ABS(COGS), "")`
13. **Receivable days (Số ngày thu tiền khách hàng)**:
    $$=\frac{\text{Receivables} \times 365}{\text{NetSales}}$$
    *Công thức Excel:* `=IFERROR((Receivables*365)/NetSales, "")`
14. **Payables days (Số ngày trả nợ người bán)**:
    $$=\frac{\text{Payables} \times 365}{|\text{COGS}|}$$
    *Công thức Excel:* `=IFERROR((Payables*365)/ABS(COGS), "")`

---

### 3. Nhóm Thanh Khoản (Liquidity)
15. **Current ratio (Thanh toán hiện hành)**:
    $$=\frac{\text{CurrentAssets}}{\text{CurrentLiabilities}}$$
    *Công thức Excel:* `=IFERROR(CurrentAssets/CurrentLiabilities, "")`
16. **Acid test ratio (Thanh toán nhanh trừ tồn kho)**:
    $$=\frac{\text{CurrentAssets} - \text{Inventory}}{\text{CurrentLiabilities}}$$
    *Công thức Excel:* `=IFERROR((CurrentAssets - Inventory)/CurrentLiabilities, "")`
17. **Cash ratio (Thanh toán bằng tiền)**:
    $$=\frac{\text{Cash}}{\text{CurrentLiabilities}}$$
    *Công thức Excel:* `=IFERROR(Cash/CurrentLiabilities, "")`
18. **Quick ratio (Thanh toán tức thời)**:
    $$=\frac{\text{Cash} + \text{ShortTermInvest} + \text{Receivables}}{\text{CurrentLiabilities}}$$
    *Công thức Excel:* `=IFERROR((Cash + ShortTermInvest + Receivables)/CurrentLiabilities, "")`

---

### 4. Nhóm Đòn Bẩy & Cơ Cấu Vốn (Solvency & Capital Structure)
19. **Total debt/equity ratio (Hệ số Nợ / Vốn CSH - D/E)**:
    $$=\frac{\text{TotalDebt}}{\text{TotalEquity}}$$
    *Công thức Excel:* `=IFERROR(TotalDebt/TotalEquity, "")`  
    *(TotalDebt tự động cộng Vay ngắn hạn + Vay dài hạn)*
20. **Total equity/debt ratio**:
    $$=\frac{\text{TotalEquity}}{\text{TotalDebt}}$$
    *Công thức Excel:* `=IFERROR(TotalEquity/TotalDebt, "")`
21. **Equity/(Debt+Equity)**:
    $$=\frac{\text{TotalEquity}}{\text{TotalDebt} + \text{TotalEquity}}$$
    *Công thức Excel:* `=IFERROR(TotalEquity/(TotalDebt + TotalEquity), "")`
22. **Debt/(Debt+Equity)**:
    $$=\frac{\text{TotalDebt}}{\text{TotalDebt} + \text{TotalEquity}}$$
    *Công thức Excel:* `=IFERROR(TotalDebt/(TotalDebt + TotalEquity), "")`
23. **Debt to assets ratio (%) (Hệ số nợ trên tổng tài sản)**:
    $$=\frac{\text{TotalDebt}}{\text{TotalAssets}}$$
    *Công thức Excel:* `=IFERROR(TotalDebt/TotalAssets, "")`
24. **Interest coverage ratio (Hệ số chi trả lãi vay)**:
    $$=\frac{\text{EBIT}}{\text{InterestExp}}$$
    *Công thức Excel:* `=IFERROR(EBIT/InterestExp, "")`

---

## PHẦN IV: HƯỚNG DẪN THỰC THI NHANH
1. Mở file Excel chứa bảng số liệu.
2. Quét chọn (Select) toàn bộ vùng bảng dữ liệu (từ dòng Header đến hết dòng cuối cùng).
3. Nhấn **Alt + F8** $\rightarrow$ Chọn `CalculateRatios_SelectedRange` $\rightarrow$ Nhấn **Run**.
4. Chọn chế độ tính:
   - `1`: Chỉ tính Bảng cân đối kế toán.
   - `2`: Chỉ tính Báo cáo kết quả kinh doanh.
   - `3`: Tính toàn bộ bảng tổng hợp (Combined).
5. Kiểm tra bảng popup xác nhận các dòng chỉ tiêu đã tìm thấy $\rightarrow$ Nhấn **Yes**.
6. Chọn ô đích để dán kết quả $\rightarrow$ Toàn bộ công thức Excel động tuyệt đối (`Address(True, True)`) sẽ được sinh ra chuẩn xác 100%!

