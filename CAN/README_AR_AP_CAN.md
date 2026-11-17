# AR_AP_CAN.bas Workflow Walkthrough

This document provides a narrative explanation of the `AR_AP_CAN.bas` macro script, which automates the consolidation of Accounts Receivable (AR) and Accounts Payable (AP) for the CAN entity.

## Overview
The script extracts specific records from the "AR" and "AP" sheets of the current workbook and combines them into a newly created "Merged" sheet. It standardizes the format by adding a "Kind" (AR/AP) and "Entity" (CAN) identifier to each record.

## Main Macro

### 1. `CAN_merge()`
- **Role:** The primary sub responsible for reading AR and AP data and merging them.
- **How it runs:**
  1. **Sheet Initialization:** It attempts to locate the "AR" and "AP" sheets within `ThisWorkbook`. If either is missing, it alerts the user and exits the subroutine.
  2. **Result Sheet Creation:** It adds a new worksheet at the end of the workbook and names it "Merged". It sets up the header row: `Kind`, `Entity`, `Cus.no`, `Name`, and `Ending Balance`.
  3. **AR Data Extraction:** Iterates through the "AR" sheet starting from row 2 until the last populated row in column A. If column A is not blank, it copies the Customer Number, Name, and Ending Balance to the "Merged" sheet. It hardcodes "AR" as the Kind and "CAN" as the Entity.
  4. **AP Data Extraction:** Iterates through the "AP" sheet in the same manner. It copies Vendor Number (to `Cus.no`), Name, and Ending Balance to the "Merged" sheet, hardcoding "AP" as the Kind and "CAN" as the Entity.
  5. **Finalization:** Auto-fits columns A through E in the "Merged" sheet and displays a success message box indicating where the results are located.
- **Inputs:** Data from existing sheets named "AR" and "AP".
- **Output:** A newly created "Merged" sheet containing the consolidated AR and AP records for the CAN entity.
