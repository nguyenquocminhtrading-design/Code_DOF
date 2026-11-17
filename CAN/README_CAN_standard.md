# CAN_standard.vb Workflow Walkthrough

This document provides a narrative explanation of the `CAN_standard.vb` macro script, which automatically generates a customized "003-CAN" Balance Sheet and Profit/Loss report directly from the underlying data sheets.

## Overview
The main purpose of this script is to generate the "003-CAN" report sheet entirely from the "B Sheet", "IS" (Income Statement), and "TB" (Trial Balance) sheets. This eliminates the need for an intermediate "004- CAN" sheet, streamlining the reporting process.

## Main Macro

### 1. `Create003CAN()`
- **Role:** The core sub that orchestrates the creation of the "003-CAN" report.
- **How it runs:**
  1. **User Input & Validation:** Prompts the user to enter a specific month (e.g., "2026.07"). It then verifies the existence of the "B Sheet" and "IS" sheets; if missing, it aborts. It also checks for the "TB" sheet (used for EBITDA interest add-back) but only warns if it's missing.
  2. **Sheet Preparation:** Deletes any existing "003-CAN" sheet to ensure a fresh start, then creates a new sheet named "003-CAN".
  3. **Title and Headers:** Sets up the main headers (BALANCE SHEET, ASSETS, CURRENT ASSETS, etc.) and records the user-input month in the designated cell.
  4. **Balance Sheet Construction:**
     - Uses an array of account categories (e.g., Cash, Accounts receivable, Inventory).
     - Iterates through the array and calls the `WriteLine` helper function to populate the "003-CAN" sheet with formulas that pull the corresponding SUM from the "B Sheet".
     - Adds subtotal formulas (e.g., total current assets, total liabilities) and a final equity check calculation (`Check = SHAREHOLDER'S EQUITY - ASSETS`).
  5. **Profit and Loss Construction:**
     - Constructs the P&L section sequentially (Sales, Cost of Sales, Gross Profit, Operating Expenses, G&A Expenses, etc.).
     - Uses `WriteLine` to insert `SUMIFS` formulas referencing the "IS" sheet.
     - Calculates sub-totals (Total Sales Expenses, Total Operating Expenses) and major metrics like Operating Profit/Loss and Earnings Before Income Taxes (EBIT).
  6. **EBITDA Calculation:**
     - Checks if the "TB" sheet exists. If it does, it dynamically sums interest metrics directly from the trial balance for specific account codes (800210, 800220, 800230).
     - Calculates the final EBITDA by combining EBIT, amortizations, interest, and other specific adjustments.
  7. **Formatting:** Applies bold formatting to major headers and auto-fits the columns for readability.
- **Inputs:** User prompt for the reporting month. Data is pulled from "B Sheet", "IS", and "TB" worksheets.
- **Output:** A newly created, fully formulated, and formatted "003-CAN" worksheet, alongside a confirmation message.

## Helper Function

### 2. `WriteLine(wsNew As Worksheet, sourceSheet As String, ByVal r As Long, ByVal label As String)`
- **Role:** A modular helper that writes a specific line item's label and dynamically generates its `SUMIFS` formula.
- **How it runs:**
  - Places the `label` in column B of the specified row `r`.
  - Checks the `sourceSheet`:
    - If "B Sheet", it constructs a formula to sum data from the "B Sheet" matching the `label` where a specific criteria column is "TRUE".
    - If "IS", it constructs a formula to sum data from the "IS" sheet using similar logic.
- **Inputs:** 
  - `wsNew` (Worksheet): The target sheet ("003-CAN").
  - `sourceSheet` (String): The source data sheet ("B Sheet" or "IS").
  - `r` (Long): The target row number.
  - `label` (String): The name of the line item to match.
- **Output:** Writes the label and corresponding excel formula directly into the worksheet.
