# AR_AP.vb Workflow Walkthrough

This document provides a top-to-bottom narrative explanation of the `AR_AP.vb` macro script, detailing its functions, how they run, and their roles in the overall process.

## Overview
The primary goal of this script is to merge Accounts Receivable (AR) and Accounts Payable (AP) data from three separate source Excel files into a single, consolidated "Merged" sheet within the current workbook.

## Helper Functions

### 1. `ConvertToNumber(ByVal val As Variant) As Variant`
- **Role:** Safely converts a given value to a numeric format if possible, otherwise returns it as a cleaned string.
- **How it runs:** 
  1. Checks if the input is `Null` or `Empty`.
  2. Trims and cleans the input to remove non-printable characters.
  3. Uses `IsNumeric()` to check if the string can be converted.
  4. Returns the converted number (using `CDbl`) or the original cleaned string.
- **Inputs:** `val` (Variant) - The cell value to be converted.
- **Output:** The converted numeric value or the original cleaned string (Variant).

### 2. `GetWorkbook(ByVal filePath As String, ByRef currentWB As Workbook) As Workbook`
- **Role:** Retrieves a workbook reference, opening the file if it isn't already open.
- **How it runs:**
  1. If the `filePath` is empty, it returns the current workbook.
  2. Otherwise, it attempts to assign the workbook using `Workbooks(filePath)`.
  3. If it fails (meaning the file isn't open), it opens the file using `Workbooks.Open(filePath)`.
- **Inputs:** `filePath` (String) representing the source file path, and `currentWB` (Workbook) as a fallback.
- **Output:** A `Workbook` object reference to the requested file.

## Main Macro

### 3. `MergeThreeSourcesWithAccount_Fixed()`
- **Role:** The main execution sub that drives the entire data merging process.
- **How it runs:**
  1. **Configuration:** Defines constants for the three source files (CAN, DOAU, and AR AP loan), including file paths, sheet names, and column indices for Kind, Entity, Name, Ending Balance, and Account.
  2. **Initialization:** Calls `GetWorkbook()` to securely retrieve or open all three source workbooks. If any fail to open, it halts execution and alerts the user.
  3. **Result Sheet Preparation:** Deletes any existing "Merged" sheet in the current workbook and creates a fresh one. It sets up the headers (Kind, Entity, Name, Ending Balance, Account) and formats the columns appropriately.
  4. **Data Extraction & Merging:**
     - **Source 1 (CAN):** Iterates through rows, copying Kind, Entity, Name, and Ending Balance. Leaves the Account column empty. Uses `ConvertToNumber` on the balance.
     - **Source 2 (DOAU):** Similar to Source 1, it copies the same columns and leaves the Account column empty.
     - **Source 3 (AR AP loan):** Copies the standard columns but *also* populates the Account column using `ConvertToNumber` on the account value.
  5. **Finalization:** Auto-fits the columns in the "Merged" sheet and displays a completion message to the user.
- **Inputs:** None (relies on hardcoded constants for file paths).
- **Output:** A newly generated sheet named "Merged" containing consolidated records from the three source workbooks, and a popup message confirming completion.
