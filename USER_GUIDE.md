# Nagarkot Bank Statement Reconciliation Tool User Guide

## Introduction
The Nagarkot Bank Statement Reconciliation Tool is a premium-branded, high-safety desktop application designed to streamline the parsing of bank statements, mapping withdrawals to their corresponding ledger accounts, and exporting clean CSV files compliant with the Logisys G/L ledger schema.

---

## How to Use

### 1. Launching the App
* **Using the Compiled Executable (For End-Users)**:
  Double-click the `Bank_Statement_Reco.exe` file inside the `dist/` directory. No console window will appear, launching directly into the fullscreen Nagarkot-branded window.
* **Using Python Source (For Developers)**:
  Activate the project's virtual environment and execute the script:
  ```powershell
  .\venv\Scripts\activate
  python Bank_Statement_Reco.py
  ```

---

### 2. The Workflow (Step-by-Step)

1. **Select Statement**: Click the **Browse Excel...** button. Select a statement file (`.xlsx` or legacy `.xls`).
   - *Note: The file must contain standard columns for Narration/Description and Withdrawal/Debit.*
2. **Execute Processing**: Click the primary **PROCESS STATEMENT** button. The tool automatically detects column headers, skips summary/header noise, resolves payee mappings, and displays the matching transactions in the interactive grid.
   - *Note: The parser automatically ignores deposits, empty lines, divider lines starting with `*`, and stops parsing immediately when it detects bank summary keywords like "Opening Balance", "Statement Summary", or "End of Statement".*
3. **Audit Results**: 
   - **Review Items**: Transactions that match predefined payee keywords display with a checkmark (`✓`) in the flag column. Unmatched items display with a warning (`⚠`) indicating they are categorized as `SUSPENSE - V`.
   - **Duplicate Auditing**: Payments on the same date with the same amount will automatically bold and highlight in a soft indigo background to indicate warning duplicates.
4. **Make Inline Corrections (Optional)**: Double-click any field (Date, Chq./Ref.No., Payee Name, Withdrawal Amt, Mode, or Flag) in the grid to make corrections:
   - Double-clicking **Payee Name** and typing a new value will automatically recalculate whether the payee is a known keyword and re-map the G/L mapping if appropriate.
   - Double-clicking **Withdrawal Amt** requires entering a valid positive numeric float.
5. **Manage Rows (Optional)**: Select one or more rows and press the **Delete** key on your keyboard (or right-click to choose **Delete Row**) to remove transactions before export.
6. **Export Logisys CSV**: Click **Export CSV** to save the formatted result into a clean uploadable spreadsheet.
   - *Note: If the tool still detects duplicate dates and amounts in your active table, it will present a warnings prompt to confirm export.*

---

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Browse Excel... Button** | Opens a file selector to locate statement files. | `.xlsx`, `.xls` files |
| **PROCESS STATEMENT Button** | Parses the selected Excel statement and populates the table. | Button triggers immediately |
| **Export CSV Button** | Saves verified transactions as a Logisys-compatible CSV upload file. | Saves default format: `Bank_Withdrawals_DDMonYY_HHMMSS.csv` |
| **Clear Data Button** | Resets the currently loaded file, transaction list, status logs, and summary counters. | Restores to fresh state |
| **Date Column** | Displays the statement transaction date. Double-click to edit. | `DD/MM/YY` (automatically formatted to `DD-Mon-YYYY` in output) |
| **Chq./Ref.No. Column** | Reference number of transaction. Double-click to edit. | String/Integer |
| **Payee Name Column** | Name of target recipient mapped by search keywords. Double-click to edit. | Plain text |
| **Withdrawal Amt Column** | The withdrawal transaction amount. Double-click to edit. | Numeric float (e.g. `1250.00`) |
| **Mode Column** | Determined mode of payment. NEFT if reference contains letters, else Others. | `NEFT` or `Others` |
| **Flag Column** | Mapped state flag showing matching audit details. | `✓` (Matched) or `⚠` (Suspense) |

---

## Troubleshooting & Validations

If you see an error, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| **Could not find standard bank statement headers (Narration, Withdrawal/Debit).** | The uploaded Excel sheet does not contain column headers matching standard bank keywords. | Ensure the sheet has standard headers (e.g., Description/Particulars/Narration and Withdrawal/Debit). |
| **Required columns (Narration/Description and Withdrawal/Debit) are missing.** | The statement parsed but required fields did not match column indices correctly. | Open Excel and check that description and withdrawal columns are properly labeled. |
| **Please select an Excel file first.** | Triggered **PROCESS STATEMENT** without browsing a file. | Click **Browse Excel...** to select your bank statement before processing. |
| **Invalid Amount: Please enter a valid number.** | User typed an invalid non-numeric string into the Withdrawal Amt cell. | Input only standard positive numeric floats (e.g. `500.50`, no letters or symbols). |
| **Duplicate Amounts Detected: There are duplicate amounts in the statement. Do you still want to export the CSV?** | Audit found identical Date and Amount pairs in the export dataset. | Click **No** to return and delete/edit duplicate rows, or **Yes** to export anyway. |
| **Export Error: [System Exception Details]** | The file is locked, open in Excel, or permissions prevent writing. | Close the target CSV file in Excel and check write permissions for the destination directory. |
