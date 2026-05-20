# Bank Statement Reconciliation Tool

An advanced, premium-branded desktop utility designed to streamline bank statement reconciliation and preparation for Logisys G/L ledger entries.

## Key Features

* **Dual-Engine Excel Parsing**: Automated format detection handling both modern `.xlsx` statements (via `openpyxl`) and legacy `.xls` statements (via `xlrd`).
* **Intelligent Summary Row Elimination**: Automatically detects and stops parsing at the HDFC statement summary/footer rows to avoid dirtying withdrawals data.
* **Payee Mapping & Suspense Classification**: Checks statement narrations against a predefined library of major logistics payees, automatically mapping transactions to their correct G/L ledger accounts or marking unmatched items as `SUSPENSE - V`.
* **Interactive Live Grid**:
  * **Double-Click Inline Editing**: Correct dates, cheque numbers, payee names, withdrawal amounts, or payment modes directly within the spreadsheet cell.
  * **Interactive Duplicate Detection**: Duplicate payments on the same date with the same amount are immediately bolded for high visibility.
  * **Delete Controls**: Remove individual records manually using the **Delete key** on your keyboard or using the right-click context menu.
* **Logisys-Compliant CSV Export**: Exports perfectly structured files formatted to the strict schema requirements of Logisys accounting software.

---

## Installation & Setup

### 1. Clone or Download Project
```bash
git clone <repository-url>
cd "Bank Statement Reco"
```

### 2. Create Virtual Environment
```bash
python -m venv venv
```

### 3. Activate Virtual Environment (MANDATORY)

* **Windows (PowerShell)**:
  ```powershell
  .\venv\Scripts\activate
  ```
* **Windows (Command Prompt)**:
  ```cmd
  venv\Scripts\activate
  ```
* **Mac/Linux**:
  ```bash
  source venv/bin/activate
  ```

### 4. Install Dependencies
Ensure you see `(venv)` at the beginning of your terminal line, then install the pinned project requirements:
```bash
pip install -r requirements.txt
```

---

## Running the Application
With the virtual environment active, run the main GUI entrypoint:
```bash
python Bank_Statement_Reco.py
```

---

## Build Executable (PyInstaller)
To package the desktop application into a standalone Windows executable containing all static assets:

1. **Install PyInstaller** (Ensure your virtual environment is active):
   ```bash
   pip install pyinstaller
   ```
2. **Build using the Spec File**:
   ```bash
   pyinstaller Bank_Statement_Reco.spec
   ```
3. **Locate Executable**:
   The standalone, zero-console window application will be generated under the `dist/` directory.

---

## File Structure
* `Bank_Statement_Reco.py` - Core application logic, GUI layout, statement parser, and mapping rules.
* `logo.png` - Nagarkot premium brand asset used in header.
* `requirements.txt` - Python project dependencies.
* `.gitignore` - Standard configuration ignoring system and output cache.

---

## Technical Specifications & Mapping Rules

### Active Payee Mapping:
Transactions containing these logistics payee names are automatically resolved to their legal G/L account numbers:
* ALLCARGO TERMINALS LIMITED
* AMEYA LOGISTICS PVT. LTD.
* ASHTE LOGISTICS PVT LTD
* AVANT CAREER PVT. LTD.
* BALMER LAWRIER & CO LTD
* CENTRAL WAREHOUSING CORPORATION
* DACHSER INDIA PVT LTD
* DHL LOGISTICS PVT LTD
* DIVVYA CPP PRIVATE LIMITED
* GATEWAY DISTRIPARKS LTD.
* HAPAG LLOYD INDIA PVT LTD
* HASTI PETRO CHEMICAL & SHIPPING LTD.
* JAY BHARAT STATIONERY
* JWC LOGISTICS PARK PVT.LTD.
* JWR LOGISTICS PVT LTD
* KALE LOGISTICS SOLUTIONS PVT LTD
* MUMBAI CARGO SERVICE CENTRE
* NAVKAR CORPORATION LTD
* Sarveshwar Logistics Services Pvt Ltd.
* SEABIRD MARINE SERVICES PVT. LTD
* SPEEDY MULTIMODES LTD

All unmapped entries are designated under `SUSPENSE - V` for manual review in the preview table.
