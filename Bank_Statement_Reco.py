import os
import sys
import csv
import datetime
import threading
from pathlib import Path

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl

# ---------------------------------------------------------
# CONFIGURATION & MAPPINGS
# ---------------------------------------------------------

# Payee mappings: keyword -> full Payee Name
# Sorted by keyword length descending so longer/more-specific keywords
# match before shorter/generic ones (e.g. "CENTRAL WAREHOUSING" before "CWC").
# Use the shortest reliable fragment that uniquely identifies the payee,
# especially for cheque narrations which are truncated and often misspelled.
#
# IMPORTANT: Only payees that exist in Logisys are mapped here.
# All other withdrawals go to SUSPENSE - V for manual review.
# Format: (search_keyword, payee_name, gl_name)
PAYEE_MAPPINGS = [
    # --- ALLCARGO TERMINALS LIMITED ---
    ("ALLCARGO",            "ALLCARGO TERMINALS LIMITED",                            "ALLCARGO TERMINALS LIMITED - V"),

    # --- AMEYA LOGISTICS PVT. LTD. ---
    ("AMEYA LOG",           "AMEYA LOGISTICS PVT. LTD.",                             "AMEYA LOGISTICS PVT. LTD. - V"),

    # --- ASHTE LOGISTICS PVT LTD ---
    # Bank truncates & misspells: ASTHE, ASIITE, ASHTE LOGSITICS, ASHTE LOGISTIS
    ("ASHTE",               "ASHTE LOGISTICS PVT LTD",                               "ASHTE LOGISTICS PVT LTD - V"),
    ("ASTHE",               "ASHTE LOGISTICS PVT LTD",                               "ASHTE LOGISTICS PVT LTD - V"),
    ("ASIITE",              "ASHTE LOGISTICS PVT LTD",                               "ASHTE LOGISTICS PVT LTD - V"),

    # --- AVANT CAREER PVT. LTD. ---
    ("AVANT CAREER",        "AVANT CAREER PVT. LTD.",                                "AVANT CAREER PVT. LTD. - V"),

    # --- BALMER LAWRIER & CO LTD ---
    ("BALMER LAWR",         "BALMER LAWRIER & CO LTD",                               "BALMER LAWRIER & CO LTD - V"),

    # --- CENTRAL WAREHOUSING CORPORATION (CWC - IMPEX PARK ---
    ("CENTRALWAREHOUSING",  "CENTRAL WAREHOUSING CORPORATION (CWC - IMPEX PARK",     "CENTRAL WAREHOUSING CORPORATION (CWC - IMPEX PARK - V"),
    ("CENTRAL WAREHOUSING", "CENTRAL WAREHOUSING CORPORATION (CWC - IMPEX PARK",     "CENTRAL WAREHOUSING CORPORATION (CWC - IMPEX PARK - V"),
    ("CENTRAL WH",          "CENTRAL WAREHOUSING CORPORATION (CWC - IMPEX PARK",     "CENTRAL WAREHOUSING CORPORATION (CWC - IMPEX PARK - V"),
    ("CWC",                 "CENTRAL WAREHOUSING CORPORATION (CWC - IMPEX PARK",     "CENTRAL WAREHOUSING CORPORATION (CWC - IMPEX PARK - V"),

    # --- DACHSER INDIA PVT LTD ---
    ("DACHSER",             "DACHSER INDIA PVT LTD",                                 "DACHSER INDIA PVT LTD - V"),

    # --- DHL LOGISTICS PVT LTD ---
    ("DHL LOG",             "DHL LOGISTICS PVT LTD",                                 "DHL LOGISTICS PVT LTD - V"),

    # --- DIVVYA CPP PRIVATE LIMITED ---
    # Bank misspells: DIVVYA CFP, DIVYA CPP
    ("DIVVYA C",            "DIVVYA CPP PRIVATE LIMITED",                            "DIVVYA CPP PRIVATE LIMITED - V"),
    ("DIVYA CPP",           "DIVVYA CPP PRIVATE LIMITED",                            "DIVVYA CPP PRIVATE LIMITED - V"),
    ("DIVYA C",             "DIVVYA CPP PRIVATE LIMITED",                            "DIVVYA CPP PRIVATE LIMITED - V"),

    # --- GATEWAY DISTRIPARKS LTD. ---
    ("GATEWAY DISTRI",      "GATEWAY DISTRIPARKS LTD.",                              "GATEWAY DISTRIPARKS LTD. - V"),

    # --- HAPAG LLOYD INDIA PVT LTD ---
    ("HAPAG",               "HAPAG LLOYD INDIA PVT LTD",                             "HAPAG LLOYD INDIA PVT LTD - V"),

    # --- HASTI PETRO CHEMICAL & SHIPPING LTD. ---
    ("HASTI PETRO",         "HASTI PETRO CHEMICAL & SHIPPING LTD.",                  "HASTI PETRO CHEMICAL & SHIPPING LTD. - V"),

    # --- JAY BHARAT STATIONERY ---
    ("JAY BHARAT",          "JAY BHARAT STATIONERY",                                 "JAY BHARAT STATIONERY - V"),

    # --- JWC LOGISTICS PARK PVT.LTD. ---
    ("JWC LOGISTICS",       "JWC LOGISTICS PARK PVT.LTD.",                           "JWC LOGISTICS PARK PVT.LTD. - V"),

    # --- JWR LOGISTICS PVT LTD ---
    ("JWR LOGISTICS",       "JWR LOGISTICS PVT LTD",                                "JWR LOGISTICS PVT LTD - V"),

    # --- KALE LOGISTICS SOLUTIONS PVT LTD ---
    ("KALE LOGISTICS",      "KALE LOGISTICS SOLUTIONS PVT LTD",                     "KALE LOGISTICS SOLUTIONS PVT LTD - V"),

    # --- MUMBAI CARGO SERVICE CENTRE AIRPORT PRIVATE LIMITED ---
    ("MUMBAI CARGO",        "MUMBAI CARGO SERVICE CENTRE AIRPORT PRIVATE LIMITED",   "MUMBAI CARGO SERVICE CENTRE AIRPORT PRIVATE LIMITED - V"),

    # --- NAVKAR CORPORATION LTD ---
    ("NAVKAR CORP",         "NAVKAR CORPORATION LTD",                                "NAVKAR CORPORATION LTD - V"),

    # --- Sarveshwar Logistics Services Pvt Ltd. - SEZ ---
    ("SARVESHWAR LOG",      "Sarveshwar Logistics Services Pvt Ltd. - SEZ",          "Sarveshwar Logistics Services Pvt Ltd. - SEZ - V"),
    ("SARVESHWAR",          "Sarveshwar Logistics Services Pvt Ltd. - SEZ",          "Sarveshwar Logistics Services Pvt Ltd. - SEZ - V"),

    # --- SEABIRD MARINE SERVICES PVT. LTD ---
    ("SEABIRD MARINE",      "SEABIRD MARINE SERVICES PVT. LTD",                     "SEABIRD MARINE SERVICES PVT. LTD - V"),

    # --- SPEEDY MULTIMODES LTD ---
    ("SPEEDY MULTIMODES",   "SPEEDY MULTIMODES LTD",                                "SPEEDY MULTIMODES LTD - V"),
]

FIXED_BANK_NAME = "HDFC BANK-GOVANDI-CC-773862"
DEFAULT_PAYEE = "SUSPENSE - V"

# ---------------------------------------------------------
# GLOBAL STATE
# ---------------------------------------------------------
selected_excel = ""
processed_results = []


# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------

def match_payee(narration):
    """Search the narration for known payee keywords.
    Returns (payee_name, gl_name) tuple.
    """
    if not narration:
        return DEFAULT_PAYEE, "SUSPENSE - V"

    narration_upper = str(narration).upper()

    for keyword, full_name, gl_name in PAYEE_MAPPINGS:
        if keyword.upper() in narration_upper:
            return full_name, gl_name

    return DEFAULT_PAYEE, "SUSPENSE - V"


def format_ref_no(ref_val):
    """Convert a ref number cell value to a clean string.
    Excel/.xls may store long numeric refs as floats, producing scientific
    notation like 6.13818E+11. This handles both:
    - Python float type (from xlrd/openpyxl reading numeric cells)
    - String already in scientific notation (e.g. '6.13818E+11')
    """
    if ref_val is None or ref_val == "":
        return ""
    # If it's a float, convert to int (drops .0)
    if isinstance(ref_val, float):
        if ref_val == int(ref_val):
            return str(int(ref_val))
        return str(ref_val)
    # If it's a string that looks like scientific notation, parse it
    s = str(ref_val).strip()
    if "E+" in s.upper() or "E-" in s.upper():
        try:
            num = float(s)
            if num == int(num):
                return str(int(num))
            return f"{num:.0f}"
        except ValueError:
            pass
    return s


def detect_payment_mode(ref_no):
    """Determine payment mode from the reference number.
    - If ref starts with alphabetic characters (e.g. HDFCR..., HDFCH..., NBT5...) -> NEFT
    - If ref is all digits (e.g. 0000000000030678) -> Others (cheque)
    """
    ref_clean = str(ref_no).strip()
    if not ref_clean:
        return "NEFT"  # default

    # Check if any letter exists in the leading portion
    has_leading_alpha = ref_clean[0].isalpha()

    if has_leading_alpha:
        return "NEFT"
    else:
        return "Others"


# ---------------------------------------------------------
# DATA PROCESSING
# ---------------------------------------------------------

def process_bank_statement(filepath):
    """Reads the Bank Statement Excel and maps it to Logisys format."""
    global processed_results
    processed_results.clear()

    try:
        filepath_lower = str(filepath).lower()
        if filepath_lower.endswith('.xls') and not filepath_lower.endswith('.xlsx'):
            import xlrd
            wb = xlrd.open_workbook(filepath)
            sheet = wb.sheet_by_index(0)
            total_rows = sheet.nrows

            def get_row(rx):
                return sheet.row_values(rx)
        else:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            sheet = wb.active
            # Cache all rows for consistent 0-based indexing
            all_rows = list(sheet.iter_rows(values_only=True))
            total_rows = len(all_rows)

            def get_row(rx):
                return list(all_rows[rx])

        # 1. Find the header row (looking for 'withdrawal'/'debit' AND 'narration')
        headers = []
        header_row_idx = None

        for rx in range(total_rows):
            row = get_row(rx)
            row_strs = [str(cell).strip().lower() if cell else "" for cell in row]
            has_narration = any(
                "narration" in c or "description" in c or "particulars" in c
                for c in row_strs
            )
            has_withdrawal = any("withdrawal" in c or "debit" in c for c in row_strs)
            if has_narration and has_withdrawal:
                headers = row_strs
                header_row_idx = rx
                break

        if not headers:
            raise ValueError(
                "Could not find standard bank statement headers "
                "(Narration, Withdrawal/Debit)."
            )

        # 2. Identify key column indices
        narration_idx = next(
            (i for i, h in enumerate(headers)
             if "narration" in h or "description" in h or "particulars" in h),
            -1
        )
        withdrawal_idx = next(
            (i for i, h in enumerate(headers) if "withdrawal" in h or "debit" in h),
            -1
        )
        date_idx = next(
            (i for i, h in enumerate(headers)
             if "date" in h and "value" not in h),
            -1
        )
        ref_idx = next(
            (i for i, h in enumerate(headers) if "chq" in h or "ref" in h),
            -1
        )

        if narration_idx == -1 or withdrawal_idx == -1:
            raise ValueError(
                "Required columns (Narration/Description and Withdrawal/Debit) "
                "are missing."
            )

        # 3. Process data rows (starting after the header)
        serial = 0

        for rx in range(header_row_idx + 1, total_rows):
            row = get_row(rx)

            # Skip empty rows or separator rows (e.g. rows full of ********)
            first_cell = str(row[0]).strip() if row[0] is not None else ""
            if not first_cell or first_cell.startswith("*"):
                continue

            # Stop parsing if we reach the summary section at the bottom
            row_strs_lower = [
                str(cell).strip().lower() for cell in row if cell is not None
            ]
            if any(
                kw in c
                for c in row_strs_lower
                for kw in ("statement summary", "opening balance",
                           "closing bal", "end of statement")
            ):
                break

            try:
                withdrawal_amt = row[withdrawal_idx]
                narration = row[narration_idx]
                date_val = row[date_idx] if date_idx != -1 else ""
                ref_num_raw = row[ref_idx] if ref_idx != -1 else ""
                ref_num = format_ref_no(ref_num_raw)
            except IndexError:
                continue

            # Skip non-withdrawal rows (deposits / empty)
            if withdrawal_amt is None or withdrawal_amt == "" or withdrawal_amt == 0:
                continue

            try:
                amount = float(withdrawal_amt)
                if amount <= 0:
                    continue
            except (ValueError, TypeError):
                continue

            # Parse statement date (DD/MM/YY) -> DD-Mon-YYYY for Logisys
            stmt_date = ""
            date_str = str(date_val).strip()
            if date_str:
                try:
                    dt = datetime.datetime.strptime(date_str, "%d/%m/%y")
                    stmt_date = dt.strftime("%d-%b-%Y")
                except ValueError:
                    # Fallback: use as-is if format differs
                    stmt_date = date_str

            # Match Payee (returns payee_name, gl_name)
            payee_name, gl_name = match_payee(narration)
            is_known = payee_name != DEFAULT_PAYEE

            # Determine Mode of Payment from ref number
            mode_of_payment = detect_payment_mode(ref_num)

            # Narration logic:
            # Known payee  -> "BEING AMOUNT PAID TO <Payee Name>"
            # Unknown payee -> "BEING AMOUNT PAID TO <original narration>"
            if is_known:
                narration_text = f"BEING AMOUNT PAID TO {payee_name}"
            else:
                narration_text = f"BEING AMOUNT PAID TO {str(narration).strip()}"

            # Serial Record ID
            serial += 1

            # Build CSV Row
            row_data = {
                "Record ID": serial,
                "Entry Date": stmt_date,
                "Posting Date": stmt_date,
                "Cash / Bank Name": FIXED_BANK_NAME,
                "Amount": round(amount, 2),
                "Currency": "INR",
                "Ex. Rate": 1,
                "Mode Of Payment": mode_of_payment,
                "Payment Ref No.": ref_num,
                "Payment Ref. Date": stmt_date,
                "Payee Name": payee_name,
                "Payee Type": "",
                "Vendor Ref. No": "",
                "Vendor Ref. Date": "",
                "Narration": narration_text,
                "Charge / GL": "G/L",
                "Charge/GL Name": gl_name,
                "Charge/GL Amount": round(amount, 2),
                "DR/CR": "DR",
                "Receipt Entry - Receipt No.": "",
                "Receipt Entry - Date": "",
                "Receipt Entry - Amount": "",
                "Cost Center": "",
                "Branch": "HO",
                "Charge Narration": "",
                "TaxGroup": "GSTIN",
                "Tax Type": "Exempt",
                "SAC/HSN": "",
                "Taxcode1": "",
                "Taxcode1 Amt": "",
                "Taxcode2": "",
                "Taxcode2 Amt": "",
                "Taxcode3": "",
                "Taxcode3 Amt": "",
                "Taxcode4": "",
                "Taxcode4 Amt": "",
                "Avail Tax Credit": "NO",
                "LOB": "",
                "Ref. Type": "",
                "Ref. No.": "",
                "Ref. Amount": "",
                "WH Tax Organization": "",
                "WH Tax Code": "",
                "WH Tax Percentage": "",
                "WH Tax Taxable": "",
                "Start Date": "",
                "End Date": "",
                "CC Code": "",
                # UI-only helper fields (excluded from CSV export via extrasaction='ignore')
                "_Date": date_str,
                "_RefNo": ref_num,
                "_OriginalNarration": str(narration),
                "_Flag": "⚠" if not is_known else "✓",
                "_ModeOfPayment": mode_of_payment,
            }
            processed_results.append(row_data)

        return True, ""
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------
# GUI
# ---------------------------------------------------------
BRAND_BLUE = "#1F3F6E"
BRAND_BG = "#F4F6F8"
ACCENT_RED = "#D8232A"
DARK_TEXT = "#1E1E1E"
MUTED_GRAY = "#6B7280"
PANEL_WHITE = "#FFFFFF"
BORDER_GRAY = "#E5E7EB"
HOVER_BLUE = "#2A528F"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bank Statement to Logisys CSV")

        try:
            self.state('zoomed')
        except Exception:
            self.geometry("1100x750")

        self.configure(bg=BRAND_BG)

        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TFrame', background=BRAND_BG)
        style.configure('TLabel', background=BRAND_BG, foreground=DARK_TEXT,
                         font=("Segoe UI", 10))
        
        # Primary Action Button Styling
        style.configure('Action.TButton', background=BRAND_BLUE,
                         foreground="white", font=("Segoe UI", 10, "bold"),
                         borderwidth=1, bordercolor=BRAND_BLUE,
                         padding=8)
        style.map('Action.TButton',
                  background=[('disabled', BORDER_GRAY), ('active', HOVER_BLUE)],
                  foreground=[('disabled', MUTED_GRAY), ('active', 'white')],
                  bordercolor=[('disabled', BORDER_GRAY), ('active', HOVER_BLUE)])
        
        # Secondary Action Button Styling
        style.configure('Secondary.TButton', background=PANEL_WHITE,
                         foreground=BRAND_BLUE, font=("Segoe UI", 10),
                         borderwidth=1, bordercolor=BORDER_GRAY,
                         padding=5)
        style.map('Secondary.TButton',
                  background=[('disabled', PANEL_WHITE), ('active', BRAND_BG)],
                  foreground=[('disabled', MUTED_GRAY), ('active', HOVER_BLUE)],
                  bordercolor=[('disabled', BORDER_GRAY), ('active', BORDER_GRAY)])

        style.configure('Treeview', background=PANEL_WHITE,
                         foreground=DARK_TEXT, fieldbackground=PANEL_WHITE,
                         rowheight=28, borderwidth=0)
        style.configure('Treeview.Heading', background=BRAND_BLUE,
                         foreground="white",
                         font=("Segoe UI", 10, "bold"), relief="flat")
        style.map('Treeview.Heading', 
                  background=[('active', HOVER_BLUE)],
                  foreground=[('active', 'white')])
        style.map('Treeview', background=[('selected', HOVER_BLUE)],
                   foreground=[('selected', 'white')])

        # ---------------------------------------------------------
        # HEADER
        # ---------------------------------------------------------
        header_bg = tk.Frame(self, bg=PANEL_WHITE)
        header_bg.pack(fill=tk.X)
        header_bg.columnconfigure(0, weight=1)
        header_bg.columnconfigure(1, weight=0)
        header_bg.columnconfigure(2, weight=1)

        logo_frame = tk.Frame(header_bg, bg=PANEL_WHITE)
        logo_frame.grid(row=0, column=0, sticky="w", padx=20, pady=15)

        logo_path = resource_path("logo.png")
        if os.path.exists(logo_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path)
                ratio = 20 / img.height
                img = img.resize(
                    (int(img.width * ratio), 20), Image.Resampling.LANCZOS
                )
                self.logo_img = ImageTk.PhotoImage(img)
                tk.Label(logo_frame, image=self.logo_img,
                         bg=PANEL_WHITE).pack(side=tk.LEFT)
            except Exception:
                tk.Label(logo_frame, text="NAGARKOT",
                         font=("Segoe UI", 14, "bold"),
                         bg=PANEL_WHITE, fg=BRAND_BLUE).pack(side=tk.LEFT)
        else:
            tk.Label(logo_frame, text="NAGARKOT",
                     font=("Segoe UI", 14, "bold"),
                     bg=PANEL_WHITE, fg=BRAND_BLUE).pack(side=tk.LEFT)

        title_frame = tk.Frame(header_bg, bg=PANEL_WHITE)
        title_frame.grid(row=0, column=1, pady=15)

        # Force column 2 to balance column 0 exactly for absolute horizontal centering
        dummy_right = tk.Frame(header_bg, bg=PANEL_WHITE, width=1)
        dummy_right.grid(row=0, column=2, sticky="nsew")
        tk.Label(
            title_frame,
            text="Bank Statement Extractor (Withdrawals)",
            font=("Segoe UI", 16, "bold"), bg=PANEL_WHITE, fg=BRAND_BLUE
        ).pack()
        tk.Label(
            title_frame,
            text="Converts Bank Statements to Logisys Payment CSV",
            font=("Segoe UI", 10), bg=PANEL_WHITE, fg=MUTED_GRAY
        ).pack(pady=(2, 0))

        tk.Frame(self, bg=BORDER_GRAY, height=1).pack(fill=tk.X)

        # ---------------------------------------------------------
        # BODY
        # ---------------------------------------------------------
        body_frame = tk.Frame(self, bg=BRAND_BG)
        body_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=20)

        # Top Controls
        controls_frame = tk.Frame(body_frame, bg=BRAND_BG)
        controls_frame.pack(fill=tk.X, pady=(0, 15))

        sel_frame = tk.Frame(controls_frame, bg=BRAND_BG)
        sel_frame.pack(side=tk.LEFT)
        ttk.Button(sel_frame, text="Browse Excel...",
                    style="Secondary.TButton",
                    command=self.browse_file).pack(side=tk.LEFT, padx=(0, 10))
        self.lbl_file = tk.Label(
            sel_frame, text="Bank Statement: None selected",
            bg=BRAND_BG, fg=DARK_TEXT, font=("Segoe UI", 10)
        )
        self.lbl_file.pack(side=tk.LEFT)

        action_frame = tk.Frame(controls_frame, bg=BRAND_BG)
        action_frame.pack(side=tk.RIGHT)
        self.lbl_status = tk.Label(
            action_frame, text="Ready.", bg=BRAND_BG, fg=MUTED_GRAY,
            font=("Segoe UI", 10, "italic")
        )
        self.lbl_status.pack(side=tk.LEFT, padx=(0, 15))
        self.btn_process = ttk.Button(
            action_frame, text="PROCESS STATEMENT",
            style="Action.TButton", command=self.run_process,
            state=tk.DISABLED
        )
        self.btn_process.pack(side=tk.LEFT)
        self.btn_export = ttk.Button(
            action_frame, text="Export CSV",
            style="Action.TButton", command=self.export_csv,
            state=tk.DISABLED
        )
        self.btn_export.pack(side=tk.LEFT, padx=(10, 0))

        # Preview Table
        table_container = tk.Frame(
            body_frame, bg=PANEL_WHITE,
            highlightbackground=BORDER_GRAY, highlightthickness=1
        )
        table_container.pack(fill=tk.BOTH, expand=True)

        columns = ("Date", "Chq./Ref.No.", "Payee Name",
                    "Withdrawal Amt", "Mode", "Flag")
        self.tree = ttk.Treeview(
            table_container, columns=columns,
            show="headings", selectmode="none"
        )

        self.tree.heading("Date", text="Date")
        self.tree.heading("Chq./Ref.No.", text="Chq./Ref.No.")
        self.tree.heading("Payee Name", text="Payee Name")
        self.tree.heading("Withdrawal Amt", text="Withdrawal Amt")
        self.tree.heading("Mode", text="Mode")
        self.tree.heading("Flag", text="Flag")

        self.tree.column("Date", width=90, anchor=tk.CENTER)
        self.tree.column("Chq./Ref.No.", width=180, anchor=tk.CENTER)
        self.tree.column("Payee Name", width=300, anchor=tk.W)
        self.tree.column("Withdrawal Amt", width=120, anchor=tk.CENTER)
        self.tree.column("Mode", width=80, anchor=tk.CENTER)
        self.tree.column("Flag", width=60, anchor=tk.CENTER)

        self.tree.tag_configure('oddrow', background=PANEL_WHITE, foreground=DARK_TEXT, font=("Segoe UI", 10))
        self.tree.tag_configure('evenrow', background=BRAND_BG, foreground=DARK_TEXT, font=("Segoe UI", 10))
        
        # Duplicate rows (Indigo/Blue alert state)
        self.tree.tag_configure('duplicate_odd', background="#EDF2FF", foreground="#364FC7", font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure('duplicate_even', background="#DBE4FF", foreground="#364FC7", font=("Segoe UI", 10, "bold"))

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                        padx=(1, 0), pady=1)
        scrollbar = ttk.Scrollbar(
            table_container, orient=tk.VERTICAL, command=self.tree.yview
        )
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=1, padx=(0, 1))

        # Enable double-click editing, manual row delete, and right-click menu
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", self.delete_selected_row)
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Delete Row", command=self.delete_selected_row_context)
        self.tree.bind("<Button-3>", self.show_context_menu)

        # ---------------------------------------------------------
        # APP FOOTER
        # ---------------------------------------------------------
        app_footer = tk.Frame(self, bg=BRAND_BG)
        app_footer.pack(fill=tk.X, padx=40, pady=(0, 10))

        self.btn_clear = ttk.Button(
            app_footer, text="Clear Data",
            style="Secondary.TButton", command=self.clear_data
        )
        self.btn_clear.pack(side=tk.LEFT)

        self.lbl_summary = tk.Label(
            app_footer, text="", bg=BRAND_BG, fg=DARK_TEXT,
            font=("Segoe UI", 10, "bold")
        )
        self.lbl_summary.pack(side=tk.LEFT, padx=20)

        # ---------------------------------------------------------
        # BRANDING FOOTER
        # ---------------------------------------------------------
        brand_footer = tk.Frame(self, bg=BRAND_BG)
        brand_footer.pack(fill=tk.X, side=tk.BOTTOM, padx=20, pady=10)
        tk.Label(
            brand_footer, text="Nagarkot Forwarders Pvt. Ltd. \u00A9",
            font=("Segoe UI", 8), bg=BRAND_BG, fg=MUTED_GRAY
        ).pack(side=tk.LEFT)

    # ---------------------------------------------------------
    # CALLBACKS
    # ---------------------------------------------------------
    def browse_file(self):
        global selected_excel
        file = filedialog.askopenfilename(
            title="Select Bank Statement",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file:
            selected_excel = file
            self.lbl_file.config(text=f"File: {Path(file).name}")
            self.btn_process.config(state=tk.NORMAL)

    def clear_data(self):
        global selected_excel, processed_results
        selected_excel = ""
        processed_results.clear()
        self.lbl_file.config(text="Bank Statement: None selected")
        self.btn_process.config(state=tk.DISABLED)
        self.btn_export.config(state=tk.DISABLED)
        self.tree.delete(*self.tree.get_children())
        self.lbl_status.config(text="Ready.")
        self.lbl_summary.config(text="")

    def run_process(self):
        if not selected_excel:
            messagebox.showerror("Error", "Please select an Excel file first.")
            return

        self.btn_process.config(state=tk.DISABLED)
        self.btn_export.config(state=tk.DISABLED)
        self.tree.delete(*self.tree.get_children())
        self.lbl_status.config(text="Processing...")

        threading.Thread(target=self._process_thread, daemon=True).start()

    def _process_thread(self):
        success, err_msg = process_bank_statement(selected_excel)

        if not success:
            self.after(0, lambda: messagebox.showerror(
                "Processing Error", err_msg
            ))
            self.after(0, lambda: self.lbl_status.config(text="Failed."))
        else:
            self.after(0, self._populate_tree)

        self.after(0, lambda: self.btn_process.config(state=tk.NORMAL))

    def _populate_tree(self):
        match_count = 0
        suspense_count = 0
        total_amount = 0.0

        # Calculate amount and date counts to detect duplicates accurately (Option A)
        from collections import Counter
        duplicate_keys = [
            (row.get("_Date", ""), row["Amount"])
            for row in processed_results
        ]
        amount_counts = Counter(duplicate_keys)

        for i, row in enumerate(processed_results):
            key = (row.get("_Date", ""), row["Amount"])
            is_dup = amount_counts[key] > 1
            
            # Determine appropriate tag based on even/odd and status
            suffix = 'even' if i % 2 == 0 else 'odd'
            if is_dup:
                tag = f'duplicate_{suffix}'
            else:
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'

            self.tree.insert("", tk.END, iid=str(i), values=(
                row.get("_Date", ""),
                row.get("_RefNo", ""),
                row.get("Payee Name", ""),
                f'{row["Amount"]:,.2f}',
                row.get("_ModeOfPayment", ""),
                row["_Flag"]
            ), tags=(tag,))

            total_amount += row["Amount"]
            if row["_Flag"] == "✓":
                match_count += 1
            else:
                suspense_count += 1

        self.lbl_status.config(
            text=f"Done. Extracted {len(processed_results)} withdrawal "
                 f"transactions."
        )
        self.lbl_summary.config(
            text=f"✓ Matched: {match_count}   |   "
                 f"⚠ Suspense: {suspense_count}"
        )

        if processed_results:
            self.btn_export.config(state=tk.NORMAL)

    def on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)

        if not row_id or not column:
            return

        x, y, width, height = self.tree.bbox(row_id, column)
        col_idx = int(column[1:]) - 1
        col_name = self.tree["columns"][col_idx]
        current_val = self.tree.item(row_id, "values")[col_idx]

        entry = ttk.Entry(self.tree)
        entry.insert(0, current_val)
        entry.select_range(0, tk.END)
        entry.focus_set()
        entry.place(x=x, y=y, width=width, height=height)

        def save_edit(event_obj=None):
            new_val = entry.get().strip()
            entry.destroy()

            if new_val == current_val:
                return

            global processed_results
            row_idx = int(row_id)
            if row_idx >= len(processed_results):
                return

            row_data = processed_results[row_idx]

            if col_name == "Date":
                row_data["_Date"] = new_val
                row_data["Entry Date"] = new_val
                row_data["Posting Date"] = new_val
                row_data["Payment Ref. Date"] = new_val
            elif col_name == "Chq./Ref.No.":
                row_data["_RefNo"] = new_val
                row_data["Payment Ref No."] = new_val
            elif col_name == "Payee Name":
                row_data["Payee Name"] = new_val
                # Check if this payee name is known to mapGL
                is_known = False
                matched_gl = "SUSPENSE - V"
                for keyword, full_name, gl_name in PAYEE_MAPPINGS:
                    if keyword.upper() in new_val.upper() or new_val.upper() == full_name.upper():
                        is_known = True
                        matched_gl = gl_name
                        break
                row_data["Charge/GL Name"] = matched_gl
                row_data["_Flag"] = "✓" if is_known else "⚠"
                row_data["Narration"] = f"BEING AMOUNT PAID TO {new_val}"
            elif col_name == "Withdrawal Amt":
                try:
                    amt_str = new_val.replace(",", "")
                    amt = float(amt_str)
                    row_data["Amount"] = amt
                    row_data["Charge/GL Amount"] = amt
                except ValueError:
                    messagebox.showerror("Invalid Amount", "Please enter a valid number.")
                    return
            elif col_name == "Mode":
                row_data["_ModeOfPayment"] = new_val
                row_data["Mode Of Payment"] = new_val
            elif col_name == "Flag":
                row_data["_Flag"] = new_val

            self.tree.delete(*self.tree.get_children())
            self._populate_tree()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", lambda e: save_edit())
        entry.bind("<Escape>", lambda e: entry.destroy())

    def delete_selected_row(self, event):
        selected_items = self.tree.selection()
        if not selected_items:
            return

        if messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete the selected {len(selected_items)} row(s)?"
        ):
            global processed_results
            indices_to_delete = sorted([int(item) for item in selected_items], reverse=True)
            for idx in indices_to_delete:
                if 0 <= idx < len(processed_results):
                    processed_results.pop(idx)

            # Recalculate Record IDs
            for i, row in enumerate(processed_results, start=1):
                row["Record ID"] = i

            self.tree.delete(*self.tree.get_children())
            self._populate_tree()

    def show_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            self.context_menu.post(event.x_root, event.y_root)

    def delete_selected_row_context(self):
        self.delete_selected_row(None)



    def export_csv(self):
        if not processed_results:
            return

        # Check for duplicate transactions (Date + Amount)
        from collections import Counter
        duplicate_keys = [
            (row.get("_Date", ""), row["Amount"])
            for row in processed_results
        ]
        amount_counts = Counter(duplicate_keys)
        has_duplicates = any(count > 1 for count in amount_counts.values())

        if has_duplicates:
            ans = messagebox.askyesno(
                "Duplicate Amounts Detected",
                "There are duplicate amounts in the statement.\n"
                "Do you still want to export the CSV?",
                icon=messagebox.WARNING
            )
            if not ans:
                return

        export_dir = os.path.join(os.path.abspath("."), "CSV Output")
        os.makedirs(export_dir, exist_ok=True)

        fpath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialdir=export_dir,
            initialfile=(
                f"Bank_Withdrawals_"
                f"{datetime.datetime.now().strftime('%d%b%y_%H%M%S')}.csv"
            ),
            title="Export CSV"
        )

        if not fpath:
            return

        # Standard columns in Logisys order
        csv_headers = [
            "Record ID", "Entry Date", "Posting Date",
            "Cash / Bank Name", "Amount", "Currency", "Ex. Rate",
            "Mode Of Payment", "Payment Ref No.", "Payment Ref. Date",
            "Payee Name", "Payee Type", "Vendor Ref. No",
            "Vendor Ref. Date", "Narration", "Charge / GL",
            "Charge/GL Name", "Charge/GL Amount", "DR/CR",
            "Receipt Entry - Receipt No.", "Receipt Entry - Date",
            "Receipt Entry - Amount", "Cost Center", "Branch",
            "Charge Narration", "TaxGroup", "Tax Type", "SAC/HSN",
            "Taxcode1", "Taxcode1 Amt", "Taxcode2", "Taxcode2 Amt",
            "Taxcode3", "Taxcode3 Amt", "Taxcode4", "Taxcode4 Amt",
            "Avail Tax Credit", "LOB", "Ref. Type", "Ref. No.",
            "Ref. Amount", "WH Tax Organization", "WH Tax Code",
            "WH Tax Percentage", "WH Tax Taxable", "Start Date",
            "End Date", "CC Code"
        ]

        try:
            with open(fpath, 'w', newline='', encoding='utf-8') as f:
                # extrasaction='ignore' skips internal _Flag, _Date etc.
                writer = csv.DictWriter(
                    f, fieldnames=csv_headers, extrasaction='ignore'
                )
                writer.writeheader()
                for row in processed_results:
                    writer.writerow(row)

            messagebox.showinfo(
                "Success",
                f"Exported {len(processed_results)} withdrawals to:\n{fpath}"
            )
        except Exception as e:
            messagebox.showerror("Export Error", str(e))


if __name__ == "__main__":
    # Quick dependency check
    try:
        import xlrd  # noqa: F401
    except ImportError:
        print(
            "WARNING: 'xlrd' is not installed. "
            "Legacy .xls files will not work.\n"
            "Install with: pip install xlrd"
        )

    app = App()
    app.mainloop()