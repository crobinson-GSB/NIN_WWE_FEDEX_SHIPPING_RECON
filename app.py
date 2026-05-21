import streamlit as st
import pandas as pd
import re
import io
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="GSB Digital — Shipping Reconciliation",
    page_icon="📦",
    layout="centered"
)

st.markdown("""
<style>
    .stApp { background-color: #f9f9f9; }
    .block-container { max-width: 760px; padding-top: 2rem; }
    h1 { color: #E8601C !important; font-size: 1.6rem !important; }
    h3 { color: #2B2B2B !important; font-size: 1rem !important; font-weight: 600 !important; }
    p { color: #2B2B2B !important; }
    div[data-testid="stFileUploader"] {
        background-color: #ffffff !important;
        border: 1.5px dashed #CCCCCC !important;
        border-radius: 8px !important;
        padding: 0.5rem !important;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #E8601C !important;
    }
    div[data-testid="stFileUploader"] label {
        font-weight: 600;
        color: #2B2B2B !important;
    }
    div[data-testid="stFileUploader"] small {
        color: #777777 !important;
    }
    div[data-testid="stMarkdownContainer"] h3 {
        border-left: 4px solid #E8601C;
        padding-left: 10px;
    }
    div[data-testid="stExpander"] {
        border: 2px solid #E8601C !important;
        border-radius: 6px !important;
    }
    div[data-testid="stButton"] {
        display: flex;
        justify-content: center;
    }
    .stButton button {
        background-color: #E8601C !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 0.75rem 2rem !important;
        border-radius: 6px !important;
        width: 100% !important;
        font-size: 1rem !important;
    }
    .stButton button:hover { background-color: #c94f14 !important; }
    div[data-testid="stAlert"] {
        background-color: #FDF3EE !important;
        border: 1px solid #E8601C !important;
        border-radius: 6px !important;
        color: #2B2B2B !important;
    }
    div[data-testid="stFileUploader"] section {
        background-color: #2B2B2B !important;
        border-color: #444444 !important;
    }
    div[data-testid="stFileUploader"] section:hover {
        border-color: #E8601C !important;
    }
    div[data-testid="stFileUploader"] section button {
        background-color: #E8601C !important;
        color: white !important;
        border: none !important;
    }
    div[data-testid="stFileUploader"] section button:hover {
        background-color: #c94f14 !important;
    }
    div[data-testid="stFileUploader"] section span {
        color: #AAAAAA !important;
    }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div style='padding-top: 2rem;'>", unsafe_allow_html=True)
st.image("https://gsbdigital.com/wp-content/uploads/2018/06/GSB_DigitalLogo-2015-300x44.png", width=200)
st.markdown("</div>", unsafe_allow_html=True)
st.markdown("# Shipping Reconciliation Tool")
st.markdown("Automatically match Vision exports against vendor invoices and identify shipping cost discrepancies.")

with st.expander("How to use this tool", expanded=False):
    st.markdown("""
**What this tool does**
Compares your Printsmith Vision shipping export against vendor invoices from NIN (courier) and WWE/UPS. It cleans and matches invoice numbers automatically, then flags any discrepancies between what was recorded in Vision and what the vendor actually charged.

**Step 1 — Export from Vision**
In Printsmith Vision, run your Digital Shipping report for the billing period. Export it as a .txt or .xlsx file.

**Step 2 — Download vendor invoices**
- **NIN:** Download your courier invoice from the NIN portal as an .xls file
- **WWE/UPS:** Download your UPS invoice from the WWE portal as an .xls file
- **FedEx:** Download your FedEx invoice if applicable

**Step 3 — Set the billing period**
Enter the start and end dates for the period you are reconciling. This will be used to name the downloaded report.

**Step 4 — Upload the files**
Drop each file into the correct upload box below. The Vision export and at least one vendor invoice (NIN or WWE) are required.

**Step 5 — Run the reconciliation**
Click "Run Reconciliation." The tool will process the files and show a summary of results.

**Step 6 — Download the report**
Click the download button to get your Excel report. It includes five tabs:
- **Summary** — overall counts, full financial totals, and mismatch impact broken out separately
- **Mismatches** — invoices where Vision cost and vendor cost don't match (action required)
- **Matched** — invoices that reconciled cleanly
- **Not in Vision** — vendor charges with no matching Vision entry
- **Vision Only** — Vision entries with no vendor invoice received yet

**Tips**
- Invoice numbers are matched automatically — no manual cleaning needed
- Run this each billing cycle when vendor invoices arrive
- The Mismatches tab is your primary action list
""")

st.markdown("<div style='margin-top: 0.25rem;'></div>", unsafe_allow_html=True)

# --- Billing period date range ---
st.markdown("### Billing Period")
dcol1, dcol2 = st.columns(2)
with dcol1:
    period_start = st.date_input("Period Start", value=None, key="period_start")
with dcol2:
    period_end = st.date_input("Period End", value=None, key="period_end")

st.markdown("### Vision Export")
vision_file = st.file_uploader(
    "Vision Report (.txt or .xlsx export from Printsmith Vision)",
    type=["txt", "xlsx"],
    key="vision"
)

st.markdown("### Vendor Invoices")
col1, col2 = st.columns(2)
with col1:
    nin_file = st.file_uploader("NIN — Courier (.xls or .xlsx)", type=["xls", "xlsx"], key="nin")
with col2:
    wwe_file = st.file_uploader("WWE — UPS (.xls or .xlsx)", type=["xls", "xlsx"], key="wwe")

col3, col4 = st.columns(2)
with col3:
    fedex_file = st.file_uploader("FedEx (optional)", type=["xls", "xlsx", "csv"], key="fedex")
with col4:
    extra_file = st.file_uploader("Additional vendor (optional)", type=["xls", "xlsx", "csv"], key="extra")

st.markdown("<hr style='border: none; border-top: 1px solid #DDDDDD; margin: 1.5rem 0;'>", unsafe_allow_html=True)

# --- Helpers ---

def clean_key(val):
    if pd.isna(val):
        return ''
    s = str(val).strip()
    s = re.sub(r'\.0$', '', s)
    s = re.sub(r'^[Dd]', '', s)
    s = s.split('/')[0]
    s = s.replace('-', '').replace(' ', '').strip()
    return s.upper()

def read_excel_any(file):
    """Read .xls or .xlsx without assuming engine."""
    if file.name.lower().endswith('.xlsx'):
        return pd.read_excel(file)
    else:
        return pd.read_excel(file, engine='xlrd')

REQUIRED_VISION  = {'Invoice', 'Cost', 'Amount', 'Sales Rep', 'Description', 'Pickup Date'}
REQUIRED_NIN     = {'Auth', 'AmountCharged', 'InvoiceNumber', 'OrderNumber'}
REQUIRED_WWE     = {'Billing Reference 1', 'Charge Total', 'Invoice #', 'Airbill #'}

def validate_columns(df, required, label):
    missing = required - set(df.columns)
    if missing:
        st.error(
            f"The file uploaded for **{label}** doesn't look right. "
            f"Expected columns not found: {', '.join(sorted(missing))}. "
            f"Make sure you're uploading the correct file in the correct box."
        )
        return False
    return True

def load_vision(file):
    if file.name.endswith('.txt'):
        try:
            df = pd.read_csv(file, sep='\t', encoding='cp1252')
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(file, sep='\t')
    else:
        df = pd.read_excel(file)
    if not validate_columns(df, REQUIRED_VISION, 'Vision Export'):
        return None
    df['Cost'] = pd.to_numeric(df['Cost'], errors='coerce').fillna(0)
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
    df['key'] = df['Invoice'].apply(clean_key)
    return df.groupby('key').agg(
        vision_invoice=('Invoice', 'first'),
        vision_sales_rep=('Sales Rep', 'first'),
        vision_billed=('Amount', 'sum'),
        vision_cost=('Cost', 'sum'),
        vision_description=('Description', 'first'),
        vision_date=('Pickup Date', 'first')
    ).reset_index()

def load_nin(file):
    df = read_excel_any(file)
    if not validate_columns(df, REQUIRED_NIN, 'NIN — Courier'):
        return None
    df['AmountCharged'] = pd.to_numeric(df['AmountCharged'], errors='coerce').fillna(0)
    df['key'] = df['Auth'].apply(clean_key)
    return df.groupby('key').agg(
        nin_invoice_num=('InvoiceNumber', 'first'),
        nin_actual_cost=('AmountCharged', 'sum'),
        nin_shipments=('OrderNumber', 'count')
    ).reset_index()

def load_wwe(file):
    df = read_excel_any(file)
    if not validate_columns(df, REQUIRED_WWE, 'WWE — UPS'):
        return None
    df['Charge Total'] = pd.to_numeric(df['Charge Total'], errors='coerce').fillna(0)
    df['key'] = df['Billing Reference 1'].apply(clean_key)
    return df.groupby('key').agg(
        wwe_invoice_num=('Invoice #', 'first'),
        wwe_actual_cost=('Charge Total', 'sum'),
        wwe_shipments=('Airbill #', 'count')
    ).reset_index()

def assign_status(row):
    has_wwe = pd.notna(row.get('wwe_actual_cost'))
    has_nin = pd.notna(row.get('nin_actual_cost'))
    has_vis = pd.notna(row.get('vision_cost')) and row.get('vision_cost', 0) != 0
    if not has_vis and (has_wwe or has_nin):
        return "NOT IN VISION"
    if has_wwe and pd.notna(row.get('wwe_diff')) and abs(row['wwe_diff']) > 0.01:
        return "MISMATCH"
    if has_nin and pd.notna(row.get('nin_diff')) and abs(row['nin_diff']) > 0.01:
        return "MISMATCH"
    if (has_wwe or has_nin) and has_vis:
        return "MATCH"
    return "VISION ONLY"

# --- Excel styling constants ---
ORANGE    = "E8601C"
DARK      = "2B2B2B"
WHITE     = "FFFFFF"
RED_FILL  = "FFCCCC"
GREEN_FILL= "CCFFCC"
YELLOW_FILL="FFF2CC"
GRAY_FILL = "F5F5F5"
MID_GRAY  = "D9D9D9"
LIGHT_BLUE= "E8F4FD"

def tb():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(cell, bg=ORANGE):
    cell.font = Font(name="Arial", bold=True, color=WHITE, size=10)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = tb()

def body(cell, center=False, fill=None, bold=False):
    cell.font = Font(name="Arial", bold=bold, size=10)
    cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center", wrap_text=True)
    cell.border = tb()
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)

def label_cell(ws, row, col, text, bg=GRAY_FILL, bold=True, color=DARK):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name="Arial", bold=bold, size=10, color=color)
    c.fill = PatternFill("solid", fgColor=bg)
    c.border = tb()
    c.alignment = Alignment(horizontal="left", vertical="center")
    return c

def value_cell(ws, row, col, val, bold=False, color=DARK, fill=None):
    c = ws.cell(row=row, column=col, value=round(val, 2) if isinstance(val, float) else val)
    c.font = Font(name="Arial", bold=bold, size=10, color=color)
    c.number_format = '$#,##0.00'
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = tb()
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    return c

def build_excel(merged, has_nin, has_wwe, vendors_label, period_label):
    mismatches    = merged[merged['status'] == 'MISMATCH'].copy()
    matches       = merged[merged['status'] == 'MATCH'].copy()
    vision_only   = merged[merged['status'] == 'VISION ONLY'].copy()
    not_in_vision = merged[merged['status'] == 'NOT IN VISION'].copy()

    # Pre-compute financial data
    all_vendor_rows = merged[merged['status'].isin(['MATCH', 'MISMATCH'])].copy()

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False

    # ── Title banner ──
    ws.merge_cells('A1:G1')
    ws['A1'] = "GSB Digital — Shipping Reconciliation Report"
    ws['A1'].font = Font(name="Arial", bold=True, size=16, color=WHITE)
    ws['A1'].fill = PatternFill("solid", fgColor=ORANGE)
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    # ── Run info row ──
    ws.merge_cells('A2:G2')
    ws['A2'] = f"Period: {period_label}   |   Vendors included: {vendors_label}   |   Run date: {date.today().strftime('%B %d, %Y')}"
    ws['A2'].font = Font(name="Arial", italic=True, size=9, color="555555")
    ws['A2'].fill = PatternFill("solid", fgColor="F2F2F2")
    ws['A2'].alignment = Alignment(horizontal="center", vertical="center")
    ws['A2'].border = tb()
    ws.row_dimensions[2].height = 18

    # ── Stat cards ──
    stats = [
        ("MISMATCHES",    len(mismatches),    "FFCCCC"),
        ("MATCHED",       len(matches),       "CCFFCC"),
        ("NOT IN VISION", len(not_in_vision), "FFF2CC"),
        ("VISION ONLY",   len(vision_only),   LIGHT_BLUE),
    ]
    ws.row_dimensions[4].height = 50
    ws.row_dimensions[5].height = 25
    for i, (label, val, bg) in enumerate(stats, 1):
        c1 = ws.cell(row=4, column=i, value=str(val))
        c1.font = Font(name="Arial", bold=True, size=22)
        c1.fill = PatternFill("solid", fgColor=bg)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = tb()
        c2 = ws.cell(row=5, column=i, value=label)
        c2.font = Font(name="Arial", bold=True, size=9, color="555555")
        c2.fill = PatternFill("solid", fgColor=bg)
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = tb()

    # ── Section helper ──
    def section_header(ws, row, text, bg=DARK):
        ws.merge_cells(f'A{row}:G{row}')
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(name="Arial", bold=True, size=11, color=WHITE)
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 22

    def col_header_row(ws, row, labels, bg="455A64"):
        for i, lbl in enumerate(labels, 1):
            c = ws.cell(row=row, column=i, value=lbl)
            c.font = Font(name="Arial", bold=True, size=9, color=WHITE)
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = tb()
        ws.row_dimensions[row].height = 18

    # ── Section 1: ALL invoices total spend ──
    section_header(ws, 7, "TOTAL SPEND — All Matched & Reconciled Invoices")

    vendor_cols = []
    if has_nin:
        vendor_cols.append(('NIN (Courier)', 'nin_actual_cost', 'nin_diff'))
    if has_wwe:
        vendor_cols.append(('WWE / UPS', 'wwe_actual_cost', 'wwe_diff'))

    col_labels = [""] + [v[0] for v in vendor_cols]
    if has_nin and has_wwe:
        col_labels.append("Combined Total")
    col_header_row(ws, 8, col_labels)

    spend_rows = [
        ("Total Vendor Cost (all invoices)",   'actual', None),
        ("Total Recorded in Vision (all)",      'vision', None),
        ("Net Difference (all invoices)",        'diff',  None),
    ]

    for r_offset, (row_label, metric, _) in enumerate(spend_rows, 9):
        label_cell(ws, r_offset, 1, row_label)
        total = 0
        for col_i, (vname, cost_col, diff_col) in enumerate(vendor_cols, 2):
            subset = all_vendor_rows[all_vendor_rows[cost_col].notna()]
            if metric == 'actual':
                val = subset[cost_col].sum()
            elif metric == 'vision':
                val = subset['vision_cost'].sum()
            else:
                val = subset[diff_col].sum()
            total += val
            is_diff = (metric == 'diff')
            color = ("CC0000" if val < 0 else ("006600" if val > 0 else DARK)) if is_diff else DARK
            value_cell(ws, r_offset, col_i, val, bold=is_diff, color=color)
        if has_nin and has_wwe:
            is_diff = (metric == 'diff')
            color = ("CC0000" if total < 0 else ("006600" if total > 0 else DARK)) if is_diff else DARK
            value_cell(ws, r_offset, len(vendor_cols) + 2, total, bold=is_diff, color=color)
        ws.row_dimensions[r_offset].height = 18

    # ── Section 2: Mismatches only ──
    section_row = 9 + len(spend_rows) + 1
    section_header(ws, section_row, "MISMATCH IMPACT — Invoices Where Vision and Vendor Costs Differ", bg="B85C00")

    col_header_row(ws, section_row + 1, col_labels, bg="B85C00")

    mismatch_metric_rows = [
        ("Total Vendor Cost (mismatches only)",  'actual'),
        ("Total Recorded in Vision (mismatches)", 'vision'),
        ("Total Discrepancy (mismatches only)",   'diff'),
    ]

    for r_offset, (row_label, metric) in enumerate(mismatch_metric_rows, section_row + 2):
        label_cell(ws, r_offset, 1, row_label)
        total = 0
        for col_i, (vname, cost_col, diff_col) in enumerate(vendor_cols, 2):
            subset = mismatches[mismatches[cost_col].notna()]
            if metric == 'actual':
                val = subset[cost_col].sum()
            elif metric == 'vision':
                val = subset['vision_cost'].sum()
            else:
                val = subset[diff_col].sum()
            total += val
            is_diff = (metric == 'diff')
            color = ("CC0000" if val < 0 else ("006600" if val > 0 else DARK)) if is_diff else DARK
            fill = "FFF0E8" if is_diff else None
            value_cell(ws, r_offset, col_i, val, bold=is_diff, color=color, fill=fill)
        if has_nin and has_wwe:
            is_diff = (metric == 'diff')
            color = ("CC0000" if total < 0 else ("006600" if total > 0 else DARK)) if is_diff else DARK
            fill = "FFF0E8" if is_diff else None
            value_cell(ws, r_offset, len(vendor_cols) + 2, total, bold=is_diff, color=color, fill=fill)
        ws.row_dimensions[r_offset].height = 18

    for i, w in enumerate([36, 18, 18, 18, 5, 5, 5], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ══ Mismatches tab ══
    ws2 = wb.create_sheet("Mismatches")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells('A1:J1')
    ws2['A1'] = "Shipping Cost Mismatches — Action Required"
    ws2['A1'].font = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws2['A1'].fill = PatternFill("solid", fgColor=ORANGE)
    ws2['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 35
    for i, h in enumerate(["GSB Invoice #","Sales Rep","Date","Description","Vendor",
                            "Vendor Invoice #","# Shipments","Vision Cost","Actual Vendor Cost","Difference ($)"], 1):
        hdr(ws2.cell(row=2, column=i, value=h))
    for i, w in enumerate([16,14,12,35,12,18,12,18,20,16], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    row_n = 3
    for _, r in mismatches.sort_values('vision_invoice').iterrows():
        vendor_pairs = []
        if has_wwe:
            vendor_pairs.append(("WWE (UPS)", r.get('wwe_actual_cost'), r.get('wwe_diff'), r.get('wwe_invoice_num',''), r.get('wwe_shipments')))
        if has_nin:
            vendor_pairs.append(("NIN (Courier)", r.get('nin_actual_cost'), r.get('nin_diff'), r.get('nin_invoice_num',''), r.get('nin_shipments')))
        for vendor, cost, diff, inv, ships in vendor_pairs:
            if pd.notna(cost):
                fill = RED_FILL if diff < -0.01 else (YELLOW_FILL if diff > 0.01 else GREEN_FILL)
                row_data = [r['vision_invoice'], r['vision_sales_rep'], r['vision_date'],
                            r['vision_description'], vendor, inv,
                            int(ships) if pd.notna(ships) else '', r['vision_cost'], cost, round(diff, 2)]
                ws2.row_dimensions[row_n].height = 18
                for ci, val in enumerate(row_data, 1):
                    c = ws2.cell(row=row_n, column=ci, value=val)
                    body(c, center=(ci in [1,3,5,6,7,8,9,10]), fill=fill if ci in [8,9,10] else None)
                    if ci in [8, 9, 10]:
                        c.number_format = '$#,##0.00'
                row_n += 1
    ws2.row_dimensions[row_n].height = 22
    ws2.cell(row=row_n, column=7, value="TOTALS").font = Font(name="Arial", bold=True, size=10)
    ws2.cell(row=row_n, column=7).alignment = Alignment(horizontal="right")
    for ci in [8, 9, 10]:
        cl = get_column_letter(ci)
        c = ws2.cell(row=row_n, column=ci, value=f'=SUM({cl}3:{cl}{row_n-1})')
        c.font = Font(name="Arial", bold=True, size=10)
        c.number_format = '$#,##0.00'
        c.fill = PatternFill("solid", fgColor=MID_GRAY)
        c.border = tb()
        c.alignment = Alignment(horizontal="center")

    # ══ Matched tab ══
    ws3 = wb.create_sheet("Matched")
    ws3.sheet_view.showGridLines = False
    ws3.merge_cells('A1:I1')
    ws3['A1'] = "Verified Matches — No Action Required"
    ws3['A1'].font = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws3['A1'].fill = PatternFill("solid", fgColor="2E7D32")
    ws3['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 35
    for i, h in enumerate(["GSB Invoice #","Sales Rep","Date","Vendor","Vendor Invoice #",
                            "Vision Cost","Vendor Cost","Difference","Status"], 1):
        hdr(ws3.cell(row=2, column=i, value=h), bg="2E7D32")
    for i, w in enumerate([16,14,12,14,18,18,16,14,10], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    row_n3 = 3
    for _, r in matches.sort_values('vision_invoice').iterrows():
        vendor_pairs = []
        if has_wwe:
            vendor_pairs.append(("WWE (UPS)", r.get('wwe_actual_cost'), r.get('wwe_diff')))
        if has_nin:
            vendor_pairs.append(("NIN (Courier)", r.get('nin_actual_cost'), r.get('nin_diff')))
        for vendor, cost, diff in vendor_pairs:
            if pd.notna(cost):
                ws3.row_dimensions[row_n3].height = 18
                for ci, val in enumerate([r['vision_invoice'], r['vision_sales_rep'], r['vision_date'],
                                          vendor, '', r['vision_cost'], cost, round(diff, 2), "✓"], 1):
                    c = ws3.cell(row=row_n3, column=ci, value=val)
                    body(c, center=(ci in [1,3,4,5,6,7,8,9]), fill=GREEN_FILL if ci == 9 else None)
                    if ci in [6, 7, 8]:
                        c.number_format = '$#,##0.00'
                row_n3 += 1

    # ══ Not in Vision tab ══
    ws4 = wb.create_sheet("Not in Vision")
    ws4.sheet_view.showGridLines = False
    ws4.merge_cells('A1:H1')
    ws4['A1'] = "Vendor Charges Not Found in Vision — Review Required"
    ws4['A1'].font = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws4['A1'].fill = PatternFill("solid", fgColor="B8860B")
    ws4['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 35
    for i, h in enumerate(["Reference Key","Vendor","Vendor Invoice #","# Shipments",
                            "Vendor Cost","Notes","Resolved?"], 1):
        hdr(ws4.cell(row=2, column=i, value=h), bg="B8860B")
    for i, w in enumerate([18,14,18,12,14,35,12], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w
    row_n4 = 3
    for _, r in not_in_vision.sort_values('key').iterrows():
        vendor_pairs = []
        if has_wwe:
            vendor_pairs.append(("WWE (UPS)", r.get('wwe_actual_cost'), r.get('wwe_shipments'), r.get('wwe_invoice_num','')))
        if has_nin:
            vendor_pairs.append(("NIN (Courier)", r.get('nin_actual_cost'), r.get('nin_shipments'), r.get('nin_invoice_num','')))
        for vendor, cost, ships, inv in vendor_pairs:
            if pd.notna(cost):
                ws4.row_dimensions[row_n4].height = 18
                for ci, val in enumerate([r['key'], vendor, inv,
                                          int(ships) if pd.notna(ships) else '',
                                          cost, "Invoice not matched in Vision", ""], 1):
                    c = ws4.cell(row=row_n4, column=ci, value=val)
                    body(c, center=(ci in [2,3,4,5,7]), fill=YELLOW_FILL if ci in [1,2,5] else None)
                    if ci == 5:
                        c.number_format = '$#,##0.00'
                row_n4 += 1

    # ══ Vision Only tab ══
    ws5 = wb.create_sheet("Vision Only")
    ws5.sheet_view.showGridLines = False
    ws5.merge_cells('A1:G1')
    ws5['A1'] = "Vision Entries with No Vendor Invoice — For Reference"
    ws5['A1'].font = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws5['A1'].fill = PatternFill("solid", fgColor="455A64")
    ws5['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 35
    for i, h in enumerate(["GSB Invoice #","Sales Rep","Date","Description",
                            "Amount Billed","Cost Recorded","Notes"], 1):
        hdr(ws5.cell(row=2, column=i, value=h), bg="455A64")
    for i, w in enumerate([16,14,12,40,16,16,30], 1):
        ws5.column_dimensions[get_column_letter(i)].width = w
    row_n5 = 3
    for _, r in vision_only.sort_values('vision_invoice').iterrows():
        ws5.row_dimensions[row_n5].height = 18
        for ci, val in enumerate([r['vision_invoice'], r['vision_sales_rep'], r['vision_date'],
                                   r['vision_description'], r['vision_billed'], r['vision_cost'], ""], 1):
            c = ws5.cell(row=row_n5, column=ci, value=val)
            body(c, center=(ci in [1,2,3,5,6]))
            if ci in [5, 6]:
                c.number_format = '$#,##0.00'
        row_n5 += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Ready check ──
ready = vision_file and (nin_file or wwe_file)

if st.button("Run Reconciliation", disabled=not ready):
    with st.spinner("Matching invoices and calculating discrepancies..."):
        try:
            has_nin = nin_file is not None
            has_wwe = wwe_file is not None

            vision_agg = load_vision(vision_file)
            if vision_agg is None:
                st.stop()

            merged = vision_agg.copy()

            if has_wwe:
                wwe_agg = load_wwe(wwe_file)
                if wwe_agg is None:
                    st.stop()
                merged = merged.merge(wwe_agg, on='key', how='outer')

            if has_nin:
                nin_agg = load_nin(nin_file)
                if nin_agg is None:
                    st.stop()
                merged = merged.merge(nin_agg, on='key', how='outer')

            merged = merged[merged['key'] != ''].copy()

            if has_wwe:
                merged['wwe_diff'] = (merged['wwe_actual_cost'] - merged['vision_cost']).round(2)
            if has_nin:
                merged['nin_diff'] = (merged['nin_actual_cost'] - merged['vision_cost']).round(2)

            merged['status'] = merged.apply(assign_status, axis=1)

            mismatches    = merged[merged['status'] == 'MISMATCH']
            matches       = merged[merged['status'] == 'MATCH']
            not_in_vision = merged[merged['status'] == 'NOT IN VISION']
            vision_only   = merged[merged['status'] == 'VISION ONLY']

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Mismatches",    len(mismatches))
            col2.metric("Matched",       len(matches))
            col3.metric("Not in Vision", len(not_in_vision))
            col4.metric("Vision Only",   len(vision_only))

            # Build labels for filename and Excel
            vendor_parts = []
            if has_nin: vendor_parts.append("NIN")
            if has_wwe: vendor_parts.append("WWE")
            vendors_label = " + ".join(vendor_parts)

            if period_start and period_end:
                period_label = f"{period_start.strftime('%b %d %Y')} – {period_end.strftime('%b %d %Y')}"
                period_slug  = f"{period_start.strftime('%Y%m%d')}_to_{period_end.strftime('%Y%m%d')}"
            else:
                period_label = "Period not specified"
                period_slug  = date.today().strftime('%Y%m%d')

            filename = f"GSB_Shipping_Recon_{period_slug}.xlsx"

            excel_buf = build_excel(merged, has_nin, has_wwe, vendors_label, period_label)
            st.success("Reconciliation complete. Download your report below.")
            st.download_button(
                label="Download Reconciliation Report (.xlsx)",
                data=excel_buf,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.info("Make sure your files match the expected format and try again.")

elif not ready:
    if not vision_file:
        st.info("Upload the Vision export to get started. Then add at least one vendor invoice (NIN or WWE).")
    else:
        st.info("Vision file uploaded. Now add at least one vendor invoice (NIN or WWE) to run the reconciliation.")        background-color: #E8601C !important;
        color: white !important;
        border: none !important;
    }
    div[data-testid="stFileUploader"] section button:hover {
        background-color: #c94f14 !important;
    }
    div[data-testid="stFileUploader"] section span {
        color: #AAAAAA !important;
    }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div style='padding-top: 2rem;'>", unsafe_allow_html=True)
st.image("https://gsbdigital.com/wp-content/uploads/2018/06/GSB_DigitalLogo-2015-300x44.png", width=200)
st.markdown("</div>", unsafe_allow_html=True)
st.markdown("# Shipping Reconciliation Tool")
st.markdown("Automatically match Vision exports against vendor invoices and identify shipping cost discrepancies.")

with st.expander("How to use this tool", expanded=False):
    st.markdown("""
**What this tool does**
Compares your Printsmith Vision shipping export against vendor invoices from NIN (courier) and WWE/UPS. It cleans and matches invoice numbers automatically, then flags any discrepancies between what was recorded in Vision and what the vendor actually charged.

**Step 1 — Export from Vision**
In Printsmith Vision, run your Digital Shipping report for the billing period. Export it as a .txt or .xlsx file.

**Step 2 — Download vendor invoices**
- **NIN:** Download your courier invoice from the NIN portal as an .xls file
- **WWE/UPS:** Download your UPS invoice from the WWE portal as an .xls file
- **FedEx:** Download your FedEx invoice if applicable

**Step 3 — Set the billing period**
Enter the start and end dates for the period you are reconciling. This will be used to name the downloaded report.

**Step 4 — Upload the files**
Drop each file into the correct upload box below. The Vision export and at least one vendor invoice (NIN or WWE) are required.

**Step 5 — Run the reconciliation**
Click "Run Reconciliation." The tool will process the files and show a summary of results.

**Step 6 — Download the report**
Click the download button to get your Excel report. It includes five tabs:
- **Summary** — overall counts, full financial totals, and mismatch impact broken out separately
- **Mismatches** — invoices where Vision cost and vendor cost don't match (action required)
- **Matched** — invoices that reconciled cleanly
- **Not in Vision** — vendor charges with no matching Vision entry
- **Vision Only** — Vision entries with no vendor invoice received yet

**Tips**
- Invoice numbers are matched automatically — no manual cleaning needed
- Run this each billing cycle when vendor invoices arrive
- The Mismatches tab is your primary action list
""")

st.markdown("<div style='margin-top: 0.25rem;'></div>", unsafe_allow_html=True)

# --- Billing period date range ---
st.markdown("### Billing Period")
dcol1, dcol2 = st.columns(2)
with dcol1:
    period_start = st.date_input("Period Start", value=None, key="period_start")
with dcol2:
    period_end = st.date_input("Period End", value=None, key="period_end")

st.markdown("### Vision Export")
vision_file = st.file_uploader(
    "Vision Report (.txt or .xlsx export from Printsmith Vision)",
    type=["txt", "xlsx"],
    key="vision"
)

st.markdown("### Vendor Invoices")
col1, col2 = st.columns(2)
with col1:
    nin_file = st.file_uploader("NIN — Courier (.xls or .xlsx)", type=["xls", "xlsx"], key="nin")
with col2:
    wwe_file = st.file_uploader("WWE — UPS (.xls or .xlsx)", type=["xls", "xlsx"], key="wwe")

col3, col4 = st.columns(2)
with col3:
    fedex_file = st.file_uploader("FedEx (optional)", type=["xls", "xlsx", "csv"], key="fedex")
with col4:
    extra_file = st.file_uploader("Additional vendor (optional)", type=["xls", "xlsx", "csv"], key="extra")

st.markdown("<hr style='border: none; border-top: 1px solid #DDDDDD; margin: 1.5rem 0;'>", unsafe_allow_html=True)

# --- Helpers ---

def clean_key(val):
    if pd.isna(val):
        return ''
    s = str(val).strip()
    s = re.sub(r'\.0$', '', s)
    s = re.sub(r'^[Dd]', '', s)
    s = s.split('/')[0]
    s = s.replace('-', '').replace(' ', '').strip()
    return s.upper()

def read_excel_any(file):
    """Read .xls or .xlsx without assuming engine."""
    if file.name.lower().endswith('.xlsx'):
        return pd.read_excel(file)
    else:
        return pd.read_excel(file, engine='xlrd')

REQUIRED_VISION  = {'Invoice', 'Cost', 'Amount', 'Sales Rep', 'Description', 'Pickup Date'}
REQUIRED_NIN     = {'Auth', 'AmountCharged', 'InvoiceNumber', 'OrderNumber'}
REQUIRED_WWE     = {'Billing Reference 1', 'Charge Total', 'Invoice #', 'Airbill #'}

def validate_columns(df, required, label):
    missing = required - set(df.columns)
    if missing:
        st.error(
            f"The file uploaded for **{label}** doesn't look right. "
            f"Expected columns not found: {', '.join(sorted(missing))}. "
            f"Make sure you're uploading the correct file in the correct box."
        )
        return False
    return True

def load_vision(file):
    if file.name.endswith('.txt'):
        try:
            df = pd.read_csv(file, sep='\t', encoding='cp1252')
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(file, sep='\t')
    else:
        df = pd.read_excel(file)
    if not validate_columns(df, REQUIRED_VISION, 'Vision Export'):
        return None
    df['Cost'] = pd.to_numeric(df['Cost'], errors='coerce').fillna(0)
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
    df['key'] = df['Invoice'].apply(clean_key)
    return df.groupby('key').agg(
        vision_invoice=('Invoice', 'first'),
        vision_sales_rep=('Sales Rep', 'first'),
        vision_billed=('Amount', 'sum'),
        vision_cost=('Cost', 'sum'),
        vision_description=('Description', 'first'),
        vision_date=('Pickup Date', 'first')
    ).reset_index()

def load_nin(file):
    df = read_excel_any(file)
    if not validate_columns(df, REQUIRED_NIN, 'NIN — Courier'):
        return None
    df['AmountCharged'] = pd.to_numeric(df['AmountCharged'], errors='coerce').fillna(0)
    df['key'] = df['Auth'].apply(clean_key)
    return df.groupby('key').agg(
        nin_invoice_num=('InvoiceNumber', 'first'),
        nin_actual_cost=('AmountCharged', 'sum'),
        nin_shipments=('OrderNumber', 'count')
    ).reset_index()

def load_wwe(file):
    df = read_excel_any(file)
    if not validate_columns(df, REQUIRED_WWE, 'WWE — UPS'):
        return None
    df['Charge Total'] = pd.to_numeric(df['Charge Total'], errors='coerce').fillna(0)
    df['key'] = df['Billing Reference 1'].apply(clean_key)
    return df.groupby('key').agg(
        wwe_invoice_num=('Invoice #', 'first'),
        wwe_actual_cost=('Charge Total', 'sum'),
        wwe_shipments=('Airbill #', 'count')
    ).reset_index()

def assign_status(row):
    has_wwe = pd.notna(row.get('wwe_actual_cost'))
    has_nin = pd.notna(row.get('nin_actual_cost'))
    has_vis = pd.notna(row.get('vision_cost')) and row.get('vision_cost', 0) != 0
    if not has_vis and (has_wwe or has_nin):
        return "NOT IN VISION"
    if has_wwe and pd.notna(row.get('wwe_diff')) and abs(row['wwe_diff']) > 0.01:
        return "MISMATCH"
    if has_nin and pd.notna(row.get('nin_diff')) and abs(row['nin_diff']) > 0.01:
        return "MISMATCH"
    if (has_wwe or has_nin) and has_vis:
        return "MATCH"
    return "VISION ONLY"

# --- Excel styling constants ---
ORANGE    = "E8601C"
DARK      = "2B2B2B"
WHITE     = "FFFFFF"
RED_FILL  = "FFCCCC"
GREEN_FILL= "CCFFCC"
YELLOW_FILL="FFF2CC"
GRAY_FILL = "F5F5F5"
MID_GRAY  = "D9D9D9"
LIGHT_BLUE= "E8F4FD"

def tb():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(cell, bg=ORANGE):
    cell.font = Font(name="Arial", bold=True, color=WHITE, size=10)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = tb()

def body(cell, center=False, fill=None, bold=False):
    cell.font = Font(name="Arial", bold=bold, size=10)
    cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center", wrap_text=True)
    cell.border = tb()
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)

def label_cell(ws, row, col, text, bg=GRAY_FILL, bold=True, color=DARK):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name="Arial", bold=bold, size=10, color=color)
    c.fill = PatternFill("solid", fgColor=bg)
    c.border = tb()
    c.alignment = Alignment(horizontal="left", vertical="center")
    return c

def value_cell(ws, row, col, val, bold=False, color=DARK, fill=None):
    c = ws.cell(row=row, column=col, value=round(val, 2) if isinstance(val, float) else val)
    c.font = Font(name="Arial", bold=bold, size=10, color=color)
    c.number_format = '$#,##0.00'
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = tb()
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    return c

def build_excel(merged, has_nin, has_wwe, vendors_label, period_label):
    mismatches    = merged[merged['status'] == 'MISMATCH'].copy()
    matches       = merged[merged['status'] == 'MATCH'].copy()
    vision_only   = merged[merged['status'] == 'VISION ONLY'].copy()
    not_in_vision = merged[merged['status'] == 'NOT IN VISION'].copy()

    # Pre-compute financial data
    all_vendor_rows = merged[merged['status'].isin(['MATCH', 'MISMATCH'])].copy()

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False

    # ── Title banner ──
    ws.merge_cells('A1:G1')
    ws['A1'] = "GSB Digital — Shipping Reconciliation Report"
    ws['A1'].font = Font(name="Arial", bold=True, size=16, color=WHITE)
    ws['A1'].fill = PatternFill("solid", fgColor=ORANGE)
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    # ── Run info row ──
    ws.merge_cells('A2:G2')
    ws['A2'] = f"Period: {period_label}   |   Vendors included: {vendors_label}   |   Run date: {date.today().strftime('%B %d, %Y')}"
    ws['A2'].font = Font(name="Arial", italic=True, size=9, color="555555")
    ws['A2'].fill = PatternFill("solid", fgColor="F2F2F2")
    ws['A2'].alignment = Alignment(horizontal="center", vertical="center")
    ws['A2'].border = tb()
    ws.row_dimensions[2].height = 18

    # ── Stat cards ──
    stats = [
        ("MISMATCHES",    len(mismatches),    "FFCCCC"),
        ("MATCHED",       len(matches),       "CCFFCC"),
        ("NOT IN VISION", len(not_in_vision), "FFF2CC"),
        ("VISION ONLY",   len(vision_only),   LIGHT_BLUE),
    ]
    ws.row_dimensions[4].height = 50
    ws.row_dimensions[5].height = 25
    for i, (label, val, bg) in enumerate(stats, 1):
        c1 = ws.cell(row=4, column=i, value=str(val))
        c1.font = Font(name="Arial", bold=True, size=22)
        c1.fill = PatternFill("solid", fgColor=bg)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = tb()
        c2 = ws.cell(row=5, column=i, value=label)
        c2.font = Font(name="Arial", bold=True, size=9, color="555555")
        c2.fill = PatternFill("solid", fgColor=bg)
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = tb()

    # ── Section helper ──
    def section_header(ws, row, text, bg=DARK):
        ws.merge_cells(f'A{row}:G{row}')
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(name="Arial", bold=True, size=11, color=WHITE)
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 22

    def col_header_row(ws, row, labels, bg="455A64"):
        for i, lbl in enumerate(labels, 1):
            c = ws.cell(row=row, column=i, value=lbl)
            c.font = Font(name="Arial", bold=True, size=9, color=WHITE)
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = tb()
        ws.row_dimensions[row].height = 18

    # ── Section 1: ALL invoices total spend ──
    section_header(ws, 7, "TOTAL SPEND — All Matched & Reconciled Invoices")

    vendor_cols = []
    if has_nin:
        vendor_cols.append(('NIN (Courier)', 'nin_actual_cost', 'nin_diff'))
    if has_wwe:
        vendor_cols.append(('WWE / UPS', 'wwe_actual_cost', 'wwe_diff'))

    col_labels = [""] + [v[0] for v in vendor_cols]
    if has_nin and has_wwe:
        col_labels.append("Combined Total")
    col_header_row(ws, 8, col_labels)

    spend_rows = [
        ("Total Vendor Cost (all invoices)",   'actual', None),
        ("Total Recorded in Vision (all)",      'vision', None),
        ("Net Difference (all invoices)",        'diff',  None),
    ]

    for r_offset, (row_label, metric, _) in enumerate(spend_rows, 9):
        label_cell(ws, r_offset, 1, row_label)
        total = 0
        for col_i, (vname, cost_col, diff_col) in enumerate(vendor_cols, 2):
            subset = all_vendor_rows[all_vendor_rows[cost_col].notna()]
            if metric == 'actual':
                val = subset[cost_col].sum()
            elif metric == 'vision':
                val = subset['vision_cost'].sum()
            else:
                val = subset[diff_col].sum()
            total += val
            is_diff = (metric == 'diff')
            color = ("CC0000" if val < 0 else ("006600" if val > 0 else DARK)) if is_diff else DARK
            value_cell(ws, r_offset, col_i, val, bold=is_diff, color=color)
        if has_nin and has_wwe:
            is_diff = (metric == 'diff')
            color = ("CC0000" if total < 0 else ("006600" if total > 0 else DARK)) if is_diff else DARK
            value_cell(ws, r_offset, len(vendor_cols) + 2, total, bold=is_diff, color=color)
        ws.row_dimensions[r_offset].height = 18

    # ── Section 2: Mismatches only ──
    section_row = 9 + len(spend_rows) + 1
    section_header(ws, section_row, "MISMATCH IMPACT — Invoices Where Vision and Vendor Costs Differ", bg="B85C00")

    col_header_row(ws, section_row + 1, col_labels, bg="B85C00")

    mismatch_metric_rows = [
        ("Total Vendor Cost (mismatches only)",  'actual'),
        ("Total Recorded in Vision (mismatches)", 'vision'),
        ("Total Discrepancy (mismatches only)",   'diff'),
    ]

    for r_offset, (row_label, metric) in enumerate(mismatch_metric_rows, section_row + 2):
        label_cell(ws, r_offset, 1, row_label)
        total = 0
        for col_i, (vname, cost_col, diff_col) in enumerate(vendor_cols, 2):
            subset = mismatches[mismatches[cost_col].notna()]
            if metric == 'actual':
                val = subset[cost_col].sum()
            elif metric == 'vision':
                val = subset['vision_cost'].sum()
            else:
                val = subset[diff_col].sum()
            total += val
            is_diff = (metric == 'diff')
            color = ("CC0000" if val < 0 else ("006600" if val > 0 else DARK)) if is_diff else DARK
            fill = "FFF0E8" if is_diff else None
            value_cell(ws, r_offset, col_i, val, bold=is_diff, color=color, fill=fill)
        if has_nin and has_wwe:
            is_diff = (metric == 'diff')
            color = ("CC0000" if total < 0 else ("006600" if total > 0 else DARK)) if is_diff else DARK
            fill = "FFF0E8" if is_diff else None
            value_cell(ws, r_offset, len(vendor_cols) + 2, total, bold=is_diff, color=color, fill=fill)
        ws.row_dimensions[r_offset].height = 18

    for i, w in enumerate([36, 18, 18, 18, 5, 5, 5], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ══ Mismatches tab ══
    ws2 = wb.create_sheet("Mismatches")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells('A1:J1')
    ws2['A1'] = "Shipping Cost Mismatches — Action Required"
    ws2['A1'].font = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws2['A1'].fill = PatternFill("solid", fgColor=ORANGE)
    ws2['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 35
    for i, h in enumerate(["GSB Invoice #","Sales Rep","Date","Description","Vendor",
                            "Vendor Invoice #","# Shipments","Vision Cost","Actual Vendor Cost","Difference ($)"], 1):
        hdr(ws2.cell(row=2, column=i, value=h))
    for i, w in enumerate([16,14,12,35,12,18,12,18,20,16], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    row_n = 3
    for _, r in mismatches.sort_values('vision_invoice').iterrows():
        vendor_pairs = []
        if has_wwe:
            vendor_pairs.append(("WWE (UPS)", r.get('wwe_actual_cost'), r.get('wwe_diff'), r.get('wwe_invoice_num',''), r.get('wwe_shipments')))
        if has_nin:
            vendor_pairs.append(("NIN (Courier)", r.get('nin_actual_cost'), r.get('nin_diff'), r.get('nin_invoice_num',''), r.get('nin_shipments')))
        for vendor, cost, diff, inv, ships in vendor_pairs:
            if pd.notna(cost):
                fill = RED_FILL if diff < -0.01 else (YELLOW_FILL if diff > 0.01 else GREEN_FILL)
                row_data = [r['vision_invoice'], r['vision_sales_rep'], r['vision_date'],
                            r['vision_description'], vendor, inv,
                            int(ships) if pd.notna(ships) else '', r['vision_cost'], cost, round(diff, 2)]
                ws2.row_dimensions[row_n].height = 18
                for ci, val in enumerate(row_data, 1):
                    c = ws2.cell(row=row_n, column=ci, value=val)
                    body(c, center=(ci in [1,3,5,6,7,8,9,10]), fill=fill if ci in [8,9,10] else None)
                    if ci in [8, 9, 10]:
                        c.number_format = '$#,##0.00'
                row_n += 1
    ws2.row_dimensions[row_n].height = 22
    ws2.cell(row=row_n, column=7, value="TOTALS").font = Font(name="Arial", bold=True, size=10)
    ws2.cell(row=row_n, column=7).alignment = Alignment(horizontal="right")
    for ci in [8, 9, 10]:
        cl = get_column_letter(ci)
        c = ws2.cell(row=row_n, column=ci, value=f'=SUM({cl}3:{cl}{row_n-1})')
        c.font = Font(name="Arial", bold=True, size=10)
        c.number_format = '$#,##0.00'
        c.fill = PatternFill("solid", fgColor=MID_GRAY)
        c.border = tb()
        c.alignment = Alignment(horizontal="center")

    # ══ Matched tab ══
    ws3 = wb.create_sheet("Matched")
    ws3.sheet_view.showGridLines = False
    ws3.merge_cells('A1:I1')
    ws3['A1'] = "Verified Matches — No Action Required"
    ws3['A1'].font = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws3['A1'].fill = PatternFill("solid", fgColor="2E7D32")
    ws3['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 35
    for i, h in enumerate(["GSB Invoice #","Sales Rep","Date","Vendor","Vendor Invoice #",
                            "Vision Cost","Vendor Cost","Difference","Status"], 1):
        hdr(ws3.cell(row=2, column=i, value=h), bg="2E7D32")
    for i, w in enumerate([16,14,12,14,18,18,16,14,10], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    row_n3 = 3
    for _, r in matches.sort_values('vision_invoice').iterrows():
        vendor_pairs = []
        if has_wwe:
            vendor_pairs.append(("WWE (UPS)", r.get('wwe_actual_cost'), r.get('wwe_diff')))
        if has_nin:
            vendor_pairs.append(("NIN (Courier)", r.get('nin_actual_cost'), r.get('nin_diff')))
        for vendor, cost, diff in vendor_pairs:
            if pd.notna(cost):
                ws3.row_dimensions[row_n3].height = 18
                for ci, val in enumerate([r['vision_invoice'], r['vision_sales_rep'], r['vision_date'],
                                          vendor, '', r['vision_cost'], cost, round(diff, 2), "✓"], 1):
                    c = ws3.cell(row=row_n3, column=ci, value=val)
                    body(c, center=(ci in [1,3,4,5,6,7,8,9]), fill=GREEN_FILL if ci == 9 else None)
                    if ci in [6, 7, 8]:
                        c.number_format = '$#,##0.00'
                row_n3 += 1

    # ══ Not in Vision tab ══
    ws4 = wb.create_sheet("Not in Vision")
    ws4.sheet_view.showGridLines = False
    ws4.merge_cells('A1:H1')
    ws4['A1'] = "Vendor Charges Not Found in Vision — Review Required"
    ws4['A1'].font = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws4['A1'].fill = PatternFill("solid", fgColor="B8860B")
    ws4['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 35
    for i, h in enumerate(["Reference Key","Vendor","Vendor Invoice #","# Shipments",
                            "Vendor Cost","Notes","Resolved?"], 1):
        hdr(ws4.cell(row=2, column=i, value=h), bg="B8860B")
    for i, w in enumerate([18,14,18,12,14,35,12], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w
    row_n4 = 3
    for _, r in not_in_vision.sort_values('key').iterrows():
        vendor_pairs = []
        if has_wwe:
            vendor_pairs.append(("WWE (UPS)", r.get('wwe_actual_cost'), r.get('wwe_shipments'), r.get('wwe_invoice_num','')))
        if has_nin:
            vendor_pairs.append(("NIN (Courier)", r.get('nin_actual_cost'), r.get('nin_shipments'), r.get('nin_invoice_num','')))
        for vendor, cost, ships, inv in vendor_pairs:
            if pd.notna(cost):
                ws4.row_dimensions[row_n4].height = 18
                for ci, val in enumerate([r['key'], vendor, inv,
                                          int(ships) if pd.notna(ships) else '',
                                          cost, "Invoice not matched in Vision", ""], 1):
                    c = ws4.cell(row=row_n4, column=ci, value=val)
                    body(c, center=(ci in [2,3,4,5,7]), fill=YELLOW_FILL if ci in [1,2,5] else None)
                    if ci == 5:
                        c.number_format = '$#,##0.00'
                row_n4 += 1

    # ══ Vision Only tab ══
    ws5 = wb.create_sheet("Vision Only")
    ws5.sheet_view.showGridLines = False
    ws5.merge_cells('A1:G1')
    ws5['A1'] = "Vision Entries with No Vendor Invoice — For Reference"
    ws5['A1'].font = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws5['A1'].fill = PatternFill("solid", fgColor="455A64")
    ws5['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 35
    for i, h in enumerate(["GSB Invoice #","Sales Rep","Date","Description",
                            "Amount Billed","Cost Recorded","Notes"], 1):
        hdr(ws5.cell(row=2, column=i, value=h), bg="455A64")
    for i, w in enumerate([16,14,12,40,16,16,30], 1):
        ws5.column_dimensions[get_column_letter(i)].width = w
    row_n5 = 3
    for _, r in vision_only.sort_values('vision_invoice').iterrows():
        ws5.row_dimensions[row_n5].height = 18
        for ci, val in enumerate([r['vision_invoice'], r['vision_sales_rep'], r['vision_date'],
                                   r['vision_description'], r['vision_billed'], r['vision_cost'], ""], 1):
            c = ws5.cell(row=row_n5, column=ci, value=val)
            body(c, center=(ci in [1,2,3,5,6]))
            if ci in [5, 6]:
                c.number_format = '$#,##0.00'
        row_n5 += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Ready check ──
ready = vision_file and (nin_file or wwe_file)

if st.button("Run Reconciliation", disabled=not ready):
    with st.spinner("Matching invoices and calculating discrepancies..."):
        try:
            has_nin = nin_file is not None
            has_wwe = wwe_file is not None

            vision_agg = load_vision(vision_file)
            if vision_agg is None:
                st.stop()

            merged = vision_agg.copy()

            if has_wwe:
                wwe_agg = load_wwe(wwe_file)
                if wwe_agg is None:
                    st.stop()
                merged = merged.merge(wwe_agg, on='key', how='outer')

            if has_nin:
                nin_agg = load_nin(nin_file)
                if nin_agg is None:
                    st.stop()
                merged = merged.merge(nin_agg, on='key', how='outer')

            merged = merged[merged['key'] != ''].copy()

            if has_wwe:
                merged['wwe_diff'] = (merged['wwe_actual_cost'] - merged['vision_cost']).round(2)
            if has_nin:
                merged['nin_diff'] = (merged['nin_actual_cost'] - merged['vision_cost']).round(2)

            merged['status'] = merged.apply(assign_status, axis=1)

            mismatches    = merged[merged['status'] == 'MISMATCH']
            matches       = merged[merged['status'] == 'MATCH']
            not_in_vision = merged[merged['status'] == 'NOT IN VISION']
            vision_only   = merged[merged['status'] == 'VISION ONLY']

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Mismatches",    len(mismatches))
            col2.metric("Matched",       len(matches))
            col3.metric("Not in Vision", len(not_in_vision))
            col4.metric("Vision Only",   len(vision_only))

            # Build labels for filename and Excel
            vendor_parts = []
            if has_nin: vendor_parts.append("NIN")
            if has_wwe: vendor_parts.append("WWE")
            vendors_label = " + ".join(vendor_parts)

            if period_start and period_end:
                period_label = f"{period_start.strftime('%b %d %Y')} – {period_end.strftime('%b %d %Y')}"
                period_slug  = f"{period_start.strftime('%Y%m%d')}_to_{period_end.strftime('%Y%m%d')}"
            else:
                period_label = "Period not specified"
                period_slug  = date.today().strftime('%Y%m%d')

            filename = f"GSB_Shipping_Recon_{period_slug}.xlsx"

            excel_buf = build_excel(merged, has_nin, has_wwe, vendors_label, period_label)
            st.success("Reconciliation complete. Download your report below.")
            st.download_button(
                label="Download Reconciliation Report (.xlsx)",
                data=excel_buf,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.info("Make sure your files match the expected format and try again.")

elif not ready:
    if not vision_file:
        st.info("Upload the Vision export to get started. Then add at least one vendor invoice (NIN or WWE).")
    else:
        st.info("Vision file uploaded. Now add at least one vendor invoice (NIN or WWE) to run the reconciliation.")
