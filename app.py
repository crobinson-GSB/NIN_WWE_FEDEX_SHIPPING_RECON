import streamlit as st
import pandas as pd
import re
import io
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
    .stUploadedFile { border-color: #E8601C !important; }
    div[data-testid="stFileUploader"] label { font-weight: 600; color: #2B2B2B; }
    .stButton button { background-color: #E8601C !important; color: white !important;
                       border: none !important; font-weight: 600 !important;
                       padding: 0.6rem 2rem !important; border-radius: 6px !important; }
    .stButton button:hover { background-color: #c94f14 !important; }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div style='padding-top: 2rem;'>", unsafe_allow_html=True)
st.image("https://gsbdigital.com/wp-content/uploads/2018/06/GSB_DigitalLogo-2015-300x44.png", width=200)
st.markdown("</div>", unsafe_allow_html=True)
st.markdown("# Shipping Reconciliation Tool")
st.markdown("Automatically match Vision exports against vendor invoices and identify shipping cost discrepancies.")
st.divider()

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

**Step 3 — Upload the files**
Drop each file into the correct upload box below. The Vision export and at least one vendor invoice (NIN or WWE) are required. FedEx and additional vendors are optional.

**Step 4 — Run the reconciliation**
Click "Run Reconciliation." The tool will process the files and show a summary of results.

**Step 5 — Download the report**
Click the download button to get your Excel report. It includes five tabs:
- **Summary** — overall counts and financial impact
- **Mismatches** — invoices where Vision cost and vendor cost don't match (action required)
- **Matched** — invoices that reconciled cleanly
- **Not in Vision** — vendor charges with no matching Vision entry
- **Vision Only** — Vision entries with no vendor invoice received yet

**Tips**
- Invoice numbers are matched automatically — no manual cleaning needed
- Run this each billing cycle when vendor invoices arrive
- The Mismatches tab is your primary action list
""")
st.divider()

st.markdown("### Vision Export")
vision_file = st.file_uploader(
    "Vision Report (.txt or .xlsx export from Printsmith Vision)",
    type=["txt", "xlsx"],
    key="vision"
)

st.markdown("### Vendor Invoices")
col1, col2 = st.columns(2)
with col1:
    nin_file = st.file_uploader("NIN — Courier (.xls)", type=["xls", "xlsx"], key="nin")
with col2:
    wwe_file = st.file_uploader("WWE — UPS (.xls)", type=["xls", "xlsx"], key="wwe")

col3, col4 = st.columns(2)
with col3:
    fedex_file = st.file_uploader("FedEx (optional)", type=["xls", "xlsx", "csv"], key="fedex")
with col4:
    extra_file = st.file_uploader("Additional vendor (optional)", type=["xls", "xlsx", "csv"], key="extra")

st.divider()

def clean_key(val):
    if pd.isna(val):
        return ''
    s = str(val).strip()
    s = re.sub(r'\.0$', '', s)
    s = re.sub(r'^[Dd]', '', s)
    s = s.split('/')[0]
    s = s.replace('-', '').replace(' ', '').strip()
    return s.upper()

def load_vision(file):
    if file.name.endswith('.txt'):
        df = pd.read_csv(file, sep='\t')
    else:
        df = pd.read_excel(file)
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
    df = pd.read_excel(file, engine='xlrd')
    df['AmountCharged'] = pd.to_numeric(df['AmountCharged'], errors='coerce').fillna(0)
    df['key'] = df['Auth'].apply(clean_key)
    return df.groupby('key').agg(
        nin_invoice_num=('InvoiceNumber', 'first'),
        nin_actual_cost=('AmountCharged', 'sum'),
        nin_shipments=('OrderNumber', 'count')
    ).reset_index()

def load_wwe(file):
    df = pd.read_excel(file, engine='xlrd')
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

ORANGE = "E8601C"
DARK = "2B2B2B"
WHITE = "FFFFFF"
RED_FILL = "FFCCCC"
GREEN_FILL = "CCFFCC"
YELLOW_FILL = "FFF2CC"
GRAY_FILL = "F5F5F5"
MID_GRAY = "D9D9D9"

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

def build_excel(merged):
    mismatches = merged[merged['status'] == 'MISMATCH'].copy()
    matches = merged[merged['status'] == 'MATCH'].copy()
    vision_only = merged[merged['status'] == 'VISION ONLY'].copy()
    not_in_vision = merged[merged['status'] == 'NOT IN VISION'].copy()

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:F1')
    ws['A1'] = "GSB Digital — Shipping Reconciliation Report"
    ws['A1'].font = Font(name="Arial", bold=True, size=16, color=WHITE)
    ws['A1'].fill = PatternFill("solid", fgColor=ORANGE)
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    stats = [("MISMATCHES", len(mismatches), "FFCCCC"), ("MATCHED", len(matches), "CCFFCC"),
             ("NOT IN VISION", len(not_in_vision), "FFF2CC"), ("VISION ONLY", len(vision_only), "E8F4FD")]
    ws.row_dimensions[3].height = 50
    ws.row_dimensions[4].height = 25
    for i, (label, val, bg) in enumerate(stats, 1):
        c1 = ws.cell(row=3, column=i, value=str(val))
        c1.font = Font(name="Arial", bold=True, size=22)
        c1.fill = PatternFill("solid", fgColor=bg)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = tb()
        c2 = ws.cell(row=4, column=i, value=label)
        c2.font = Font(name="Arial", bold=True, size=9, color="555555")
        c2.fill = PatternFill("solid", fgColor=bg)
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = tb()

    nin_disc = mismatches[mismatches['nin_actual_cost'].notna()]
    wwe_disc = mismatches[mismatches['wwe_actual_cost'].notna()]
    ws['A6'] = "Financial Impact"
    ws['A6'].font = Font(name="Arial", bold=True, size=11, color=WHITE)
    ws['A6'].fill = PatternFill("solid", fgColor=DARK)
    ws.merge_cells('A6:F6')
    ws['A6'].alignment = Alignment(horizontal="left", vertical="center")
    fin_rows = [("Total Vendor Cost", nin_disc['nin_actual_cost'].sum(), wwe_disc['wwe_actual_cost'].sum()),
                ("Total in Vision", nin_disc['vision_cost'].sum(), wwe_disc['vision_cost'].sum()),
                ("Total Discrepancy", nin_disc['nin_diff'].sum(), wwe_disc['wwe_diff'].sum())]
    for r, (label, nin_v, wwe_v) in enumerate(fin_rows, 8):
        ws.cell(row=r, column=1, value=label).font = Font(name="Arial", bold=True, size=10)
        ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=GRAY_FILL)
        ws.cell(row=r, column=1).border = tb()
        for col, val in [(2, nin_v), (3, wwe_v), (4, nin_v + wwe_v)]:
            c = ws.cell(row=r, column=col, value=round(val, 2))
            c.number_format = '$#,##0.00'
            c.alignment = Alignment(horizontal="center")
            c.border = tb()
            if label == "Total Discrepancy":
                c.font = Font(name="Arial", bold=True, size=10,
                              color="CC0000" if val < 0 else ("006600" if val > 0 else DARK))
    for i, w in enumerate([30, 18, 18, 18, 5, 5], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

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
        for vendor, cost, diff, inv, ships in [
            ("WWE (UPS)", r.get('wwe_actual_cost'), r.get('wwe_diff'), r.get('wwe_invoice_num',''), r.get('wwe_shipments')),
            ("NIN (Courier)", r.get('nin_actual_cost'), r.get('nin_diff'), r.get('nin_invoice_num',''), r.get('nin_shipments')),
        ]:
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
        for vendor, cost, diff in [
            ("WWE (UPS)", r.get('wwe_actual_cost'), r.get('wwe_diff')),
            ("NIN (Courier)", r.get('nin_actual_cost'), r.get('nin_diff')),
        ]:
            if pd.notna(cost):
                ws3.row_dimensions[row_n3].height = 18
                for ci, val in enumerate([r['vision_invoice'], r['vision_sales_rep'], r['vision_date'],
                                          vendor, '', r['vision_cost'], cost, round(diff, 2), "✓"], 1):
                    c = ws3.cell(row=row_n3, column=ci, value=val)
                    body(c, center=(ci in [1,3,4,5,6,7,8,9]), fill=GREEN_FILL if ci == 9 else None)
                    if ci in [6, 7, 8]:
                        c.number_format = '$#,##0.00'
                row_n3 += 1

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
        for vendor, cost, ships, inv in [
            ("WWE (UPS)", r.get('wwe_actual_cost'), r.get('wwe_shipments'), r.get('wwe_invoice_num','')),
            ("NIN (Courier)", r.get('nin_actual_cost'), r.get('nin_shipments'), r.get('nin_invoice_num','')),
        ]:
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

ready = vision_file and nin_file and wwe_file

if st.button("Run Reconciliation", disabled=not ready):
    with st.spinner("Matching invoices and calculating discrepancies..."):
        try:
            vision_agg = load_vision(vision_file)
            nin_agg    = load_nin(nin_file)
            wwe_agg    = load_wwe(wwe_file)

            merged = vision_agg.merge(wwe_agg, on='key', how='outer')
            merged = merged.merge(nin_agg, on='key', how='outer')
            merged = merged[merged['key'] != ''].copy()

            merged['wwe_diff'] = (merged['wwe_actual_cost'] - merged['vision_cost']).round(2)
            merged['nin_diff'] = (merged['nin_actual_cost'] - merged['vision_cost']).round(2)
            merged['status']   = merged.apply(assign_status, axis=1)

            mismatches    = merged[merged['status'] == 'MISMATCH']
            matches       = merged[merged['status'] == 'MATCH']
            not_in_vision = merged[merged['status'] == 'NOT IN VISION']
            vision_only   = merged[merged['status'] == 'VISION ONLY']

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Mismatches",    len(mismatches))
            col2.metric("Matched",       len(matches))
            col3.metric("Not in Vision", len(not_in_vision))
            col4.metric("Vision Only",   len(vision_only))

            excel_buf = build_excel(merged)
            st.success("Reconciliation complete. Download your report below.")
            st.download_button(
                label="Download Reconciliation Report (.xlsx)",
                data=excel_buf,
                file_name="GSB_Shipping_Reconciliation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.info("Make sure your files match the expected format and try again.")

elif not ready:
    st.info("Upload the Vision report, NIN invoice, and WWE invoice to get started.")
