import io
import re
import streamlit as st
import docx
import pdfplumber
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# --- Custom Styling & Theme Configuration (Light Blue Palette) ---
st.set_page_config(page_title="KIST Document Generator", page_icon="📄", layout="centered")

st.markdown("""
    <style>
    .stApp {
        background-color: #f4f8fb;
    }
    .stButton>button {
        background-color: #4682B4;
        color: white;
        border-radius: 6px;
        border: none;
        padding: 0.5rem 2rem;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #356a95;
        color: white;
    }
    h1, h2, h3 {
        color: #2c5d88;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .black-text {
        color: #000000 !important;
        font-weight: normal;
    }
    div[data-testid="stFileUploader"] {
        border: 2px dashed #b0c4de;
        padding: 20px;
        border-radius: 10px;
        background-color: #ffffff;
    }
    </style>
""", unsafe_allow_html=True)

# --- Helper functions for Word Table Styling ---
def set_cell_margins(cell, top=40, bottom=40, left=100, right=100):
    """Reduced padding parameters to optimize row heights and keep page counts minimal"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        '<w:tblBorders %s>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E5E5E5"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E5E5E5"/>'
        '</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

# --- Document Generation Function ---
def generate_cash_sheet(invoices, total_amount):
    doc = Document()

    # Compressed Margins for maximizing page utilization
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)

    # ---------------- PAGE 1 ----------------
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(4)
    run = title.add_run("KIST DAY CASH SHEET")
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.name = 'Arial'

    meta_p = doc.add_paragraph()
    meta_p.paragraph_format.space_after = Pt(6)
    meta_p.add_run("Date : ___________________    Route : ___________________    No.Bill : ___________________").font.name = 'Arial'

    headers = ["No", "Invoice Number", "Shop Name", "Amount", "BNW", "Cancel", "Adjust", "Dis", "Cash", "Credit", "Cheque", "Rtn"]

    table = doc.add_table(rows=1, cols=12)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)

    hdr_cells = table.rows[0].cells
    for i, heading_text in enumerate(headers):
        hdr_cells[i].text = heading_text
        set_cell_background(hdr_cells[i], "FFFFFF")  # Normal white background
        set_cell_margins(hdr_cells[i], top=60, bottom=60)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True                        # Black and bold fonts
            r.font.size = Pt(8.5)
            r.font.color.rgb = docx.shared.RGBColor(0, 0, 0)

    idx = 1
    for item in invoices:
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(item['invoice'])
        row_cells[2].text = ""
        row_cells[3].text = f"{item['amount']:.2f}"

        for c_idx, cell in enumerate(row_cells):
            set_cell_margins(cell, top=40, bottom=40) # Reduced row padding
            p = cell.paragraphs[0]
            if p.runs:
                p.runs[0].font.size = Pt(8.5)
            if c_idx in [0, 1]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif c_idx == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        idx += 1

    st_cells = table.add_row().cells
    st_cells[0].text = "ST"
    st_cells[1].text = "System Sale"
    st_cells[2].text = ""
    st_cells[3].text = f"{total_amount:,.2f}"

    for c_idx, cell in enumerate(st_cells):
        set_cell_margins(cell, top=40, bottom=40)
        set_cell_background(cell, "EBF2F8")
        p = cell.paragraphs[0]
        if p.runs:
            p.runs[0].font.bold = True
            p.runs[0].font.size = Pt(8.5)
        if c_idx in [0, 1]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif c_idx == 3:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for _ in range(5):
        extra_cells = table.add_row().cells
        extra_cells[0].text = str(idx)
        for c_idx, cell in enumerate(extra_cells):
            set_cell_margins(cell, top=80, bottom=80)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            if cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].font.size = Pt(8.5)
        idx += 1

    col_widths = [Inches(0.35), Inches(1.1), Inches(1.6), Inches(0.9), Inches(0.35), Inches(0.45), Inches(0.5), Inches(0.45), Inches(0.5), Inches(0.5), Inches(0.5), Inches(0.4)]
    for row in table.rows:
        for i, width in enumerate(col_widths):
            row.cells[i].width = width

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(6)
    p_spacer.paragraph_format.space_after = Pt(0)
    
    sales_labels = ["Biscuits sale : _______________________", "Nectar Sale : _______________________", "Water Sale : _______________________", "Total Sale : _______________________"]
    for label in sales_labels:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.add_run(label).font.name = 'Arial'

    # ---------------- PAGE 2 ----------------
    doc.add_page_break()

    p2_title = doc.add_paragraph()
    p2_title.paragraph_format.space_before = Pt(0)
    p2_title.paragraph_format.space_after = Pt(4)
    run = p2_title.add_run("Cash Receivables")
    run.font.size = Pt(12)
    run.font.bold = True

    rec_headers = ["NO", "Bill Date", "Shop", "Credit Amount", "Pay Amount", "Balance"]
    rec_table = doc.add_table(rows=1, cols=6)
    rec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rec_table)

    hdr_cells = rec_table.rows[0].cells
    for i, h_text in enumerate(rec_headers):
        hdr_cells[i].text = h_text
        set_cell_background(hdr_cells[i], "FFFFFF")  # Normal white background
        set_cell_margins(hdr_cells[i], top=60, bottom=60)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.bold = True                # Black and bold fonts
            p.runs[0].font.color.rgb = docx.shared.RGBColor(0, 0, 0)
            p.runs[0].font.size = Pt(8.5)

    for r_idx in range(1, 11):
        row_cells = rec_table.add_row().cells
        row_cells[0].text = str(r_idx)
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for cell in row_cells:
            set_cell_margins(cell, top=40, bottom=40)
            if cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].font.size = Pt(8.5)

    rec_widths = [Inches(0.4), Inches(1.2), Inches(2.2), Inches(1.2), Inches(1.2), Inches(1.3)]
    for row in rec_table.rows:
        for i, w in enumerate(rec_widths):
            row.cells[i].width = w

    p_bal = doc.add_paragraph()
    p_bal.paragraph_format.space_before = Pt(8)
    p_bal.paragraph_format.space_after = Pt(4)
    run_bal = p_bal.add_run("Cash Sheet Balancing")
    run_bal.font.size = Pt(12)
    run_bal.font.bold = True

    master_table = doc.add_table(rows=1, cols=2)
    master_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    cell_left = master_table.rows[0].cells[0]
    cell_left.width = Inches(3.8)
    sub_table_left = cell_left.add_table(rows=0, cols=2)
    set_table_borders(sub_table_left)

    # Corrected "Total Returnt" to "Total Return"
    left_rows = [
        ("System Sale", f"{total_amount:,.2f}"), ("FOC", ""), ("Total Cancel", ""), ("Balance (1)", ""),
        ("Total Discounts", ""), ("Total Adjust", ""), ("Total Return", ""), ("Balance (2)", ""),
        ("Total Cash", ""), ("Total Credit", ""), ("Total Cheques", ""), ("Balance (3)", "")
    ]

    for item, val in left_rows:
        row_cells = sub_table_left.add_row().cells
        row_cells[0].text = item
        row_cells[1].text = val
        set_cell_margins(row_cells[0], top=40, bottom=40)
        set_cell_margins(row_cells[1], top=40, bottom=40)
        
        p0 = row_cells[0].paragraphs[0]
        p1 = row_cells[1].paragraphs[0]
        if p0.runs: p0.runs[0].font.size = Pt(8.5)
        if p1.runs: p1.runs[0].font.size = Pt(8.5)

        if "Balance" in item:
            set_cell_background(row_cells[0], "F0F4F8")
            set_cell_background(row_cells[1], "F0F4F8")
            if p0.runs:
                p0.runs[0].font.bold = True
        row_cells[0].width = Inches(2.3)
        row_cells[1].width = Inches(1.5)
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    cell_right = master_table.rows[0].cells[1]
    cell_right.width = Inches(3.7)
    sub_table_right = cell_right.add_table(rows=1, cols=2)
    set_table_borders(sub_table_right)

    hdr_r = sub_table_right.rows[0].cells
    hdr_r[0].text = "Cash Balance"
    hdr_r[1].text = ""
    set_cell_background(hdr_r[0], "FFFFFF")  # Normal white background
    set_cell_background(hdr_r[1], "FFFFFF")
    set_cell_margins(hdr_r[0], top=40, bottom=40)
    set_cell_margins(hdr_r[1], top=40, bottom=40)
    if hdr_r[0].paragraphs[0].runs:
        hdr_r[0].paragraphs[0].runs[0].font.bold = True
        hdr_r[0].paragraphs[0].runs[0].font.color.rgb = docx.shared.RGBColor(0, 0, 0) # Black Font
        hdr_r[0].paragraphs[0].runs[0].font.size = Pt(8.5)

    right_rows = [("Total Day Cash", False), ("Total Credit Received", False), ("Total Expenses", True), ("Banked Value.", False)]
    for item, is_expense in right_rows:
        row_cells = sub_table_right.add_row().cells
        row_cells[0].text = item
        set_cell_margins(row_cells[0], top=40, bottom=40)
        set_cell_margins(row_cells[1], top=40, bottom=40)
        
        if row_cells[0].paragraphs[0].runs: row_cells[0].paragraphs[0].runs[0].font.size = Pt(8.5)
        row_cells[0].width = Inches(2.2)
        row_cells[1].width = Inches(1.5)
        if is_expense:
            set_cell_margins(row_cells[0], top=40, bottom=240)  # Tighter layout sizing

    p_calc = doc.add_paragraph()
    p_calc.paragraph_format.space_before = Pt(8)
    p_calc.paragraph_format.space_after = Pt(4)
    run_calc = p_calc.add_run("Calculating cash")
    run_calc.font.size = Pt(12)
    run_calc.font.bold = True

    calc_table = doc.add_table(rows=1, cols=2)
    calc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(calc_table)

    hdr_c = calc_table.rows[0].cells
    hdr_c[0].text = "Cash Analitics"
    hdr_c[1].text = "Valuve"
    set_cell_background(hdr_c[0], "FFFFFF")  # Normal white background
    set_cell_background(hdr_c[1], "FFFFFF")
    for cell in hdr_c:
        set_cell_margins(cell, top=60, bottom=60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.bold = True                # Black and bold fonts
            p.runs[0].font.color.rgb = docx.shared.RGBColor(0, 0, 0)
            p.runs[0].font.size = Pt(8.5)

    denominations = ["20", "50", "100", "500", "1000", "2000", "5000", "coins", "Total"]
    for denom in denominations:
        row_cells = calc_table.add_row().cells
        row_cells[0].text = denom
        set_cell_margins(row_cells[0], top=40, bottom=40)
        set_cell_margins(row_cells[1], top=40, bottom=40)
        
        p0 = row_cells[0].paragraphs[0]
        p1 = row_cells[1].paragraphs[0]
        if p0.runs: p0.runs[0].font.size = Pt(8.5)
        if p1.runs: p1.runs[0].font.size = Pt(8.5)
        
        row_cells[0].width = Inches(3.75)
        row_cells[1].width = Inches(3.75)
        if denom == "Total":
            set_cell_background(row_cells[0], "EBF2F8")
            set_cell_background(row_cells[1], "EBF2F8")
            if p0.runs:
                p0.runs[0].font.bold = True

    p_foot = doc.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(8)
    p_foot.add_run("Distance Travelled : ___________________    KM : ___________________    OOT : ___________________").font.name = 'Arial'
    if p_foot.runs:
        p_foot.runs[0].font.size = Pt(9)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# --- Native pdfplumber Raw Text Extraction Engine ---
def parse_pdf_file(uploaded_file):
    invoices = []
    grand_total = 0.0
    full_text = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text_content = page.extract_text()
            if text_content:
                full_text += text_content + "\n"

    invoice_section = re.search(r'Net Amt \(LKR\)(.*)', full_text, re.DOTALL)
    if not invoice_section:
        return [], 0.0

    invoice_block_text = invoice_section.group(1)

    matches = re.finditer(r'\b(TI\d{6})\b.*?([\d,]+\.\d{2})', invoice_block_text, re.DOTALL)

    seen_invoices = set()
    for match in matches:
        inv_code = match.group(1).strip()
        amt_val = float(match.group(2).replace(',', ''))

        if inv_code not in seen_invoices:
            invoices.append({'invoice': inv_code, 'amount': amt_val})
            seen_invoices.add(inv_code)

    total_match = re.search(r'Total\s*(?:\|\s*)?([\d,]+\.\d{2})', invoice_block_text)
    if total_match:
        grand_total = float(total_match.group(1).replace(',', ''))
    elif invoices:
        grand_total = sum(i['amount'] for i in invoices)

    return invoices, grand_total

# --- Main App Execution Interface ---
st.title("📄 KIST Sheet Custom Generation Engine")

# Formatted with explicit HTML and a class selector to force a standard black color tone
st.markdown('<p class="black-text">Upload the picklist PDF file below to instantly extract data and generate the structured layout.</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Select input Picklist PDF file", type=["pdf"])

if uploaded_file is not None:
    invoices, total_amt = parse_pdf_file(uploaded_file)

    if invoices:
        st.success(f"Successfully processed {len(invoices)} invoices records from PDF. Identified System Sale Value: LKR {total_amt:,.2f}")

        with st.spinner("Compiling structural layout patterns..."):
            doc_stream = generate_cash_sheet(invoices, total_amt)

        st.download_button(
            label="📥 Download Tailored Word Document",
            data=doc_stream,
            file_name="KIST_Day_Cash_Sheet.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("Could not parse invoices structures. Please verify that the PDF contains valid invoice entries.")
