import io
import re
import streamlit as st
import pdfplumber

# ReportLab core engines for precise programmatic PDF layout
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- Custom Styling & Theme Configuration for Streamlit ---
st.set_page_config(page_title="KIST Document Generator", page_icon="📄", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #f4f8fb; }
    .stButton>button {
        background-color: #4682B4; color: white; border-radius: 6px;
        border: none; padding: 0.5rem 2rem; font-weight: bold;
    }
    .stButton>button:hover { background-color: #356a95; color: white; }
    div[data-testid="stFileUploader"] {
        border: 2px dashed #b0c4de; padding: 20px; border-radius: 10px; background-color: #ffffff;
    }
    </style>
""", unsafe_allow_html=True)

# --- ReportLab Pure-Python PDF Generation Layout ---
def generate_cash_sheet_pdf(invoices, total_amount):
    buffer = io.BytesIO()
    
    # Rule: > 35 invoices allows 3 pages, else strictly force a 2-page fit
    is_large_dataset = len(invoices) > 35
    
    # Margins dynamically adjusted based on the required page budget rule
    margin_val = 36 if is_large_dataset else 26
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=margin_val, leftMargin=margin_val, topMargin=margin_val, bottomMargin=margin_val
    )
    body_elements = []
    
    # Text Typography Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold',
        fontSize=16, leading=20, alignment=1, spaceAfter=6, textColor=colors.HexColor("#000000")
    )
    meta_style = ParagraphStyle(
        'MetaStyle', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=9.5, leading=13, spaceAfter=10, textColor=colors.HexColor("#000000")
    )
    cell_style = ParagraphStyle(
        'CellStyle', parent=styles['Normal'], fontName='Helvetica', 
        fontSize=8, leading=10, alignment=1
    )
    cell_bold = ParagraphStyle(
        'CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', 
        fontSize=8.5, leading=10.5, alignment=1
    )
    cell_right = ParagraphStyle(
        'CellRight', parent=styles['Normal'], fontName='Helvetica', 
        fontSize=8.5, leading=10.5, alignment=2
    )

    # ==================== DATA ENTRIES ====================
    body_elements.append(Paragraph("KIST DAY CASH SHEET", title_style))
    body_elements.append(Paragraph("Date : ___________________    Route : ___________________    No.Bill : ___________________", meta_style))
    
    headers = ["No", "Invoice Number", "Shop Name", "Amount", "BNW", "Cancel", "Adjust", "Dis", "Cash", "Credit", "Cheque", "Rtn"]
    table_data = [[Paragraph(h, cell_bold) for h in headers]]
    
    # Inject active rows parsed from raw file
    idx = 1
    for item in invoices:
        table_data.append([
            Paragraph(str(idx), cell_style), Paragraph(item['invoice'], cell_style), Paragraph("", cell_style),
            Paragraph(f"{item['amount']:.2f}", cell_right),
            Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style),
            Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style)
        ])
        idx += 1
        
    # Standard blank spacer rows for manual adjustments
    for _ in range(5):
        table_data.append([Paragraph(str(idx), cell_style)] + [Paragraph("", cell_style) for _ in range(11)])
        idx += 1
        
    # System Sale summary anchor line put directly at the end after the 5 extra rows
    table_data.append([
        Paragraph("ST", cell_bold), Paragraph("System Sale", cell_bold), Paragraph("", cell_style),
        Paragraph(f"{total_amount:,.2f}", ParagraphStyle('ST_R', parent=cell_right, fontName='Helvetica-Bold')),
        Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style),
        Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style)
    ])
    
    # Perfectly Balanced Explicit Point Column Widths
    col_widths = [22, 62, 90, 50, 32, 42, 42, 30, 45, 45, 45, 32]
    
    main_table_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5.5),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#ebf2f8")), 
    ])
    
    body_elements.append(Table(table_data, colWidths=col_widths, style=main_table_style))
    body_elements.append(Spacer(1, 10))
    
    sales_labels = [
        "Biscuits sale : _______________________", "Nectar Sale : _______________________",
        "Water Sale : _______________________", "Total Sale : _______________________"
    ]
    for label in sales_labels:
        body_elements.append(Paragraph(label, meta_style))
        
    doc.build(body_elements)
    buffer.seek(0)
    return buffer

# --- Secondary Fallback Parser Matrix (Triggers if Audit Mismatch Detected) ---
def parse_pdf_via_coordinates_fallback(uploaded_file):
    invoices = []
    invoice_seen = set()
    raw_words = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words(y_tolerance=3, x_tolerance=3)
            for w in words:
                raw_words.append({
                    'text': w['text'].strip(),
                    'top': round(w['top'], 1),
                    'x0': round(w['x0'], 1),
                    'page': page.page_number
                })
                
    page_groups = {}
    for w in raw_words:
        p = w['page']
        if p not in page_groups: page_groups[p] = []
        page_groups[p].append(w)
        
    structured_rows = []
    for p, words_on_page in page_groups.items():
        words_on_page.sort(key=lambda x: (x['top'], x['x0']))
        current_row = []
        current_top = -100.0
        for w in words_on_page:
            if abs(w['top'] - current_top) <= 3.0:
                current_row.append(w)
            else:
                if current_row: structured_rows.append(current_row)
                current_row = [w]
                current_top = w['top']
        if current_row: structured_rows.append(current_row)

    for r_idx, row in enumerate(structured_rows):
        line_text = " ".join([w['text'] for w in row])
        last_word = row[-1]['text']
        if re.search(r'^\d[\d,]*\.\d{2}$', last_word) and not "total" in line_text.lower():
            target_code = None
            for check_idx in range(max(0, r_idx - 3), r_idx):
                check_text = " ".join([w['text'] for w in structured_rows[check_idx]])
                all_digits = re.findall(r'\d+', check_text)
                if all_digits:
                    candidate = max(all_digits, key=len)
                    if len(candidate) >= 3:
                        target_code = candidate
                        break
            if not target_code:
                all_digits = re.findall(r'\d+', line_text.replace(last_word, ''))
                if all_digits: target_code = max(all_digits, key=len)
                
            if target_code and len(target_code) >= 2:
                try:
                    amt_val = float(last_word.replace(",", ""))
                    slice_len = 4
                    final_slice = target_code[-slice_len:]
                    while final_slice in invoice_seen and slice_len < len(target_code):
                        slice_len += 1
                        final_slice = target_code[-slice_len:]
                    if final_slice not in invoice_seen:
                        invoices.append({"invoice": final_slice, "amount": amt_val})
                        invoice_seen.add(final_slice)
                except ValueError:
                    pass
    return invoices

# --- Primary Stateful Line-by-Line Multi-Format Text Extraction Engine ---
def parse_pdf_file(uploaded_file):
    invoices = []
    total_amount = 0.0
    invoice_seen = set()
    expected_entry_count = 0

    composite_invoice_pattern = re.compile(r'\b\d{2}[A-Z]{3}_\d{4,}[\w\d_]*\b')
    standard_invoice_pattern = re.compile(r'\b(IN|TI)\d{5,7}\b')
    isolated_id_pattern = re.compile(r'^\b\d{3,4}\b$')
    amount_pattern = re.compile(r'\d[\d,]*\.\d{2}')

    with pdfplumber.open(uploaded_file) as pdf:
        active_prefix = None
        active_sub_id = None

        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue

            lines = text.splitlines()
            for line in lines:
                line_str = line.strip()
                if not line_str: continue

                # Count rows that look like entry rows (start with index digit followed by data blocks)
                if re.match(r'^\b\d{1,2}\b\s+\b(TI|IN|\d{2}[A-Z]{3})\b|^\b\d{1,2}\b$', line_str):
                    expected_entry_count += 1

                if line_str.lower().startswith("total") or "grand total" in line_str.lower():
                    amts = amount_pattern.findall(line_str)
                    if amts:
                        try:
                            total_amount = float(amts[-1].replace(",", ""))
                        except ValueError: pass
                    continue

                comp_match = composite_invoice_pattern.search(line_str)
                if comp_match:
                    active_prefix = comp_match.group()
                    line_remainder = line_str.replace(active_prefix, "").strip()
                    remainder_numbers = re.findall(r'\b\d{3,4}\b', line_remainder)
                    if remainder_numbers: active_sub_id = remainder_numbers[0]
                    continue

                std_match = standard_invoice_pattern.search(line_str)
                if std_match:
                    active_prefix = std_match.group()
                    active_sub_id = None

                iso_match = isolated_id_pattern.search(line_str)
                if iso_match and active_prefix and not active_prefix.startswith(("IN", "TI")):
                    active_sub_id = iso_match.group()
                    continue

                amts = amount_pattern.findall(line_str)
                if amts and active_prefix:
                    if "net amt" in line_str.lower() or "mrp" in line_str.lower(): continue

                    if active_prefix.startswith(("IN", "TI")):
                        invoice_identity = active_prefix
                    elif active_sub_id:
                        invoice_identity = active_sub_id
                    else:
                        digits_found = re.findall(r'\d+', active_prefix)
                        invoice_identity = digits_found[-1] if digits_found else active_prefix

                    try:
                        amt_val = float(amts[-1].replace(",", ""))
                        if len(invoice_identity) >= 2:
                            slice_length = 4
                            final_invoice_slice = invoice_identity[-slice_length:]
                            while final_invoice_slice in invoice_seen and slice_length < len(invoice_identity):
                                slice_length += 1
                                final_invoice_slice = invoice_identity[-slice_length:]

                            invoices.append({"invoice": final_invoice_slice, "amount": amt_val})
                            invoice_seen.add(final_invoice_slice)
                            
                        active_sub_id = None
                        if active_prefix.startswith(("IN", "TI")): active_prefix = None
                    except ValueError: pass

    # ==================== THE AUDIT VERIFICATION ENGINE ====================
    calculated_sum = sum(i["amount"] for i in invoices)
    
    # Audit Rule Trigger: If counts or sums don't match, run the fallback scanner to recover missing text fields
    if len(invoices) < expected_entry_count or (total_amount > 0.0 and abs(calculated_sum - total_amount) > 0.05):
        fallback_invoices = parse_pdf_via_coordinates_fallback(uploaded_file)
        if len(fallback_invoices) > len(invoices):
            invoices = fallback_invoices
            calculated_sum = sum(i["amount"] for i in invoices)

    if total_amount == 0.0 and invoices:
        total_amount = calculated_sum

    return invoices, total_amount

# --- Main App Interface ---
st.html("""
    <div style="margin-bottom: 1.5rem;">
        <h1 style="color: #000000 !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin-bottom: 0.5rem; font-size: 2.25rem; font-weight: 700;">
            📄 KIST Sheet Custom Generation Engine
        </h1>
        <p style="color: #000000 !important; font-size: 1rem; font-weight: normal; margin: 0;">
            Upload the picklist PDF file below to instantly extract data and generate the structured layout.
        </p>
    </div>
""")

uploaded_file = st.file_uploader("Select input Picklist PDF file", type=["pdf"])

if uploaded_file is not None:
    invoices, total_amt = parse_pdf_file(uploaded_file)

    if invoices:
        st.success(f"Audit Complete: Processed {len(invoices)} entries. System Sale Value: LKR {total_amt:,.2f}")

        with st.spinner("Compiling structural printable layout..."):
            pdf_data = generate_cash_sheet_pdf(invoices, total_amt)

        st.download_button(
            label="📥 Download Tailored Cash Sheet PDF",
            data=pdf_data,
            file_name="KIST_Day_Cash_Sheet.pdf",
            mime="application/pdf"
        )
    else:
        st.error("Could not parse invoices structures. Please verify that the PDF contains valid invoice entries.")
