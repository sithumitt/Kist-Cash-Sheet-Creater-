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
    
    # Standard printable margin layout values (0.5 inch / 36 points)
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
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
        fontSize=8.5, leading=10.5, alignment=1
    )
    cell_bold = ParagraphStyle(
        'CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', 
        fontSize=8.5, leading=10.5, alignment=1
    )
    cell_right = ParagraphStyle(
        'CellRight', parent=styles['Normal'], fontName='Helvetica', 
        fontSize=8.5, leading=10.5, alignment=2
    )

    # ==================== PAGE 1 DATA ENTRIES ====================
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
        
    # System Sale summary anchor line
    table_data.append([
        Paragraph("ST", cell_bold), Paragraph("System Sale", cell_bold), Paragraph("", cell_style),
        Paragraph(f"{total_amount:,.2f}", ParagraphStyle('ST_R', parent=cell_right, fontName='Helvetica-Bold')),
        Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style),
        Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style)
    ])
    
    # Full horizontal dynamic mapping width span configurations (Total 523 points)
    printable_width = 523
    col_widths = [
        printable_width * 0.04,  # No
        printable_width * 0.14,  # Invoice Number space
        printable_width * 0.10,  # Shop Name
        printable_width * 0.09,  # Amount
        printable_width * 0.05,  # BNW
        printable_width * 0.065, # Cancel
        printable_width * 0.065, # Adjust
        printable_width * 0.05,  # Dis
        printable_width * 0.12,  # Cash
        printable_width * 0.12,  # Credit
        printable_width * 0.12,  # Cheque
        printable_width * 0.04   # Rtn
    ]
    
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

# --- Universal Spatial Coordinate Text Alignment Parsing Engine ---
def parse_pdf_file(uploaded_file):
    invoices = []
    grand_total = 0.0
    
    raw_words = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words(y_tolerance=3, x_tolerance=3)
            for w in words:
                raw_words.append({
                    'text': w['text'].strip(),
                    'top': round(w['top'], 1),
                    'bottom': round(w['bottom'], 1),
                    'x0': round(w['x0'], 1),
                    'page': page.page_number
                })
                
    if not raw_words:
        return [], 0.0

    page_groups = {}
    for w in raw_words:
        p = w['page']
        if p not in page_groups:
            page_groups[p] = []
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
                if current_row:
                    structured_rows.append(current_row)
                current_row = [w]
                current_top = w['top']
        if current_row:
            structured_rows.append(current_row)

    for r_idx, row in enumerate(structured_rows):
        line_text = " ".join([w['text'] for w in row])
        last_word = row[-1]['text']
        amt_match = re.search(r'^\d[\d,]*\.\d{2}$', last_word)
        
        if amt_match:
            if "total" in line_text.lower():
                if "grand total" in line_text.lower() or line_text.lower().startswith("total"):
                    try:
                        grand_total = float(last_word.replace(',', ''))
                    except ValueError:
                        pass
                continue
            
            if not re.search(r'\b\d{3}R\d{2}\b', line_text):
                continue
                
            target_invoice_code = None
            
            for check_idx in range(max(0, r_idx - 3), r_idx):
                check_row = structured_rows[check_idx]
                check_text = " ".join([w['text'] for w in check_row])
                
                if re.search(r'\b\d{2}[A-Z]{3}_\d{4}', check_text) or re.search(r'\b\d{2}[A-Z]{3}\b', check_text):
                    all_digits = re.findall(r'\d+', check_text)
                    if all_digits:
                        selected_digits = max(all_digits, key=len)
                        
                        if len(selected_digits) < 3 and len(all_digits) > 1:
                            sorted_digits = sorted(all_digits, key=len, reverse=True)
                            selected_digits = sorted_digits[0]
                            
                        target_invoice_code = selected_digits[-4:]
                        break
                        
            if not target_invoice_code:
                all_digits = re.findall(r'\d+', line_text.replace(last_word, ''))
                filtered_digits = [d for d in all_digits if len(d) <= 4 and d != "730"]
                if filtered_digits:
                    target_invoice_code = filtered_digits[-1][-4:]
            
            if target_invoice_code:
                try:
                    amt_val = float(last_word.replace(',', ''))
                    if not any(item['invoice'] == target_invoice_code for item in invoices):
                        invoices.append({'invoice': target_invoice_code, 'amount': amt_val})
                except ValueError:
                    pass

    if invoices and grand_total == 0.0:
        grand_total = sum(i['amount'] for i in invoices)

    return invoices, grand_total

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
        st.success(f"Successfully processed {len(invoices)} invoices records from PDF. Identified System Sale Value: LKR {total_amt:,.2f}")

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
