import io
import re
import streamlit as st
import pdfplumber

# ReportLab core engines for precise programmatic PDF layout
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
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
    # Tight margins (0.35 inch / 25 points) to guarantee 2-page fit
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25
    )
    story = []
    
    # Text Typography Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold',
        fontSize=14, leading=16, alignment=1, spaceAfter=4, textColor=colors.HexColor("#000000")
    )
    meta_style = ParagraphStyle(
        'MetaStyle', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=8.5, leading=10, spaceAfter=8, textColor=colors.HexColor("#000000")
    )
    section_style = ParagraphStyle(
        'SectionStyle', parent=styles['Heading2'], fontName='Helvetica-Bold',
        fontSize=11, leading=13, spaceBefore=4, spaceAfter=4, textColor=colors.HexColor("#000000")
    )
    cell_style = ParagraphStyle(
        'CellStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=8.5, alignment=1
    )
    cell_bold = ParagraphStyle(
        'CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, leading=8.5, alignment=1
    )
    cell_left = ParagraphStyle(
        'CellLeft', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=8.5, alignment=0
    )
    cell_right = ParagraphStyle(
        'CellRight', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=8.5, alignment=2
    )

    # ==================== PAGE 1 ====================
    story.append(Paragraph("KIST DAY CASH SHEET", title_style))
    story.append(Paragraph("Date : ___________________    Route : ___________________    No.Bill : ___________________", meta_style))
    
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
    
    # Column measurements tuning - Invoice/Shop contracted; Cash/Credit/Cheque expanded
    # Total printable width is 562 points (Letter width 612 - 50 margins)
    col_widths = [20, 50, 60, 55, 30, 36, 36, 30, 65, 65, 65, 50]
    
    # Ultra-compressed padding settings (TOP/BOTTOM padding to 2pt maps precisely to digit heights)
    main_table_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#ebf2f8")), # highlight ST row
    ])
    
    story.append(Table(table_data, colWidths=col_widths, style=main_table_style))
    story.append(Spacer(1, 6))
    
    sales_labels = [
        "Biscuits sale : _______________________", "Nectar Sale : _______________________",
        "Water Sale : _______________________", "Total Sale : _______________________"
    ]
    for label in sales_labels:
        story.append(Paragraph(label, meta_style))
        
    # ==================== PAGE 2 (STRICT LAYOUT BOUNDARY) ====================
    story.append(PageBreak())
    
    story.append(Paragraph("Cash Receivables", section_style))
    rec_headers = ["NO", "Bill Date", "Shop", "Credit Amount", "Pay Amount", "Balance"]
    rec_data = [[Paragraph(rh, cell_bold) for rh in rec_headers]]
    for r_idx in range(1, 11):
        rec_data.append([Paragraph(str(r_idx), cell_style)] + [Paragraph("", cell_style) for _ in range(5)])
        
    rec_widths = [30, 80, 202, 85, 85, 80] # Total 562
    rec_table_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ])
    story.append(Table(rec_data, colWidths=rec_widths, style=rec_table_style))
    story.append(Spacer(1, 6))
    
    story.append(Paragraph("Cash Sheet Balancing", section_style))
    
    # Build System Balancing (Left side metrics block)
    left_rows = [
        ("System Sale", f"{total_amount:,.2f}"), ("FOC", ""), ("Total Cancel", ""), ("Balance (1)", ""),
        ("Total Discounts", ""), ("Total Adjust", ""), ("Total Return", ""), ("Balance (2)", ""),
        ("Total Cash", ""), ("Total Credit", ""), ("Total Cheques", ""), ("Balance (3)", "")
    ]
    left_table_data = []
    left_style_actions = []
    for r_idx, (item, val) in enumerate(left_rows):
        is_bal = "Balance" in item or item == "System Sale"
        p_style = cell_bold if is_bal else cell_left
        v_style = ParagraphStyle('v', parent=cell_right, fontName='Helvetica-Bold' if is_bal else 'Helvetica')
        left_table_data.append([Paragraph(item, p_style), Paragraph(val, v_style)])
        if "Balance" in item:
            left_style_actions.append(('BACKGROUND', (0, r_idx), (1, r_idx), colors.HexColor("#f0f4f8")))
            
    # Build Cash Balance (Right side tracking block)
    right_table_data = [
        [Paragraph("Cash Balance", cell_bold), Paragraph("", cell_style)],
        [Paragraph("Total Day Cash", cell_left), Paragraph("", cell_style)],
        [Paragraph("Total Credit Received", cell_left), Paragraph("", cell_style)],
        [Paragraph("Total Expenses", cell_left), Paragraph("", cell_style)],
        [Paragraph("Banked Value.", cell_left), Paragraph("", cell_style)]
    ]
    
    # side-by-side composite grid alignment (Left width: 275pt, Right width: 275pt + 12pt clear spacer span)
    left_table = Table(left_table_data, colWidths=[165, 110])
    left_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ('TOPPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ] + left_style_actions))
    
    right_table = Table(right_table_data, colWidths=[165, 110])
    right_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('TOPPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,3), (1,3), 32), # Height spacer specifically configured for Manual Expense Line
    ]))
    
    master_balancing_table = Table([[left_table, right_table]], colWidths=[278, 284])
    master_balancing_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(master_balancing_table)
    story.append(Spacer(1, 4))
    
    story.append(Paragraph("Calculating cash", section_style))
    denom_headers = [Paragraph("Cash Analitics", cell_bold), Paragraph("Valuve", cell_bold)]
    denom_data = [denom_headers]
    denominations = ["20", "50", "100", "500", "1000", "2000", "5000", "coins", "Total"]
    for denom in denominations:
        p_style = cell_bold if denom == "Total" else cell_style
        denom_data.append([Paragraph(denom, p_style), Paragraph("", cell_style)])
        
    denom_table = Table(denom_data, colWidths=[281, 281])
    denom_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ('TOPPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#ebf2f8")),
    ]))
    story.append(denom_table)
    story.append(Spacer(1, 6))
    
    story.append(Paragraph("Distance Travelled : ___________________    KM : ___________________    OOT : ___________________", meta_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# --- Native pdfplumber Raw Text Extraction Engine ---
def parse_pdf_file(uploaded_file):
    invoices = []
    grand_total = 0.0
    full_text = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text_content = page.extract_text()
            if text_content: full_text += text_content + "\n"

    invoice_section = re.search(r'Net Amt \(LKR\)(.*)', full_text, re.DOTALL)
    if not invoice_section: return [], 0.0

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

# --- Main App Interface ---
# Injects explicit inline styles to force both the title and label text to render completely black
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

        with st.spinner("Compiling ultra-compact printable document structure..."):
            pdf_data = generate_cash_sheet_pdf(invoices, total_amt)

        st.download_button(
            label="📥 Download Tailored Cash Sheet PDF",
            data=pdf_data,
            file_name="KIST_Day_Cash_Sheet.pdf",
            mime="application/pdf"
        )
    else:
        st.error("Could not parse invoices structures. Please verify that the PDF contains valid invoice entries.")
