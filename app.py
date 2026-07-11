import io
import re
import streamlit as st
import pdfplumber

# ReportLab core engines for precise programmatic PDF layout
from reportlab.lib.pagesizes import A4
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
        fontSize=16 if is_large_dataset else 14, leading=20 if is_large_dataset else 16, 
        alignment=1, spaceAfter=6, textColor=colors.HexColor("#000000")
    )
    meta_style = ParagraphStyle(
        'MetaStyle', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=9.5 if is_large_dataset else 8.5, leading=13 if is_large_dataset else 10, 
        spaceAfter=10 if is_large_dataset else 6, textColor=colors.HexColor("#000000")
    )
    section_style = ParagraphStyle(
        'SectionStyle', parent=styles['Heading2'], fontName='Helvetica-Bold',
        fontSize=12 if is_large_dataset else 11, leading=15 if is_large_dataset else 13, 
        spaceBefore=8 if is_large_dataset else 4, spaceAfter=6 if is_large_dataset else 4, textColor=colors.HexColor("#000000")
    )
    cell_style = ParagraphStyle(
        'CellStyle', parent=styles['Normal'], fontName='Helvetica', 
        fontSize=8.5 if is_large_dataset else 7.5, leading=10.5 if is_large_dataset else 8.5, alignment=1
    )
    cell_bold = ParagraphStyle(
        'CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', 
        fontSize=8.5 if is_large_dataset else 7.5, leading=10.5 if is_large_dataset else 8.5, alignment=1
    )
    cell_left = ParagraphStyle(
        'CellLeft', parent=styles['Normal'], fontName='Helvetica', 
        fontSize=8.5 if is_large_dataset else 7.5, leading=10.5 if is_large_dataset else 8.5, alignment=0
    )
    cell_right = ParagraphStyle(
        'CellRight', parent=styles['Normal'], fontName='Helvetica', 
        fontSize=8.5 if is_large_dataset else 7.5, leading=10.5 if is_large_dataset else 8.5, alignment=2
    )

    # ==================== PAGE 1 ====================
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
    
    # Printable horizontal width mapping proportions 
    printable_width = 523 if is_large_dataset else 543
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
    
    padding_val = 5.5 if is_large_dataset else 2.5
    main_table_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), padding_val),
        ('BOTTOMPADDING', (0,0), (-1,-1), padding_val),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#ebf2f8")), 
    ])
    
    body_elements.append(Table(table_data, colWidths=col_widths, style=main_table_style))
    body_elements.append(Spacer(1, 10 if is_large_dataset else 6))
    
    sales_labels = [
        "Biscuits sale : _______________________", "Nectar Sale : _______________________",
        "Water Sale : _______________________", "Total Sale : _______________________"
    ]
    for label in sales_labels:
        body_elements.append(Paragraph(label, meta_style))
        
    # ==================== SEGREGATION BOUNDARY ====================
    body_elements.append(PageBreak())
    
    body_elements.append(Paragraph("Cash Receivables", section_style))
    rec_headers = ["NO", "Bill Date", "Shop", "Credit Amount", "Pay Amount", "Balance"]
    rec_data = [[Paragraph(rh, cell_bold) for rh in rec_headers]]
    for r_idx in range(1, 11):
        rec_data.append([Paragraph(str(r_idx), cell_style)] + [Paragraph("", cell_style) for _ in range(5)])
        
    rec_widths = [
        printable_width * 0.06,
        printable_width * 0.14,
        printable_width * 0.35,
        printable_width * 0.15,
        printable_width * 0.15,
        printable_width * 0.15
    ]
    rec_padding = 4.5 if is_large_dataset else 2.5
    rec_table_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ('TOPPADDING', (0,0), (-1,-1), rec_padding),
        ('BOTTOMPADDING', (0,0), (-1,-1), rec_padding),
    ])
    body_elements.append(Table(rec_data, colWidths=rec_widths, style=rec_table_style))
    body_elements.append(Spacer(1, 8 if is_large_dataset else 5))
    
    body_elements.append(Paragraph("Cash Sheet Balancing", section_style))
    
    left_rows = [
        ("System Sale", f"{total_amount:,.2f}"), ("FOC", ""), ("Total Cancel", ""), ("Balance (1)", ""),
        ("Total Discounts", ""), ("Total Adjust", ""), ("Total Return", ""), ("Balance (2)", ""),
        ("Total Cash", ""), ("Total Credit", ""), ("Total Cheques", ""), ("Balance (3)", "")
    ]
    
    right_rows_data = [
        ("Cash Balance", ""),
        ("Total Day Cash", ""),
        ("Total Credit Received", ""),
        ("Total Expenses", ""),
        ("Banked Value.", "")
    ]

    if is_large_dataset:
        left_table_data = []
        left_style_actions = []
        for r_idx, (item, val) in enumerate(left_rows):
            is_bal = "Balance" in item or item == "System Sale"
            p_style = cell_bold if is_bal else cell_left
            v_style = ParagraphStyle('v', parent=cell_right, fontName='Helvetica-Bold' if is_bal else 'Helvetica', fontSize=8.5)
            left_table_data.append([Paragraph(item, p_style), Paragraph(val, v_style)])
            if "Balance" in item:
                left_style_actions.append(('BACKGROUND', (0, r_idx), (1, r_idx), colors.HexColor("#f0f4f8")))
                
        right_table_data = []
        for r_idx, (item, val) in enumerate(right_rows_data):
            p_style = cell_bold if r_idx == 0 else cell_left
            right_table_data.append([Paragraph(item, p_style), Paragraph(val, cell_style)])

        w_left = printable_width * 0.49
        w_right = printable_width * 0.49
        
        left_table = Table(left_table_data, colWidths=[w_left*0.58, w_left*0.42])
        left_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
            ('TOPPADDING', (0,0), (-1,-1), 4.5), ('BOTTOMPADDING', (0,0), (-1,-1), 4.5),
        ] + left_style_actions))
        
        right_table = Table(right_table_data, colWidths=[w_right*0.58, w_right*0.42])
        right_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
            ('BACKGROUND', (0,0), (-1,0), colors.white),
            ('TOPPADDING', (0,0), (-1,-1), 4.5), ('BOTTOMPADDING', (0,0), (-1,-1), 4.5),
            ('BOTTOMPADDING', (0,3), (1,3), 36),
        ]))
        
        master_balancing_table = Table([[left_table, right_table]], colWidths=[printable_width*0.50, printable_width*0.50])
        master_balancing_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        body_elements.append(master_balancing_table)
    else:
        bal_table_data = []
        bal_style_actions = []
        combined_rows = []
        for item, val in left_rows:
            combined_rows.append((item, val, "Balance" in item or item == "System Sale", False))
        for r_idx, (item, val) in enumerate(right_rows_data):
            combined_rows.append((item, val, False, r_idx == 0))
            
        for r_idx, (item, val, is_bal, is_hdr) in enumerate(combined_rows):
            if is_hdr:
                p_style = cell_bold
                v_style = cell_style
                bal_style_actions.append(('BACKGROUND', (0, r_idx), (1, r_idx), colors.white))
            elif is_bal:
                p_style = cell_bold
                v_style = ParagraphStyle('v', parent=cell_right, fontName='Helvetica-Bold', fontSize=7.5)
                bal_style_actions.append(('BACKGROUND', (0, r_idx), (1, r_idx), colors.HexColor("#f0f4f8")))
            else:
                p_style = cell_left
                v_style = cell_right
                
            bal_table_data.append([Paragraph(item, p_style), Paragraph(val, v_style)])
            if item == "Total Expenses":
                bal_style_actions.append(('BOTTOMPADDING', (0, r_idx), (1, r_idx), 26))
                
        bal_table = Table(bal_table_data, colWidths=[printable_width * 0.58, printable_width * 0.42])
        bal_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ] + bal_style_actions))
        body_elements.append(bal_table)

    body_elements.append(Spacer(1, 8 if is_large_dataset else 5))
    
    body_elements.append(Paragraph("Calculating cash", section_style))
    denom_headers = [Paragraph("Cash Analitics", cell_bold), Paragraph("Valuve", cell_bold)]
    denom_data = [denom_headers]
    denominations = ["20", "50", "100", "500", "1000", "2000", "5000", "coins", "Total"]
    for denom in denominations:
        p_style = cell_bold if denom == "Total" else cell_style
        denom_data.append([Paragraph(denom, p_style), Paragraph("", cell_style)])
        
    denom_table = Table(denom_data, colWidths=[printable_width*0.50, printable_width*0.50])
    denom_padding = 3.5 if is_large_dataset else 2.2
    denom_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ('TOPPADDING', (0,0), (-1,-1), denom_padding), ('BOTTOMPADDING', (0,0), (-1,-1), denom_padding),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#ebf2f8")),
    ]))
    body_elements.append(denom_table)
    body_elements.append(Spacer(1, 10 if is_large_dataset else 6))
    
    body_elements.append(Paragraph("Distance Travelled : ___________________    KM : ___________________    OOT : ___________________", meta_style))
    
    doc.build(body_elements)
    buffer.seek(0)
    return buffer

# --- Universal Line-by-Line Multi-Format Text Extraction Engine ---
def parse_pdf_file(uploaded_file):
    invoices = []
    grand_total = 0.0
    full_text = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text_content = page.extract_text()
            if text_content: 
                full_text += text_content + "\n"

    # Split invoice content segments following the Net Amt column block layout
    invoice_section = re.search(r'Net Amt \(LKR\)(.*)', full_text, re.DOTALL)
    if not invoice_section: 
        return [], 0.0

    invoice_block_text = invoice_section.group(1)
    lines = invoice_block_text.split('\n')
    
    # Pre-compiled regex pattern to capture values at the end of valid ledger rows
    amount_pattern = re.compile(r'([\d,]+\.\d{2})\s*$')
    
    # Sliding History Buffer to look back into multiline outputs safely
    history_lines = []

    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
            
        history_lines.append(line_str)
        # Keep only the last 4 rows in active loop memory context
        if len(history_lines) > 4:
            history_lines.pop(0)
            
        # Parse the money value at the right end of the active row string
        amt_match = amount_pattern.search(line_str)
        if amt_match:
            amt_text = amt_match.group(1)
            
            # Skip overall summary total lines
            if "total" in line_str.lower():
                if "grand total" in line_str.lower() or line_str.lower().startswith("total"):
                    try:
                        grand_total = float(amt_text.replace(',', ''))
                    except ValueError:
                        pass
                continue
                
            # Fallback evaluation context targets history lines to grab the segmented invoice numbers
            final_invoice = None
            
            # Walk backward from current row index to evaluate previous string segments
            for hist_line in reversed(history_lines):
                # Clean out current amount field to prevent extracting its decimal integers
                clean_hist = hist_line.replace(amt_text, '').strip()
                
                # Capture digits clusters
                numbers = re.findall(r'\d+', clean_hist)
                if numbers:
                    # Find the primary sequential identifier (longer digit sets take priority over row indexes)
                    candidate = max(numbers, key=len)
                    # Slices strictly the last 4 digits
                    if len(candidate) >= 3:
                        final_invoice = candidate[-4:]
                        break
            
            if final_invoice:
                try:
                    amt_val = float(amt_text.replace(',', ''))
                    # Protect data array elements from duplicated value injection rows
                    if not any(i['invoice'] == final_invoice for i in invoices):
                        invoices.append({'invoice': final_invoice, 'amount': amt_val})
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
