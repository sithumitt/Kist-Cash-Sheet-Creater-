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
    
    # Full horizontal layout width span configurations (Total 523 points)
    printable_width = 523
    col_widths = [
        printable_width * 0.04,  # No
        printable_width * 0.24,  # Expanded Invoice Number space for long tracking IDs
        printable_width * 0.05,  # Shop Name
        printable_width * 0.10,  # Amount
        printable_width * 0.04,  # BNW
        printable_width * 0.06,  # Cancel
        printable_width * 0.06,  # Adjust
        printable_width * 0.04,  # Dis
        printable_width * 0.11,  # Cash
        printable_width * 0.11,  # Credit
        printable_width * 0.11,  # Cheque
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

# --- Stateful Line-by-Line Multi-Format Text Extraction Engine ---
def parse_pdf_file(uploaded_file):
    invoices = []
    total_amount = 0.0
    invoice_seen = set()

    # Pattern 1: Catch composite complex prefix dates (e.g. 26JUL_1201_17300100009)
    composite_invoice_pattern = re.compile(r'\b\d{2}[A-Z]{3}_\d{4,}[\w\d_]*\b')
    # Pattern 2: Catch standard prefixes (e.g. IN008868, TI009403)
    standard_invoice_pattern = re.compile(r'\b(IN|TI)\d{5,7}\b')
    # Pattern 3: Catch floating isolated identifier lines (e.g. 368 or 387)
    isolated_id_pattern = re.compile(r'^\b\d{3,4}\b$')
    # Pattern 4: Strict regex filter for line monetary values at trailing edges
    amount_pattern = re.compile(r'\d[\d,]*\.\d{2}')

    with pdfplumber.open(uploaded_file) as pdf:
        # State tracking registers for multi-line bridging
        active_prefix = None
        active_sub_id = None

        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = text.splitlines()
            for line in lines:
                line_str = line.strip()
                if not line_str:
                    continue

                # Parse and track overall grand summary numbers from page footers
                if line_str.lower().startswith("total") or "grand total" in line_str.lower():
                    amts = amount_pattern.findall(line_str)
                    if amts:
                        try:
                            total_amount = float(amts[-1].replace(",", ""))
                        except ValueError:
                            pass
                    continue

                # State Step 1: Detect composite string rows (e.g. 26JUL_1201_17300100009)
                comp_match = composite_invoice_pattern.search(line_str)
                if comp_match:
                    active_prefix = comp_match.group()
                    # Check if sub-identifiers are combined on the same string block row
                    line_remainder = line_str.replace(active_prefix, "").strip()
                    remainder_numbers = re.findall(r'\b\d{3,4}\b', line_remainder)
                    if remainder_numbers:
                        active_sub_id = remainder_numbers[0]
                    continue

                # State Step 2: Catch standalone standard structures (e.g. IN008868)
                std_match = standard_invoice_pattern.search(line_str)
                if std_match:
                    active_prefix = std_match.group()
                    active_sub_id = None  # No sub-ID matching layer needed for clean alpha codes

                # State Step 3: Catch floating numeric sequence splits (e.g. 368 / 387 line segments)
                iso_match = isolated_id_pattern.search(line_str)
                if iso_match and active_prefix and not active_prefix.startswith(("IN", "TI")):
                    active_sub_id = iso_match.group()
                    continue

                # State Step 4: Final ledger row value confirmation mapping
                amts = amount_pattern.findall(line_str)
                if amts and active_prefix:
                    # Guard selection to ensure parsing context doesn't read pricing headers
                    if "net amt" in line_str.lower() or "mrp" in line_str.lower():
                        continue

                    # Select the definitive identifier string block value
                    if active_prefix.startswith(("IN", "TI")):
                        invoice_identity = active_prefix
                    elif active_sub_id:
                        invoice_identity = f"{active_prefix}\n{active_sub_id}"
                    else:
                        invoice_identity = active_prefix

                    try:
                        amt_val = float(amts[-1].replace(",", ""))
                        
                        # De-duplicate entries and drop any empty/corrupted noise strings
                        if invoice_identity not in invoice_seen and len(invoice_identity) > 2:
                            invoices.append({
                                "invoice": invoice_identity,
                                "amount": amt_val
                            })
                            invoice_seen.add(invoice_identity)
                            
                        # Reset tracking states for the next loop sequence block iteration
                        active_sub_id = None
                        # Keep active_prefix as fallback unless standard invoice matched
                        if active_prefix.startswith(("IN", "TI")):
                            active_prefix = None
                    except ValueError:
                        pass

    # Safe math computation fallback
    if total_amount == 0.0 and invoices:
        total_amount = sum(i["amount"] for i in invoices)

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
