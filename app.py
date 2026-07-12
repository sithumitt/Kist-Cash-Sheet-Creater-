import io
import os
import re
import streamlit as st
import pdfplumber

# ReportLab core engines for precise programmatic PDF layout
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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

# =====================================================================
# SINHALA FONT REGISTRATION
# The sample Word document uses Sinhala labels/headers, so we need a
# Unicode font that actually contains Sinhala glyphs (Helvetica does
# not). Put NotoSansSinhala-Regular.ttf / -Bold.ttf in a "fonts/"
# folder next to this script (both files are provided alongside this
# code). We fall back to a couple of common system locations too.
# =====================================================================
FONT_REGULAR_NAME = "SinhalaFont"
FONT_BOLD_NAME = "SinhalaFont-Bold"

_CANDIDATE_DIRS = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts"),
    "/usr/share/fonts/truetype/sinhala",
]

def _find_font(filename_options):
    for d in _CANDIDATE_DIRS:
        for name in filename_options:
            path = os.path.join(d, name)
            if os.path.isfile(path):
                return path
    return None

_regular_path = _find_font(["NotoSansSinhala-Regular.ttf", "lklug.ttf"])
_bold_path = _find_font(["NotoSansSinhala-Bold.ttf", "lklug.ttf"])

if _regular_path:
    pdfmetrics.registerFont(TTFont(FONT_REGULAR_NAME, _regular_path))
else:
    FONT_REGULAR_NAME = "Helvetica"  # last-resort fallback (Sinhala glyphs will be missing)

if _bold_path:
    pdfmetrics.registerFont(TTFont(FONT_BOLD_NAME, _bold_path))
else:
    FONT_BOLD_NAME = "Helvetica-Bold"

# Sinhala "pre-base" vowel signs (ෙ, ේ, ෛ) are stored AFTER their base
# consonant in Unicode but must be drawn BEFORE it. ReportLab has no
# complex-script shaping engine, so we reorder them ourselves.
_PREBASE_VOWELS = "\u0DD9\u0DDA\u0DDB"  # ෙ ේ ෛ
_reorder_pattern = re.compile(r'(.)([' + _PREBASE_VOWELS + r'])')

def sn(text):
    """Fix up Sinhala pre-base vowel sign order for correct rendering."""
    if not text:
        return text
    return _reorder_pattern.sub(lambda m: m.group(2) + m.group(1), text)


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
        'TitleStyle', parent=styles['Heading1'], fontName=FONT_BOLD_NAME,
        fontSize=16, leading=20, alignment=1, spaceAfter=6, textColor=colors.HexColor("#000000")
    )
    meta_style = ParagraphStyle(
        'MetaStyle', parent=styles['Normal'], fontName=FONT_BOLD_NAME,
        fontSize=9.5, leading=13, spaceAfter=10, textColor=colors.HexColor("#000000")
    )
    section_style = ParagraphStyle(
        'SectionStyle', parent=styles['Normal'], fontName=FONT_BOLD_NAME,
        fontSize=10, leading=12, spaceBefore=6, spaceAfter=4, textColor=colors.HexColor("#000000")
    )
    cell_style = ParagraphStyle(
        'CellStyle', parent=styles['Normal'], fontName=FONT_REGULAR_NAME,
        fontSize=8, leading=10, alignment=1
    )
    cell_bold = ParagraphStyle(
        'CellBold', parent=styles['Normal'], fontName=FONT_BOLD_NAME,
        fontSize=8.5, leading=10.5, alignment=1
    )
    cell_right = ParagraphStyle(
        'CellRight', parent=styles['Normal'], fontName=FONT_REGULAR_NAME,
        fontSize=8.5, leading=10.5, alignment=2
    )
    label_style = ParagraphStyle(
        'LabelStyle', parent=styles['Normal'], fontName=FONT_REGULAR_NAME,
        fontSize=7.8, leading=9.5, alignment=0, leftIndent=4
    )
    label_bold_style = ParagraphStyle(
        'LabelBoldStyle', parent=styles['Normal'], fontName=FONT_BOLD_NAME,
        fontSize=7.8, leading=9.5, alignment=0, leftIndent=4
    )

    # ==================== PAGE 1 : DAY CASH SHEET ====================
    body_elements.append(Paragraph("KIST DAY CASH SHEET", title_style))
    body_elements.append(Paragraph(
        sn("දිනය : ___________________&nbsp;&nbsp;&nbsp;මාර්ගය : ___________________&nbsp;&nbsp;&nbsp;බිල්පත් ගණන : ___________________"),
        meta_style
    ))

    # Column headers exactly as they appear in the KIST sample Word document
    headers = [
        "No", sn("බිල් අංකය"), sn("කඩේ නම"), sn("ප්‍රමාණය"), "BNW",
        sn("කෑන්සල්"), sn("ඇඩ්ජස්ට්"), sn("ඩීස්"), sn("මුදල්"),
        sn("ණය"), sn("චෙක්"), sn("රිටන්")
    ]
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
        Paragraph("ST", cell_bold), Paragraph(sn("System Sale"), cell_bold), Paragraph("", cell_style),
        Paragraph(f"{total_amount:,.2f}", ParagraphStyle('ST_R', parent=cell_right, fontName=FONT_BOLD_NAME)),
        Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style),
        Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style), Paragraph("", cell_style)
    ])

    # Perfectly Balanced Explicit Point Column Widths (Total width fits A4 printable profile seamlessly)
    col_widths = [
        22,  # No
        62,  # Invoice Number (Perfect for clean non-wrapped header display)
        90,  # Shop Name (Maintains an open handwriting buffer space)
        50,  # Amount
        32,  # BNW (Balanced for up to 4 digits without stacking header text)
        42,  # Cancel (Stops breaking into 'Cance l')
        42,  # Adjust
        30,  # Dis
        45,  # Cash
        45,  # Credit
        45,  # Cheque
        32   # Rtn (Stops breaking into stacked individual characters)
    ]

    main_table_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#aaaaaa")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 5.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#ebf2f8")),
    ])

    body_elements.append(Table(table_data, colWidths=col_widths, style=main_table_style))
    body_elements.append(Spacer(1, 10))

    sales_labels = [
        sn("බිස්කට් සෙල් : _______________________"),
        sn("නෙට්ටා සෙල් : _______________________"),
        sn("වතුර සෙල් : _______________________"),
        sn("මුළු සේල් : _______________________"),
    ]
    for label in sales_labels:
        body_elements.append(Paragraph(label, meta_style))

    # ==================== PAGE 2 ====================
    body_elements.append(PageBreak())
    page2_elements = build_page_two(
        section_style, cell_style, cell_bold, cell_right,
        label_style, label_bold_style, meta_style
    )
    body_elements.extend(page2_elements)

    doc.build(body_elements)
    buffer.seek(0)
    return buffer


def build_page_two(section_style, cell_style, cell_bold, cell_right,
                    label_style, label_bold_style, meta_style):
    """Builds the second page of the KIST cash sheet:
       1) Credit-bill payments-received table
       2) Cash balancing summary table
       3) Cash-note counting table
       4) Distance-travelled footer line
       Mirrors the layout of the supplied sample Word document.
    """
    elements = []

    # ---- 1) ණය බිල් වලට මුදල් ලැබීම (Credit bill payments received) ----
    elements.append(Paragraph(sn("ණය බිල් වලට මුදල් ලැබීම"), section_style))

    credit_headers = [
        "NO", sn("බිල්පත් දිනය"), sn("කඩේ නම"), sn("ණය"),
        sn("මුදල් ගෙවූ ප්‍රමාණය"), sn("ඉතිරිය")
    ]
    credit_table_data = [[Paragraph(h, cell_bold) for h in credit_headers]]
    for i in range(1, 11):
        credit_table_data.append(
            [Paragraph(str(i), cell_style)] + [Paragraph("", cell_style) for _ in range(5)]
        )

    credit_col_widths = [28, 75, 140, 65, 90, 85]
    credit_style = TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#aaaaaa")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ])
    elements.append(Table(credit_table_data, colWidths=credit_col_widths, style=credit_style))
    elements.append(Spacer(1, 8))

    # ---- 2) මුදල් පත්‍ර බැලන්ස් කිරීම (Cash balancing) ----
    elements.append(Paragraph(sn("මුදල් පත්‍ර බැලන්ස් කිරීම"), section_style))

    balance_rows_labels = [
        (sn("සිස්ටම් සෙල් එක"), True),
        (sn("FOC"), False),
        (sn("මුළු කැන්සල්"), False),
        (sn("ශේෂය (1)"), True),
        (sn("මුළු වට්ටම්"), False),
        (sn("මුළු ඇඩ්ජස්ට්"), False),
        (sn("මුළු රිටන්"), False),
        (sn("ශේෂය (2)"), True),
        (sn("මුළු මුදල්"), False),
        (sn("මුළු ණය"), False),
        (sn("මුළු චෙක්පත්"), False),
        (sn("ශේෂය (3)"), True),
        (sn("මුදල් ශේෂය"), True),
        (sn("මුළු දින මුදල්"), False),
        (sn("ලැබුණු මුළු ණය"), False),
        (sn("මුළු වියදම්"), False),
        (sn("බැංකුගත වටිනාකම"), False),
    ]
    balance_table_data = []
    for label, is_bold in balance_rows_labels:
        style = label_bold_style if is_bold else label_style
        balance_table_data.append([Paragraph(label, style), Paragraph("", cell_style)])

    balance_col_widths = [300, 183]
    balance_style = TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#aaaaaa")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2.2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.2),
    ])
    elements.append(Table(balance_table_data, colWidths=balance_col_widths, style=balance_style))
    elements.append(Spacer(1, 8))

    # ---- 3) මුදල් ගණනය කිරීම (Cash note counting) ----
    elements.append(Paragraph(sn("මුදල් ගණනය කිරීම"), section_style))

    note_headers = [sn("මුදල් නෝට්ටු"), sn("වටිනාකම")]
    note_values = ["20", "50", "100", "500", "1000", "2000", "5000", sn("coins"), sn("Total")]
    note_table_data = [[Paragraph(h, cell_bold) for h in note_headers]]
    for i, note in enumerate(note_values):
        style = cell_bold if note == sn("Total") else cell_style
        note_table_data.append([Paragraph(note, style), Paragraph("", cell_style)])

    note_col_widths = [240, 240]
    note_style = TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#aaaaaa")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#ebf2f8")),
    ])
    elements.append(Table(note_table_data, colWidths=note_col_widths, style=note_style))
    elements.append(Spacer(1, 8))

    # ---- 4) Footer: distance travelled ----
    elements.append(Paragraph(
        sn("ගමන් කළ දුර : ___________ &nbsp;&nbsp;&nbsp; KM : ___________ &nbsp;&nbsp;&nbsp; OOT : ___________"),
        meta_style
    ))

    return elements


# --- Stateful Line-by-Line Multi-Format Text Extraction Engine ---
def parse_pdf_file(uploaded_file):
    invoices = []
    total_amount = 0.0
    invoice_seen = set()

    composite_invoice_pattern = re.compile(r'\b\d{2}[A-Z]{3}_\d{4,}[\w\d_]*\b')
    standard_invoice_pattern = re.compile(r'\b(IN|TI)\d{5,7}\b')
    isolated_id_pattern = re.compile(r'^\b\d{3,4}\b$')
    amount_pattern = re.compile(r'\d[\d,]*\.\d{2}')

    with pdfplumber.open(uploaded_file) as pdf:
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

                if line_str.lower().startswith("total") or "grand total" in line_str.lower():
                    amts = amount_pattern.findall(line_str)
                    if amts:
                        try:
                            total_amount = float(amts[-1].replace(",", ""))
                        except ValueError:
                            pass
                    continue

                comp_match = composite_invoice_pattern.search(line_str)
                if comp_match:
                    active_prefix = comp_match.group()
                    line_remainder = line_str.replace(active_prefix, "").strip()
                    remainder_numbers = re.findall(r'\b\d{3,4}\b', line_remainder)
                    if remainder_numbers:
                        active_sub_id = remainder_numbers[0]
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
                    if "net amt" in line_str.lower() or "mrp" in line_str.lower():
                        continue

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

                            invoices.append({
                                "invoice": final_invoice_slice,
                                "amount": amt_val
                            })
                            invoice_seen.add(final_invoice_slice)

                        active_sub_id = None
                        if active_prefix.startswith(("IN", "TI")):
                            active_prefix = None
                    except ValueError:
                        pass

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
