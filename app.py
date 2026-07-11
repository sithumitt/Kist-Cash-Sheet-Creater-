import io
import re
import streamlit as st
import pdfplumber
from weasyprint import HTML

# --- Custom Styling & Theme Configuration ---
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

# --- PDF Generation Function (HTML to WeasyPrint PDF Workflow) ---
def generate_cash_sheet_pdf(invoices, total_amount):
    # Constructing rows for parsed data dynamically
    invoice_rows_html = ""
    idx = 1
    for item in invoices:
        invoice_rows_html += f"""
        <tr>
            <td>{idx}</td>
            <td>{item['invoice']}</td>
            <td></td>
            <td class="align-right">{item['amount']:.2f}</td>
            <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
        </tr>"""
        idx += 1
        
    # Standard dynamic empty lines for manual ledger input additions
    for _ in range(5):
        invoice_rows_html += f"""
        <tr>
            <td>{idx}</td>
            <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
        </tr>"""
        idx += 1

    html_template = f"""<!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
    @page {{
        size: A4 portrait;
        margin: 8mm 8mm;
        background-color: #ffffff;
    }}
    * {{ box-sizing: border-box; }}
    body {{
        margin: 0; padding: 0;
        font-family: 'Arial', sans-serif;
        color: #000000; font-size: 7.5pt;
        line-height: 1.1;
    }}
    .title {{
        text-align: center; font-size: 13pt; font-weight: bold; margin-bottom: 3px;
    }}
    .meta-info {{ font-size: 8.5pt; margin-bottom: 6px; font-weight: bold; }}
    table {{
        width: 100%; border-collapse: collapse; margin-bottom: 6px; table-layout: fixed;
    }}
    th, td {{
        border: 0.5px solid #aaaaaa; padding: 2px 2px; text-align: center;
        vertical-align: middle; overflow: hidden; white-space: nowrap;
    }}
    th {{ background-color: #ffffff; font-weight: bold; color: #000000; font-size: 7.5pt; }}
    .align-left {{ text-align: left; padding-left: 3px; }}
    .align-right {{ text-align: right; padding-right: 3px; }}
    .bg-st {{ background-color: #ebf2f8; font-weight: bold; }}
    .bg-balance {{ background-color: #f0f4f8; font-weight: bold; }}
    
    /* Configured layout widths matching prompt requirements */
    .col-no {{ width: 4.5%; }}
    .col-inv {{ width: 10%; }}      /* Reduced width */
    .col-shop {{ width: 11.5%; }}    /* Reduced width */
    .col-amt {{ width: 10%; }}      /* Reduced width */
    .col-bnw {{ width: 5%; }}
    .col-cancel {{ width: 6.5%; }}
    .col-adjust {{ width: 6.5%; }}
    .col-dis {{ width: 5.5%; }}
    .col-cash {{ width: 11.5%; }}   /* Expanded width */
    .col-credit {{ width: 11.5%; }} /* Expanded width */
    .col-cheque {{ width: 11.5%; }} /* Expanded width */
    .col-rtn {{ width: 6.5%; }}

    .sales-summary {{ font-size: 8.5pt; font-weight: bold; margin-top: 5px; margin-bottom: 12px; }}
    .sales-line {{ margin-bottom: 3px; }}
    .page-break {{ page-break-before: always; }}
    .section-title {{ font-size: 10pt; font-weight: bold; margin-top: 2px; margin-bottom: 4px; }}
    .flex-table-container {{ display: table; width: 100%; }}
    .flex-table-cell {{ display: table-cell; width: 50%; vertical-align: top; }}
    .flex-table-cell:first-child {{ padding-right: 6px; }}
    .flex-table-cell:last-child {{ padding-left: 6px; }}
    </style>
    </head>
    <body>

        <!-- ================= PAGE 1 ================= -->
        <div class="title">KIST DAY CASH SHEET</div>
        <div class="meta-info">
            Date : ___________________ &nbsp;&nbsp;&nbsp;&nbsp; Route : ___________________ &nbsp;&nbsp;&nbsp;&nbsp; No.Bill : ___________________
        </div>

        <table>
            <thead>
                <tr>
                    <th class="col-no">No</th>
                    <th class="col-inv">Invoice Number</th>
                    <th class="col-shop">Shop Name</th>
                    <th class="col-amt">Amount</th>
                    <th class="col-bnw">BNW</th>
                    <th class="col-cancel">Cancel</th>
                    <th class="col-adjust">Adjust</th>
                    <th class="col-dis">Dis</th>
                    <th class="col-cash">Cash</th>
                    <th class="col-credit">Credit</th>
                    <th class="col-cheque">Cheque</th>
                    <th class="col-rtn">Rtn</th>
                </tr>
            </thead>
            <tbody>
                {invoice_rows_html}
                <tr class="bg-st">
                    <td>ST</td>
                    <td>System Sale</td>
                    <td></td>
                    <td class="align-right">{total_amount:,.2f}</td>
                    <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
                </tr>
            </tbody>
        </table>

        <div class="sales-summary">
            <div class="sales-line">Biscuits sale : _______________________</div>
            <div class="sales-line">Nectar Sale : _______________________</div>
            <div class="sales-line">Water Sale : _______________________</div>
            <div class="sales-line">Total Sale : _______________________</div>
        </div>

        <!-- ================= PAGE 2 ================= -->
        <div class="page-break"></div>
        
        <div class="section-title">Cash Receivables</div>
        <table>
            <thead>
                <tr>
                    <th style="width: 6%;">NO</th>
                    <th style="width: 15%;">Bill Date</th>
                    <th style="width: 34%;">Shop</th>
                    <th style="width: 15%;">Credit Amount</th>
                    <th style="width: 15%;">Pay Amount</th>
                    <th style="width: 15%;">Balance</th>
                </tr>
            </thead>
            <tbody>
                <tr><td>1</td><td></td><td></td><td></td><td></td><td></td></tr>
                <tr><td>2</td><td></td><td></td><td></td><td></td><td></td></tr>
                <tr><td>3</td><td></td><td></td><td></td><td></td><td></td></tr>
                <tr><td>4</td><td></td><td></td><td></td><td></td><td></td></tr>
                <tr><td>5</td><td></td><td></td><td></td><td></td><td></td></tr>
                <tr><td>6</td><td></td><td></td><td></td><td></td><td></td></tr>
                <tr><td>7</td><td></td><td></td><td></td><td></td><td></td></tr>
                <tr><td>8</td><td></td><td></td><td></td><td></td><td></td></tr>
                <tr><td>9</td><td></td><td></td><td></td><td></td><td></td></tr>
                <tr><td>10</td><td></td><td></td><td></td><td></td><td></td></tr>
            </tbody>
        </table>

        <div class="section-title">Cash Sheet Balancing</div>
        <div class="flex-table-container">
            <div class="flex-table-cell">
                <table>
                    <tbody>
                        <tr><td class="align-left" style="width: 60%;">System Sale</td><td class="align-right" style="width: 40%;">{total_amount:,.2f}</td></tr>
                        <tr><td class="align-left">FOC</td><td></td></tr>
                        <tr><td class="align-left">Total Cancel</td><td></td></tr>
                        <tr class="bg-balance"><td class="align-left">Balance (1)</td><td></td></tr>
                        <tr><td class="align-left">Total Discounts</td><td></td></tr>
                        <tr><td class="align-left">Total Adjust</td><td></td></tr>
                        <tr><td class="align-left">Total Return</td><td></td></tr>
                        <tr class="bg-balance"><td class="align-left">Balance (2)</td><td></td></tr>
                        <tr><td class="align-left">Total Cash</td><td></td></tr>
                        <tr><td class="align-left">Total Credit</td><td></td></tr>
                        <tr><td class="align-left">Total Cheques</td><td></td></tr>
                        <tr class="bg-balance"><td class="align-left">Balance (3)</td><td></td></tr>
                    </tbody>
                </table>
            </div>
            <div class="flex-table-cell">
                <table>
                    <thead>
                        <tr>
                            <th class="align-left" style="width: 60%;">Cash Balance</th>
                            <th style="width: 40%;"></th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr><td class="align-left">Total Day Cash</td><td></td></tr>
                        <tr><td class="align-left">Total Credit Received</td><td></td></tr>
                        <tr><td class="align-left" style="padding-bottom: 45px;">Total Expenses</td><td></td></tr>
                        <tr><td class="align-left">Banked Value.</td><td></td></tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div class="section-title">Calculating cash</div>
        <table>
            <thead>
                <tr>
                    <th style="width: 50%;">Cash Analitics</th>
                    <th style="width: 50%;">Valuve</th>
                </tr>
            </thead>
            <tbody>
                <tr><td>20</td><td></td></tr>
                <tr><td>50</td><td></td></tr>
                <tr><td>100</td><td></td></tr>
                <tr><td>500</td><td></td></tr>
                <tr><td>1000</td><td></td></tr>
                <tr><td>2000</td><td></td></tr>
                <tr><td>5000</td><td></td></tr>
                <tr><td>coins</td><td></td></tr>
                <tr class="bg-st"><td>Total</td><td></td></tr>
            </tbody>
        </table>

        <div class="meta-info" style="margin-top: 10px; font-size: 8pt;">
            Distance Travelled : ___________________ &nbsp;&nbsp;&nbsp;&nbsp; KM : ___________________ &nbsp;&nbsp;&nbsp;&nbsp; OOT : ___________________
        </div>
    </body>
    </html>
    """
    
    pdf_stream = io.BytesIO()
    HTML(string=html_template).write_pdf(pdf_stream)
    pdf_stream.seek(0)
    return pdf_stream

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

# --- Main App Interface ---
st.title("📄 KIST Sheet Custom Generation Engine")
st.markdown('<p class="black-text">Upload the picklist PDF file below to instantly extract data and generate the structured layout.</p>', unsafe_allow_html=True)

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
