import json
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

MONTH = "Feb- 2026"
DATE = "28-02-2026"
MONTH_LABEL = "Feb'26"
MONTHLY = 16500
MONTH_DAYS = 28
SIG_DIR = 'signatures/'

# Same SIG_MAP as generate_invoices.js — keys must match service_provider names in emp_data.json
SIG_MAP = {
  "Ahmed Azam":                   "Ahmed_Azam-removebg-preview.png",
  "Anand Ganesh  Chopade":        "Anand Ganesh Chopade.png",
  "Arundathi Jalagam":            "Arundathi_Jalagam-removebg-preview.png",
  "Bonagiri Rehana":              "Bonagiri_Rehana-removebg-preview.png",
  "Chouta Keerthana":             "Chouta_Keerthana-removebg-preview.png",
  "Deshapaga Raghavendar":        "Deshapaga_Raghavendar-removebg-preview.png",
  "Dhenumakonda Lavanya":         "Dhenumakonda_Lavanya-removebg-preview.png",
  "Diravath Mounika":             "Diravath_Mounika-removebg-preview.png",
  "Gaja Bala Narayana":           "Gaja Bala Narayana-removebg-preview.png",
  "Gopala Saritha":               "Gopala_Saritha-removebg-preview.png",
  "Gudise Hemanth Kumar":         "Gudise_Hemanth_Kumar-removebg-preview.png",
  "Jetti Hima Sindhu":            "Jetti_Hima_Sindhu-removebg-preview.png",
  "K THIRUPATHAIAH":              "K_THIRUPATHAIAH-removebg-preview.png",
  "Kanche Srisailam":             "Kanche_Srisailam-removebg-preview.png",
  "Kasireddy Harish":             "Kasireddy_Harish-removebg-preview.png",
  "Kasturi Sathish":              "Kasturi_Sathish-removebg-preview.png",
  "KATRAVATH MANGESH":            "KATRAVATH_MANGESH-removebg-preview.png",
  "Katravath Mohan Rathod":       "Katravath_Mohan_Rathod-removebg-preview.png",
  "Katravath Radhika":            "Katravath_Radhika-removebg-preview.png",
  "Kodavath Swamy":               "Kodavath_Swamy-removebg-preview.png",
  "KOLLA SUDHAKAR":               "KOLLA_SUDHAKAR-removebg-preview.png",
  "Kunduru Upender Reddy":        "Kunduru_Upender_Reddy-removebg-preview.png",
  "M JATHIN SAI":                 "M_JATHIN_SAI-removebg-preview.png",
  "M Sunitha":                    "M_SUNITHA-removebg-preview.png",
  "M Swarupa":                    "M_Swarupa-removebg-preview.png",
  "Meka Esthar Rani":             "Meka Esthar Rani-removebg-preview.png",
  "Mamidi Raj kumar":             "Mamidi_Raj_kumar-removebg-preview.png",
  "Md Mainoddin":                 "Md_Mainoddin-removebg-preview.png",
  "Mekala Anusha":                "MEKALA_ANUSHA-removebg-preview.png",
  "MOHAMMAD NAZIYA":              "MOHAMMAD_NAZIYA-removebg-preview.png",
  "Mohammad Shaheen Begum":       "Mohammad Shaheen Begum-removebg-preview.png",
  "Mohammed Yaqoob khan":         "Mohammed_Yaqoob_Khan-removebg-preview.png",
  "Mudavath Balakoti":            "Mudavath_Balakoti-removebg-preview.png",
  "MUDAVATH KIRAN":               "MUDAVATH_KIRAN-removebg-preview.png",
  "Mudavath Ramesh":              "Mudavath_Ramesh-removebg-preview.png",
  "Neeli Sreevani":               "Neeli SreevaniSignature remove.png",
  "Padira Radhika":               "Padira_Radhika-removebg-preview.png",
  "Pandiri Punith kumar":         "Pandiri_Punith_kumar-removebg-preview.png",
  "Poojari Ramu":                 "Pujari_Ramu-removebg-preview.png",
  "Porandla Chandar":             "Porandla_Chandar-removebg-preview.png",
  "Ramavath Saimahesh Nayak":     "Ramavath_Saimahesh_Nayak-removebg-preview.png",
  "Ramavath Uma Mahesh":          "Ramavath_Uma_Mahesh-removebg-preview.png",
  "Ranabotu Saidi Reddy":         "Ranabotu_Saidi_Reddy-removebg-preview.png",
  "Sadde Sindhura":               "Sadde_Sindhura-removebg-preview.png",
  "Shaikh abdul Avesh":           "Shaikh_abdul_Avesh-removebg-preview.png",
  "Shivarathri Swapna":           "Shivarathri_Swapna-removebg-preview.png",
  "Tandra Sabastin":              "Tandra_Sabastin-removebg-preview.png",
  "Tabassum Afreen":              "Thabasum_Afreen-removebg-preview.png",
  "Thatipally Manoj Kumar":       "Thatipally_Manoj_Kumar-removebg-preview.png",
  "Thurpati vijay Baskar":        "Thurpati_vijay_Baskar-removebg-preview.png",
  "Ushanolla Ravali":             "Ushanolla_Ravali-removebg-preview.png",
  "Ushanula Ramya":               "Ushanula_Ramya-removebg-preview.png",
  "Chejerla Nagavamsidhar Reddy": "Chejerla_Nagavamsidhar_Reddy-removebg-preview.png",
}

def get_sig_path(emp):
    name_key = emp.get('service_provider') or emp.get('name', '')
    # Exact match
    filename = SIG_MAP.get(name_key)
    # Case-insensitive fallback
    if not filename:
        for k, v in SIG_MAP.items():
            if k.strip().lower() == name_key.strip().lower():
                filename = v
                break
    # Fallback to sig_filename field from JSON
    if not filename:
        filename = emp.get('sig_filename')
    if not filename:
        return None
    for candidate in [os.path.join(SIG_DIR, filename), filename]:
        if os.path.exists(candidate):
            return candidate
    print(f"  [WARN] Signature not found for \"{name_key}\" (tried: {filename})")
    return None

def rupees(n):
    import locale
    n = round(n)
    s = str(n)
    if len(s) <= 3:
        return s
    result = s[-3:]
    s = s[:-3]
    while len(s) > 2:
        result = s[-2:] + ',' + result
        s = s[:-2]
    if s:
        result = s + ',' + result
    return result

def make_styles():
    normal = ParagraphStyle('normal', fontName='Helvetica', fontSize=10, leading=14)
    bold = ParagraphStyle('bold', fontName='Helvetica-Bold', fontSize=10, leading=14)
    small = ParagraphStyle('small', fontName='Helvetica', fontSize=9, leading=12)
    small_bold = ParagraphStyle('small_bold', fontName='Helvetica-Bold', fontSize=9, leading=12)
    return normal, bold, small, small_bold

def build_pdf(emp, output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=20*mm, leftMargin=20*mm,
        topMargin=15*mm, bottomMargin=15*mm
    )

    normal, bold, small, small_bold = make_styles()
    story = []

    def p(text, style=normal, space_after=4):
        story.append(Paragraph(text, style))
        if space_after:
            story.append(Spacer(1, space_after))

    # Header
    p(f"<b>Date: - {DATE}</b>", bold, 6)
    p("")
    p("<b>TO,</b>", bold, 2)
    p("<b>CUBE Highways Technologies Private Limited,</b>", bold, 2)
    p("3rd Floor, GMR Aero Towers – 2,", normal, 2)
    p("Mamidipally Village, Saroor Nagar Mandal,", normal, 2)
    p("Ranga Reddy, Hyderabad, Telangana - 500108", normal, 6)

    p("<b>GST No- </b>36AAKCC7533R1ZW", bold, 2)
    p("<b>PAN No- </b>AAKCC7533R", bold, 6)
    p("<b>Sir,</b>", bold, 6)

    total = emp['total_amount']
    p(f"<b>Subject: </b>Consultant fee for <b>{MONTH}</b> data processing &amp; Analysis "
      f"<b>Rs.{rupees(total)}/-</b> per month. The commercials are mentioned below.", normal, 8)

    # Fee table
    rate = MONTHLY / MONTH_DAYS
    projs = emp.get('projects', [])

    header_bg = colors.Color(0.741, 0.843, 0.933)  # BDD7EE
    total_bg = colors.Color(0.886, 0.937, 0.855)   # E2EFDA

    col_widths = [65*mm, 32*mm, 65*mm, 35*mm]
    table_data = [
        [
            Paragraph(f"<b>Particulars</b>", small_bold),
            Paragraph(f"<b>No. of working days in {MONTH_LABEL}</b>", small_bold),
            Paragraph("<b>WBS Elements</b>", small_bold),
            Paragraph("<b>Payable Amount (Rs.)</b>", small_bold),
        ]
    ]

    if not projs:
        table_data.append([
            Paragraph("Consultant fee – Data Processing", small),
            Paragraph(str(emp['attendance']), small),
            Paragraph("", small),
            Paragraph(rupees(total), small),
        ])
        for _ in range(2):
            table_data.append(["", "", "", ""])
    else:
        for idx, proj in enumerate(projs):
            amt = round(proj['days'] * rate)
            table_data.append([
                Paragraph("Consultant fee – Data Processing" if idx == 0 else "", small),
                Paragraph(str(proj['days']), small),
                Paragraph(proj.get('wbs', ''), small),
                Paragraph(rupees(amt), small),
            ])
        while len(table_data) < 4:
            table_data.append(["", "", "", ""])

    # Total row
    table_data.append([
        "",
        Paragraph(str(emp['attendance']), small_bold),
        Paragraph("<b>Total Pay</b>", small_bold),
        Paragraph(f"<b>Rs. {rupees(total)}/-</b>", small_bold),
    ])

    n_data_rows = len(table_data) - 1  # excluding header
    total_row_idx = len(table_data) - 1

    fee_table = Table(table_data, colWidths=col_widths)
    style_cmds = [
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), header_bg),
        ('BACKGROUND', (2, total_row_idx), (-1, total_row_idx), total_bg),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
    ]
    fee_table.setStyle(TableStyle(style_cmds))
    story.append(fee_table)
    story.append(Spacer(1, 8))

    p("Thanking you and always assuring you of our best services.", normal, 6)
    p("<b>Yours faithfully</b>", bold, 6)
    p("<b>Authorised Signature</b>", bold, 4)

    sig_path = get_sig_path(emp)
    if sig_path:
        sig_img = Image(sig_path, width=45*mm, height=18*mm)
        sig_img.hAlign = "LEFT"
        story.append(sig_img)
        story.append(Spacer(1, 4))
    else:
        story.append(Spacer(1, 22))  # blank space if no signature

    p(f"<b>Service Provider: </b>{emp.get('service_provider') or emp.get('name', '')}", bold, 2)
    p(f"<b>Address: </b>{emp.get('address', '')}", bold, 2)
    p(f"<b>Email- </b>{emp.get('email', '')}", bold, 2)
    p(f"<b>Contact No. </b>{emp.get('contact', '')}", bold, 2)
    p(f"<b>PAN No- </b>{emp.get('pan', '')}", bold, 6)

    p("<b>Bank details below:</b>", bold, 4)

    bank_data = [
        [
            Paragraph("<b>Account-Name</b>", small_bold),
            Paragraph("<b>Bank Name</b>", small_bold),
            Paragraph("<b>Bank Account Number</b>", small_bold),
            Paragraph("<b>IFSC Code</b>", small_bold),
        ],
        [
            Paragraph(emp.get('account_name', ''), small),
            Paragraph(emp.get('bank_name', ''), small),
            Paragraph(str(emp.get('account_number', '')), small),
            Paragraph(emp.get('ifsc', ''), small),
        ]
    ]
    bank_col_widths = [49*mm, 49*mm, 54*mm, 45*mm]
    bank_table = Table(bank_data, colWidths=bank_col_widths)
    bank_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), header_bg),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(bank_table)

    doc.build(story)


# Load employee data
employees = json.load(open('emp_data.json'))

# Create output folder
pdf_dir = 'individual_pdfs'
os.makedirs(pdf_dir, exist_ok=True)

for emp in employees:
    name = emp.get('service_provider') or emp.get('name', 'Unknown')
    safe_name = name.replace(' ', '_').replace('/', '_').replace('\\', '_')
    out_path = os.path.join(pdf_dir, f"{safe_name}_Invoice_Feb2026.pdf")
    build_pdf(emp, out_path)
    print(f"Generated: {out_path}")

print(f"\nDone! {len(employees)} PDFs created in '{pdf_dir}/'")
