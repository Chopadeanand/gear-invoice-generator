const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, ImageRun, PageBreak } = require('docx');
const fs   = require('fs');
const path = require('path');

const employees = JSON.parse(fs.readFileSync('emp_data.json', 'utf8'));

const MONTH       = "Feb- 2026";
const DATE        = "28-02-2026";
const MONTH_LABEL = "Feb'26";
const SIG_DIR     = 'signatures/';  // put PNGs here when running locally

// ── Hardcoded name → signature filename map ──────────────────────────────────
// Keys are the service_provider names exactly as they appear in emp_data.json.
// Update SIG_DIR above to the folder where you keep the PNG files.
const SIG_MAP = {
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
};
// ─────────────────────────────────────────────────────────────────────────────

function rupees(n) {
  return new Intl.NumberFormat('en-IN').format(Math.round(n));
}

const bdr  = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const bdrs = { top: bdr, bottom: bdr, left: bdr, right: bdr };

function boldRun(text, sz=20)   { return new TextRun({ text, bold: true,  size: sz, font: "Arial" }); }
function normRun(text, sz=20)   { return new TextRun({ text, bold: false, size: sz, font: "Arial" }); }
function para(children, align=AlignmentType.LEFT, spacing={after:80}) {
  return new Paragraph({ children, alignment: align, spacing });
}

function mkCell(children, opts={}) {
  const { width=2250, shade=null, align=AlignmentType.LEFT, colspan=1 } = opts;
  return new TableCell({
    children: Array.isArray(children) ? children : [para(Array.isArray(children)?children:[normRun(String(children))], align, {after:40})],
    width: { size: width, type: WidthType.DXA },
    borders: bdrs,
    columnSpan: colspan,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
  });
}

// Resolves signature filename for an employee:
// 1. Check SIG_MAP by service_provider name (hardcoded, most reliable)
// 2. Fall back to sig_filename from emp_data.json
// 3. Fall back to name field
function getSigFilename(emp) {
  const nameKey = emp.service_provider || emp.name || '';
  if (SIG_MAP[nameKey]) return SIG_MAP[nameKey];
  // Try case-insensitive match
  const lowerKey = nameKey.trim().toLowerCase();
  for (const [k, v] of Object.entries(SIG_MAP)) {
    if (k.trim().toLowerCase() === lowerKey) return v;
  }
  // Fall back to whatever was in the JSON
  return emp.sig_filename || null;
}

function tryLoadSig(emp) {
  const sigFilename = getSigFilename(emp);
  if (!sigFilename) return null;
  const tryPaths = [
    path.join(SIG_DIR, sigFilename),
    path.join('/mnt/user-data/uploads/', sigFilename),
    sigFilename,  // absolute or relative path as-is
  ];
  for (const p of tryPaths) {
    if (fs.existsSync(p)) {
      try {
        const data = fs.readFileSync(p);
        return new ImageRun({
          data,
          transformation: { width: 120, height: 50 },
          type: 'png',
        });
      } catch(e) { /* skip */ }
    }
  }
  console.warn(`  [WARN] Signature not found for "${emp.service_provider || emp.name}" (tried: ${sigFilename})`);
  return null;
}

function buildPage(emp) {
  const projs    = emp.projects || [];
  const totalAmt = emp.total_amount;
  const MONTHLY = 16500, MONTH_DAYS = 28;
  const rate    = MONTHLY / MONTH_DAYS;

  // ── Fee table rows ──
  const tableRows = [
    new TableRow({ children: [
      mkCell([para([boldRun('Particulars')],          AlignmentType.CENTER, {after:40})], { width:3500, shade:"BDD7EE" }),
      mkCell([para([boldRun(`No. of working\ndays in ${MONTH_LABEL}`)], AlignmentType.CENTER, {after:40})], { width:1800, shade:"BDD7EE" }),
      mkCell([para([boldRun('WBS Elements')],         AlignmentType.CENTER, {after:40})], { width:4000, shade:"BDD7EE" }),
      mkCell([para([boldRun('Payable Amount\n(Rs.)')],AlignmentType.CENTER, {after:40})], { width:2000, shade:"BDD7EE" }),
    ], tableHeader: true }),
  ];

  if (projs.length === 0) {
    tableRows.push(new TableRow({ children: [
      mkCell("Consultant fee – Data Processing", { width:3500 }),
      mkCell(String(emp.attendance), { width:1800, align:AlignmentType.CENTER }),
      mkCell('', { width:4000 }),
      mkCell(rupees(totalAmt), { width:2000, align:AlignmentType.RIGHT }),
    ]}));
    for (let i=0;i<2;i++) tableRows.push(new TableRow({ children:[
      mkCell('',{width:3500}),mkCell('',{width:1800}),mkCell('',{width:4000}),mkCell('',{width:2000})
    ]}));
  } else {
    projs.forEach((p, idx) => {
      const amt = Math.round(p.days * rate);
      tableRows.push(new TableRow({ children: [
        mkCell(idx===0 ? "Consultant fee – Data Processing" : '', { width:3500 }),
        mkCell(String(p.days), { width:1800, align:AlignmentType.CENTER }),
        mkCell(p.wbs, { width:4000 }),
        mkCell(rupees(amt), { width:2000, align:AlignmentType.RIGHT }),
      ]}));
    });
    while (tableRows.length < 4) tableRows.push(new TableRow({ children:[
      mkCell('',{width:3500}),mkCell('',{width:1800}),mkCell('',{width:4000}),mkCell('',{width:2000})
    ]}));
  }

  // Total row
  tableRows.push(new TableRow({ children: [
    mkCell('', { width:3500 }),
    mkCell(String(emp.attendance), { width:1800, align:AlignmentType.CENTER }),
    mkCell([para([boldRun('Total Pay')], AlignmentType.CENTER, {after:40})], { width:4000, shade:"E2EFDA" }),
    mkCell([para([boldRun(`Rs. ${rupees(totalAmt)}/-`)], AlignmentType.RIGHT, {after:40})], { width:2000, shade:"E2EFDA" }),
  ]}));

  const feeTable = new Table({
    width: { size: 11300, type: WidthType.DXA },
    columnWidths: [3500, 1800, 4000, 2000],
    rows: tableRows,
  });

  // ── Bank details table ──
  const bankTable = new Table({
    width: { size: 11300, type: WidthType.DXA },
    columnWidths: [2800, 2700, 3000, 2800],
    rows: [
      new TableRow({ children: [
        mkCell([para([boldRun('Account-Name')],       AlignmentType.CENTER,{after:40})],{width:2800,shade:"BDD7EE"}),
        mkCell([para([boldRun('Bank Name')],           AlignmentType.CENTER,{after:40})],{width:2700,shade:"BDD7EE"}),
        mkCell([para([boldRun('Bank Account Number')], AlignmentType.CENTER,{after:40})],{width:3000,shade:"BDD7EE"}),
        mkCell([para([boldRun('IFSC Code')],           AlignmentType.CENTER,{after:40})],{width:2800,shade:"BDD7EE"}),
      ]}),
      new TableRow({ children: [
        mkCell(emp.account_name    || '', {width:2800}),
        mkCell(emp.bank_name       || '', {width:2700}),
        mkCell(emp.account_number  || '', {width:3000}),
        mkCell(emp.ifsc            || '', {width:2800}),
      ]}),
    ],
  });

  // ── Signature ──
  const sigImg = tryLoadSig(emp);
  const sigPara = sigImg
    ? new Paragraph({ children: [sigImg], spacing:{after:40} })
    : para([normRun('')], AlignmentType.LEFT, {after:60});

  return [
    para([boldRun(`Date: - ${DATE}`)]),
    para([]),
    para([boldRun('TO,')]),
    para([boldRun('CUBE Highways Technologies Private Limited,')]),
    para([normRun('3rd Floor, GMR Aero Towers – 2,')]),
    para([normRun('Mamidipally Village, Saroor Nagar Mandal,')]),
    para([normRun('Ranga Reddy, Hyderabad, Telangana - 500108')]),
    para([]),
    para([boldRun('GST No- '), normRun('36AAKCC7533R1ZW')]),
    para([boldRun('PAN No- '), normRun('AAKCC7533R')]),
    para([]),
    para([boldRun('Sir,')]),
    para([]),
    para([
      boldRun('Subject: '),
      normRun('Consultant fee for '),
      boldRun(MONTH),
      normRun(' data processing & Analysis '),
      boldRun(`Rs.${rupees(totalAmt)}/-`),
      normRun(' per month. The commercials are mentioned below.'),
    ]),
    para([]),
    feeTable,
    para([]),
    para([normRun('Thanking you and always assuring you of our best services.')]),
    para([]),
    para([boldRun('Yours faithfully')]),
    para([]),
    para([boldRun('Authorised Signature')]),
    sigPara,
    para([]),
    para([boldRun('Service Provider: '), normRun(emp.service_provider || emp.name || '')]),
    para([boldRun('Address: '),          normRun(emp.address          || '')]),
    para([boldRun('Email- '),            normRun(emp.email            || '')]),
    para([boldRun('Contact No. '),       normRun(emp.contact          || '')]),
    para([boldRun('PAN No- '),           normRun(emp.pan              || '')]),
    para([]),
    para([boldRun('Bank details below:')]),
    para([]),
    bankTable,
  ];
}

// Build all pages
const allContent = [];
employees.forEach((emp, idx) => {
  const page = buildPage(emp);
  if (idx < employees.length - 1) {
    page.push(new Paragraph({ children: [new PageBreak()] }));
  }
  allContent.push(...page);
});

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 720, right: 720, bottom: 720, left: 720 },
      },
    },
    children: allContent,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('Employee_Invoices_new.docx', buf);
  console.log('Done: Employee_Invoices_new.docx');
}).catch(err => { console.error(err); process.exit(1); });
