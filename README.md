# Gear Invoice Generator 🚀

Flask/Streamlit app to generate employee invoices (Word, PDF, Excel) from Excel input + signatures.

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.app/)
[![Deploy](https://img.shields.io/badge/Streamlit-Deploy-brightgreen)](https://share.streamlit.io/)

## Quick Deploy on Streamlit Cloud

1. Go to [share.streamlit.app](https://share.streamlit.app)
2. Sign in with GitHub
3. Click "Create app" → Select this repo → `app_streamlit.py` → Deploy

Free hosting! Upload `emp_data_input.xlsx` → download invoices/PDFs/timesheets.

## Local Setup

```bash
pip install -r requirements.txt
npm install
streamlit run app_streamlit.py
```

Or Flask: `python app.py` (deprecated).

## Scripts

- `generate_output.py`: Excel/JSON processing
- `generate_invoices.js`: Word docs (Node.js + docx)
- `generate_pdf_invoices.py`: Individual PDFs (ReportLab + signatures)

## Signatures

`Signatures/` folder (~50 PNGs) auto-matched by name.

## Outputs

- `Salary_TimeSheet_Output_new.xlsx`
- `Employee_Invoices_new.docx`
- `individual_pdfs/*.pdf` (zipped download)

**Input**: `emp_data_input.xlsx` (2 sheets: attendance + details).

See commits for full history.
