import streamlit as st
import subprocess
import os
import sys
import zipfile
import io

st.set_page_config(
    page_title="Invoice Generator",
    page_icon="📄",
    layout="centered"
)

st.markdown("""
<style>
    .main { background: linear-gradient(135deg, #f4f7f6 0%, #e8f5f2 100%); }
    .block-container { max-width: 560px; padding-top: 2rem; }
    h1 { color: #1a1a2e !important; font-size: 26px !important; text-align: center; }
    .stDownloadButton > button {
        background-color: #009578 !important; color: white !important;
        border-radius: 8px !important; font-weight: 600 !important;
        width: 100%; padding: 12px !important; border: none !important;
    }
    .stDownloadButton > button:hover { background-color: #007b63 !important; }
    .stButton > button {
        background-color: #009578 !important; color: white !important;
        border-radius: 8px !important; font-weight: 600 !important;
        width: 100%; font-size: 16px !important; padding: 14px !important; border: none !important;
    }
    .stButton > button:hover { background-color: #007b63 !important; }
    .divider { border-top: 2px solid #e8f5f2; margin: 1rem 0; }
</style>
""", unsafe_allow_html=True)

for key, default in [("generated", False), ("xlsx_bytes", None),
                     ("docx_bytes", None), ("zip_bytes", None), ("error", None)]:
    if key not in st.session_state:
        st.session_state[key] = default

st.title("📄 Invoice Generator")
st.markdown(
    "<p style='text-align:center;color:#888;font-size:14px;margin-top:-10px'>"
    "Upload your <strong>emp_data_input.xlsx</strong> to generate all outputs</p>",
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "Drop Excel file here or click to upload",
    type=["xlsx", "xls"], label_visibility="collapsed"
)

if uploaded_file:
    st.success(f"✅ {uploaded_file.name} — ready to process")

generate_clicked = st.button("⚙️ Process & Generate Files", disabled=not uploaded_file)

if generate_clicked and uploaded_file:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    PYTHON   = sys.executable
    progress = st.progress(0, text="Reading Excel data...")

    try:
        with open(os.path.join(BASE_DIR, "emp_data_input.xlsx"), "wb") as f:
            f.write(uploaded_file.getbuffer())

        def run(script, label):
            result = subprocess.run(
                [PYTHON, os.path.join(BASE_DIR, script)],
                cwd=BASE_DIR, capture_output=True, text=True
            )
            if result.returncode != 0:
                detail = result.stderr.strip() or result.stdout.strip() or "(no output)"
                raise RuntimeError(f"{label} failed:\n{detail}")

        progress.progress(20, text="Generating timesheet & JSON...")
        run("generate_output.py", "generate_output.py")

        progress.progress(50, text="Building Word invoices...")
        run("generate_invoices_py.py", "generate_invoices_py.py")  # pure Python, no Node

        progress.progress(75, text="Creating individual PDFs...")
        run("generate_pdf_invoices.py", "generate_pdf_invoices.py")

        progress.progress(95, text="Finalising outputs...")

        with open(os.path.join(BASE_DIR, "Salary_TimeSheet_Output_new.xlsx"), "rb") as f:
            st.session_state.xlsx_bytes = f.read()
        with open(os.path.join(BASE_DIR, "Employee_Invoices_new.docx"), "rb") as f:
            st.session_state.docx_bytes = f.read()

        pdf_dir = os.path.join(BASE_DIR, "individual_pdfs")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fn in sorted(os.listdir(pdf_dir)):
                if fn.endswith(".pdf"):
                    zf.write(os.path.join(pdf_dir, fn), fn)
        zip_buf.seek(0)
        st.session_state.zip_bytes = zip_buf.read()

        progress.progress(100, text="Done!")
        st.session_state.generated = True
        st.session_state.error = None

    except Exception as e:
        progress.empty()
        st.session_state.error = str(e)
        st.session_state.generated = False

if st.session_state.error:
    st.error(f"❌ {st.session_state.error}")

if st.session_state.generated:
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    st.markdown("### ✅ Files Ready — Download Below")
    st.download_button("📊  Salary_TimeSheet_Output.xlsx  ⬇",
        data=st.session_state.xlsx_bytes, file_name="Salary_TimeSheet_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("📝  Employee_Invoices.docx  ⬇",
        data=st.session_state.docx_bytes, file_name="Employee_Invoices.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.download_button("🗜️  Individual_PDF_Invoices.zip  ⬇",
        data=st.session_state.zip_bytes, file_name="Individual_PDF_Invoices.zip",
        mime="application/zip")