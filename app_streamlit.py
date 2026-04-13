import streamlit as st
import subprocess
import os
import sys
import zipfile
import io
import tempfile
import shutil
import shutil as _shutil

# Resolve node binary – Streamlit Cloud installs it via packages.txt
# Common paths on Debian/Ubuntu
_NODE_CANDIDATES = ["node", "/usr/bin/node", "/usr/local/bin/node"]
NODE = next((p for p in _NODE_CANDIDATES if _shutil.which(p)), "node")

# Auto-install node_modules if missing (needed on Streamlit Cloud)
_BASE = os.path.dirname(os.path.abspath(__file__))
if not os.path.exists(os.path.join(_BASE, "node_modules")):
    subprocess.run(["npm", "install"], cwd=_BASE, capture_output=True)

st.set_page_config(
    page_title="Invoice Generator",
    page_icon="📄",
    layout="centered"
)

# ── Custom CSS to match original look ────────────────────────────────────────
st.markdown("""
<style>
    .main { background: linear-gradient(135deg, #f4f7f6 0%, #e8f5f2 100%); }
    .block-container { max-width: 560px; padding-top: 2rem; }

    /* Title */
    h1 { color: #1a1a2e !important; font-size: 26px !important; text-align: center; }

    /* Download buttons */
    .stDownloadButton > button {
        background-color: #009578 !important;
        color: white !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        width: 100%;
        padding: 12px !important;
        border: none !important;
        transition: background 0.3s;
    }
    .stDownloadButton > button:hover {
        background-color: #007b63 !important;
    }

    /* Generate button */
    .stButton > button {
        background-color: #009578 !important;
        color: white !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        width: 100%;
        font-size: 16px !important;
        padding: 14px !important;
        border: none !important;
    }
    .stButton > button:hover { background-color: #007b63 !important; }

    .divider { border-top: 2px solid #e8f5f2; margin: 1rem 0; }
</style>
""", unsafe_allow_html=True)


# ── State init ────────────────────────────────────────────────────────────────
for key in ("generated", "xlsx_bytes", "docx_bytes", "zip_bytes", "error"):
    if key not in st.session_state:
        st.session_state[key] = None
if "generated" not in st.session_state:
    st.session_state.generated = False


# ── UI ────────────────────────────────────────────────────────────────────────
st.title("📄 Invoice Generator")
st.markdown(
    "<p style='text-align:center;color:#888;font-size:14px;margin-top:-10px'>"
    "Upload your <strong>emp_data_input.xlsx</strong> to generate all outputs</p>",
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "Drop Excel file here or click to upload",
    type=["xlsx", "xls"],
    label_visibility="collapsed"
)

if uploaded_file:
    st.success(f"✅ {uploaded_file.name} — ready to process")

generate_clicked = st.button("⚙️ Process & Generate Files", disabled=not uploaded_file)

# ── Processing ────────────────────────────────────────────────────────────────
if generate_clicked and uploaded_file:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    PYTHON   = sys.executable

    progress = st.progress(0, text="Reading Excel data...")
    status   = st.empty()
    error_box = st.empty()

    try:
        # Save uploaded Excel
        excel_path = os.path.join(BASE_DIR, "emp_data_input.xlsx")
        with open(excel_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        def run(cmd, label):
            result = subprocess.run(cmd, cwd=BASE_DIR, capture_output=True, text=True)
            if result.returncode != 0:
                detail = result.stderr.strip() or result.stdout.strip() or "(no output)"
                raise RuntimeError(f"{label} failed:\n{detail}")

        # Step 1
        progress.progress(20, text="Generating timesheet & JSON...")
        run([PYTHON, os.path.join(BASE_DIR, "generate_output.py")], "generate_output.py")

        # Step 2
        progress.progress(50, text="Building Word invoices...")
        run([NODE, os.path.join(BASE_DIR, "generate_invoices.js")], "generate_invoices.js")

        # Step 3
        progress.progress(75, text="Creating individual PDFs...")
        run([PYTHON, os.path.join(BASE_DIR, "generate_pdf_invoices.py")], "generate_pdf_invoices.py")

        progress.progress(95, text="Finalising outputs...")

        # Read output files into memory
        xlsx_path = os.path.join(BASE_DIR, "Salary_TimeSheet_Output_new.xlsx")
        docx_path = os.path.join(BASE_DIR, "Employee_Invoices_new.docx")
        pdf_dir   = os.path.join(BASE_DIR, "individual_pdfs")

        with open(xlsx_path, "rb") as f:
            st.session_state.xlsx_bytes = f.read()
        with open(docx_path, "rb") as f:
            st.session_state.docx_bytes = f.read()

        # Build ZIP in memory
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fn in sorted(os.listdir(pdf_dir)):
                if fn.endswith(".pdf"):
                    zf.write(os.path.join(pdf_dir, fn), fn)
        zip_buf.seek(0)
        st.session_state.zip_bytes = zip_buf.read()

        progress.progress(100, text="Done!")
        st.session_state.generated = True
        st.session_state.error     = None

    except Exception as e:
        progress.empty()
        st.session_state.error = str(e)
        st.session_state.generated = False

# ── Error display ─────────────────────────────────────────────────────────────
if st.session_state.error:
    st.error(f"❌ {st.session_state.error}")

# ── Download section ──────────────────────────────────────────────────────────
if st.session_state.generated:
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    st.markdown("### ✅ Files Ready — Download Below")

    st.download_button(
        label="📊  Salary_TimeSheet_Output.xlsx  ⬇",
        data=st.session_state.xlsx_bytes,
        file_name="Salary_TimeSheet_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        label="📝  Employee_Invoices.docx  ⬇",
        data=st.session_state.docx_bytes,
        file_name="Employee_Invoices.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    st.download_button(
        label="🗜️  Individual_PDF_Invoices.zip  ⬇",
        data=st.session_state.zip_bytes,
        file_name="Individual_PDF_Invoices.zip",
        mime="application/zip",
    )
