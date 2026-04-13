from flask import Flask, request, render_template_string, send_file, jsonify
import subprocess
import os
import zipfile
import io

app = Flask(__name__)

# --- FRONTEND: HTML, CSS, and JS ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Invoice Generator</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            background: linear-gradient(135deg, #f4f7f6 0%, #e8f5f2 100%);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 30px 20px;
        }
        .container {
            background: white;
            padding: 40px;
            border-radius: 16px;
            box-shadow: 0 8px 30px rgba(0,0,0,0.1);
            text-align: center;
            width: 100%;
            max-width: 500px;
        }
        h2 {
            color: #1a1a2e;
            margin-bottom: 8px;
            font-size: 24px;
        }
        .subtitle {
            color: #888;
            font-size: 14px;
            margin-bottom: 28px;
        }
        .drop-zone {
            width: 100%;
            height: 180px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 10px;
            cursor: pointer;
            color: #aaa;
            border: 3px dashed #009578;
            border-radius: 12px;
            transition: all 0.3s;
            margin-bottom: 20px;
            background: #fafafa;
        }
        .drop-zone--over {
            border-style: solid;
            background-color: #e8f5f2;
            color: #009578;
        }
        .drop-zone--has-file {
            border-color: #009578;
            border-style: solid;
            background: #e8f5f2;
            color: #009578;
        }
        .drop-zone .icon { font-size: 36px; }
        .drop-zone__input { display: none; }
        .btn {
            background-color: #009578;
            color: white;
            border: none;
            padding: 14px 24px;
            font-size: 16px;
            border-radius: 8px;
            cursor: pointer;
            transition: background 0.3s;
            width: 100%;
            font-weight: 600;
            letter-spacing: 0.3px;
        }
        .btn:hover { background-color: #007b63; }
        .btn:disabled { background-color: #ccc; cursor: not-allowed; }

        /* Progress */
        #progress-wrap {
            display: none;
            margin-top: 20px;
        }
        .progress-bar-bg {
            background: #eee;
            border-radius: 999px;
            height: 8px;
            overflow: hidden;
            margin-bottom: 8px;
        }
        .progress-bar-fill {
            height: 100%;
            background: linear-gradient(90deg, #009578, #00c49a);
            border-radius: 999px;
            width: 0%;
            transition: width 0.5s ease;
            animation: pulse 1.5s ease-in-out infinite;
        }
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50%       { opacity: 0.7; }
        }
        .progress-label {
            font-size: 13px;
            color: #666;
        }

        /* Status */
        #status {
            margin-top: 16px;
            font-size: 14px;
            color: #333;
        }

        /* Download section */
        #download-section {
            display: none;
            margin-top: 28px;
            text-align: left;
        }
        #download-section h3 {
            font-size: 16px;
            color: #1a1a2e;
            margin-bottom: 14px;
            border-bottom: 2px solid #e8f5f2;
            padding-bottom: 8px;
        }
        .download-grid {
            display: flex;
            flex-direction: column;
            gap: 10px;
        }
        .download-btn {
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 12px 16px;
            border-radius: 10px;
            border: 1.5px solid #e0e0e0;
            background: #fafafa;
            cursor: pointer;
            text-decoration: none;
            color: #1a1a2e;
            transition: all 0.2s;
            font-size: 14px;
        }
        .download-btn:hover {
            border-color: #009578;
            background: #e8f5f2;
            transform: translateY(-1px);
            box-shadow: 0 3px 10px rgba(0,149,120,0.15);
        }
        .download-btn .file-icon {
            font-size: 24px;
            flex-shrink: 0;
        }
        .download-btn .file-info { flex: 1; }
        .download-btn .file-name { font-weight: 600; }
        .download-btn .file-desc { font-size: 12px; color: #888; margin-top: 2px; }
        .download-btn .dl-arrow { font-size: 18px; color: #009578; }
    </style>
</head>
<body>
    <div class="container">
        <h2>📄 Invoice Generator</h2>
        <p class="subtitle">Upload your <strong>emp_data_input.xlsx</strong> to generate all outputs</p>

        <form id="uploadForm">
            <div class="drop-zone" id="dropZone">
                <span class="icon">📂</span>
                <span class="drop-zone__prompt">Drop Excel file here or click to upload</span>
                <input type="file" name="excel_file" id="fileInput" class="drop-zone__input" accept=".xlsx,.xls">
            </div>
            <button type="submit" class="btn" id="submitBtn" disabled>⚙️ Process & Generate Files</button>
        </form>

        <div id="progress-wrap">
            <div class="progress-bar-bg">
                <div class="progress-bar-fill" id="progressBar"></div>
            </div>
            <div class="progress-label" id="progressLabel">Starting...</div>
        </div>

        <div id="status"></div>

        <!-- Download Section -->
        <div id="download-section">
            <h3>✅ Files Ready — Download Below</h3>
            <div class="download-grid" id="downloadGrid">
                <a href="/download/xlsx" class="download-btn" download>
                    <span class="file-icon">📊</span>
                    <div class="file-info">
                        <div class="file-name">Salary_TimeSheet_Output.xlsx</div>
                        <div class="file-desc">Full salary data & timesheet (Excel)</div>
                    </div>
                    <span class="dl-arrow">⬇</span>
                </a>
                <a href="/download/docx" class="download-btn" download>
                    <span class="file-icon">📝</span>
                    <div class="file-info">
                        <div class="file-name">Employee_Invoices.docx</div>
                        <div class="file-desc">All employee invoices in one Word document</div>
                    </div>
                    <span class="dl-arrow">⬇</span>
                </a>
                <a href="/download/pdfs_zip" class="download-btn" download>
                    <span class="file-icon">🗜️</span>
                    <div class="file-info">
                        <div class="file-name">Individual_PDF_Invoices.zip</div>
                        <div class="file-desc">One PDF invoice per employee (zipped)</div>
                    </div>
                    <span class="dl-arrow">⬇</span>
                </a>
            </div>
        </div>
    </div>

    <script>
        const dropZone   = document.getElementById('dropZone');
        const fileInput  = document.getElementById('fileInput');
        const prompt     = document.querySelector('.drop-zone__prompt');
        const submitBtn  = document.getElementById('submitBtn');
        const statusDiv  = document.getElementById('status');
        const progressWrap = document.getElementById('progress-wrap');
        const progressBar  = document.getElementById('progressBar');
        const progressLabel = document.getElementById('progressLabel');
        const downloadSection = document.getElementById('download-section');

        // Drop zone click
        dropZone.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', () => {
            if (fileInput.files.length) updateThumbnail(fileInput.files[0]);
        });

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('drop-zone--over');
        });
        ['dragleave', 'dragend'].forEach(t => {
            dropZone.addEventListener(t, () => dropZone.classList.remove('drop-zone--over'));
        });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            if (e.dataTransfer.files.length) {
                fileInput.files = e.dataTransfer.files;
                updateThumbnail(e.dataTransfer.files[0]);
            }
            dropZone.classList.remove('drop-zone--over');
        });

        function updateThumbnail(file) {
            prompt.textContent = '✅ ' + file.name;
            dropZone.classList.add('drop-zone--has-file');
            submitBtn.disabled = false;
        }

        // Fake progress animation while waiting
        let progressInterval = null;
        function startProgress() {
            progressWrap.style.display = 'block';
            downloadSection.style.display = 'none';
            let pct = 0;
            const steps = [
                { to: 20, label: 'Reading Excel data...' },
                { to: 45, label: 'Generating timesheet & JSON...' },
                { to: 70, label: 'Building Word invoices...' },
                { to: 90, label: 'Creating individual PDFs...' },
                { to: 95, label: 'Finalising outputs...' },
            ];
            let stepIdx = 0;
            progressInterval = setInterval(() => {
                if (stepIdx < steps.length && pct >= (stepIdx > 0 ? steps[stepIdx-1].to : 0)) {
                    progressLabel.textContent = steps[stepIdx].label;
                    stepIdx++;
                }
                if (pct < 95) { pct += 0.8; }
                progressBar.style.width = pct + '%';
            }, 300);
        }

        function stopProgress(success) {
            clearInterval(progressInterval);
            progressBar.style.width = success ? '100%' : '0%';
            progressBar.style.animation = 'none';
            if (success) {
                progressBar.style.background = 'linear-gradient(90deg,#009578,#00c49a)';
            } else {
                progressBar.style.background = '#e74c3c';
            }
            progressLabel.textContent = success ? 'Done!' : 'Failed.';
        }

        // Form submit
        document.getElementById('uploadForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const formData = new FormData();
            formData.append('excel_file', fileInput.files[0]);

            submitBtn.disabled = true;
            submitBtn.textContent = 'Processing...';
            statusDiv.innerHTML = '';
            startProgress();

            try {
                const response = await fetch('/upload', { method: 'POST', body: formData });
                const result = await response.text();

                if (response.ok) {
                    stopProgress(true);
                    statusDiv.innerHTML = '';
                    downloadSection.style.display = 'block';
                } else {
                    stopProgress(false);
                    statusDiv.innerHTML = '<span style="color:red">❌ Error: ' + result + '</span>';
                }
            } catch (err) {
                stopProgress(false);
                statusDiv.innerHTML = '<span style="color:red">❌ Network Error.</span>';
            } finally {
                submitBtn.textContent = '⚙️ Process & Generate Files';
                submitBtn.disabled = false;
            }
        });
    </script>
</body>
</html>
"""

# ── Route: Main UI ────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


# ── Route: Upload & Process ───────────────────────────────────────────────────
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'excel_file' not in request.files:
        return "No file uploaded", 400

    file = request.files['excel_file']
    if file.filename == '':
        return "No selected file", 400

    # All scripts must run from the folder where app.py lives
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    # Detect correct python executable (handles 'python3' on Linux/Mac)
    import sys
    PYTHON = sys.executable  # always the same interpreter running Flask

    try:
        # Save uploaded Excel where the scripts expect it
        file.save(os.path.join(BASE_DIR, 'emp_data_input.xlsx'))

        def run(cmd, label):
            result = subprocess.run(
                cmd,
                cwd=BASE_DIR,          # run from app's folder
                capture_output=True,   # grab stdout + stderr
                text=True
            )
            if result.returncode != 0:
                detail = result.stderr.strip() or result.stdout.strip() or "(no output)"
                raise RuntimeError(f"{label} failed:\n{detail}")
            return result

        # Step 1: generate emp_data.json + Salary_TimeSheet_Output_new.xlsx
        run([PYTHON, 'generate_output.py'], "generate_output.py")

        # Step 2: generate Employee_Invoices_new.docx
        run(['node', 'generate_invoices.js'], "generate_invoices.js")

        # Step 3: generate individual PDFs in individual_pdfs/
        run([PYTHON, 'generate_pdf_invoices.py'], "generate_pdf_invoices.py")

        return "All invoices and timesheets generated successfully!", 200

    except RuntimeError as e:
        return str(e), 500
    except Exception as e:
        return f"An error occurred: {str(e)}", 500


# ── Route: Download Excel ─────────────────────────────────────────────────────
@app.route('/download/xlsx')
def download_xlsx():
    path = 'Salary_TimeSheet_Output_new.xlsx'
    if not os.path.exists(path):
        return "File not found. Please process an Excel file first.", 404
    return send_file(
        path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Salary_TimeSheet_Output.xlsx'
    )


# ── Route: Download DOCX ──────────────────────────────────────────────────────
@app.route('/download/docx')
def download_docx():
    path = 'Employee_Invoices_new.docx'
    if not os.path.exists(path):
        return "File not found. Please process an Excel file first.", 404
    return send_file(
        path,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='Employee_Invoices.docx'
    )


# ── Route: Download all PDFs as ZIP ──────────────────────────────────────────
@app.route('/download/pdfs_zip')
def download_pdfs_zip():
    pdf_dir = 'individual_pdfs'
    if not os.path.exists(pdf_dir):
        return "PDFs not found. Please process an Excel file first.", 404

    pdf_files = [f for f in os.listdir(pdf_dir) if f.endswith('.pdf')]
    if not pdf_files:
        return "No PDF files found.", 404

    # Build ZIP in memory
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for pdf_file in sorted(pdf_files):
            full_path = os.path.join(pdf_dir, pdf_file)
            zf.write(full_path, pdf_file)
    zip_buffer.seek(0)

    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name='Individual_PDF_Invoices.zip'
    )


if __name__ == '__main__':
    print("Starting UI. Open http://127.0.0.1:5000 in your browser.")
    app.run(debug=True, port=5000)
