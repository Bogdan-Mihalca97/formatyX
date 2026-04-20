"""
Simple Flask UI for the conference paper formatter.
Run with: python app.py
Then open http://localhost:5000 in your browser.
"""

import os
import uuid
import threading
import subprocess
import sys
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template_string

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Track job status: job_id -> {status, message, output_file, log, filename}
jobs: dict = {}


HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FormatyX</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Segoe UI', sans-serif;
    background: #f0f2f5;
    min-height: 100vh;
    display: flex;
    align-items: flex-start;
    justify-content: center;
    padding: 32px 24px;
  }

  .card {
    background: #fff;
    border-radius: 12px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.10);
    padding: 36px 40px;
    width: 100%;
    max-width: 620px;
  }

  h1 { font-size: 1.4rem; font-weight: 700; color: #1a1a2e; margin-bottom: 4px; }
  .subtitle { color: #6b7280; font-size: 0.88rem; margin-bottom: 28px; }

  /* Drop zone */
  #drop-zone {
    border: 2px dashed #c7d2fe;
    border-radius: 10px;
    background: #f5f7ff;
    padding: 30px 20px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
    margin-bottom: 16px;
  }
  #drop-zone.drag-over { border-color: #6366f1; background: #eef2ff; }
  #drop-zone .icon { font-size: 2rem; margin-bottom: 6px; }
  #drop-zone p { color: #6b7280; font-size: 0.88rem; }
  #drop-zone strong { color: #4f46e5; }
  #file-input { display: none; }

  /* File queue */
  #queue { margin-bottom: 18px; display: none; }
  #queue-title {
    font-size: 0.8rem;
    font-weight: 600;
    color: #6b7280;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 8px;
  }
  .file-row {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 9px 12px;
    border-radius: 8px;
    background: #f9fafb;
    border: 1px solid #e5e7eb;
    margin-bottom: 6px;
    font-size: 0.86rem;
  }
  .file-row .fname { flex: 1; color: #111827; font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .file-row .badge {
    font-size: 0.74rem;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 20px;
    white-space: nowrap;
  }
  .badge-pending  { background: #f3f4f6; color: #6b7280; }
  .badge-running  { background: #dbeafe; color: #1d4ed8; }
  .badge-done     { background: #dcfce7; color: #15803d; }
  .badge-error    { background: #fee2e2; color: #b91c1c; }
  .file-row .dl-link {
    font-size: 0.78rem;
    color: #16a34a;
    font-weight: 600;
    text-decoration: none;
    white-space: nowrap;
  }
  .file-row .dl-link:hover { text-decoration: underline; }
  .file-row .remove-btn {
    background: none;
    border: none;
    color: #9ca3af;
    cursor: pointer;
    font-size: 1rem;
    line-height: 1;
    padding: 0 2px;
  }
  .file-row .remove-btn:hover { color: #ef4444; }

  /* Fields */
  .field { margin-bottom: 14px; }
  label { display: block; font-size: 0.83rem; font-weight: 600; color: #374151; margin-bottom: 5px; }
  label .opt { font-weight: 400; color: #9ca3af; }
  input[type="text"], textarea {
    width: 100%;
    padding: 9px 12px;
    border: 1px solid #d1d5db;
    border-radius: 7px;
    font-size: 0.9rem;
    color: #111827;
    outline: none;
    transition: border-color 0.15s;
    font-family: inherit;
  }
  input[type="text"]:focus, textarea:focus { border-color: #6366f1; }

  /* Buttons */
  .btn-row { display: flex; gap: 10px; margin-top: 6px; }
  button.primary {
    flex: 1;
    padding: 11px;
    background: #4f46e5;
    color: #fff;
    border: none;
    border-radius: 8px;
    font-size: 0.93rem;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.2s;
  }
  button.primary:hover:not(:disabled) { background: #4338ca; }
  button.primary:disabled { background: #a5b4fc; cursor: not-allowed; }

  /* Log box */
  #log-wrap { margin-top: 16px; display: none; }
  #log-label { font-size: 0.78rem; font-weight: 600; color: #6b7280; margin-bottom: 4px; }
  #log-box {
    background: #1e1e2e;
    color: #cdd6f4;
    border-radius: 7px;
    padding: 10px 12px;
    font-family: 'Consolas', monospace;
    font-size: 0.74rem;
    max-height: 160px;
    overflow-y: auto;
    white-space: pre-wrap;
  }

  .spinner {
    display: inline-block;
    width: 11px; height: 11px;
    border: 2px solid #bfdbfe;
    border-top-color: #1d4ed8;
    border-radius: 50%;
    animation: spin 0.7s linear infinite;
    vertical-align: middle;
    margin-right: 4px;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>
<div class="card">
  <h1>FormatyX</h1>
  <p class="subtitle">Format Romanian academic papers for conference submission.</p>

  <div id="drop-zone" onclick="document.getElementById('file-input').click()">
    <div class="icon">📄</div>
    <p><strong>Click to select</strong> or drag &amp; drop one or more .docx files</p>
    <input type="file" id="file-input" accept=".docx" multiple>
  </div>

  <div id="queue">
    <div id="queue-title">Files to process</div>
    <div id="queue-list"></div>
  </div>

  <div class="field">
    <label>Authors <span class="opt">(one per line, applied to all files)</span></label>
    <textarea id="authors" rows="3" placeholder="Gheorghe Badea&#10;George Naghiu&#10;Andrei Măgureanu"></textarea>
  </div>

  <div class="field">
    <label>English Title <span class="opt">(optional — auto-translated per file if empty)</span></label>
    <input type="text" id="title-en" placeholder="Leave blank to auto-translate">
  </div>

  <div class="btn-row">
    <button class="primary" id="run-btn" disabled>Format All</button>
  </div>

  <div id="log-wrap">
    <div id="log-label">Log — <span id="log-filename"></span></div>
    <div id="log-box"></div>
  </div>
</div>

<script>
  const dropZone   = document.getElementById('drop-zone');
  const fileInput  = document.getElementById('file-input');
  const queueEl    = document.getElementById('queue');
  const queueList  = document.getElementById('queue-list');
  const runBtn     = document.getElementById('run-btn');
  const logWrap    = document.getElementById('log-wrap');
  const logBox     = document.getElementById('log-box');
  const logFilename = document.getElementById('log-filename');

  // fileQueue: [{file, id, status, jobId}]
  let fileQueue = [];
  let processing = false;

  // --- Drag & drop / file select ---
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    addFiles(Array.from(e.dataTransfer.files));
  });
  fileInput.addEventListener('change', () => {
    addFiles(Array.from(fileInput.files));
    fileInput.value = '';
  });

  function addFiles(files) {
    const valid = files.filter(f => f.name.endsWith('.docx'));
    if (valid.length !== files.length) alert('Only .docx files are supported — others were ignored.');
    valid.forEach(f => {
      // skip duplicates by name
      if (fileQueue.some(q => q.file.name === f.name)) return;
      fileQueue.push({ file: f, id: Math.random().toString(36).slice(2), status: 'pending', jobId: null });
    });
    renderQueue();
  }

  function renderQueue() {
    if (fileQueue.length === 0) {
      queueEl.style.display = 'none';
      runBtn.disabled = true;
      return;
    }
    queueEl.style.display = 'block';
    queueList.innerHTML = '';
    fileQueue.forEach(item => {
      const row = document.createElement('div');
      row.className = 'file-row';
      row.id = 'row-' + item.id;

      const name = document.createElement('span');
      name.className = 'fname';
      name.title = item.file.name;
      name.textContent = item.file.name;

      const badge = document.createElement('span');
      badge.className = 'badge badge-' + item.status;
      badge.id = 'badge-' + item.id;
      const labels = { pending: 'Pending', running: '⏳ Processing', done: '✓ Done', error: '✗ Error' };
      badge.innerHTML = item.status === 'running' ? '<span class="spinner"></span>Processing' : labels[item.status];

      row.appendChild(name);
      row.appendChild(badge);

      if (item.status === 'done' && item.jobId) {
        const dl = document.createElement('a');
        dl.className = 'dl-link';
        dl.href = '/download/' + item.jobId;
        dl.textContent = 'Download';
        row.appendChild(dl);
      }

      if (item.status === 'pending') {
        const rm = document.createElement('button');
        rm.className = 'remove-btn';
        rm.title = 'Remove';
        rm.textContent = '×';
        rm.onclick = () => {
          fileQueue = fileQueue.filter(q => q.id !== item.id);
          renderQueue();
        };
        row.appendChild(rm);
      }

      queueList.appendChild(row);
    });

    const hasPending = fileQueue.some(q => q.status === 'pending');
    runBtn.disabled = processing || !hasPending;
  }

  runBtn.addEventListener('click', () => {
    if (processing) return;
    processing = true;
    runBtn.disabled = true;
    processNext();
  });

  async function processNext() {
    const item = fileQueue.find(q => q.status === 'pending');
    if (!item) {
      processing = false;
      renderQueue();
      return;
    }

    item.status = 'running';
    renderQueue();

    // Show log for this file
    logWrap.style.display = 'block';
    logFilename.textContent = item.file.name;
    logBox.textContent = 'Uploading…';

    const formData = new FormData();
    formData.append('file', item.file);
    const authors = document.getElementById('authors').value.trim();
    if (authors) formData.append('authors', authors);
    const titleEn = document.getElementById('title-en').value.trim();
    if (titleEn) formData.append('title_en', titleEn);

    let jobId;
    try {
      const res = await fetch('/format', { method: 'POST', body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Upload failed');
      jobId = data.job_id;
      item.jobId = jobId;
    } catch (err) {
      item.status = 'error';
      renderQueue();
      logBox.textContent = 'Error: ' + err.message;
      processing = false;
      runBtn.disabled = false;
      return;
    }

    // Poll until done
    await new Promise(resolve => {
      const timer = setInterval(async () => {
        try {
          const res = await fetch('/status/' + jobId);
          const data = await res.json();
          if (data.log) {
            logBox.textContent = data.log;
            logBox.scrollTop = logBox.scrollHeight;
          }
          if (data.status === 'done' || data.status === 'error') {
            clearInterval(timer);
            item.status = data.status;
            renderQueue();
            resolve();
          }
        } catch (e) { /* retry */ }
      }, 2000);
    });

    // Move to next file
    processNext();
  }
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/format", methods=["POST"])
def format_doc():
    if "file" not in request.files:
        return jsonify({"error": "No file provided."}), 400

    f = request.files["file"]
    if not f.filename.endswith(".docx"):
        return jsonify({"error": "Only .docx files are supported."}), 400

    job_id = str(uuid.uuid4())
    input_path = UPLOAD_DIR / f"{job_id}_{f.filename}"
    stem = Path(f.filename).stem
    output_path = OUTPUT_DIR / f"{stem}_formatted.docx"

    f.save(input_path)

    authors_raw = request.form.get("authors", "").strip()
    authors = [a.strip() for a in authors_raw.splitlines() if a.strip()]
    title_en = request.form.get("title_en", "").strip()

    jobs[job_id] = {
        "status": "running",
        "message": "",
        "output_file": str(output_path),
        "log": "",
        "filename": f.filename,
    }

    thread = threading.Thread(
        target=run_formatter,
        args=(job_id, str(input_path), str(output_path), authors, title_en),
        daemon=True,
    )
    thread.start()

    return jsonify({"job_id": job_id})


def run_formatter(job_id, input_path, output_path, authors, title_en):
    PYTHON = sys.executable
    cmd = [PYTHON, "-u", "formatter.py", input_path, "-o", output_path]
    if authors:
        cmd += ["--authors"] + authors
    if title_en:
        cmd += ["--title-en", title_en]

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=BASE_DIR,
        )
        log_lines = []
        for line in proc.stdout:
            log_lines.append(line.rstrip())
            jobs[job_id]["log"] = "\n".join(log_lines)

        proc.wait()
        if proc.returncode == 0:
            summary = next((l for l in reversed(log_lines) if l.startswith("Done!")), "")
            jobs[job_id]["status"] = "done"
            jobs[job_id]["message"] = summary
        else:
            error_tail = "\n".join(log_lines[-10:])
            jobs[job_id]["status"] = "error"
            jobs[job_id]["message"] = error_tail
    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["message"] = str(e)


@app.route("/status/<job_id>")
def job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"status": "error", "message": "Unknown job."}), 404
    return jsonify(job)


@app.route("/download/<job_id>")
def download(job_id):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return "Not ready.", 404
    output_file = job["output_file"]
    if not os.path.exists(output_file):
        return "File not found.", 404
    return send_file(output_file, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting FormatyX at http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
