"""
FormatyX — AI Paper Generator
Flask web app for generating Romanian academic conference papers using Claude.
Run with: python generator_app.py
Open http://localhost:5001
"""

import os
import uuid
import threading
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template_string
from generator import generate_section, check_grammar, build_docx, SECTIONS, SECTION_KEYS

app = Flask(__name__)
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# job_id -> { status, sections: {key: text}, current, error, meta }
jobs: dict = {}


# ─────────────────────────────────────────────────────────────────────────────
# HTML
# ─────────────────────────────────────────────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="ro">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FormatyX — Paper Generator</title>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: 'Segoe UI', sans-serif;
  background: #f0f2f5;
  min-height: 100vh;
  padding: 32px 24px;
}

.layout {
  display: flex;
  gap: 24px;
  max-width: 1300px;
  margin: 0 auto;
  align-items: flex-start;
}

/* ── Sidebar ── */
.sidebar {
  width: 340px;
  flex-shrink: 0;
  background: #fff;
  border-radius: 12px;
  box-shadow: 0 4px 24px rgba(0,0,0,0.10);
  padding: 28px 28px 24px;
  position: sticky;
  top: 32px;
}

.sidebar h1 { font-size: 1.3rem; font-weight: 700; color: #1a1a2e; margin-bottom: 2px; }
.sidebar .subtitle { color: #6b7280; font-size: 0.82rem; margin-bottom: 22px; }

.field { margin-bottom: 14px; }
label { display: block; font-size: 0.82rem; font-weight: 600; color: #374151; margin-bottom: 5px; }
label .opt { font-weight: 400; color: #9ca3af; }
input[type="text"], textarea, select {
  width: 100%;
  padding: 8px 10px;
  border: 1px solid #d1d5db;
  border-radius: 7px;
  font-size: 0.88rem;
  color: #111827;
  outline: none;
  transition: border-color 0.15s;
  font-family: inherit;
  background: #fff;
}
input[type="text"]:focus, textarea:focus, select:focus { border-color: #6366f1; }

button.primary {
  width: 100%;
  padding: 11px;
  background: #4f46e5;
  color: #fff;
  border: none;
  border-radius: 8px;
  font-size: 0.93rem;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.2s;
  margin-top: 6px;
}
button.primary:hover:not(:disabled) { background: #4338ca; }
button.primary:disabled { background: #a5b4fc; cursor: not-allowed; }

button.secondary {
  padding: 6px 12px;
  background: #f3f4f6;
  color: #374151;
  border: 1px solid #d1d5db;
  border-radius: 6px;
  font-size: 0.78rem;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.15s;
}
button.secondary:hover:not(:disabled) { background: #e5e7eb; }
button.secondary:disabled { opacity: 0.5; cursor: not-allowed; }

button.download-btn {
  width: 100%;
  padding: 10px;
  background: #16a34a;
  color: #fff;
  border: none;
  border-radius: 8px;
  font-size: 0.9rem;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.2s;
  margin-top: 12px;
  display: none;
}
button.download-btn:hover { background: #15803d; }

.progress-wrap { margin-top: 14px; display: none; }
.progress-label { font-size: 0.78rem; color: #6b7280; margin-bottom: 5px; }
.progress-bar-bg { background: #e5e7eb; border-radius: 99px; height: 6px; }
.progress-bar { background: #6366f1; border-radius: 99px; height: 6px; width: 0%; transition: width 0.4s; }
.progress-text { font-size: 0.75rem; color: #6b7280; margin-top: 4px; }

/* ── Main panel ── */
.main {
  flex: 1;
  min-width: 0;
}

.empty-state {
  background: #fff;
  border-radius: 12px;
  box-shadow: 0 4px 24px rgba(0,0,0,0.07);
  padding: 60px 40px;
  text-align: center;
  color: #9ca3af;
}
.empty-state .icon { font-size: 3rem; margin-bottom: 12px; }
.empty-state p { font-size: 0.9rem; }

/* ── Section card ── */
.section-card {
  background: #fff;
  border-radius: 10px;
  box-shadow: 0 2px 12px rgba(0,0,0,0.07);
  margin-bottom: 16px;
  overflow: hidden;
  animation: fadeIn 0.3s ease;
}
@keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }

.card-header {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 12px 16px;
  border-bottom: 1px solid #f3f4f6;
  background: #fafafa;
}
.card-header .section-label {
  font-size: 0.85rem;
  font-weight: 700;
  color: #1a1a2e;
  flex: 1;
}
.card-header .card-actions { display: flex; gap: 6px; }

.card-body { padding: 14px 16px; }

.section-textarea {
  width: 100%;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  padding: 10px 12px;
  font-size: 0.86rem;
  font-family: 'Segoe UI', sans-serif;
  color: #111827;
  resize: vertical;
  min-height: 80px;
  outline: none;
  transition: border-color 0.15s;
  line-height: 1.6;
}
.section-textarea:focus { border-color: #6366f1; }
.section-textarea.generating {
  background: #fafafa;
  color: #9ca3af;
  font-style: italic;
}

.grammar-results {
  margin-top: 10px;
  display: none;
}
.grammar-issue {
  background: #fef9c3;
  border: 1px solid #fde68a;
  border-radius: 6px;
  padding: 7px 10px;
  margin-bottom: 6px;
  font-size: 0.78rem;
}
.grammar-issue .gi-msg { color: #92400e; font-weight: 600; margin-bottom: 2px; }
.grammar-issue .gi-ctx { color: #78716c; font-family: monospace; }
.grammar-issue .gi-fix { color: #15803d; margin-top: 2px; }
.grammar-ok { color: #15803d; font-size: 0.78rem; font-weight: 600; margin-top: 6px; }

.spinner {
  display: inline-block;
  width: 11px; height: 11px;
  border: 2px solid #c7d2fe;
  border-top-color: #6366f1;
  border-radius: 50%;
  animation: spin 0.7s linear infinite;
  vertical-align: middle;
}
@keyframes spin { to { transform: rotate(360deg); } }

.status-dot {
  width: 8px; height: 8px;
  border-radius: 50%;
  flex-shrink: 0;
}
.dot-done { background: #22c55e; }
.dot-generating { background: #6366f1; animation: pulse 1s infinite; }
.dot-pending { background: #d1d5db; }
@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
</style>
</head>
<body>
<div class="layout">

  <!-- ── Sidebar ── -->
  <div class="sidebar">
    <h1>FormatyX</h1>
    <p class="subtitle">Generare lucrări academice cu AI</p>

    <div class="field">
      <label>Subiect / Temă <span class="opt">*</span></label>
      <input type="text" id="topic" placeholder="ex: Microcogenerare pe biomasă">
    </div>
    <div class="field">
      <label>Domeniu <span class="opt">*</span></label>
      <input type="text" id="domain" placeholder="ex: Inginerie energetică">
    </div>
    <div class="field">
      <label>Obiective cercetare</label>
      <textarea id="objectives" rows="3" placeholder="ex: Analiza eficienței exergetice a unui sistem de microcogenerare pe biomasă..."></textarea>
    </div>
    <div class="field">
      <label>Cuvinte cheie sugerate <span class="opt">(opțional)</span></label>
      <input type="text" id="keywords" placeholder="ex: biomasă, cogenerare, exergie">
    </div>
    <div class="field">
      <label>Autori <span class="opt">(unul pe linie, opțional)</span></label>
      <textarea id="authors" rows="2" placeholder="Ion Popescu&#10;Maria Ionescu"></textarea>
    </div>

    <button class="primary" id="generate-btn" onclick="startGeneration()">Generează lucrarea</button>
    <button class="download-btn" id="download-btn" onclick="downloadDoc()">Descarcă DOCX</button>

    <div class="progress-wrap" id="progress-wrap">
      <div class="progress-label">Progres generare</div>
      <div class="progress-bar-bg"><div class="progress-bar" id="progress-bar"></div></div>
      <div class="progress-text" id="progress-text"></div>
    </div>
  </div>

  <!-- ── Main panel ── -->
  <div class="main" id="main">
    <div class="empty-state" id="empty-state">
      <div class="icon">📝</div>
      <p>Completează formularul și apasă <strong>Generează lucrarea</strong><br>pentru a începe.</p>
    </div>
    <div id="sections-wrap"></div>
  </div>

</div>

<script>
const SECTIONS = """ + str([{"key": s["key"], "label": s["label"]} for s in SECTIONS]) + """;

let currentJobId = null;
let pollTimer = null;
let renderedKeys = new Set();

function startGeneration() {
  const topic = document.getElementById('topic').value.trim();
  const domain = document.getElementById('domain').value.trim();
  if (!topic || !domain) { alert('Completează Subiectul și Domeniul.'); return; }

  const body = {
    topic,
    domain,
    objectives: document.getElementById('objectives').value.trim(),
    keywords: document.getElementById('keywords').value.trim(),
    authors: document.getElementById('authors').value.trim(),
  };

  document.getElementById('generate-btn').disabled = true;
  document.getElementById('download-btn').style.display = 'none';
  document.getElementById('empty-state').style.display = 'none';
  document.getElementById('sections-wrap').innerHTML = '';
  document.getElementById('progress-wrap').style.display = 'block';
  renderedKeys = new Set();

  fetch('/generate', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(body),
  })
  .then(r => r.json())
  .then(data => {
    if (data.error) { alert(data.error); resetUI(); return; }
    currentJobId = data.job_id;
    pollStatus();
  })
  .catch(e => { alert('Eroare: ' + e); resetUI(); });
}

function pollStatus() {
  if (!currentJobId) return;
  fetch('/status/' + currentJobId)
    .then(r => r.json())
    .then(data => {
      updateProgress(data);
      renderNewSections(data.sections || {});

      if (data.status === 'done' || data.status === 'error') {
        document.getElementById('generate-btn').disabled = false;
        if (data.status === 'done') {
          document.getElementById('download-btn').style.display = 'block';
        }
        if (data.status === 'error') {
          alert('Eroare la generare: ' + (data.error || 'necunoscută'));
        }
      } else {
        pollTimer = setTimeout(pollStatus, 2500);
      }
    })
    .catch(() => { pollTimer = setTimeout(pollStatus, 3000); });
}

function updateProgress(data) {
  const done = Object.keys(data.sections || {}).length;
  const total = SECTIONS.length;
  const pct = Math.round((done / total) * 100);
  document.getElementById('progress-bar').style.width = pct + '%';
  const cur = data.current ? SECTIONS.find(s => s.key === data.current) : null;
  document.getElementById('progress-text').textContent =
    data.status === 'done' ? `Completat (${total}/${total} secțiuni)` :
    cur ? `Generez: ${cur.label} (${done}/${total})` : `${done}/${total} secțiuni`;
}

function renderNewSections(sections) {
  for (const sec of SECTIONS) {
    if (sections[sec.key] !== undefined && !renderedKeys.has(sec.key)) {
      renderedKeys.add(sec.key);
      appendSectionCard(sec.key, sec.label, sections[sec.key]);
    }
  }
}

function appendSectionCard(key, label, text) {
  const wrap = document.getElementById('sections-wrap');
  const rows = Math.max(4, Math.ceil(text.length / 80));

  const card = document.createElement('div');
  card.className = 'section-card';
  card.id = 'card-' + key;
  card.innerHTML = `
    <div class="card-header">
      <span class="status-dot dot-done" id="dot-${key}"></span>
      <span class="section-label">${label}</span>
      <div class="card-actions">
        <button class="secondary" onclick="regenerateSection('${key}')">Regenerează</button>
        <button class="secondary" onclick="grammarCheck('${key}')">Verifică gramatică</button>
      </div>
    </div>
    <div class="card-body">
      <textarea class="section-textarea" id="ta-${key}" rows="${rows}">${escHtml(text)}</textarea>
      <div class="grammar-results" id="gr-${key}"></div>
    </div>
  `;
  wrap.appendChild(card);
}

function escHtml(t) {
  return t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function regenerateSection(key) {
  if (!currentJobId) return;
  const ta = document.getElementById('ta-' + key);
  ta.value = 'Generez...';
  ta.className = 'section-textarea generating';
  document.getElementById('dot-' + key).className = 'status-dot dot-generating';

  // Collect current texts of all sections as edited context
  const overrides = {};
  for (const s of SECTIONS) {
    const el = document.getElementById('ta-' + s.key);
    if (el && el.value && el.value !== 'Generez...') overrides[s.key] = el.value;
  }

  fetch('/regenerate', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ job_id: currentJobId, key, overrides }),
  })
  .then(r => r.json())
  .then(data => {
    if (data.error) { ta.value = '[Eroare: ' + data.error + ']'; }
    else { ta.value = data.text; }
    ta.className = 'section-textarea';
    const rows = Math.max(4, Math.ceil(ta.value.length / 80));
    ta.rows = rows;
    document.getElementById('dot-' + key).className = 'status-dot dot-done';
  })
  .catch(e => {
    ta.value = '[Eroare rețea]';
    ta.className = 'section-textarea';
    document.getElementById('dot-' + key).className = 'status-dot dot-done';
  });
}

function grammarCheck(key) {
  const ta = document.getElementById('ta-' + key);
  const gr = document.getElementById('gr-' + key);
  const text = ta.value.trim();
  if (!text) return;

  gr.style.display = 'block';
  gr.innerHTML = '<span class="spinner"></span> Verificare gramaticală...';

  fetch('/grammar', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ text }),
  })
  .then(r => r.json())
  .then(data => {
    const issues = data.issues || [];
    if (!issues.length) {
      gr.innerHTML = '<div class="grammar-ok">✓ Nicio problemă gramaticală detectată.</div>';
    } else {
      gr.innerHTML = issues.map(i => `
        <div class="grammar-issue">
          <div class="gi-msg">${escHtml(i.message)}</div>
          <div class="gi-ctx">"${escHtml(i.context)}"</div>
          ${i.replacements.length ? `<div class="gi-fix">Sugestii: ${i.replacements.map(r => `<strong>${escHtml(r)}</strong>`).join(', ')}</div>` : ''}
        </div>
      `).join('');
    }
  })
  .catch(() => { gr.innerHTML = '<div class="grammar-issue"><div class="gi-msg">Eroare la verificare.</div></div>'; });
}

function downloadDoc() {
  if (!currentJobId) return;
  // Collect all current textarea values
  const sections = {};
  for (const s of SECTIONS) {
    const el = document.getElementById('ta-' + s.key);
    if (el) sections[s.key] = el.value;
  }
  const authors = document.getElementById('authors').value.trim();

  fetch('/download/' + currentJobId, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ sections, authors }),
  })
  .then(r => {
    if (!r.ok) throw new Error('Download failed');
    return r.blob();
  })
  .then(blob => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'paper_generated.docx';
    a.click();
    URL.revokeObjectURL(url);
  })
  .catch(e => alert('Eroare la descărcare: ' + e));
}

function resetUI() {
  document.getElementById('generate-btn').disabled = false;
  document.getElementById('progress-wrap').style.display = 'none';
}
</script>
</body>
</html>
"""


# ─────────────────────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    topic = (data.get("topic") or "").strip()
    domain = (data.get("domain") or "").strip()
    if not topic or not domain:
        return jsonify({"error": "Topic and domain are required."}), 400

    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "status": "running",
        "sections": {},
        "current": None,
        "error": None,
        "meta": {
            "topic": topic,
            "domain": domain,
            "objectives": data.get("objectives", ""),
            "keywords": data.get("keywords", ""),
            "authors": data.get("authors", ""),
        },
    }

    thread = threading.Thread(
        target=_run_generation, args=(job_id,), daemon=True
    )
    thread.start()
    return jsonify({"job_id": job_id})


def _run_generation(job_id: str):
    job = jobs[job_id]
    meta = job["meta"]
    generated = {}

    try:
        for section in SECTIONS:
            key = section["key"]
            job["current"] = key
            text = generate_section(
                key=key,
                topic=meta["topic"],
                domain=meta["domain"],
                objectives=meta["objectives"],
                keywords=meta["keywords"],
                generated=generated,
            )
            generated[key] = text
            job["sections"][key] = text

        job["status"] = "done"
        job["current"] = None
    except Exception as e:
        job["status"] = "error"
        job["error"] = str(e)


@app.route("/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job."}), 404
    return jsonify({
        "status": job["status"],
        "sections": job["sections"],
        "current": job["current"],
        "error": job["error"],
    })


@app.route("/regenerate", methods=["POST"])
def regenerate():
    data = request.get_json()
    job_id = data.get("job_id")
    key = data.get("key")
    overrides = data.get("overrides", {})

    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job."}), 404

    meta = job["meta"]
    # Use overrides (user-edited content) as context
    context = {**job["sections"], **overrides}
    # Remove the section being regenerated so it doesn't self-reference
    context.pop(key, None)

    try:
        text = generate_section(
            key=key,
            topic=meta["topic"],
            domain=meta["domain"],
            objectives=meta["objectives"],
            keywords=meta["keywords"],
            generated=context,
        )
        job["sections"][key] = text
        return jsonify({"text": text})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/grammar", methods=["POST"])
def grammar():
    data = request.get_json()
    text = (data.get("text") or "").strip()
    if not text:
        return jsonify({"issues": []})
    issues = check_grammar(text)
    return jsonify({"issues": issues})


@app.route("/download/<job_id>", methods=["POST"])
def download(job_id):
    job = jobs.get(job_id)
    if not job:
        return "Unknown job.", 404

    data = request.get_json()
    sections = data.get("sections", job["sections"])
    authors_raw = data.get("authors", job["meta"].get("authors", ""))
    authors = [a.strip() for a in authors_raw.splitlines() if a.strip()]

    output_path = str(OUTPUT_DIR / f"paper_{job_id}.docx")
    try:
        build_docx({**sections, "authors": authors}, output_path)
        return send_file(output_path, as_attachment=True, download_name="paper_generated.docx")
    except Exception as e:
        return str(e), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT_GENERATOR", 5001))
    print(f"Starting FormatyX Paper Generator at http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
