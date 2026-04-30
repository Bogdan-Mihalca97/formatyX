"""
FormatyX — Conference Paper Formatter + AI Paper Generator
Run with: python app.py  →  http://localhost:5000
"""

import os
import uuid
import threading
import subprocess
import sys
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template_string
from generator import generate_section, check_grammar, build_docx, SECTIONS, SECTION_KEYS

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Formatter jobs: job_id -> {status, message, output_file, log, filename}
fmt_jobs: dict = {}

# Generator jobs: job_id -> {status, sections, current, error, meta}
gen_jobs: dict = {}


# ─────────────────────────────────────────────────────────────────────────────
# HTML
# ─────────────────────────────────────────────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="ro">
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
}

/* ── Top nav ── */
.topnav {
  background: #1a1a2e;
  padding: 0 32px;
  display: flex;
  align-items: center;
  gap: 0;
  height: 52px;
  position: sticky;
  top: 0;
  z-index: 100;
  box-shadow: 0 2px 8px rgba(0,0,0,0.2);
}
.topnav .brand {
  font-size: 1.15rem;
  font-weight: 800;
  color: #fff;
  letter-spacing: -0.5px;
  margin-right: 28px;
  user-select: none;
}
.topnav .brand span { color: #818cf8; }
.tab-btn {
  height: 52px;
  padding: 0 20px;
  background: none;
  border: none;
  border-bottom: 3px solid transparent;
  color: #94a3b8;
  font-size: 0.88rem;
  font-weight: 600;
  cursor: pointer;
  transition: color 0.15s, border-color 0.15s;
}
.tab-btn:hover { color: #e2e8f0; }
.tab-btn.active { color: #fff; border-bottom-color: #818cf8; }

/* ── Page panels ── */
.panel { display: none; padding: 32px 24px; }
.panel.active { display: block; }

/* ────────────────────────────────────────────────
   FORMATTER PANEL
──────────────────────────────────────────────── */
.fmt-wrap {
  max-width: 620px;
  margin: 0 auto;
  background: #fff;
  border-radius: 12px;
  box-shadow: 0 4px 24px rgba(0,0,0,0.10);
  padding: 32px 36px;
}

.fmt-wrap h2 { font-size: 1.1rem; font-weight: 700; color: #1a1a2e; margin-bottom: 4px; }
.fmt-wrap .subtitle { color: #6b7280; font-size: 0.84rem; margin-bottom: 24px; }

#drop-zone {
  border: 2px dashed #c7d2fe;
  border-radius: 10px;
  background: #f5f7ff;
  padding: 28px 20px;
  text-align: center;
  cursor: pointer;
  transition: border-color 0.2s, background 0.2s;
  margin-bottom: 14px;
}
#drop-zone.drag-over { border-color: #6366f1; background: #eef2ff; }
#drop-zone .icon { font-size: 1.8rem; margin-bottom: 6px; }
#drop-zone p { color: #6b7280; font-size: 0.86rem; }
#drop-zone strong { color: #4f46e5; }
#file-input { display: none; }

#queue { margin-bottom: 16px; display: none; }
#queue-title { font-size: 0.78rem; font-weight: 600; color: #6b7280; text-transform: uppercase; letter-spacing:.05em; margin-bottom: 8px; }
.file-row { display: flex; align-items: center; gap: 10px; padding: 9px 12px; border-radius: 8px; background: #f9fafb; border: 1px solid #e5e7eb; margin-bottom: 6px; font-size: 0.86rem; }
.file-row .fname { flex: 1; color: #111827; font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.badge { font-size: 0.74rem; font-weight: 600; padding: 2px 8px; border-radius: 20px; white-space: nowrap; }
.badge-pending { background: #f3f4f6; color: #6b7280; }
.badge-running { background: #dbeafe; color: #1d4ed8; }
.badge-done { background: #dcfce7; color: #15803d; }
.badge-error { background: #fee2e2; color: #b91c1c; }
.dl-link { font-size: 0.78rem; color: #16a34a; font-weight: 600; text-decoration: none; white-space: nowrap; }
.dl-link:hover { text-decoration: underline; }
.remove-btn { background: none; border: none; color: #9ca3af; cursor: pointer; font-size: 1rem; line-height: 1; padding: 0 2px; }
.remove-btn:hover { color: #ef4444; }

.field { margin-bottom: 14px; }
label { display: block; font-size: 0.82rem; font-weight: 600; color: #374151; margin-bottom: 5px; }
label .opt { font-weight: 400; color: #9ca3af; }
input[type="text"], textarea, select {
  width: 100%; padding: 8px 11px; border: 1px solid #d1d5db; border-radius: 7px;
  font-size: 0.88rem; color: #111827; outline: none; transition: border-color 0.15s;
  font-family: inherit; background: #fff;
}
input[type="text"]:focus, textarea:focus { border-color: #6366f1; }

button.primary {
  flex: 1; padding: 11px; background: #4f46e5; color: #fff; border: none;
  border-radius: 8px; font-size: 0.92rem; font-weight: 600; cursor: pointer; transition: background 0.2s;
}
button.primary:hover:not(:disabled) { background: #4338ca; }
button.primary:disabled { background: #a5b4fc; cursor: not-allowed; }

.btn-row { display: flex; gap: 10px; margin-top: 6px; }

#log-wrap { margin-top: 14px; display: none; }
#log-label { font-size: 0.78rem; font-weight: 600; color: #6b7280; margin-bottom: 4px; }
#log-box {
  background: #1e1e2e; color: #cdd6f4; border-radius: 7px; padding: 10px 12px;
  font-family: 'Consolas', monospace; font-size: 0.73rem; max-height: 150px;
  overflow-y: auto; white-space: pre-wrap;
}

/* ────────────────────────────────────────────────
   GENERATOR PANEL
──────────────────────────────────────────────── */
.gen-layout {
  display: flex;
  gap: 24px;
  max-width: 1320px;
  margin: 0 auto;
  align-items: flex-start;
}

.gen-sidebar {
  width: 320px;
  flex-shrink: 0;
  background: #fff;
  border-radius: 12px;
  box-shadow: 0 4px 24px rgba(0,0,0,0.10);
  padding: 26px 26px 22px;
  position: sticky;
  top: 84px;
}
.gen-sidebar h2 { font-size: 1.05rem; font-weight: 700; color: #1a1a2e; margin-bottom: 4px; }
.gen-sidebar .subtitle { color: #6b7280; font-size: 0.82rem; margin-bottom: 20px; }

button.gen-primary {
  width: 100%; padding: 11px; background: #4f46e5; color: #fff; border: none;
  border-radius: 8px; font-size: 0.92rem; font-weight: 600; cursor: pointer;
  transition: background 0.2s; margin-top: 6px;
}
button.gen-primary:hover:not(:disabled) { background: #4338ca; }
button.gen-primary:disabled { background: #a5b4fc; cursor: not-allowed; }

button.secondary {
  padding: 6px 12px; background: #f3f4f6; color: #374151; border: 1px solid #d1d5db;
  border-radius: 6px; font-size: 0.78rem; font-weight: 600; cursor: pointer; transition: background 0.15s;
}
button.secondary:hover:not(:disabled) { background: #e5e7eb; }
button.secondary:disabled { opacity: 0.5; cursor: not-allowed; }

.download-btn {
  width: 100%; padding: 10px; background: #16a34a; color: #fff; border: none;
  border-radius: 8px; font-size: 0.9rem; font-weight: 600; cursor: pointer;
  transition: background 0.2s; margin-top: 10px; display: none;
}
.download-btn:hover { background: #15803d; }

.progress-wrap { margin-top: 14px; display: none; }
.progress-label { font-size: 0.78rem; color: #6b7280; margin-bottom: 5px; }
.progress-bar-bg { background: #e5e7eb; border-radius: 99px; height: 6px; }
.progress-bar { background: #6366f1; border-radius: 99px; height: 6px; width: 0%; transition: width 0.4s; }
.progress-text { font-size: 0.74rem; color: #6b7280; margin-top: 4px; }

.gen-main { flex: 1; min-width: 0; }

.empty-state {
  background: #fff; border-radius: 12px; box-shadow: 0 4px 24px rgba(0,0,0,0.07);
  padding: 60px 40px; text-align: center; color: #9ca3af;
}
.empty-state .icon { font-size: 3rem; margin-bottom: 12px; }
.empty-state p { font-size: 0.9rem; }

.section-card {
  background: #fff; border-radius: 10px; box-shadow: 0 2px 12px rgba(0,0,0,0.07);
  margin-bottom: 14px; overflow: hidden; animation: fadeIn 0.3s ease;
}
@keyframes fadeIn { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }

.card-header {
  display: flex; align-items: center; gap: 10px;
  padding: 11px 14px; border-bottom: 1px solid #f3f4f6; background: #fafafa;
}
.card-header .section-label { font-size: 0.84rem; font-weight: 700; color: #1a1a2e; flex: 1; }
.card-header .card-actions { display: flex; gap: 6px; }

.card-body { padding: 12px 14px; }

.section-textarea {
  width: 100%; border: 1px solid #e5e7eb; border-radius: 6px; padding: 9px 11px;
  font-size: 0.85rem; font-family: 'Segoe UI', sans-serif; color: #111827;
  resize: vertical; min-height: 70px; outline: none; transition: border-color 0.15s; line-height: 1.6;
}
.section-textarea:focus { border-color: #6366f1; }
.section-textarea.generating { background: #fafafa; color: #9ca3af; font-style: italic; }

.grammar-results { margin-top: 8px; display: none; }
.grammar-issue { background: #fef9c3; border: 1px solid #fde68a; border-radius: 6px; padding: 6px 10px; margin-bottom: 5px; font-size: 0.77rem; }
.gi-msg { color: #92400e; font-weight: 600; margin-bottom: 2px; }
.gi-ctx { color: #78716c; font-family: monospace; }
.gi-fix { color: #15803d; margin-top: 2px; }
.grammar-ok { color: #15803d; font-size: 0.77rem; font-weight: 600; margin-top: 5px; }

.status-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.dot-done { background: #22c55e; }
.dot-generating { background: #6366f1; animation: pulse 1s infinite; }
.dot-na { background: #d1d5db; }
.card-na { opacity: 0.6; }
@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }

.spinner {
  display: inline-block; width: 11px; height: 11px;
  border: 2px solid #c7d2fe; border-top-color: #6366f1;
  border-radius: 50%; animation: spin 0.7s linear infinite; vertical-align: middle;
}
@keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>

<!-- ── Top nav ── -->
<nav class="topnav">
  <div class="brand">Formaty<span>X</span></div>
  <button class="tab-btn active" onclick="switchTab('formatter')">Formatter</button>
  <button class="tab-btn" onclick="switchTab('generator')">AI Generator</button>
</nav>

<!-- ══════════════════════════════════════════════
     FORMATTER PANEL
══════════════════════════════════════════════ -->
<div class="panel active" id="panel-formatter">
<div class="fmt-wrap">
  <h2>Formatter</h2>
  <p class="subtitle">Formatează lucrări academice românești pentru conferință.</p>

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
    <label>Authors <span class="opt">(one per line)</span></label>
    <textarea id="authors" rows="3" placeholder="Gheorghe Badea&#10;George Naghiu&#10;Andrei Măgureanu"></textarea>
  </div>
  <div class="field">
    <label>English Title <span class="opt">(optional — auto-translated if empty)</span></label>
    <input type="text" id="title-en" placeholder="Leave blank to auto-translate">
  </div>
  <div class="field">
    <label style="display:flex;align-items:center;gap:8px;cursor:pointer;">
      <input type="checkbox" id="fast-mode" style="width:auto;margin:0;">
      Fast mode <span class="opt">(skip diacritics restoration — document is already correctly formatted)</span>
    </label>
  </div>
  <div class="field">
    <label style="display:flex;align-items:center;gap:8px;cursor:pointer;">
      <input type="checkbox" id="expand-abbrev" style="width:auto;margin:0;">
      Expand abbreviations <span class="opt">(replace CEE, EMS, etc. with their full defined names)</span>
    </label>
  </div>

  <div class="btn-row">
    <button class="primary" id="run-btn" disabled onclick="startFormatting()">Format All</button>
  </div>

  <div id="log-wrap">
    <div id="log-label">Log — <span id="log-filename"></span></div>
    <div id="log-box"></div>
  </div>
</div>
</div>

<!-- ══════════════════════════════════════════════
     GENERATOR PANEL
══════════════════════════════════════════════ -->
<div class="panel" id="panel-generator">
<div class="gen-layout">

  <div class="gen-sidebar">
    <h2>AI Generator</h2>
    <p class="subtitle">Generare lucrări academice cu Claude AI.</p>

    <div class="field">
      <label>Subiect / Temă *</label>
      <input type="text" id="g-topic" placeholder="ex: Microcogenerare pe biomasă">
    </div>
    <div class="field">
      <label>Domeniu *</label>
      <input type="text" id="g-domain" placeholder="ex: Inginerie energetică">
    </div>
    <div class="field">
      <label>Obiective cercetare</label>
      <textarea id="g-objectives" rows="3" placeholder="ex: Analiza eficienței exergetice a unui sistem de microcogenerare..."></textarea>
    </div>
    <div class="field">
      <label>Cuvinte cheie sugerate <span class="opt">(opțional)</span></label>
      <input type="text" id="g-keywords" placeholder="ex: biomasă, cogenerare, exergie">
    </div>
    <div class="field">
      <label>Autori <span class="opt">(unul pe linie)</span></label>
      <textarea id="g-authors" rows="2" placeholder="Ion Popescu&#10;Maria Ionescu"></textarea>
    </div>

    <button class="gen-primary" id="gen-btn" onclick="startGeneration()">Generează lucrarea</button>
    <button class="download-btn" id="download-btn" onclick="downloadDoc()">Descarcă DOCX</button>

    <div class="progress-wrap" id="progress-wrap">
      <div class="progress-label">Progres generare</div>
      <div class="progress-bar-bg"><div class="progress-bar" id="progress-bar"></div></div>
      <div class="progress-text" id="progress-text"></div>
    </div>
  </div>

  <div class="gen-main" id="gen-main">
    <div class="empty-state" id="empty-state">
      <div class="icon">📝</div>
      <p>Completează formularul și apasă <strong>Generează lucrarea</strong> pentru a începe.</p>
    </div>
    <div id="sections-wrap"></div>
  </div>

</div>
</div>

<script>
/* ────────────────────────────────────────────────
   TAB SWITCHING
──────────────────────────────────────────────── */
function switchTab(name) {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('panel-' + name).classList.add('active');
  event.currentTarget.classList.add('active');
}

/* ────────────────────────────────────────────────
   FORMATTER
──────────────────────────────────────────────── */
const dropZone  = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const queueEl   = document.getElementById('queue');
const queueList = document.getElementById('queue-list');
const runBtn    = document.getElementById('run-btn');
const logWrap   = document.getElementById('log-wrap');
const logBox    = document.getElementById('log-box');

let fileQueue = [];
let fmtProcessing = false;

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('drag-over'); addFiles(Array.from(e.dataTransfer.files)); });
fileInput.addEventListener('change', () => { addFiles(Array.from(fileInput.files)); fileInput.value = ''; });

function addFiles(files) {
  const valid = files.filter(f => f.name.endsWith('.docx'));
  if (valid.length !== files.length) alert('Only .docx files are supported.');
  valid.forEach(f => {
    if (fileQueue.some(q => q.file.name === f.name)) return;
    fileQueue.push({ file: f, id: Math.random().toString(36).slice(2), status: 'pending', jobId: null });
  });
  renderQueue();
}

function renderQueue() {
  if (!fileQueue.length) { queueEl.style.display = 'none'; runBtn.disabled = true; return; }
  queueEl.style.display = 'block';
  queueList.innerHTML = '';
  fileQueue.forEach(item => {
    const row = document.createElement('div');
    row.className = 'file-row';
    const labels = { pending: 'Pending', running: 'Processing', done: '✓ Done', error: '✗ Error' };
    const badgeCls = { pending: 'badge-pending', running: 'badge-running', done: 'badge-done', error: 'badge-error' };
    row.innerHTML = `
      <span class="fname" title="${item.file.name}">${item.file.name}</span>
      <span class="badge ${badgeCls[item.status]}" id="badge-${item.id}">${labels[item.status]}</span>
      ${item.status === 'done' && item.jobId ? `<a class="dl-link" href="/fmt/download/${item.jobId}">Download</a>` : ''}
      ${item.status === 'pending' ? `<button class="remove-btn" onclick="removeFile('${item.id}')">×</button>` : ''}
    `;
    queueList.appendChild(row);
  });
  runBtn.disabled = fmtProcessing || !fileQueue.some(q => q.status === 'pending');
}

function removeFile(id) { fileQueue = fileQueue.filter(q => q.id !== id); renderQueue(); }

function startFormatting() {
  if (fmtProcessing) return;
  fmtProcessing = true; runBtn.disabled = true;
  processFmtNext();
}

async function processFmtNext() {
  const item = fileQueue.find(q => q.status === 'pending');
  if (!item) { fmtProcessing = false; renderQueue(); return; }
  item.status = 'running'; renderQueue();
  logWrap.style.display = 'block';
  document.getElementById('log-filename').textContent = item.file.name;
  logBox.textContent = 'Uploading…';

  const fd = new FormData();
  fd.append('file', item.file);
  const authors = document.getElementById('authors').value.trim();
  if (authors) fd.append('authors', authors);
  const titleEn = document.getElementById('title-en').value.trim();
  if (titleEn) fd.append('title_en', titleEn);
  if (document.getElementById('fast-mode').checked) fd.append('fast_mode', 'on');
  if (document.getElementById('expand-abbrev').checked) fd.append('expand_abbreviations', 'on');

  let jobId;
  try {
    const r = await fetch('/fmt/format', { method: 'POST', body: fd });
    const d = await r.json();
    if (!r.ok) throw new Error(d.error || 'Upload failed');
    jobId = d.job_id; item.jobId = jobId;
  } catch (e) {
    item.status = 'error'; renderQueue(); logBox.textContent = 'Error: ' + e.message;
    fmtProcessing = false; runBtn.disabled = false; return;
  }

  await new Promise(resolve => {
    const t = setInterval(async () => {
      try {
        const r = await fetch('/fmt/status/' + jobId);
        const d = await r.json();
        if (d.log) { logBox.textContent = d.log; logBox.scrollTop = logBox.scrollHeight; }
        if (d.status === 'done' || d.status === 'error') {
          clearInterval(t); item.status = d.status; renderQueue(); resolve();
        }
      } catch(e) {}
    }, 2000);
  });
  processFmtNext();
}

/* ────────────────────────────────────────────────
   GENERATOR
──────────────────────────────────────────────── */
const GEN_SECTIONS = """ + str([{"key": s["key"], "label": s["label"]} for s in SECTIONS]) + """;

let currentJobId = null;
let renderedKeys = new Set();

function startGeneration() {
  const topic = document.getElementById('g-topic').value.trim();
  const domain = document.getElementById('g-domain').value.trim();
  if (!topic || !domain) { alert('Completează Subiectul și Domeniul.'); return; }

  document.getElementById('gen-btn').disabled = true;
  document.getElementById('download-btn').style.display = 'none';
  document.getElementById('empty-state').style.display = 'none';
  document.getElementById('sections-wrap').innerHTML = '';
  document.getElementById('progress-wrap').style.display = 'block';
  renderedKeys = new Set();

  fetch('/gen/generate', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      topic, domain,
      objectives: document.getElementById('g-objectives').value.trim(),
      keywords:   document.getElementById('g-keywords').value.trim(),
      authors:    document.getElementById('g-authors').value.trim(),
    }),
  })
  .then(r => r.json())
  .then(d => {
    if (d.error) { alert(d.error); resetGenUI(); return; }
    currentJobId = d.job_id;
    pollGen();
  })
  .catch(e => { alert('Eroare: ' + e); resetGenUI(); });
}

function pollGen() {
  if (!currentJobId) return;
  fetch('/gen/status/' + currentJobId)
    .then(r => r.json())
    .then(d => {
      updateGenProgress(d);
      renderNewSections(d.sections || {});
      if (d.status === 'done' || d.status === 'error') {
        document.getElementById('gen-btn').disabled = false;
        if (d.status === 'done') document.getElementById('download-btn').style.display = 'block';
        if (d.status === 'error') alert('Eroare: ' + (d.error || 'necunoscută'));
      } else {
        setTimeout(pollGen, 2500);
      }
    })
    .catch(() => setTimeout(pollGen, 3000));
}

function updateGenProgress(d) {
  const done = Object.keys(d.sections || {}).length;
  const total = GEN_SECTIONS.length;
  document.getElementById('progress-bar').style.width = Math.round(done/total*100) + '%';
  const cur = d.current ? GEN_SECTIONS.find(s => s.key === d.current) : null;
  document.getElementById('progress-text').textContent =
    d.status === 'done' ? `Completat (${total}/${total} secțiuni)` :
    cur ? `Generez: ${cur.label} (${done}/${total})` : `${done}/${total} secțiuni`;
}

function renderNewSections(sections) {
  for (const sec of GEN_SECTIONS) {
    if (sections[sec.key] !== undefined && !renderedKeys.has(sec.key)) {
      renderedKeys.add(sec.key);
      appendSectionCard(sec.key, sec.label, sections[sec.key]);
    }
  }
}

function isNA(t) { return !t || t.trim().toUpperCase() === 'N/A'; }

function appendSectionCard(key, label, text) {
  const wrap = document.getElementById('sections-wrap');
  const na = isNA(text);
  const card = document.createElement('div');
  card.className = 'section-card' + (na ? ' card-na' : '');
  card.id = 'card-' + key;

  if (na) {
    card.innerHTML = `
      <div class="card-header">
        <span class="status-dot dot-na" id="dot-${key}"></span>
        <span class="section-label" style="color:#9ca3af">${label}</span>
        <span style="font-size:0.73rem;color:#9ca3af;margin-left:auto;margin-right:8px">Secțiune omisă (N/A)</span>
        <button class="secondary" onclick="regenerateSection('${key}')">Generează</button>
      </div>
      <textarea class="section-textarea" id="ta-${key}" rows="1" style="display:none">N/A</textarea>
      <div class="grammar-results" id="gr-${key}"></div>
    `;
  } else {
    const rows = Math.max(4, Math.ceil(text.length / 80));
    card.innerHTML = `
      <div class="card-header">
        <span class="status-dot dot-done" id="dot-${key}"></span>
        <span class="section-label">${label}</span>
        <div class="card-actions">
          <button class="secondary" onclick="regenerateSection('${key}')">Regenerează</button>
          <button class="secondary" onclick="grammarCheck('${key}')">Gramatică</button>
        </div>
      </div>
      <div class="card-body">
        <textarea class="section-textarea" id="ta-${key}" rows="${rows}">${escHtml(text)}</textarea>
        <div class="grammar-results" id="gr-${key}"></div>
      </div>
    `;
  }
  wrap.appendChild(card);
}

function refreshCard(key, text) {
  const existing = document.getElementById('card-' + key);
  if (existing) existing.remove();
  renderedKeys.delete(key);
  const label = GEN_SECTIONS.find(s => s.key === key).label;
  appendSectionCard(key, label, text);
  // Re-sort cards
  const wrap = document.getElementById('sections-wrap');
  const order = GEN_SECTIONS.map(s => s.key);
  [...wrap.children].sort((a,b) => order.indexOf(a.id.replace('card-','')) - order.indexOf(b.id.replace('card-',''))).forEach(c => wrap.appendChild(c));
  renderedKeys.add(key);
}

function regenerateSection(key) {
  if (!currentJobId) return;
  const overrides = {};
  for (const s of GEN_SECTIONS) {
    const el = document.getElementById('ta-' + s.key);
    if (el && el.value && el.value !== 'Generez...') overrides[s.key] = el.value;
  }
  const ta = document.getElementById('ta-' + key);
  if (ta) { ta.value = 'Generez...'; ta.style.display = ''; ta.className = 'section-textarea generating'; }
  document.getElementById('dot-' + key).className = 'status-dot dot-generating';

  fetch('/gen/regenerate', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ job_id: currentJobId, key, overrides }),
  })
  .then(r => r.json())
  .then(d => refreshCard(key, d.error ? '[Eroare: ' + d.error + ']' : d.text))
  .catch(() => refreshCard(key, '[Eroare rețea]'));
}

function grammarCheck(key) {
  const ta = document.getElementById('ta-' + key);
  const gr = document.getElementById('gr-' + key);
  const text = ta ? ta.value.trim() : '';
  if (!text || isNA(text)) return;
  gr.style.display = 'block';
  gr.innerHTML = '<span class="spinner"></span> Verificare gramaticală...';
  fetch('/gen/grammar', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ text }),
  })
  .then(r => r.json())
  .then(d => {
    const issues = d.issues || [];
    gr.innerHTML = !issues.length
      ? '<div class="grammar-ok">✓ Nicio problemă detectată.</div>'
      : issues.map(i => `<div class="grammar-issue"><div class="gi-msg">${escHtml(i.message)}</div><div class="gi-ctx">"${escHtml(i.context)}"</div>${i.replacements.length?`<div class="gi-fix">Sugestii: ${i.replacements.map(r=>`<strong>${escHtml(r)}</strong>`).join(', ')}</div>`:''}</div>`).join('');
  })
  .catch(() => { gr.innerHTML = '<div class="grammar-issue"><div class="gi-msg">Eroare la verificare.</div></div>'; });
}

function downloadDoc() {
  if (!currentJobId) return;
  const sections = {};
  for (const s of GEN_SECTIONS) {
    const el = document.getElementById('ta-' + s.key);
    if (el) sections[s.key] = el.value;
  }
  const authors = document.getElementById('g-authors').value.trim();
  fetch('/gen/download/' + currentJobId, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ sections, authors }),
  })
  .then(r => { if (!r.ok) throw new Error('Download failed'); return r.blob(); })
  .then(blob => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = 'paper_generated.docx'; a.click();
    URL.revokeObjectURL(url);
  })
  .catch(e => alert('Eroare: ' + e));
}

function resetGenUI() {
  document.getElementById('gen-btn').disabled = false;
  document.getElementById('progress-wrap').style.display = 'none';
}

function escHtml(t) {
  return String(t).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
</script>
</body>
</html>
"""


# ─────────────────────────────────────────────────────────────────────────────
# FORMATTER ROUTES  (/fmt/...)
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/fmt/format", methods=["POST"])
def fmt_format():
    if "file" not in request.files:
        return jsonify({"error": "No file provided."}), 400
    f = request.files["file"]
    if not f.filename.endswith(".docx"):
        return jsonify({"error": "Only .docx files are supported."}), 400

    job_id = str(uuid.uuid4())
    input_path = UPLOAD_DIR / f"{job_id}_{f.filename}"
    output_path = OUTPUT_DIR / f"{Path(f.filename).stem}_formatted.docx"
    f.save(input_path)

    authors_raw = request.form.get("authors", "").strip()
    authors = [a.strip() for a in authors_raw.splitlines() if a.strip()]
    title_en = request.form.get("title_en", "").strip()
    fast_mode = request.form.get("fast_mode") == "on"
    expand_abbrev = request.form.get("expand_abbreviations") == "on"

    fmt_jobs[job_id] = {"status": "running", "message": "", "output_file": str(output_path), "log": "", "filename": f.filename}
    threading.Thread(target=_run_formatter, args=(job_id, str(input_path), str(output_path), authors, title_en, fast_mode, expand_abbrev), daemon=True).start()
    return jsonify({"job_id": job_id})


def _run_formatter(job_id, input_path, output_path, authors, title_en, fast_mode=False, expand_abbrev=False):
    cmd = [sys.executable, "-u", "formatter.py", input_path, "-o", output_path]
    if authors:
        cmd += ["--authors"] + authors
    if title_en:
        cmd += ["--title-en", title_en]
    if fast_mode:
        cmd.append("--skip-diacritics")
    if expand_abbrev:
        cmd.append("--expand-abbreviations")
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, encoding="utf-8", errors="replace", cwd=BASE_DIR)
        log_lines = []
        for line in proc.stdout:
            log_lines.append(line.rstrip())
            fmt_jobs[job_id]["log"] = "\n".join(log_lines)
        proc.wait()
        if proc.returncode == 0:
            summary = next((l for l in reversed(log_lines) if l.startswith("Done!")), "")
            fmt_jobs[job_id]["status"] = "done"
            fmt_jobs[job_id]["message"] = summary
        else:
            fmt_jobs[job_id]["status"] = "error"
            fmt_jobs[job_id]["message"] = "\n".join(log_lines[-10:])
    except Exception as e:
        fmt_jobs[job_id]["status"] = "error"
        fmt_jobs[job_id]["message"] = str(e)


@app.route("/fmt/status/<job_id>")
def fmt_status(job_id):
    job = fmt_jobs.get(job_id)
    if not job:
        return jsonify({"status": "error", "message": "Unknown job."}), 404
    return jsonify(job)


@app.route("/fmt/download/<job_id>")
def fmt_download(job_id):
    job = fmt_jobs.get(job_id)
    if not job or job["status"] != "done":
        return "Not ready.", 404
    if not os.path.exists(job["output_file"]):
        return "File not found.", 404
    return send_file(job["output_file"], as_attachment=True)


# ─────────────────────────────────────────────────────────────────────────────
# GENERATOR ROUTES  (/gen/...)
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/gen/generate", methods=["POST"])
def gen_generate():
    data = request.get_json()
    topic = (data.get("topic") or "").strip()
    domain = (data.get("domain") or "").strip()
    if not topic or not domain:
        return jsonify({"error": "Topic and domain are required."}), 400

    job_id = str(uuid.uuid4())
    gen_jobs[job_id] = {
        "status": "running", "sections": {}, "current": None, "error": None,
        "meta": {"topic": topic, "domain": domain,
                 "objectives": data.get("objectives", ""),
                 "keywords": data.get("keywords", ""),
                 "authors": data.get("authors", "")},
    }
    threading.Thread(target=_run_generation, args=(job_id,), daemon=True).start()
    return jsonify({"job_id": job_id})


def _run_generation(job_id):
    job = gen_jobs[job_id]
    meta = job["meta"]
    generated = {}
    try:
        for section in SECTIONS:
            key = section["key"]
            job["current"] = key
            text = generate_section(key=key, topic=meta["topic"], domain=meta["domain"],
                                    objectives=meta["objectives"], keywords=meta["keywords"],
                                    generated=generated)
            generated[key] = text
            job["sections"][key] = text
        job["status"] = "done"
        job["current"] = None
    except Exception as e:
        job["status"] = "error"
        job["error"] = str(e)


@app.route("/gen/status/<job_id>")
def gen_status(job_id):
    job = gen_jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job."}), 404
    return jsonify({"status": job["status"], "sections": job["sections"],
                    "current": job["current"], "error": job["error"]})


@app.route("/gen/regenerate", methods=["POST"])
def gen_regenerate():
    data = request.get_json()
    job = gen_jobs.get(data.get("job_id"))
    if not job:
        return jsonify({"error": "Unknown job."}), 404
    key = data.get("key")
    context = {**job["sections"], **data.get("overrides", {})}
    context.pop(key, None)
    meta = job["meta"]
    try:
        text = generate_section(key=key, topic=meta["topic"], domain=meta["domain"],
                                objectives=meta["objectives"], keywords=meta["keywords"],
                                generated=context)
        job["sections"][key] = text
        return jsonify({"text": text})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/gen/grammar", methods=["POST"])
def gen_grammar():
    data = request.get_json()
    text = (data.get("text") or "").strip()
    return jsonify({"issues": check_grammar(text) if text else []})


@app.route("/gen/download/<job_id>", methods=["POST"])
def gen_download(job_id):
    job = gen_jobs.get(job_id)
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
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting FormatyX at http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
