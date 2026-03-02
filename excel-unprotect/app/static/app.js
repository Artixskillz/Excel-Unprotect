/* ── Element refs ──────────────────────────────────────────────── */
const dropZone        = document.getElementById('dropZone');
const fileInput       = document.getElementById('fileInput');

const uploadView      = document.getElementById('uploadView');
const processingView  = document.getElementById('processingView');
const resultView      = document.getElementById('resultView');
const errorView       = document.getElementById('errorView');

const processingFile  = document.getElementById('processingFilename');
const resultHeadline  = document.getElementById('resultHeadline');
const resultFilename  = document.getElementById('resultFilename');
const statSheets      = document.getElementById('statSheets');
const statWorkbook    = document.getElementById('statWorkbook');
const downloadBtn     = document.getElementById('downloadBtn');
const errorMessage    = document.getElementById('errorMessage');

const emptyState      = document.getElementById('emptyState');
const tableWrap       = document.getElementById('tableWrap');
const historyBody     = document.getElementById('historyBody');
const historyCount    = document.getElementById('historyCount');

/* ── Drop-zone wiring ──────────────────────────────────────────── */
dropZone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', e => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
});

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('dragging');
});

dropZone.addEventListener('dragleave', e => {
  if (!dropZone.contains(e.relatedTarget)) {
    dropZone.classList.remove('dragging');
  }
});

dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragging');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

/* ── Core upload flow ──────────────────────────────────────────── */
async function handleFile(file) {
  showView('processing');
  processingFile.textContent = file.name;

  const body = new FormData();
  body.append('file', file);

  try {
    const res  = await fetch('/api/upload', { method: 'POST', body });
    const data = await res.json();

    if (!res.ok) {
      showError(data.detail ?? 'Upload failed. Please try again.');
      return;
    }

    showResult(data);
    loadHistory();          // refresh the table
  } catch (err) {
    showError('Network error – ' + err.message);
  }
}

/* ── Result rendering ──────────────────────────────────────────── */
function showResult(data) {
  const total = data.sheets_unprotected + (data.workbook_unprotected ? 1 : 0);

  resultHeadline.textContent =
    total > 0 ? 'Protection Removed!' : 'File Processed';

  resultFilename.textContent = data.filename;

  statSheets.textContent  = data.sheets_unprotected;
  statWorkbook.textContent = data.workbook_unprotected ? 'Yes' : 'No';

  // Wire download button
  downloadBtn.onclick = () => triggerDownload(data.id, data.filename);

  showView('result');
}

function showError(msg) {
  errorMessage.textContent = msg;
  showView('error');
}

/* ── View switcher ─────────────────────────────────────────────── */
function showView(name) {
  uploadView.classList.toggle    ('hidden', name !== 'upload');
  processingView.classList.toggle('hidden', name !== 'processing');
  resultView.classList.toggle    ('hidden', name !== 'result');
  errorView.classList.toggle     ('hidden', name !== 'error');
}

/* ── Reset ─────────────────────────────────────────────────────── */
function resetUpload() {
  fileInput.value = '';
  showView('upload');
}

/* ── Download helper ───────────────────────────────────────────── */
function triggerDownload(id, originalName) {
  const a = document.createElement('a');
  a.href     = `/api/download/${id}`;
  a.download = `unlocked_${originalName}`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

/* ── History ───────────────────────────────────────────────────── */
async function loadHistory() {
  try {
    const res   = await fetch('/api/files');
    const files = await res.json();
    renderHistory(files);
  } catch (err) {
    console.error('History load failed:', err);
  }
}

function renderHistory(files) {
  if (files.length === 0) {
    emptyState.classList.remove('hidden');
    tableWrap.classList.add('hidden');
    historyCount.textContent = '';
    return;
  }

  emptyState.classList.add('hidden');
  tableWrap.classList.remove('hidden');
  historyCount.textContent = files.length;

  historyBody.innerHTML = files.map(f => `
    <tr>
      <td>
        <div class="file-cell">
          <div class="xl-icon">
            ${xlSvg()}
          </div>
          <span class="file-name-text" title="${esc(f.original_filename)}">
            ${esc(f.original_filename)}
          </span>
        </div>
      </td>
      <td class="date-cell">${fmtDate(f.upload_time)}</td>
      <td class="size-cell">${fmtSize(f.file_size_bytes)}</td>
      <td>
        ${f.sheets_unprotected > 0
          ? `<span class="badge badge-yes">✓ ${f.sheets_unprotected} sheet${f.sheets_unprotected !== 1 ? 's' : ''}</span>`
          : `<span class="badge badge-no">None</span>`}
      </td>
      <td>
        ${f.workbook_unprotected
          ? `<span class="badge badge-yes">✓ Yes</span>`
          : `<span class="badge badge-no">No</span>`}
      </td>
      <td>
        <button class="dl-btn" onclick="triggerDownload('${esc(f.id)}', '${esc(f.original_filename)}')">
          ${dlSvg()} Download
        </button>
      </td>
    </tr>
  `).join('');
}

/* ── Formatting helpers ────────────────────────────────────────── */
function fmtDate(iso) {
  if (!iso) return '—';
  return new Date(iso + 'Z').toLocaleString(undefined, {
    month:   'short',
    day:     'numeric',
    year:    'numeric',
    hour:    '2-digit',
    minute:  '2-digit',
  });
}

function fmtSize(bytes) {
  if (!bytes) return '—';
  if (bytes < 1_024)           return bytes + ' B';
  if (bytes < 1_048_576)       return (bytes / 1_024).toFixed(1) + ' KB';
  return (bytes / 1_048_576).toFixed(1) + ' MB';
}

function esc(str) {
  return String(str)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;');
}

/* ── Inline SVGs ───────────────────────────────────────────────── */
function xlSvg() {
  return `<svg width="16" height="16" viewBox="0 0 24 24" fill="none">
    <path d="M8 7 L12.5 12 L8 17 H10.8 L13.5 13.3 L16.2 17 H19 L14.5 12 L19 7 H16.2 L13.5 10.7 L10.8 7 Z"
          fill="white"/>
  </svg>`;
}

function dlSvg() {
  return `<svg width="13" height="13" viewBox="0 0 24 24" fill="none"
    stroke="currentColor" stroke-width="2.5" stroke-linecap="round">
    <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
    <polyline points="7 10 12 15 17 10"/>
    <line x1="12" y1="15" x2="12" y2="3"/>
  </svg>`;
}

/* ── Boot ──────────────────────────────────────────────────────── */
showView('upload');
loadHistory();
