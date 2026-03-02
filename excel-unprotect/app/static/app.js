/* ── Element refs ──────────────────────────────────────────────── */
const dropZone       = document.getElementById('dropZone');
const fileInput      = document.getElementById('fileInput');

const uploadView     = document.getElementById('uploadView');
const processingView = document.getElementById('processingView');
const resultView     = document.getElementById('resultView');
const errorView      = document.getElementById('errorView');

const processingFile = document.getElementById('processingFilename');
const resultHeadline = document.getElementById('resultHeadline');
const resultFilename = document.getElementById('resultFilename');
const statSheets     = document.getElementById('statSheets');
const statWorkbook   = document.getElementById('statWorkbook');
const sheetNamesList = document.getElementById('sheetNamesList');
const sheetNamesTags = document.getElementById('sheetNamesTags');
const downloadBtn    = document.getElementById('downloadBtn');
const errorMessage   = document.getElementById('errorMessage');

const emptyState     = document.getElementById('emptyState');
const tableWrap      = document.getElementById('tableWrap');
const historyBody    = document.getElementById('historyBody');
const historyCount   = document.getElementById('historyCount');
const purgeBtn       = document.getElementById('purgeBtn');

const confirmModal   = document.getElementById('confirmModal');
const confirmTitle   = document.getElementById('confirmTitle');
const confirmMsg     = document.getElementById('confirmMessage');
const confirmOkBtn   = document.getElementById('confirmOkBtn');
const confirmCancelBtn = document.getElementById('confirmCancelBtn');
const themeToggle    = document.getElementById('themeToggle');

/* ── Dark mode ─────────────────────────────────────────────────── */
function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('theme', theme);
}

function initTheme() {
  const saved = localStorage.getItem('theme');
  if (saved === 'dark' || saved === 'light') {
    applyTheme(saved);
  } else if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
    applyTheme('dark');
  } else {
    applyTheme('light');
  }
}

themeToggle.addEventListener('click', () => {
  const current = document.documentElement.getAttribute('data-theme');
  applyTheme(current === 'dark' ? 'light' : 'dark');
});

/* ── Confirmation modal ────────────────────────────────────────── */
function showConfirm(title, message, okLabel, onConfirm) {
  confirmTitle.textContent  = title;
  confirmMsg.textContent    = message;
  confirmOkBtn.textContent  = okLabel;
  confirmModal.classList.remove('hidden');

  confirmOkBtn.onclick = () => {
    confirmModal.classList.add('hidden');
    onConfirm();
  };
}

confirmCancelBtn.onclick = () => confirmModal.classList.add('hidden');

// Close modal on overlay click
confirmModal.addEventListener('click', e => {
  if (e.target === confirmModal) confirmModal.classList.add('hidden');
});

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
  if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('dragging');
});

dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragging');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
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
    if (!res.ok) { showError(data.detail ?? 'Upload failed. Please try again.'); return; }
    showResult(data);
    loadHistory();
  } catch (err) {
    showError('Network error – ' + err.message);
  }
}

/* ── Result rendering ──────────────────────────────────────────── */
function showResult(data) {
  const total = data.sheets_unprotected + (data.workbook_unprotected ? 1 : 0);
  resultHeadline.textContent = total > 0 ? 'Protection Removed!' : 'File Processed';
  resultFilename.textContent = data.filename;
  statSheets.textContent     = data.sheets_unprotected;
  statWorkbook.textContent   = data.workbook_unprotected ? 'Yes' : 'No';

  // Sheet names
  const names = data.unlocked_sheet_names || [];
  if (names.length > 0) {
    sheetNamesTags.innerHTML = names.map(n => `<span class="sheet-tag">${esc(n)}</span>`).join('');
    sheetNamesList.classList.remove('hidden');
  } else {
    sheetNamesList.classList.add('hidden');
  }

  downloadBtn.onclick = () => triggerDownload(`/api/download/${data.id}`, `unlocked_${data.filename}`);
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

function resetUpload() {
  fileInput.value = '';
  showView('upload');
}

/* ── Download helper ───────────────────────────────────────────── */
function triggerDownload(url, filename) {
  const a = document.createElement('a');
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

/* ── Delete single file ────────────────────────────────────────── */
async function deleteFile(id, filename) {
  showConfirm(
    'Delete File?',
    `"${filename}" and its unlocked version will be permanently removed from the server.`,
    'Yes, Delete',
    async () => {
      try {
        const res = await fetch(`/api/files/${id}`, { method: 'DELETE' });
        if (!res.ok) { alert('Delete failed.'); return; }
        loadHistory();
      } catch (err) {
        alert('Network error: ' + err.message);
      }
    }
  );
}

/* ── Purge all ─────────────────────────────────────────────────── */
function confirmPurgeAll() {
  const count = parseInt(historyCount.textContent, 10) || 0;
  showConfirm(
    'Clear All History?',
    `This will permanently delete all ${count} processed file${count !== 1 ? 's' : ''} and their originals from the server.`,
    'Yes, Delete Everything',
    async () => {
      try {
        const res = await fetch('/api/files', { method: 'DELETE' });
        if (!res.ok) { alert('Purge failed.'); return; }
        loadHistory();
      } catch (err) {
        alert('Network error: ' + err.message);
      }
    }
  );
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
  // Show / hide the "Clear All" button
  purgeBtn.style.display = files.length > 0 ? '' : 'none';

  if (files.length === 0) {
    emptyState.classList.remove('hidden');
    tableWrap.classList.add('hidden');
    historyCount.textContent = '';
    return;
  }

  emptyState.classList.add('hidden');
  tableWrap.classList.remove('hidden');
  historyCount.textContent = files.length;

  historyBody.innerHTML = files.map(f => {
    const names = f.unlocked_sheet_names || [];
    const sheetBadge = f.sheets_unprotected > 0
      ? `<span class="badge badge-yes">✓ ${f.sheets_unprotected} sheet${f.sheets_unprotected !== 1 ? 's' : ''}</span>`
      : `<span class="badge badge-no">None</span>`;

    const sheetNameTags = names.length > 0
      ? `<div class="sheet-names-inline">${names.map(n => `<span class="sheet-tag-sm">${esc(n)}</span>`).join('')}</div>`
      : '';

    const wbBadge = f.workbook_unprotected
      ? `<span class="badge badge-yes">✓ Yes</span>`
      : `<span class="badge badge-no">No</span>`;

    const origBtn = f.has_original
      ? `<button class="dl-btn-sm orig" onclick="triggerDownload('/api/download/${esc(f.id)}/original', '${esc(f.original_filename)}')">${dlSvg()} Original</button>`
      : `<button class="dl-btn-sm orig" disabled title="Original not stored">${dlSvg()} Original</button>`;

    return `
      <tr>
        <td>
          <div class="file-cell">
            <div class="xl-icon">${xlSvg()}</div>
            <span class="file-name-text" title="${esc(f.original_filename)}">${esc(f.original_filename)}</span>
          </div>
        </td>
        <td class="date-cell">${fmtDate(f.upload_time)}</td>
        <td class="size-cell">${fmtSize(f.file_size_bytes)}</td>
        <td>
          ${sheetBadge}
          ${sheetNameTags}
        </td>
        <td>${wbBadge}</td>
        <td>
          <div class="dl-buttons">
            ${origBtn}
            <button class="dl-btn-sm" onclick="triggerDownload('/api/download/${esc(f.id)}', 'unlocked_${esc(f.original_filename)}')">${lockOpenSvg()} Unlocked</button>
          </div>
        </td>
        <td>
          <button class="delete-btn" title="Delete this entry" onclick="deleteFile('${esc(f.id)}', '${esc(f.original_filename)}')">${trashSvg()}</button>
        </td>
      </tr>`;
  }).join('');
}

/* ── Formatting helpers ────────────────────────────────────────── */
function fmtDate(iso) {
  if (!iso) return '—';
  return new Date(iso + 'Z').toLocaleString(undefined, {
    month: 'short', day: 'numeric', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });
}

function fmtSize(bytes) {
  if (!bytes) return '—';
  if (bytes < 1_024)       return bytes + ' B';
  if (bytes < 1_048_576)   return (bytes / 1_024).toFixed(1) + ' KB';
  return (bytes / 1_048_576).toFixed(1) + ' MB';
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/* ── Inline SVGs ───────────────────────────────────────────────── */
function xlSvg() {
  return `<svg width="16" height="16" viewBox="0 0 24 24" fill="none">
    <path d="M8 7L12.5 12L8 17H10.8L13.5 13.3L16.2 17H19L14.5 12L19 7H16.2L13.5 10.7L10.8 7Z" fill="white"/>
  </svg>`;
}

function dlSvg() {
  return `<svg width="12" height="12" viewBox="0 0 24 24" fill="none"
    stroke="currentColor" stroke-width="2.5" stroke-linecap="round">
    <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
    <polyline points="7 10 12 15 17 10"/>
    <line x1="12" y1="15" x2="12" y2="3"/>
  </svg>`;
}

function lockOpenSvg() {
  return `<svg width="12" height="12" viewBox="0 0 24 24" fill="none"
    stroke="currentColor" stroke-width="2.5" stroke-linecap="round">
    <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
    <path d="M7 11V7a5 5 0 019.9-1"/>
  </svg>`;
}

function trashSvg() {
  return `<svg width="14" height="14" viewBox="0 0 24 24" fill="none"
    stroke="currentColor" stroke-width="2" stroke-linecap="round">
    <polyline points="3 6 5 6 21 6"/>
    <path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/>
    <path d="M10 11v6M14 11v6"/>
    <path d="M9 6V4a1 1 0 011-1h4a1 1 0 011 1v2"/>
  </svg>`;
}

/* ── Boot ──────────────────────────────────────────────────────── */
initTheme();
showView('upload');
loadHistory();
