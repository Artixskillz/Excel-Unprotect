/* ═══════════════════════════════════════════════════════════════
   Excel Unprotect — app.js
   ═══════════════════════════════════════════════════════════════ */

/* ── State ─────────────────────────────────────────────────────── */
let allFiles    = [];           // full list from server
let sortCol     = 'upload_time';
let sortDir     = 'desc';
let searchQuery = '';
let deferredInstallPrompt = null;

/* ── Element refs ──────────────────────────────────────────────── */
const dropZone          = document.getElementById('dropZone');
const fileInput         = document.getElementById('fileInput');
const dzHeading         = document.getElementById('dzHeading');

const uploadView        = document.getElementById('uploadView');
const processingView    = document.getElementById('processingView');
const batchView         = document.getElementById('batchView');
const resultView        = document.getElementById('resultView');
const errorView         = document.getElementById('errorView');

const processingFile    = document.getElementById('processingFilename');
const resultHeadline    = document.getElementById('resultHeadline');
const resultFilename    = document.getElementById('resultFilename');
const statSheets        = document.getElementById('statSheets');
const statWorkbook      = document.getElementById('statWorkbook');
const sheetNamesList    = document.getElementById('sheetNamesList');
const sheetNamesTags    = document.getElementById('sheetNamesTags');
const downloadBtn       = document.getElementById('downloadBtn');
const errorMessage      = document.getElementById('errorMessage');

// Batch
const batchTotal        = document.getElementById('batchTotal');
const batchStatusText   = document.getElementById('batchStatusText');
const batchProgressFill = document.getElementById('batchProgressFill');
const batchList         = document.getElementById('batchList');
const batchDoneRow      = document.getElementById('batchDoneRow');

// History
const emptyState        = document.getElementById('emptyState');
const noResultsState    = document.getElementById('noResultsState');
const noResultsHint     = document.getElementById('noResultsHint');
const tableWrap         = document.getElementById('tableWrap');
const historyBody       = document.getElementById('historyBody');
const historyCount      = document.getElementById('historyCount');
const purgeBtn          = document.getElementById('purgeBtn');
const searchInput       = document.getElementById('searchInput');

// Modal
const confirmModal      = document.getElementById('confirmModal');
const confirmTitle      = document.getElementById('confirmTitle');
const confirmMsg        = document.getElementById('confirmMessage');
const confirmOkBtn      = document.getElementById('confirmOkBtn');
const confirmCancelBtn  = document.getElementById('confirmCancelBtn');

// Header
const themeToggle       = document.getElementById('themeToggle');
const installBtn        = document.getElementById('installBtn');


/* ── PWA — Service Worker ───────────────────────────────────────── */
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('/sw.js').catch(err =>
    console.warn('SW registration failed:', err)
  );
}

/* ── PWA — Install prompt ───────────────────────────────────────── */
window.addEventListener('beforeinstallprompt', e => {
  e.preventDefault();
  deferredInstallPrompt = e;
  installBtn.classList.add('visible');
});

window.addEventListener('appinstalled', () => {
  installBtn.classList.remove('visible');
  deferredInstallPrompt = null;
});

installBtn.addEventListener('click', async () => {
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  const { outcome } = await deferredInstallPrompt.userChoice;
  if (outcome === 'accepted') installBtn.classList.remove('visible');
  deferredInstallPrompt = null;
});


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
  confirmTitle.textContent = title;
  confirmMsg.textContent   = message;
  confirmOkBtn.textContent = okLabel;
  confirmModal.classList.remove('hidden');

  confirmOkBtn.onclick = () => {
    confirmModal.classList.add('hidden');
    onConfirm();
  };
}

confirmCancelBtn.onclick = () => confirmModal.classList.add('hidden');
confirmModal.addEventListener('click', e => {
  if (e.target === confirmModal) confirmModal.classList.add('hidden');
});


/* ── Drop-zone wiring (enhanced) ───────────────────────────────── */
// Prevent browser from opening files dropped anywhere on the page
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop',     e => e.preventDefault());

dropZone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', e => {
  if (e.target.files.length) handleFiles(Array.from(e.target.files));
});

// Use a counter to handle nested dragenter/dragleave correctly
let dragCounter = 0;

dropZone.addEventListener('dragenter', e => {
  e.preventDefault();
  dragCounter++;
  if (dragCounter === 1) {
    dropZone.classList.add('dragging');
    dzHeading.textContent = 'Release to drop!';
  }
});

dropZone.addEventListener('dragleave', () => {
  dragCounter--;
  if (dragCounter <= 0) {
    dragCounter = 0;
    dropZone.classList.remove('dragging');
    dzHeading.textContent = 'Drop your Excel files here';
  }
});

dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dragCounter = 0;
  dropZone.classList.remove('dragging');
  dzHeading.textContent = 'Drop your Excel files here';
  const files = Array.from(e.dataTransfer.files);
  if (files.length) handleFiles(files);
});


/* ── File dispatch ─────────────────────────────────────────────── */
function handleFiles(files) {
  // Filter to only allowed extensions
  const allowed = ['.xlsx', '.xlsm', '.xltx', '.xltm'];
  const valid   = files.filter(f => allowed.includes(extOf(f.name)));

  if (valid.length === 0) {
    showError('No supported Excel files selected. Please choose .xlsx, .xlsm, .xltx, or .xltm files.');
    return;
  }

  if (valid.length === 1) {
    handleSingleFile(valid[0]);
  } else {
    handleBatch(valid);
  }
}

function extOf(filename) {
  const idx = filename.lastIndexOf('.');
  return idx >= 0 ? filename.slice(idx).toLowerCase() : '';
}


/* ── Single file flow ──────────────────────────────────────────── */
async function handleSingleFile(file) {
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
    showError('Network error — ' + err.message);
  }
}


/* ── Batch upload flow ─────────────────────────────────────────── */
async function handleBatch(files) {
  showView('batch');
  batchTotal.textContent = files.length;
  batchDoneRow.classList.add('hidden');
  batchProgressFill.style.width = '0%';
  batchStatusText.textContent   = `0 of ${files.length} complete`;

  // Build the queue list
  batchList.innerHTML = files.map((f, i) => `
    <div class="batch-item" id="bi-${i}">
      <div class="batch-item-left">
        <div class="xl-icon">${xlSvg()}</div>
        <span class="batch-item-name" title="${esc(f.name)}">${esc(f.name)}</span>
      </div>
      <span class="batch-item-status pending" id="bi-${i}-status">Pending</span>
    </div>
  `).join('');

  // Process sequentially
  let done = 0;
  for (let i = 0; i < files.length; i++) {
    setBatchStatus(i, 'processing', null);

    try {
      const body = new FormData();
      body.append('file', files[i]);
      const res  = await fetch('/api/upload', { method: 'POST', body });
      const data = await res.json();

      if (res.ok) {
        const n   = data.sheets_unprotected;
        const wb  = data.workbook_unprotected;
        const msg = (n > 0 || wb)
          ? `✓ ${n} sheet${n !== 1 ? 's' : ''}${wb ? ' + workbook' : ''} unlocked`
          : '✓ No protections found';
        setBatchStatus(i, 'done', msg);
      } else {
        setBatchStatus(i, 'error', data.detail ?? 'Failed');
      }
    } catch {
      setBatchStatus(i, 'error', 'Network error');
    }

    done++;
    batchProgressFill.style.width = `${(done / files.length) * 100}%`;
    batchStatusText.textContent   = `${done} of ${files.length} complete`;
  }

  batchDoneRow.classList.remove('hidden');
  loadHistory();
}

function setBatchStatus(index, status, text) {
  const el = document.getElementById(`bi-${index}-status`);
  if (!el) return;
  el.className = `batch-item-status ${status}`;

  if (status === 'processing') {
    el.innerHTML = `<div class="batch-item-spinner"></div> Processing…`;
  } else {
    el.textContent = text;
  }
}


/* ── Result rendering ──────────────────────────────────────────── */
function showResult(data) {
  const total = data.sheets_unprotected + (data.workbook_unprotected ? 1 : 0);
  resultHeadline.textContent = total > 0 ? 'Protection Removed!' : 'File Processed';
  resultFilename.textContent = data.filename;
  statSheets.textContent     = data.sheets_unprotected;
  statWorkbook.textContent   = data.workbook_unprotected ? 'Yes' : 'No';

  const names = data.unlocked_sheet_names || [];
  if (names.length > 0) {
    sheetNamesTags.innerHTML = names.map(n => `<span class="sheet-tag">${esc(n)}</span>`).join('');
    sheetNamesList.classList.remove('hidden');
  } else {
    sheetNamesList.classList.add('hidden');
  }

  downloadBtn.onclick = () =>
    triggerDownload(`/api/download/${data.id}`, `unlocked_${data.filename}`);
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
  batchView.classList.toggle     ('hidden', name !== 'batch');
  resultView.classList.toggle    ('hidden', name !== 'result');
  errorView.classList.toggle     ('hidden', name !== 'error');
}

function resetUpload() {
  fileInput.value = '';
  dragCounter = 0;
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
  const count = allFiles.length;
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


/* ── Sort & Search ─────────────────────────────────────────────── */
function sortBy(col) {
  if (sortCol === col) {
    sortDir = sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    sortCol = col;
    // Dates and sizes default to newest/largest first; names default to A-Z
    sortDir = (col === 'upload_time' || col === 'file_size_bytes' || col === 'sheets_unprotected')
      ? 'desc' : 'asc';
  }
  updateSortIcons();
  renderFilteredHistory();
}

function updateSortIcons() {
  const cols = ['original_filename', 'upload_time', 'file_size_bytes', 'sheets_unprotected'];
  cols.forEach(col => {
    const el = document.getElementById(`sort-${col}`);
    if (!el) return;
    if (col === sortCol) {
      el.textContent = sortDir === 'asc' ? '↑' : '↓';
      el.classList.add('active');
    } else {
      el.textContent = '↕';
      el.classList.remove('active');
    }
  });
}

searchInput.addEventListener('input', () => {
  searchQuery = searchInput.value.trim().toLowerCase();
  renderFilteredHistory();
});

function renderFilteredHistory() {
  // Filter
  let list = allFiles;
  if (searchQuery) {
    list = allFiles.filter(f =>
      f.original_filename.toLowerCase().includes(searchQuery)
    );
  }

  // Sort
  const dir = sortDir === 'asc' ? 1 : -1;
  list = [...list].sort((a, b) => {
    const va = a[sortCol] ?? '';
    const vb = b[sortCol] ?? '';
    if (typeof va === 'number') return (va - vb) * dir;
    return String(va).localeCompare(String(vb)) * dir;
  });

  renderHistory(list);
}


/* ── History load ──────────────────────────────────────────────── */
async function loadHistory() {
  try {
    const res   = await fetch('/api/files');
    allFiles    = await res.json();
    renderFilteredHistory();
  } catch (err) {
    console.error('History load failed:', err);
  }
}


/* ── History render ────────────────────────────────────────────── */
function renderHistory(files) {
  const hasAny     = allFiles.length > 0;
  const hasResults = files.length > 0;

  purgeBtn.style.display = hasAny ? '' : 'none';
  historyCount.textContent = hasAny ? allFiles.length : '';

  // Nothing in DB at all
  if (!hasAny) {
    emptyState.classList.remove('hidden');
    noResultsState.classList.add('hidden');
    tableWrap.classList.add('hidden');
    return;
  }

  emptyState.classList.add('hidden');

  // Has data but search returned nothing
  if (!hasResults) {
    noResultsState.classList.remove('hidden');
    noResultsHint.textContent = `No files matching "${searchQuery}"`;
    tableWrap.classList.add('hidden');
    return;
  }

  noResultsState.classList.add('hidden');
  tableWrap.classList.remove('hidden');

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
      ? `<button class="dl-btn-sm orig" onclick="triggerDownload('/api/download/${esc(f.id)}/original','${esc(f.original_filename)}')">${dlSvg()} Original</button>`
      : `<button class="dl-btn-sm orig" disabled title="Original not stored">${dlSvg()} Original</button>`;

    return `
      <tr>
        <td data-label="Filename">
          <div class="file-cell">
            <div class="xl-icon">${xlSvg()}</div>
            <span class="file-name-text" title="${esc(f.original_filename)}">${esc(f.original_filename)}</span>
          </div>
        </td>
        <td data-label="Date"   class="date-cell">${fmtDate(f.upload_time)}</td>
        <td data-label="Size"   class="size-cell">${fmtSize(f.file_size_bytes)}</td>
        <td data-label="Sheets">${sheetBadge}${sheetNameTags}</td>
        <td data-label="Workbook">${wbBadge}</td>
        <td data-label="Download">
          <div class="dl-buttons">
            ${origBtn}
            <button class="dl-btn-sm" onclick="triggerDownload('/api/download/${esc(f.id)}','unlocked_${esc(f.original_filename)}')">${lockOpenSvg()} Unlocked</button>
          </div>
        </td>
        <td data-label="">
          <button class="delete-btn" title="Delete entry" onclick="deleteFile('${esc(f.id)}','${esc(f.original_filename)}')">${trashSvg()}</button>
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
  if (bytes < 1_024)     return bytes + ' B';
  if (bytes < 1_048_576) return (bytes / 1_024).toFixed(1) + ' KB';
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
updateSortIcons();
showView('upload');
loadHistory();
