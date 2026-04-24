/* ==========================================
   ZePdf — app.js
   CloudConvert API integration + full UI
   API key is injected via Netlify env variable
   CLOUDCONVERT_API_KEY (see _config.js / build)
   ========================================== */

// ──────────────────────────────────────────
// API KEY  — injected by Netlify at build time
// The _config.js file sets window.ZEPDF_CONFIG
// ──────────────────────────────────────────
const API_KEY = (window.ZEPDF_CONFIG && window.ZEPDF_CONFIG.apiKey) || '';

// ──────────────────────────────────────────
// STATE
// ──────────────────────────────────────────
const state = {
  files: [],
  history: JSON.parse(localStorage.getItem('zepdf_history') || '[]'),
  convertedToday: parseInt(localStorage.getItem('zepdf_today_count') || '0'),
  todayKey: localStorage.getItem('zepdf_today_key') || '',
};

// ──────────────────────────────────────────
// DOM REFS
// ──────────────────────────────────────────
const dropZone      = document.getElementById('dropZone');
const fileInput     = document.getElementById('fileInput');
const browseBtn     = document.getElementById('browseBtn');
const fileQueue     = document.getElementById('fileQueue');
const queueList     = document.getElementById('queueList');
const convertBtn    = document.getElementById('convertBtn');
const globalFormat  = document.getElementById('globalFormat');
const clearAllBtn   = document.getElementById('clearAllBtn');
const themeToggle   = document.getElementById('themeToggle');
const historyBtn    = document.getElementById('historyBtn');
const historyPanel  = document.getElementById('historyPanel');
const historyOverlay= document.getElementById('historyOverlay');
const closeHistoryBtn= document.getElementById('closeHistoryBtn');
const clearHistoryBtn= document.getElementById('clearHistoryBtn');
const historyList   = document.getElementById('historyList');
const historyCount  = document.getElementById('historyCount');
const statConverted = document.getElementById('statConverted');
const toastContainer= document.getElementById('toastContainer');

// ──────────────────────────────────────────
// THEME
// ──────────────────────────────────────────
const savedTheme = localStorage.getItem('zepdf_theme') || 'dark';
document.documentElement.setAttribute('data-theme', savedTheme);

themeToggle.addEventListener('click', () => {
  const current = document.documentElement.getAttribute('data-theme');
  const next = current === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem('zepdf_theme', next);
});


// ──────────────────────────────────────────
// FILE ICON HELPER
// ──────────────────────────────────────────
function getFileType(ext) {
  const e = ext.replace('.','').toLowerCase();
  if (e === 'pdf') return { label: 'PDF', cls: 'type-pdf' };
  if (['docx','doc','rtf','odt','txt'].includes(e)) return { label: e.toUpperCase(), cls: 'type-word' };
  if (['xlsx','xls','ods'].includes(e)) return { label: e.toUpperCase(), cls: 'type-excel' };
  if (['pptx','ppt','odp'].includes(e)) return { label: e.toUpperCase(), cls: 'type-ppt' };
  if (['jpg','jpeg','png','gif','webp'].includes(e)) return { label: e.toUpperCase(), cls: 'type-image' };
  if (['epub','mobi'].includes(e)) return { label: e.toUpperCase(), cls: 'type-ebook' };
  return { label: e.toUpperCase(), cls: 'type-other' };
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

// ──────────────────────────────────────────
// DRAG & DROP / FILE SELECTION
// ──────────────────────────────────────────
const MAX_SIZE = 200 * 1024 * 1024; // 200 MB
const ACCEPTED = ['.pdf','.docx','.doc','.pptx','.ppt','.xlsx','.xls','.rtf','.odt','.odp','.ods','.epub','.mobi','.jpg','.jpeg','.png'];

dropZone.addEventListener('click', () => fileInput.click());
browseBtn.addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });
fileInput.addEventListener('change', (e) => handleFiles(Array.from(e.target.files)));

dropZone.addEventListener('dragenter', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragover',  (e) => { e.preventDefault(); });
dropZone.addEventListener('dragleave', (e) => { if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over'); });
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  handleFiles(Array.from(e.dataTransfer.files));
});

function handleFiles(files) {
  let added = 0;
  files.forEach(file => {
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!ACCEPTED.includes(ext)) {
      showToast(`"${file.name}" — unsupported format.`, 'error');
      return;
    }
    if (file.size > MAX_SIZE) {
      showToast(`"${file.name}" exceeds the 200 MB limit.`, 'error');
      return;
    }
    // Avoid duplicates
    if (state.files.find(f => f.file.name === file.name && f.file.size === file.size)) return;
    const id = 'f_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    state.files.push({ id, file, targetFormat: '', status: 'idle', downloadUrl: null });
    added++;
  });
  if (added) renderQueue();
}

// ──────────────────────────────────────────
// RENDER QUEUE
// ──────────────────────────────────────────
function renderQueue() {
  if (state.files.length === 0) {
    fileQueue.style.display = 'none';
    return;
  }
  fileQueue.style.display = 'block';
  queueList.innerHTML = '';

  state.files.forEach(item => {
    const ext = '.' + item.file.name.split('.').pop().toLowerCase();
    const typeInfo = getFileType(ext);

    const div = document.createElement('div');
    div.className = 'queue-item';
    div.id = 'qi_' + item.id;
    div.innerHTML = `
      <div class="file-icon ${typeInfo.cls}">${typeInfo.label.slice(0,4)}</div>
      <div class="file-info">
        <div class="file-name" title="${item.file.name}">${item.file.name}</div>
        <div class="file-size">${formatBytes(item.file.size)}</div>
      </div>
      <select class="item-format-select" id="fmt_${item.id}">
        <option value="">Auto</option>
        <optgroup label="Documents">
          <option value="pdf">PDF</option><option value="docx">DOCX</option>
          <option value="doc">DOC</option><option value="pptx">PPTX</option>
          <option value="xlsx">XLSX</option><option value="rtf">RTF</option>
          <option value="txt">TXT</option>
        </optgroup>
        <optgroup label="OpenDocument">
          <option value="odt">ODT</option><option value="odp">ODP</option><option value="ods">ODS</option>
        </optgroup>
        <optgroup label="Images">
          <option value="jpg">JPG</option><option value="png">PNG</option>
          <option value="webp">WEBP</option><option value="gif">GIF</option>
        </optgroup>
        <optgroup label="eBook">
          <option value="epub">EPUB</option><option value="mobi">MOBI</option>
        </optgroup>
      </select>
      <button class="remove-item-btn" id="rm_${item.id}" title="Remove">✕</button>
      <div class="item-progress" id="prg_${item.id}"><div class="item-progress-bar" id="bar_${item.id}"></div></div>
      <div class="item-status" id="sts_${item.id}"></div>
      <a class="download-link" id="dl_${item.id}" download>
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
          <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
        </svg>
        Download
      </a>`;
    queueList.appendChild(div);

    // Set saved format
    if (item.targetFormat) document.getElementById('fmt_' + item.id).value = item.targetFormat;

    // Per-file format change
    document.getElementById('fmt_' + item.id).addEventListener('change', (e) => {
      item.targetFormat = e.target.value;
      updateConvertBtn();
    });

    // Remove button
    document.getElementById('rm_' + item.id).addEventListener('click', () => {
      state.files = state.files.filter(f => f.id !== item.id);
      renderQueue();
    });
  });

  updateConvertBtn();
}

function updateConvertBtn() {
  const hasFormat = !!globalFormat.value || state.files.some(f => f.targetFormat);
  const hasFiles = state.files.length > 0;
  convertBtn.disabled = !(hasFiles && hasFormat);
}

globalFormat.addEventListener('change', updateConvertBtn);

clearAllBtn.addEventListener('click', () => {
  state.files = [];
  renderQueue();
  fileInput.value = '';
});

// ──────────────────────────────────────────
// CLOUDCONVERT CONVERSION
// ──────────────────────────────────────────
convertBtn.addEventListener('click', async () => {
  if (!API_KEY) {
    showToast('Conversion service not configured. Please contact the administrator.', 'error');
    return;
  }

  const filesToConvert = state.files.filter(f => f.status !== 'done');
  if (!filesToConvert.length) { showToast('All files already converted!', 'info'); return; }

  convertBtn.classList.add('loading');
  convertBtn.innerHTML = `<svg class="spin" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg> Converting…`;

  for (const item of filesToConvert) {
    const fmt = item.targetFormat || globalFormat.value;
    if (!fmt) {
      setItemStatus(item.id, 'error', 'No target format selected.');
      continue;
    }
    await convertFile(item, fmt);
  }

  convertBtn.classList.remove('loading');
  convertBtn.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/></svg> Convert Now`;
});

async function convertFile(item, targetFormat) {
  setProgress(item.id, 5);
  setItemStatus(item.id, 'active', 'Uploading…');

  try {
    // Step 1 — Create Job
    const inputFmt = item.file.name.split('.').pop().toLowerCase();
    const jobRes = await fetch('https://api.cloudconvert.com/v2/jobs', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${API_KEY}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tasks: {
          'upload-task': { operation: 'import/upload' },
          'convert-task': {
            operation: 'convert',
            input: 'upload-task',
            input_format: inputFmt,
            output_format: targetFormat,
          },
          'export-task': {
            operation: 'export/url',
            input: 'convert-task',
          },
        },
      }),
    });

    if (!jobRes.ok) {
      const err = await jobRes.json();
      throw new Error(err.message || `HTTP ${jobRes.status}`);
    }

    const jobData = await jobRes.json();
    const uploadTask = jobData.data.tasks.find(t => t.name === 'upload-task');
    setProgress(item.id, 20);
    setItemStatus(item.id, 'active', 'Uploading file…');

    // Step 2 — Upload file
    const form = new FormData();
    Object.entries(uploadTask.result.form.parameters).forEach(([k, v]) => form.append(k, v));
    form.append('file', item.file);

    const uploadRes = await fetch(uploadTask.result.form.url, { method: 'POST', body: form });
    if (!uploadRes.ok) throw new Error('Upload failed');

    setProgress(item.id, 55);
    setItemStatus(item.id, 'active', 'Converting…');

    // Step 3 — Poll for completion
    const jobId = jobData.data.id;
    const exportUrl = await pollJob(jobId);

    setProgress(item.id, 90);
    setItemStatus(item.id, 'active', 'Preparing download…');

    // Step 4 — Fetch download URL and trigger download
    const fileRes = await fetch(exportUrl);
    const blob = await fileRes.blob();
    const blobUrl = URL.createObjectURL(blob);
    const baseName = item.file.name.replace(/\.[^/.]+$/, '');
    const dlName = `${baseName}.${targetFormat}`;

    const dlLink = document.getElementById('dl_' + item.id);
    dlLink.href = blobUrl;
    dlLink.download = dlName;
    dlLink.classList.add('active');

    setProgress(item.id, 100);
    setItemStatus(item.id, 'success', `Done — ${targetFormat.toUpperCase()}`);
    item.status = 'done';

    // Save to history
    addToHistory({ name: dlName, from: inputFmt, to: targetFormat, size: item.file.size, date: new Date().toISOString() });

    // Update today counter
    updateTodayCount();
    showToast(`"${item.file.name}" converted successfully!`, 'success');

  } catch (err) {
    setProgress(item.id, 0);
    setItemStatus(item.id, 'error', err.message);
    showToast(`Failed: ${err.message}`, 'error');
    console.error(err);
  }
}

// Poll job until done
async function pollJob(jobId) {
  const maxAttempts = 60;
  for (let i = 0; i < maxAttempts; i++) {
    await sleep(2000);
    const res = await fetch(`https://api.cloudconvert.com/v2/jobs/${jobId}`, {
      headers: { 'Authorization': `Bearer ${API_KEY}` },
    });
    const data = await res.json();
    const job = data.data;

    if (job.status === 'finished') {
      const exportTask = job.tasks.find(t => t.name === 'export-task');
      return exportTask.result.files[0].url;
    }
    if (job.status === 'error') {
      const errTask = job.tasks.find(t => t.status === 'error');
      throw new Error(errTask?.message || 'Conversion error on server.');
    }
  }
  throw new Error('Conversion timed out.');
}

// ──────────────────────────────────────────
// PROGRESS / STATUS HELPERS
// ──────────────────────────────────────────
function setProgress(id, pct) {
  const prg = document.getElementById('prg_' + id);
  const bar = document.getElementById('bar_' + id);
  if (!prg || !bar) return;
  prg.classList.add('active');
  bar.style.width = pct + '%';
}

function setItemStatus(id, type, msg) {
  const el = document.getElementById('sts_' + id);
  if (!el) return;
  el.className = 'item-status active ' + type;
  el.textContent = msg;
}

// ──────────────────────────────────────────
// HISTORY
// ──────────────────────────────────────────
function addToHistory(entry) {
  state.history.unshift(entry);
  if (state.history.length > 50) state.history.pop();
  localStorage.setItem('zepdf_history', JSON.stringify(state.history));
  renderHistory();
}

function renderHistory() {
  const count = state.history.length;
  if (count > 0) {
    historyCount.textContent = count;
    historyCount.style.display = 'flex';
  } else {
    historyCount.style.display = 'none';
  }

  if (!count) {
    historyList.innerHTML = '<p class="empty-history">No conversions yet. Start converting!</p>';
    return;
  }
  historyList.innerHTML = state.history.map(h => `
    <div class="history-item">
      <div class="history-item-name">${h.name}</div>
      <div class="history-item-meta">${h.from.toUpperCase()} → ${h.to.toUpperCase()} · ${formatBytes(h.size)}</div>
      <span class="history-item-badge">${new Date(h.date).toLocaleDateString()}</span>
    </div>`).join('');
}
renderHistory();

historyBtn.addEventListener('click', () => {
  historyPanel.classList.add('open');
  historyOverlay.classList.add('open');
});
closeHistoryBtn.addEventListener('click', closeHistory);
historyOverlay.addEventListener('click', closeHistory);
function closeHistory() {
  historyPanel.classList.remove('open');
  historyOverlay.classList.remove('open');
}

clearHistoryBtn.addEventListener('click', () => {
  state.history = [];
  localStorage.removeItem('zepdf_history');
  renderHistory();
  showToast('History cleared.', 'info');
});

// ──────────────────────────────────────────
// TODAY COUNTER
// ──────────────────────────────────────────
function updateTodayCount() {
  const today = new Date().toDateString();
  if (state.todayKey !== today) {
    state.convertedToday = 0;
    state.todayKey = today;
    localStorage.setItem('zepdf_today_key', today);
  }
  state.convertedToday++;
  localStorage.setItem('zepdf_today_count', state.convertedToday);
  statConverted.textContent = state.convertedToday;
}

// Restore today's count on load
(function initTodayCount() {
  const today = new Date().toDateString();
  if (state.todayKey === today) {
    statConverted.textContent = state.convertedToday;
  } else {
    statConverted.textContent = 0;
  }
})();

// ──────────────────────────────────────────
// TOAST NOTIFICATIONS
// ──────────────────────────────────────────
function showToast(msg, type = 'info', duration = 4000) {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `<span>${msg}</span><button class="toast-dismiss">✕</button>`;
  toastContainer.appendChild(toast);
  toast.querySelector('.toast-dismiss').addEventListener('click', () => toast.remove());
  setTimeout(() => toast.remove(), duration);
}

// ──────────────────────────────────────────
// SPINNING ICON CSS (injected)
// ──────────────────────────────────────────
const spinStyle = document.createElement('style');
spinStyle.textContent = `
  .spin { animation: spin 1s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }
`;
document.head.appendChild(spinStyle);

// ──────────────────────────────────────────
// ──────────────────────────────────────────
// PDF TOOLS NAV BUTTON
// ──────────────────────────────────────────
document.getElementById('navPdfToolsBtn').addEventListener('click', () => {
  const section = document.getElementById('pdfTools');
  if (section) section.scrollIntoView({ behavior: 'smooth', block: 'start' });
});

// ──────────────────────────────────────────
// UTILS
// ──────────────────────────────────────────
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
