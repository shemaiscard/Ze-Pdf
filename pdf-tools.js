/* ==========================================
   ZePdf — pdf-tools.js
   Client-side PDF Split & Merge via pdf-lib
   No API key required — runs 100% in browser
   ========================================== */

// ──────────────────────────────────────────
// TAB SWITCHING
// ──────────────────────────────────────────
document.querySelectorAll('.pdf-tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.pdf-tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    const which = tab.dataset.tab;
    document.getElementById('panelMerge').style.display = which === 'merge' ? 'block' : 'none';
    document.getElementById('panelSplit').style.display = which === 'split'  ? 'block' : 'none';
  });
});

// ──────────────────────────────────────────
// MERGE — state + UI
// ──────────────────────────────────────────
let mergeFiles = []; // array of File objects (ordered)
let dragSrcIndex = null;

const mergeDropZone   = document.getElementById('mergeDropZone');
const mergeInput      = document.getElementById('mergeInput');
const mergeList       = document.getElementById('mergeList');
const mergeItems      = document.getElementById('mergeItems');
const mergeFileCount  = document.getElementById('mergeFileCount');
const mergePdfBtn     = document.getElementById('mergePdfBtn');
const clearMergeBtn   = document.getElementById('clearMergeBtn');

// Click to browse
mergeDropZone.addEventListener('click', () => mergeInput.click());
mergeInput.addEventListener('change', e => addMergeFiles(Array.from(e.target.files)));

// Drag & drop
mergeDropZone.addEventListener('dragenter', e => { e.preventDefault(); mergeDropZone.classList.add('drag-over'); });
mergeDropZone.addEventListener('dragover',  e => { e.preventDefault(); });
mergeDropZone.addEventListener('dragleave', e => { if (!mergeDropZone.contains(e.relatedTarget)) mergeDropZone.classList.remove('drag-over'); });
mergeDropZone.addEventListener('drop', e => {
  e.preventDefault();
  mergeDropZone.classList.remove('drag-over');
  addMergeFiles(Array.from(e.dataTransfer.files).filter(f => f.name.toLowerCase().endsWith('.pdf')));
});

clearMergeBtn.addEventListener('click', () => {
  mergeFiles = [];
  renderMergeList();
});

function addMergeFiles(files) {
  files.forEach(f => {
    if (!f.name.toLowerCase().endsWith('.pdf')) {
      showToast(`"${f.name}" is not a PDF.`, 'error');
      return;
    }
    if (!mergeFiles.find(x => x.name === f.name && x.size === f.size)) {
      mergeFiles.push(f);
    }
  });
  renderMergeList();
}

function renderMergeList() {
  if (!mergeFiles.length) {
    mergeList.style.display = 'none';
    mergePdfBtn.disabled = true;
    return;
  }
  mergeList.style.display = 'block';
  mergePdfBtn.disabled = false;
  mergeFileCount.textContent = `${mergeFiles.length} file${mergeFiles.length > 1 ? 's' : ''} selected`;
  mergeItems.innerHTML = mergeFiles.map((f, i) => `
    <div class="merge-item" draggable="true" data-index="${i}" id="mi_${i}">
      <span class="merge-drag-handle">⠿</span>
      <span class="merge-item-name" title="${f.name}">${f.name}</span>
      <span class="merge-item-size">${formatBytesTools(f.size)}</span>
      <button class="merge-remove-btn" data-remove="${i}" title="Remove">✕</button>
    </div>`).join('');

  // Remove buttons
  mergeItems.querySelectorAll('.merge-remove-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      mergeFiles.splice(parseInt(btn.dataset.remove), 1);
      renderMergeList();
    });
  });

  // Drag-to-reorder
  mergeItems.querySelectorAll('.merge-item').forEach(item => {
    item.addEventListener('dragstart', e => {
      dragSrcIndex = parseInt(item.dataset.index);
      item.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
    });
    item.addEventListener('dragend', () => {
      item.classList.remove('dragging');
      mergeItems.querySelectorAll('.merge-item').forEach(i => i.classList.remove('drag-over-item'));
    });
    item.addEventListener('dragover', e => {
      e.preventDefault();
      mergeItems.querySelectorAll('.merge-item').forEach(i => i.classList.remove('drag-over-item'));
      item.classList.add('drag-over-item');
    });
    item.addEventListener('drop', e => {
      e.preventDefault();
      const targetIndex = parseInt(item.dataset.index);
      if (dragSrcIndex === null || dragSrcIndex === targetIndex) return;
      const moved = mergeFiles.splice(dragSrcIndex, 1)[0];
      mergeFiles.splice(targetIndex, 0, moved);
      dragSrcIndex = null;
      renderMergeList();
    });
  });
}

// ── MERGE ACTION ──
mergePdfBtn.addEventListener('click', async () => {
  if (!mergeFiles.length) return;
  if (typeof PDFLib === 'undefined') {
    showToast('pdf-lib not loaded. Check your connection.', 'error');
    return;
  }

  mergePdfBtn.disabled = true;
  mergePdfBtn.innerHTML = `<svg class="spin" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg> Merging…`;

  try {
    const { PDFDocument } = PDFLib;
    const merged = await PDFDocument.create();

    for (const file of mergeFiles) {
      const bytes = await file.arrayBuffer();
      const srcDoc = await PDFDocument.load(bytes);
      const pages = await merged.copyPages(srcDoc, srcDoc.getPageIndices());
      pages.forEach(p => merged.addPage(p));
    }

    const mergedBytes = await merged.save();
    downloadBytes(mergedBytes, 'merged.pdf', 'application/pdf');

    showToast(`Merged ${mergeFiles.length} PDFs successfully!`, 'success');

    // Log to history
    addToHistory({
      name: 'merged.pdf',
      from: 'pdf (x' + mergeFiles.length + ')',
      to: 'pdf',
      size: mergedBytes.byteLength,
      date: new Date().toISOString(),
    });
    updateTodayCount();

  } catch (err) {
    showToast('Merge failed: ' + err.message, 'error');
    console.error(err);
  }

  mergePdfBtn.disabled = false;
  mergePdfBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 6H3v12h5M16 6h5v12h-5M12 3v18"/></svg> Merge & Download`;
});

// ──────────────────────────────────────────
// SPLIT — state + UI
// ──────────────────────────────────────────
let splitFile = null;
let splitPageCount = 0;

const splitDropZone = document.getElementById('splitDropZone');
const splitInput    = document.getElementById('splitInput');
const splitOptions  = document.getElementById('splitOptions');
const splitFileInfo = document.getElementById('splitFileInfo');
const splitPdfBtn   = document.getElementById('splitPdfBtn');
const splitRangeInput = document.getElementById('splitRangeInput');
const splitEveryInput = document.getElementById('splitEveryInput');
const pageRangeInput  = document.getElementById('pageRangeInput');
const everyNInput     = document.getElementById('everyNInput');
const rangeHint       = document.getElementById('rangeHint');

splitDropZone.addEventListener('click', () => splitInput.click());
splitInput.addEventListener('change', e => { if (e.target.files[0]) loadSplitFile(e.target.files[0]); });

splitDropZone.addEventListener('dragenter', e => { e.preventDefault(); splitDropZone.classList.add('drag-over'); });
splitDropZone.addEventListener('dragover',  e => { e.preventDefault(); });
splitDropZone.addEventListener('dragleave', e => { if (!splitDropZone.contains(e.relatedTarget)) splitDropZone.classList.remove('drag-over'); });
splitDropZone.addEventListener('drop', e => {
  e.preventDefault();
  splitDropZone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f && f.name.toLowerCase().endsWith('.pdf')) loadSplitFile(f);
  else showToast('Please drop a PDF file.', 'error');
});

async function loadSplitFile(file) {
  try {
    const bytes = await file.arrayBuffer();
    const { PDFDocument } = PDFLib;
    const doc = await PDFDocument.load(bytes);
    splitPageCount = doc.getPageCount();
    splitFile = file;
    splitFileInfo.textContent = `📄 ${file.name} — ${splitPageCount} page${splitPageCount > 1 ? 's' : ''} · ${formatBytesTools(file.size)}`;
    splitOptions.style.display = 'block';
    rangeHint.textContent = `Pages 1 to ${splitPageCount}`;
  } catch (err) {
    showToast('Could not read PDF: ' + err.message, 'error');
  }
}

// Toggle mode inputs
document.querySelectorAll('input[name="splitMode"]').forEach(radio => {
  radio.addEventListener('change', () => {
    splitRangeInput.style.display = radio.value === 'range' ? 'block' : 'none';
    splitEveryInput.style.display = radio.value === 'every' ? 'block' : 'none';
  });
});

// ── SPLIT ACTION ──
splitPdfBtn.addEventListener('click', async () => {
  if (!splitFile) return;
  if (typeof PDFLib === 'undefined') {
    showToast('pdf-lib not loaded. Check your connection.', 'error');
    return;
  }

  const mode = document.querySelector('input[name="splitMode"]:checked').value;
  splitPdfBtn.disabled = true;
  splitPdfBtn.innerHTML = `<svg class="spin" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg> Splitting…`;

  try {
    const { PDFDocument } = PDFLib;
    const srcBytes = await splitFile.arrayBuffer();
    const srcDoc = await PDFDocument.load(srcBytes);
    const baseName = splitFile.name.replace(/\.pdf$/i, '');

    let chunks = []; // array of 0-based page index arrays

    if (mode === 'all') {
      chunks = Array.from({ length: splitPageCount }, (_, i) => [i]);
    } else if (mode === 'range') {
      const rangeStr = pageRangeInput.value.trim();
      if (!rangeStr) { showToast('Please enter a page range.', 'error'); throw new Error('no range'); }
      const pages = parsePageRange(rangeStr, splitPageCount);
      if (!pages.length) { showToast('Invalid page range.', 'error'); throw new Error('bad range'); }
      chunks = [pages];
    } else if (mode === 'every') {
      const n = parseInt(everyNInput.value) || 1;
      for (let i = 0; i < splitPageCount; i += n) {
        chunks.push(Array.from({ length: Math.min(n, splitPageCount - i) }, (_, j) => i + j));
      }
    }

    if (chunks.length === 1) {
      // Single output — direct download
      const newDoc = await PDFDocument.create();
      const copied = await newDoc.copyPages(srcDoc, chunks[0]);
      copied.forEach(p => newDoc.addPage(p));
      const bytes = await newDoc.save();
      const suffix = mode === 'range' ? '_extracted' : '_page1';
      downloadBytes(bytes, `${baseName}${suffix}.pdf`, 'application/pdf');
    } else {
      // Multiple outputs — download as zip using JSZip if available, else sequential
      if (typeof JSZip !== 'undefined') {
        const zip = new JSZip();
        for (let ci = 0; ci < chunks.length; ci++) {
          const newDoc = await PDFDocument.create();
          const copied = await newDoc.copyPages(srcDoc, chunks[ci]);
          copied.forEach(p => newDoc.addPage(p));
          const bytes = await newDoc.save();
          const label = mode === 'all' ? `page_${chunks[ci][0] + 1}` : `part_${ci + 1}`;
          zip.file(`${baseName}_${label}.pdf`, bytes);
        }
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        downloadBlob(zipBlob, `${baseName}_split.zip`);
      } else {
        // No zip — download files one by one with small delay
        for (let ci = 0; ci < chunks.length; ci++) {
          const newDoc = await PDFDocument.create();
          const copied = await newDoc.copyPages(srcDoc, chunks[ci]);
          copied.forEach(p => newDoc.addPage(p));
          const bytes = await newDoc.save();
          const label = mode === 'all' ? `page_${chunks[ci][0] + 1}` : `part_${ci + 1}`;
          downloadBytes(bytes, `${baseName}_${label}.pdf`, 'application/pdf');
          await sleep(120);
        }
      }
    }

    showToast(`Split into ${chunks.length} file${chunks.length > 1 ? 's' : ''}!`, 'success');
    addToHistory({ name: baseName + '_split.pdf', from: 'pdf', to: 'pdf (split)', size: splitFile.size, date: new Date().toISOString() });
    updateTodayCount();

  } catch (err) {
    if (err.message !== 'no range' && err.message !== 'bad range') {
      showToast('Split failed: ' + err.message, 'error');
      console.error(err);
    }
  }

  splitPdfBtn.disabled = false;
  splitPdfBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 6h5v12h-5M8 6H3v12h5M12 3v18"/></svg> Split & Download`;
});

// ──────────────────────────────────────────
// HELPERS
// ──────────────────────────────────────────
function parsePageRange(str, total) {
  // Parses "1-3, 5, 7-9" into 0-based indices
  const indices = new Set();
  str.split(',').forEach(part => {
    part = part.trim();
    if (part.includes('-')) {
      const [a, b] = part.split('-').map(Number);
      for (let i = a; i <= b && i <= total; i++) {
        if (i >= 1) indices.add(i - 1);
      }
    } else {
      const n = parseInt(part);
      if (n >= 1 && n <= total) indices.add(n - 1);
    }
  });
  return Array.from(indices).sort((a, b) => a - b);
}

function downloadBytes(bytes, filename, mime) {
  const blob = new Blob([bytes], { type: mime });
  downloadBlob(blob, filename);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
}

function formatBytesTools(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
