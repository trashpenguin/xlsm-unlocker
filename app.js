let processedFiles = [];
let pendingFile = null;
let pendingZip = null;
let pendingSheets = []; // { path, name, protected }
let pendingWbProtected = false;

const uploadCard = document.getElementById('upload-card');
const selectCard = document.getElementById('select-card');
const resultCard = document.getElementById('result-card');
const resetBtn   = document.getElementById('reset-btn');
const dropZone   = document.getElementById('drop-zone');
const fileInput  = document.getElementById('file-input');

// ── Whole-page drag overlay ──────────────────────────────────────────────────
let dragCounter = 0;

document.addEventListener('dragenter', e => {
  if ([...( e.dataTransfer?.types ?? [])].includes('Files')) {
    e.preventDefault();
    if (++dragCounter === 1) document.body.classList.add('dragging');
  }
});
document.addEventListener('dragleave', () => {
  if (--dragCounter <= 0) { dragCounter = 0; document.body.classList.remove('dragging'); }
});
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop', e => {
  e.preventDefault();
  dragCounter = 0;
  document.body.classList.remove('dragging');
  const files = validFiles(e.dataTransfer.files);
  if (files.length) handleFiles(files);
});

// ── Paste support ────────────────────────────────────────────────────────────
document.addEventListener('paste', e => {
  const files = validFiles(e.clipboardData?.files ?? []);
  if (files.length) handleFiles(files);
});

// ── Drop zone ────────────────────────────────────────────────────────────────
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('active'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('active'));
dropZone.addEventListener('drop', () => dropZone.classList.remove('active'));
dropZone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') fileInput.click(); });
fileInput.addEventListener('change', () => {
  const files = validFiles(fileInput.files);
  if (files.length) handleFiles(files);
});

// ── Helpers ──────────────────────────────────────────────────────────────────
function validFiles(list) {
  return [...(list ?? [])].filter(f => /\.(xlsm|xlsx)$/i.test(f.name));
}

function fmtSize(b) {
  return b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.round(b / 1024) + ' KB';
}

function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function friendlyError(err) {
  const m = err.message || '';
  if (/not a zip/i.test(m))           return 'This doesn\'t appear to be a valid Excel file.';
  if (/corrupt/i.test(m))             return 'The file appears to be corrupted or incomplete.';
  if (/password|encrypt/i.test(m))    return 'This file is password-encrypted. This tool removes protection tags only, not file-level encryption.';
  return 'Something went wrong: ' + m;
}

function isValidZip(buf) {
  const b = new Uint8Array(buf, 0, 4);
  return b[0] === 0x50 && b[1] === 0x4B;
}

// ── Workbook metadata from relationships + workbook.xml ───────────────────────
async function getWorkbookMeta(zip) {
  const nameMap = {};
  const rIdToPath = {};
  const hiddenPaths = new Set();
  let hasWbProtection = false;
  try {
    const relsEntry = zip.files['xl/_rels/workbook.xml.rels'];
    const wbEntry   = zip.files['xl/workbook.xml'];
    if (!relsEntry || !wbEntry) return { nameMap, rIdToPath, hiddenPaths, hasWbProtection };

    const relsText = await relsEntry.async('text');
    const wbText   = await wbEntry.async('text');

    hasWbProtection = /<workbookProtection/i.test(wbText);

    for (const m of relsText.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/gi)) {
      const target = m[2].replace(/\.\.\//g, '');
      rIdToPath[m[1]] = 'xl/' + target;
    }

    for (const m of wbText.matchAll(/<sheet\b([^/]*?)(?:\/?>)/gi)) {
      const attrs = m[1];
      const nameM  = attrs.match(/\bname="([^"]+)"/i);
      const ridM   = attrs.match(/\br:id="([^"]+)"/i);
      const stateM = attrs.match(/\bstate="(?:hidden|veryHidden)"/i);
      if (nameM && ridM) {
        const path = rIdToPath[ridM[1]];
        if (path) {
          nameMap[path] = nameM[1];
          if (stateM) hiddenPaths.add(path);
        }
      }
    }
  } catch (_) {}
  return { nameMap, rIdToPath, hiddenPaths, hasWbProtection };
}

// ── Dispatch ─────────────────────────────────────────────────────────────────
async function handleFiles(files) {
  if (files.length === 1) {
    await scanForSelection(files[0]);
  } else {
    await processBatch(files);
  }
}

// ── Single-file: scan then show sheet selection ───────────────────────────────
async function scanForSelection(file) {
  uploadCard.style.display = 'none';
  try {
    const buf = await file.arrayBuffer();
    if (!isValidZip(buf)) throw new Error('not a zip file');

    const zip = await JSZip.loadAsync(buf);
    const { nameMap: sheetNameMap, hiddenPaths, hasWbProtection } = await getWorkbookMeta(zip);

    pendingFile        = file;
    pendingZip         = zip;
    pendingSheets      = [];
    pendingWbProtected = hasWbProtection;

    for (const [path, entry] of Object.entries(zip.files)) {
      if (entry.dir || !/^xl\/worksheets\/.+\.xml$/i.test(path)) continue;
      const text = await entry.async('text');
      pendingSheets.push({
        path,
        name: sheetNameMap[path] || ('Sheet ' + (pendingSheets.length + 1)),
        protected: /<sheetProtection/i.test(text),
        hidden: hiddenPaths.has(path),
      });
    }

    if (!pendingWbProtected && !pendingSheets.some(s => s.protected) && !pendingSheets.some(s => s.hidden)) {
      await processFile(file, zip, null);
      return;
    }

    showSelectCard(file);
  } catch (err) {
    resultCard.style.display = 'block';
    document.getElementById('file-results').innerHTML =
      `<div class="error-box" style="display:block;margin:1.5rem">${esc(friendlyError(err))}</div>`;
    resetBtn.style.display = 'block';
  }
}

function showSelectCard(file) {
  document.getElementById('select-filename').textContent = file.name;
  const list = document.getElementById('sheet-list');
  list.innerHTML = '';

  if (pendingWbProtected) {
    const item = document.createElement('label');
    item.className = 'sheet-item';
    item.innerHTML = `<input type="checkbox" checked disabled value="__wb__">
      <span class="sheet-name">Workbook structure lock</span>
      <span class="sheet-badge badge-protected">Protected</span>`;
    list.appendChild(item);
  }

  pendingSheets.filter(s => s.protected || s.hidden).forEach(sheet => {
    const item = document.createElement('label');
    item.className = 'sheet-item';
    let badgeClass, badgeText;
    if (sheet.protected && sheet.hidden) {
      badgeClass = 'badge-protected'; badgeText = 'Protected + Hidden';
    } else if (sheet.protected) {
      badgeClass = 'badge-protected'; badgeText = 'Protected';
    } else {
      badgeClass = 'badge-hidden'; badgeText = 'Hidden';
    }
    item.innerHTML = `<input type="checkbox" checked value="${esc(sheet.path)}">
      <span class="sheet-name">${esc(sheet.name)}</span>
      <span class="sheet-badge ${badgeClass}">${badgeText}</span>`;
    list.appendChild(item);
  });

  selectCard.style.display = 'block';
}

function selectAll() {
  document.querySelectorAll('#sheet-list input[type=checkbox]:not([disabled])').forEach(cb => cb.checked = true);
}
function deselectAll() {
  document.querySelectorAll('#sheet-list input[type=checkbox]:not([disabled])').forEach(cb => cb.checked = false);
}

async function proceedUnlock() {
  const selectedPaths = new Set(
    [...document.querySelectorAll('#sheet-list input[type=checkbox]:not([disabled]):checked')].map(cb => cb.value)
  );
  selectCard.style.display = 'none';
  await processFile(pendingFile, pendingZip, selectedPaths);
}

// ── Core: process single file ─────────────────────────────────────────────────
async function processFile(file, zip, selectedPaths) {
  uploadCard.style.display = 'none';
  processedFiles = [];

  resultCard.style.display = 'block';
  document.getElementById('file-results').innerHTML = '';
  document.getElementById('batch-actions').style.display = 'none';
  resetBtn.style.display = 'none';

  const resultEl = createFileResult(file.name, fmtSize(file.size));
  document.getElementById('file-results').appendChild(resultEl);

  try {
    const { blob, wbFixed, sheetNames, unhiddenNames } = await unlockZip(zip, file.name, selectedPaths, resultEl);
    const outputName = file.name.replace(/(\.\w+)$/, '_UNLOCKED$1');
    finishResult(resultEl, blob, outputName, wbFixed, sheetNames, unhiddenNames, file);
    processedFiles.push({ name: outputName, blob });
  } catch (err) {
    resultEl.querySelector('.error-box').textContent = friendlyError(err);
    resultEl.querySelector('.error-box').style.display = 'block';
    resultEl.querySelector('.progress-bar').style.display = 'none';
  }

  resetBtn.style.display = 'block';
}

// ── Core: batch ───────────────────────────────────────────────────────────────
async function processBatch(files) {
  uploadCard.style.display = 'none';
  processedFiles = [];

  resultCard.style.display = 'block';
  document.getElementById('file-results').innerHTML = '';
  document.getElementById('batch-actions').style.display = 'none';
  resetBtn.style.display = 'none';

  for (const file of files) {
    const resultEl = createFileResult(file.name, fmtSize(file.size));
    document.getElementById('file-results').appendChild(resultEl);

    try {
      const buf = await file.arrayBuffer();
      if (!isValidZip(buf)) throw new Error('not a zip file');
      const zip = await JSZip.loadAsync(buf);
      const { blob, wbFixed, sheetNames, unhiddenNames } = await unlockZip(zip, file.name, null, resultEl);
      const outputName = file.name.replace(/(\.\w+)$/, '_UNLOCKED$1');
      finishResult(resultEl, blob, outputName, wbFixed, sheetNames, unhiddenNames);
      processedFiles.push({ name: outputName, blob });
    } catch (err) {
      resultEl.querySelector('.error-box').textContent = friendlyError(err);
      resultEl.querySelector('.error-box').style.display = 'block';
      resultEl.querySelector('.progress-bar').style.display = 'none';
      resultEl.querySelector('.success-row').style.display = 'flex';
    }
  }

  if (processedFiles.length > 1) {
    document.getElementById('batch-actions').style.display = 'flex';
  }

  resetBtn.style.display = 'block';
}

// ── Core: unlock ZIP ──────────────────────────────────────────────────────────
async function unlockZip(zip, filename, selectedPaths, resultEl) {
  const log          = resultEl.querySelector('.log');
  const progressFill = resultEl.querySelector('.progress-fill');

  function addLog(type, text) {
    const li = document.createElement('li');
    li.innerHTML = (type === 'loading' ? '<span class="dot-spin"></span>'
                  : type === 'ok'      ? '<span class="dot dot-green"></span>'
                  :                      '<span class="dot dot-gray"></span>') + esc(text);
    log.appendChild(li);
    return li;
  }

  const { nameMap: sheetNameMap, rIdToPath } = await getWorkbookMeta(zip);
  const scanLog = addLog('loading', 'Scanning for protection…');

  let wbFixed = false;
  const sheetNames = [];
  const unhiddenNames = [];
  const newZip = new JSZip();

  for (const [path, entry] of Object.entries(zip.files)) {
    if (entry.dir) { newZip.folder(path); continue; }

    let content = await entry.async('uint8array');

    if (path === 'xl/workbook.xml') {
      let text = new TextDecoder().decode(content);
      if (/<workbookProtection/i.test(text)) {
        text = text.replace(/<workbookProtection[^>]*\/?>/gi, '');
        wbFixed = true;
      }
      text = text.replace(/<sheet\b([^>]*?)(\/?>)/gi, (match, attrs, closing) => {
        if (!/\bstate="(?:hidden|veryHidden)"/i.test(attrs)) return match;
        const ridM = attrs.match(/\br:id="([^"]+)"/i);
        if (!ridM) return match;
        const sheetPath = rIdToPath[ridM[1]];
        const shouldUnhide = selectedPaths === null || (sheetPath && selectedPaths.has(sheetPath));
        if (!shouldUnhide) return match;
        const nameM = attrs.match(/\bname="([^"]+)"/i);
        unhiddenNames.push(nameM ? nameM[1] : (sheetPath || 'Unknown'));
        const newAttrs = attrs.replace(/\s*\bstate="(?:hidden|veryHidden)"/gi, '');
        return `<sheet${newAttrs}${closing}`;
      });
      content = new TextEncoder().encode(text);
    } else if (/^xl\/worksheets\/.+\.xml$/i.test(path)) {
      const shouldProcess = selectedPaths === null || selectedPaths.has(path);
      let text = new TextDecoder().decode(content);
      if (shouldProcess && /<sheetProtection/i.test(text)) {
        const name = sheetNameMap[path] || path.match(/sheet(\d+)\.xml$/i)?.[1] || path;
        text = text.replace(/<sheetProtection[^>]*\/?>/gi, '');
        content = new TextEncoder().encode(text);
        sheetNames.push(name);
      }
    }

    newZip.file(path, content, { binary: true });
  }

  scanLog.remove();

  if (wbFixed)                    addLog('ok',   'Workbook structure lock removed');
  sheetNames.forEach(name =>      addLog('ok',   `"${name}" unlocked`));
  unhiddenNames.forEach(name =>   addLog('ok',   `"${name}" unhidden`));
  if (!wbFixed && sheetNames.length === 0 && unhiddenNames.length === 0) addLog('info', 'No protection or hidden sheets found');

  const packLog = addLog('loading', 'Repacking file…');

  const mimeType = /\.xlsm$/i.test(filename)
    ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const blob = await newZip.generateAsync(
    { type: 'blob', mimeType, compression: 'DEFLATE' },
    meta => { progressFill.style.width = meta.percent.toFixed(0) + '%'; }
  );

  packLog.remove();
  addLog('ok', 'Ready to download');

  return { blob, wbFixed, sheetNames, unhiddenNames };
}

// ── UI helpers ────────────────────────────────────────────────────────────────
function finishResult(resultEl, blob, outputName, wbFixed, sheetNames, unhiddenNames = [], originalFile = null) {
  const total = sheetNames.length + unhiddenNames.length + (wbFixed ? 1 : 0);
  const badge = resultEl.querySelector('.result-badge');
  badge.textContent = total > 0
    ? `✓ ${total} protection${total > 1 ? 's' : ''} removed`
    : '⚠ No protection found';
  badge.className = 'result-badge' + (total === 0 ? ' warn' : '');

  const dlBtn = resultEl.querySelector('.btn-download');
  dlBtn.style.display = 'inline-flex';
  dlBtn.onclick = () => triggerDownload(blob, outputName);

  if (originalFile) {
    const revertBtn = resultEl.querySelector('.btn-revert');
    revertBtn.style.display = 'inline-flex';
    revertBtn.onclick = () => triggerDownload(originalFile, originalFile.name);
  }

  resultEl.querySelector('.progress-bar').style.display = 'none';
  resultEl.querySelector('.success-row').style.display = 'flex';
}

function createFileResult(name, size) {
  const div = document.createElement('div');
  div.className = 'file-result';
  div.innerHTML = `
    <div class="file-row">
      <div class="file-icon">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z" stroke="#1D9E75" stroke-width="1.5"/><path d="M14 2v6h6" stroke="#1D9E75" stroke-width="1.5"/><path d="M8 13h8M8 17h5" stroke="#1D9E75" stroke-width="1.5" stroke-linecap="round"/></svg>
      </div>
      <div class="file-meta">
        <strong>${esc(name)}</strong>
        <span>${size}</span>
      </div>
    </div>
    <div class="progress-bar"><div class="progress-fill"></div></div>
    <ul class="log"></ul>
    <div class="error-box" style="display:none"></div>
    <div class="success-row" style="display:none">
      <div class="result-badge"></div>
      <div class="dl-actions">
        <button class="btn-revert" style="display:none">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M3 12a9 9 0 1 0 9-9 9 9 0 0 0-6.364 2.636L3 8" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M3 3v5h5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
          Original
        </button>
        <button class="btn-download" style="display:none">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M12 3v13M7 11l5 5 5-5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M5 20h14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Download
        </button>
      </div>
    </div>`;
  return div;
}

function triggerDownload(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

async function downloadAll() {
  if (!processedFiles.length) return;
  const zip = new JSZip();
  processedFiles.forEach(f => zip.file(f.name, f.blob));
  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
  triggerDownload(blob, 'unlocked_files.zip');
}

function reset() {
  processedFiles    = [];
  pendingFile       = null;
  pendingZip        = null;
  pendingSheets     = [];
  pendingWbProtected = false;

  uploadCard.style.display = 'block';
  selectCard.style.display = 'none';
  resultCard.style.display = 'none';
  resetBtn.style.display   = 'none';
  fileInput.value          = '';
  dropZone.classList.remove('active');
}
