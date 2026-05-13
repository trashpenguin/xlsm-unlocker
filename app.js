let outputBlob = null;
let outputName = '';

const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const uploadCard = document.getElementById('upload-card');
const resultCard = document.getElementById('result-card');
const resetBtn = document.getElementById('reset-btn');

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('active'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('active'));
dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('active'); handleFile(e.dataTransfer.files[0]); });
dropZone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') fileInput.click(); });
fileInput.addEventListener('change', () => handleFile(fileInput.files[0]));

function fmtSize(b) {
  return b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.round(b / 1024) + ' KB';
}

function addLog(type, text) {
  const li = document.createElement('li');
  let dot;
  if (type === 'loading') {
    dot = '<span class="dot-spin"></span>';
  } else if (type === 'ok') {
    dot = '<span class="dot dot-green"></span>';
  } else {
    dot = '<span class="dot dot-gray"></span>';
  }
  li.innerHTML = dot + text;
  li.id = 'log-' + Date.now();
  document.getElementById('log').appendChild(li);
  return li;
}

async function handleFile(file) {
  if (!file) return;
  if (!/\.(xlsm|xlsx)$/i.test(file.name)) {
    alert('Please upload an .xlsm or .xlsx file.');
    return;
  }

  uploadCard.style.display = 'none';
  resultCard.style.display = 'block';
  resetBtn.style.display = 'none';

  document.getElementById('fname').textContent = file.name;
  document.getElementById('fsize').textContent = fmtSize(file.size);
  document.getElementById('log').innerHTML = '';
  document.getElementById('error-box').style.display = 'none';
  document.getElementById('success-row').style.display = 'none';
  outputBlob = null;

  const readingLog = addLog('loading', 'Reading file…');

  try {
    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);

    readingLog.querySelector('.dot-spin').className = 'dot dot-green';
    readingLog.childNodes[1].textContent = ' File opened successfully';

    addLog('loading', 'Scanning for protection…');

    let sheetCount = 0;
    let wbFixed = false;
    const newZip = new JSZip();

    for (const [path, entry] of Object.entries(zip.files)) {
      if (entry.dir) { newZip.folder(path); continue; }

      let content = await entry.async('uint8array');

      if (path === 'xl/workbook.xml') {
        let text = new TextDecoder().decode(content);
        if (/<workbookProtection/i.test(text)) {
          text = text.replace(/<workbookProtection[^>]*\/?>/gi, '');
          content = new TextEncoder().encode(text);
          wbFixed = true;
        }
      } else if (/^xl\/worksheets\/.+\.xml$/i.test(path)) {
        let text = new TextDecoder().decode(content);
        if (/<sheetProtection/i.test(text)) {
          text = text.replace(/<sheetProtection[^>]*\/?>/gi, '');
          content = new TextEncoder().encode(text);
          sheetCount++;
        }
      }

      newZip.file(path, content, { binary: true });
    }

    document.getElementById('log').lastChild.remove();

    if (wbFixed) addLog('ok', 'Workbook structure lock removed');
    else addLog('info', 'No workbook-level lock found');

    if (sheetCount > 0) addLog('ok', `Sheet protection removed from ${sheetCount} sheet${sheetCount > 1 ? 's' : ''}`);
    else addLog('info', 'No sheet-level protection found');

    const mimeType = /\.xlsm$/i.test(file.name)
      ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    outputBlob = await newZip.generateAsync({ type: 'blob', mimeType, compression: 'DEFLATE' });
    outputName = file.name.replace(/(\.\w+)$/, '_UNLOCKED$1');

    addLog('ok', 'Done — ready to download');

    const total = sheetCount + (wbFixed ? 1 : 0);
    const badge = document.getElementById('result-badge');
    if (total > 0) {
      badge.textContent = `✓ ${total} protection${total > 1 ? 's' : ''} removed`;
      badge.className = 'result-badge';
    } else {
      badge.textContent = '⚠ No protection found';
      badge.className = 'result-badge warn';
    }

    document.getElementById('success-row').style.display = 'flex';
    resetBtn.style.display = 'block';

  } catch (err) {
    const box = document.getElementById('error-box');
    box.textContent = 'Error: ' + err.message;
    box.style.display = 'block';
    addLog('info', 'Processing failed');
    resetBtn.style.display = 'block';
  }
}

function downloadFile() {
  if (!outputBlob) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(outputBlob);
  a.download = outputName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

function reset() {
  outputBlob = null;
  uploadCard.style.display = 'block';
  resultCard.style.display = 'none';
  resetBtn.style.display = 'none';
  fileInput.value = '';
  document.getElementById('log').innerHTML = '';
  dropZone.classList.remove('active');
}
