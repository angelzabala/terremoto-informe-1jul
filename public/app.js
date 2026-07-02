const REQUEST_DELAY_MS = 2500;

const fileInput = document.getElementById('file-input');
const dropzone = document.getElementById('dropzone');
const fileLabel = document.getElementById('file-label');
const processBtn = document.getElementById('process-btn');
const downloadBtn = document.getElementById('download-btn');
const resetBtn = document.getElementById('reset-btn');
const progressWrap = document.getElementById('progress-wrap');
const progressText = document.getElementById('progress-text');
const progressPercent = document.getElementById('progress-percent');
const progressFill = document.getElementById('progress-fill');
const logBody = document.getElementById('log-body');
const statsEl = document.getElementById('stats');

let selectedFile = null;
let workbook = null;
let outputWorkbook = null;
let outputFileName = '';
let processing = false;
let stats = { ok: 0, warn: 0, err: 0 };

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function resetStats() {
  stats = { ok: 0, warn: 0, err: 0 };
  updateStats();
}

function updateStats() {
  statsEl.innerHTML = `
    <span class="stat ok">${stats.ok} encontrados</span>
    <span class="stat warn">${stats.warn} no encontrados</span>
    <span class="stat err">${stats.err} errores</span>
  `;
}

function setProgress(current, total) {
  const percent = total === 0 ? 0 : Math.round((current / total) * 100);
  progressText.textContent = `${current} / ${total}`;
  progressPercent.textContent = `${percent}%`;
  progressFill.style.width = `${percent}%`;
}

function statusClass(status) {
  if (status === 'ENCONTRADO') return 'ok';
  if (status === 'NO ENCONTRADO') return 'warn';
  return 'err';
}

function appendLog(rowNumber, cedula, result) {
  const tr = document.createElement('tr');
  const displayStatus = result.status === 'ERROR' && result.error
    ? `ERROR: ${result.error}`
    : result.status;

  const name = result.status === 'ENCONTRADO'
    ? `${result.primerApellido} ${result.segundoApellido}, ${result.nombres}`.trim()
    : '—';

  tr.innerHTML = `
    <td>${rowNumber}</td>
    <td>${cedula}</td>
    <td><span class="status-pill ${statusClass(result.status)}">${displayStatus}</span></td>
    <td>${name}</td>
  `;

  if (logBody.querySelector('.empty-row')) {
    logBody.innerHTML = '';
  }

  logBody.prepend(tr);

  if (result.status === 'ENCONTRADO') stats.ok += 1;
  else if (result.status === 'NO ENCONTRADO') stats.warn += 1;
  else stats.err += 1;

  updateStats();
}

function readWorkbookFromFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        resolve(XLSX.read(data, { type: 'array', cellDates: true }));
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function handleFile(file) {
  if (!file) return;

  const lower = file.name.toLowerCase();
  if (!lower.endsWith('.xlsx') && !lower.endsWith('.xls')) {
    alert('Selecciona un archivo Excel (.xlsx o .xls).');
    return;
  }

  selectedFile = file;
  outputWorkbook = null;
  outputFileName = `procesado-${file.name.replace(/\.xls$/i, '.xlsx')}`;
  fileLabel.textContent = file.name;
  processBtn.disabled = false;
  downloadBtn.disabled = true;
  resetBtn.disabled = false;
  progressWrap.hidden = true;
  logBody.innerHTML = '<tr class="empty-row"><td colspan="4">Listo para procesar.</td></tr>';
  resetStats();
}

async function consultCedula(cedula) {
  const response = await fetch('/api/consultar', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ cedula }),
  });

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error ?? 'Error en la consulta');
  }

  return payload;
}

async function processWorkbook() {
  if (!selectedFile || processing) return;

  processing = true;
  processBtn.disabled = true;
  downloadBtn.disabled = true;
  progressWrap.hidden = false;
  logBody.innerHTML = '';
  resetStats();

  try {
    workbook = await readWorkbookFromFile(selectedFile);
    outputWorkbook = XLSX.utils.book_new();

    const sheetStates = workbook.SheetNames.map((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        raw: true,
      });
      return { sheetName, ...ExcelUtils.prepareSheet(rows) };
    });

    const tasks = sheetStates.flatMap((sheet) =>
      sheet.tasks.map((task) => ({ ...task, sheet })),
    );

    setProgress(0, tasks.length);

    for (let index = 0; index < tasks.length; index += 1) {
      const task = tasks[index];
      const result = await consultCedula(task.cedula);

      ExcelUtils.applyConsultResult(
        task.sheet.rows,
        task.sheet.startCol,
        task.rowIndex,
        result,
      );

      appendLog(task.rowIndex + 1, task.cedula, result);
      setProgress(index + 1, tasks.length);

      if (index < tasks.length - 1) {
        await sleep(REQUEST_DELAY_MS);
      }
    }

    for (const sheet of sheetStates) {
      const worksheet = XLSX.utils.aoa_to_sheet(sheet.rows);
      XLSX.utils.book_append_sheet(outputWorkbook, worksheet, sheet.sheetName);
    }

    downloadBtn.disabled = false;
  } catch (error) {
    alert(error.message ?? 'Ocurrió un error procesando el archivo.');
  } finally {
    processing = false;
    processBtn.disabled = !selectedFile;
  }
}

function downloadResult() {
  if (!outputWorkbook) return;

  const buffer = XLSX.write(outputWorkbook, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = outputFileName;
  link.click();
  URL.revokeObjectURL(url);
}

function resetAll() {
  selectedFile = null;
  workbook = null;
  outputWorkbook = null;
  fileInput.value = '';
  fileLabel.textContent = 'Arrastra un .xlsx aquí o haz clic para seleccionar';
  processBtn.disabled = true;
  downloadBtn.disabled = true;
  resetBtn.disabled = true;
  progressWrap.hidden = true;
  setProgress(0, 0);
  logBody.innerHTML = '<tr class="empty-row"><td colspan="4">Aún no hay consultas.</td></tr>';
  resetStats();
}

fileInput.addEventListener('change', (event) => {
  handleFile(event.target.files[0]);
});

dropzone.addEventListener('dragover', (event) => {
  event.preventDefault();
  dropzone.classList.add('dragover');
});

dropzone.addEventListener('dragleave', () => {
  dropzone.classList.remove('dragover');
});

dropzone.addEventListener('drop', (event) => {
  event.preventDefault();
  dropzone.classList.remove('dragover');
  handleFile(event.dataTransfer.files[0]);
});

processBtn.addEventListener('click', processWorkbook);
downloadBtn.addEventListener('click', downloadResult);
resetBtn.addEventListener('click', resetAll);
