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
let activeFilters = { ok: true, warn: true, err: true };

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function resetFilters() {
  activeFilters = { ok: true, warn: true, err: true };
  statsEl.querySelectorAll('.stat-filter').forEach((button) => {
    button.classList.add('active');
    button.setAttribute('aria-pressed', 'true');
  });
}

function resetStats() {
  stats = { ok: 0, warn: 0, err: 0 };
  updateStats();
}

function updateStats() {
  statsEl.querySelector('[data-filter="ok"]').textContent = `${stats.ok} CIV verificadas`;
  statsEl.querySelector('[data-filter="warn"]').textContent = `${stats.warn} CIV no verificadas`;
  statsEl.querySelector('[data-filter="err"]').textContent = `${stats.err} errores`;
}

function filterCategory(status) {
  if (status === 'ENCONTRADO') return 'ok';
  if (status === 'NO ENCONTRADO') return 'warn';
  return 'err';
}

function statusClass(status) {
  return filterCategory(status);
}

function displayStatus(result) {
  if (result.status === 'ERROR' && result.error) {
    return `ERROR: ${result.error}`;
  }
  return ExcelUtils.formatEstadoCiv(result.status, result.error);
}

function applyRowFilters() {
  const rows = logBody.querySelectorAll('tr[data-filter]');
  let visibleCount = 0;

  rows.forEach((row) => {
    const visible = activeFilters[row.dataset.filter];
    row.hidden = !visible;
    if (visible) visibleCount += 1;
  });

  const filterEmptyRow = logBody.querySelector('.filter-empty-row');
  if (filterEmptyRow) filterEmptyRow.remove();

  if (rows.length > 0 && visibleCount === 0) {
    const tr = document.createElement('tr');
    tr.className = 'empty-row filter-empty-row';
    tr.innerHTML = '<td colspan="4">Ninguna categoría seleccionada. Activa al menos un filtro.</td>';
    logBody.appendChild(tr);
  }
}

function setProgress(current, total) {
  const percent = total === 0 ? 0 : Math.round((current / total) * 100);
  progressText.textContent = `${current} / ${total}`;
  progressPercent.textContent = `${percent}%`;
  progressFill.style.width = `${percent}%`;
}

function appendLog(rowNumber, cedula, result) {
  const tr = document.createElement('tr');
  const category = filterCategory(result.status);

  const name = result.status === 'ENCONTRADO'
    ? `${result.primerApellido} ${result.segundoApellido}, ${result.nombres}`.trim()
    : '—';

  tr.dataset.filter = category;
  tr.innerHTML = `
    <td>${rowNumber}</td>
    <td>${cedula}</td>
    <td><span class="status-pill ${statusClass(result.status)}">${displayStatus(result)}</span></td>
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
  applyRowFilters();
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
  resetFilters();
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
  resetFilters();

  try {
    workbook = await readWorkbookFromFile(selectedFile);

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

    const splitSheets = ExcelUtils.buildSplitSheets(sheetStates);

    outputWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      outputWorkbook,
      XLSX.utils.aoa_to_sheet(splitSheets.verifiedRows),
      splitSheets.verifiedSheetName,
    );
    XLSX.utils.book_append_sheet(
      outputWorkbook,
      XLSX.utils.aoa_to_sheet(splitSheets.pendingRows),
      splitSheets.pendingSheetName,
    );

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
  resetFilters();
}

function toggleFilter(filterKey) {
  activeFilters[filterKey] = !activeFilters[filterKey];

  const button = statsEl.querySelector(`[data-filter="${filterKey}"]`);
  button.classList.toggle('active', activeFilters[filterKey]);
  button.setAttribute('aria-pressed', String(activeFilters[filterKey]));

  applyRowFilters();
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

statsEl.addEventListener('click', (event) => {
  const button = event.target.closest('.stat-filter');
  if (!button) return;
  toggleFilter(button.dataset.filter);
});

processBtn.addEventListener('click', processWorkbook);
downloadBtn.addEventListener('click', downloadResult);
resetBtn.addEventListener('click', resetAll);
