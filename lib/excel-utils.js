const OUTPUT_COLUMNS = [
  'RIF',
  'Primer Apellido',
  'Segundo Apellido',
  'Nombres',
  'Consulta Cedula',
];

const CEDULA_ALIASES = [
  'cedula',
  'civ',
  'id',
  'identificacion',
  'documento',
  'documento de identidad',
  'ci',
  'dni',
  'numero de identidad',
  'no identificacion',
  'n identificacion',
];

const HEADER_SCAN_ROWS = 30;

function normalizeHeader(value) {
  return String(value ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function isCedulaHeader(value) {
  const normalized = normalizeHeader(value);
  if (!normalized) return false;
  if (CEDULA_ALIASES.includes(normalized)) return true;
  if (normalized.includes('cedula')) return true;
  if (normalized.includes('identificacion')) return true;
  if (normalized === 'civ') return true;
  if (normalized === 'id') return true;
  return false;
}

function normalizeCedula(raw) {
  if (raw == null || raw === '') return null;

  const value = String(raw).trim();
  if (!value) return null;

  const parts = value.split(/\s*[,;|/]\s*/);
  if (parts.length > 1) {
    const normalizedParts = parts
      .map((part) => part.replace(/\D/g, ''))
      .filter((part) => part.length >= 6);

    if (normalizedParts.length > 0) {
      return normalizedParts[0];
    }
  }

  const digits = value.replace(/\D/g, '');
  return digits || null;
}

function findHeaderRow(rows) {
  const limit = Math.min(rows.length, HEADER_SCAN_ROWS);

  for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    if (row.some((cell) => isCedulaHeader(cell))) {
      return rowIndex;
    }
  }

  return -1;
}

function findCedulaColumnIndex(headerRow) {
  for (let colIndex = 0; colIndex < headerRow.length; colIndex += 1) {
    if (isCedulaHeader(headerRow[colIndex])) {
      return colIndex;
    }
  }
  return -1;
}

function ensureRowLength(row, length) {
  while (row.length < length) {
    row.push('');
  }
  return row;
}

function prepareSheet(rows) {
  const headerRowIndex = findHeaderRow(rows);
  if (headerRowIndex === -1) {
    return { rows, headerRowIndex: -1, cedulaColIndex: -1, startCol: -1, tasks: [] };
  }

  const headerRow = ensureRowLength([...(rows[headerRowIndex] ?? [])], 1);
  const cedulaColIndex = findCedulaColumnIndex(headerRow);
  if (cedulaColIndex === -1) {
    return { rows, headerRowIndex, cedulaColIndex: -1, startCol: -1, tasks: [] };
  }

  const startCol = headerRow.length;
  OUTPUT_COLUMNS.forEach((columnName, offset) => {
    headerRow[startCol + offset] = columnName;
  });
  rows[headerRowIndex] = headerRow;

  const tasks = [];
  for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = ensureRowLength([...(rows[rowIndex] ?? [])], startCol + OUTPUT_COLUMNS.length);
    rows[rowIndex] = row;

    const cedula = normalizeCedula(row[cedulaColIndex]);
    if (cedula) {
      tasks.push({ rowIndex, cedula });
    }
  }

  return { rows, headerRowIndex, cedulaColIndex, startCol, tasks };
}

function applyConsultResult(rows, startCol, rowIndex, result) {
  const row = rows[rowIndex];
  row[startCol] = result.rif ?? '';
  row[startCol + 1] = result.primerApellido ?? '';
  row[startCol + 2] = result.segundoApellido ?? '';
  row[startCol + 3] = result.nombres ?? '';
  row[startCol + 4] = result.status === 'ERROR' && result.error
    ? `ERROR: ${result.error}`
    : result.status;
}

const api = {
  OUTPUT_COLUMNS,
  normalizeCedula,
  findHeaderRow,
  findCedulaColumnIndex,
  prepareSheet,
  applyConsultResult,
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = api;
}

if (typeof window !== 'undefined') {
  window.ExcelUtils = api;
}
