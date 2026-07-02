#!/usr/bin/env node
/**
 * CLI: consulta cédulas en archivos Excel del directorio actual.
 * Para la UI web: npm run ui
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { consultCedula, REQUEST_DELAY_MS } = require('./lib/cedula-service');
const {
  prepareSheet,
  applyConsultResult,
  buildSplitSheets,
} = require('./lib/excel-utils');

const ROOT_DIR = __dirname;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function listExcelFiles() {
  return fs
    .readdirSync(ROOT_DIR)
    .filter((fileName) => {
      const lower = fileName.toLowerCase();
      return lower.endsWith('.xlsx') && !lower.startsWith('procesado-') && !fileName.startsWith('~$');
    })
    .map((fileName) => path.join(ROOT_DIR, fileName));
}

async function processSheetState(sheetState, sheetName) {
  if (sheetState.headerRowIndex === -1 || sheetState.cedulaColIndex === -1) {
    console.warn(`  Hoja "${sheetName}": no se encontró columna de cédula, se omite.`);
    return sheetState;
  }

  console.log(
    `  Hoja "${sheetName}": encabezado fila ${sheetState.headerRowIndex + 1}, columna cédula ${sheetState.cedulaColIndex + 1}`,
  );

  for (const task of sheetState.tasks) {
    console.log(`    Fila ${task.rowIndex + 1}: consultando cédula ${task.cedula}...`);
    const result = await consultCedula(task.cedula);

    applyConsultResult(sheetState.rows, sheetState.startCol, task.rowIndex, result);

    if (result.status === 'ENCONTRADO') {
      console.log(`      OK: ${result.primerApellido} ${result.segundoApellido}, ${result.nombres}`);
    } else if (result.status === 'NO ENCONTRADO') {
      console.log('      Sin resultados');
    } else {
      console.log(`      Error: ${result.error ?? 'desconocido'}`);
    }

    await sleep(REQUEST_DELAY_MS);
  }

  return sheetState;
}

async function processWorkbook(filePath) {
  const sourceWorkbook = XLSX.readFile(filePath);
  const sheetStates = [];

  for (const sheetName of sourceWorkbook.SheetNames) {
    const worksheet = sourceWorkbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: '',
      raw: true,
    });

    const sheetState = { sheetName, ...prepareSheet(rows) };
    sheetStates.push(await processSheetState(sheetState, sheetName));
  }

  const splitSheets = buildSplitSheets(sheetStates);
  const outputWorkbook = XLSX.utils.book_new();

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

  return outputWorkbook;
}

async function main() {
  const files = listExcelFiles();

  if (files.length === 0) {
    console.log('No hay archivos .xlsx para procesar en:', ROOT_DIR);
    return;
  }

  console.log(`Archivos a procesar: ${files.length}`);

  for (const filePath of files) {
    const baseName = path.basename(filePath);
    const outputPath = path.join(ROOT_DIR, `procesado-${baseName}`);

    console.log(`\nProcesando: ${baseName}`);

    try {
      const workbook = await processWorkbook(filePath);
      XLSX.writeFile(workbook, outputPath);
      console.log(`Guardado: ${path.basename(outputPath)}`);
    } catch (error) {
      console.error(`Error procesando ${baseName}: ${error.message}`);
    }
  }

  console.log('\nFinalizado.');
}

main().catch((error) => {
  console.error(error.message ?? error);
  process.exit(1);
});
