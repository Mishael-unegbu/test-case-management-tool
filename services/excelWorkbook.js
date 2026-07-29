/**
 * Low-level Excel I/O primitives, shared by every resource-specific service
 * (projectsService.js, and future userStoriesService.js, testCasesService.js,
 * bugsService.js, etc. as those stories are approved and built).
 *
 * This module is intentionally "dumb": it knows how to open/save the workbook,
 * find a worksheet, map header names (row 1) to column numbers, and read/write
 * rows by header name. It holds NO business logic — no validation rules, no
 * audit logging, no per-resource ID policy. That keeps it reusable across every
 * sheet in the TRD without becoming a generic multi-sheet CRUD engine ahead of
 * approval (see Master AI Coding Instructions: "Do not introduce new
 * architecture without explicit approval").
 *
 * Columns are always looked up by header name, never by fixed index, so a
 * resource service never silently breaks if a column is reordered (renaming
 * columns is still forbidden per governance, but reordering happens).
 */

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const config = require('../config');

/**
 * Serializes load->mutate->save cycles against the same workbook file.
 * Every request handler does: load workbook, mutate in memory, save workbook.
 * Without this, two concurrent requests (e.g. two Create Project calls) can
 * both load before either saves, both compute the same "next ID", and the
 * second save silently overwrites the first's changes. This does not add
 * multi-user auth or any new architecture (still a single Node process,
 * still Excel-backed) — it just makes the existing per-request read-modify-
 * write cycle atomic with respect to other in-process requests.
 */
const fileLocks = new Map();

function withWorkbookLock(filePath, task) {
  const key = path.resolve(filePath || config.EXCEL_FILE_PATH);
  const previous = fileLocks.get(key) || Promise.resolve();
  const run = previous.then(task, task);
  fileLocks.set(
    key,
    run.then(
      () => undefined,
      () => undefined
    )
  );
  return run;
}

function assertWorkbookExists(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(
      `Excel workbook not found at ${filePath}. ` +
        'Set EXCEL_FILE_PATH env var if your SharePoint sync location differs.'
    );
  }
}

async function loadWorkbook(filePath = config.EXCEL_FILE_PATH) {
  assertWorkbookExists(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  return workbook;
}

async function saveWorkbook(workbook, filePath = config.EXCEL_FILE_PATH) {
  await workbook.xlsx.writeFile(filePath);
}

function getWorksheetOrThrow(workbook, sheetName) {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`Sheet "${sheetName}" was not found in the workbook. Storage structure may have changed.`);
  }
  return worksheet;
}

/** Maps header text (row 1) -> column number. Never hardcode column order. */
function getHeaderMap(worksheet) {
  const headerRow = worksheet.getRow(1);
  const map = {};
  headerRow.eachCell((cell, colNumber) => {
    const key = cell.value === null || cell.value === undefined ? '' : String(cell.value).trim();
    if (key) map[key] = colNumber;
  });
  return map;
}

function rowToObject(row, headerMap) {
  const obj = {};
  Object.entries(headerMap).forEach(([columnName, colIndex]) => {
    const cellValue = row.getCell(colIndex).value;
    obj[columnName] = cellValue === undefined ? null : cellValue;
  });
  return obj;
}

function findRowByColumnValue(worksheet, headerMap, columnName, value) {
  const colIndex = headerMap[columnName];
  if (!colIndex) return null;

  let match = null;
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // skip header
    const cellValue = row.getCell(colIndex).value;
    if (cellValue !== null && cellValue !== undefined && String(cellValue).trim() === String(value).trim()) {
      match = row;
    }
  });
  return match;
}

/** Reads all non-empty data rows (excludes header) as plain objects keyed by header name. */
function readAllRows(worksheet, headerMap) {
  const rows = [];
  const columnNames = Object.keys(headerMap);
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const isEmpty = columnNames.every((name) => {
      const v = row.getCell(headerMap[name]).value;
      return v === null || v === undefined || v === '';
    });
    if (isEmpty) return;
    rows.push(rowToObject(row, headerMap));
  });
  return rows;
}

function appendRowByHeader(worksheet, headerMap, dataObj) {
  const newRow = worksheet.addRow([]);
  Object.entries(headerMap).forEach(([columnName, colIndex]) => {
    if (Object.prototype.hasOwnProperty.call(dataObj, columnName)) {
      newRow.getCell(colIndex).value = dataObj[columnName] ?? null;
    }
  });
  newRow.commit();
  return newRow;
}

/** Computes the next simple incrementing numeric ID (1, 2, 3, ...) for a given ID column. */
function nextNumericId(existingRows, idColumnName) {
  let maxId = 0;
  for (const row of existingRows) {
    const idNum = Number(row[idColumnName]);
    if (!Number.isNaN(idNum) && idNum > maxId) maxId = idNum;
  }
  return maxId + 1;
}

module.exports = {
  loadWorkbook,
  saveWorkbook,
  getWorksheetOrThrow,
  getHeaderMap,
  rowToObject,
  findRowByColumnValue,
  readAllRows,
  appendRowByHeader,
  nextNumericId,
  withWorkbookLock,
};
