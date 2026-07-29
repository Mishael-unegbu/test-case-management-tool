// Excel-backed storage service.
// Reads/writes the shared workbook defined in config.EXCEL_FILE_PATH.
// Only touches the Projects and AuditLog sheets — per Master AI Coding
// Instructions, this story must not modify unrelated modules/sheets.
//
// Storage rules honored here:
// - Sheet names and column names are never changed.
// - AuditLog is append-only (we only ever add rows, never edit/delete).
// - Existing Project rows are never overwritten by a create operation.

const ExcelJS = require('exceljs');
const fs = require('fs');
const config = require('../config');

const PROJECTS_SHEET = 'Projects';
const AUDIT_LOG_SHEET = 'AuditLog';

const PROJECTS_COLUMNS = [
  'ProjectID',
  'ProjectName',
  'Description',
  'Status',
  'StartDate',
  'EndDate',
  'CreatedDate',
  'UpdatedDate',
];

const AUDIT_LOG_COLUMNS = [
  'AuditID',
  'EntityType',
  'EntityID',
  'Action',
  'ChangedField',
  'OldValue',
  'NewValue',
  'ChangeDate',
];

function assertWorkbookExists() {
  if (!fs.existsSync(config.EXCEL_FILE_PATH)) {
    throw new Error(
      `Excel workbook not found at ${config.EXCEL_FILE_PATH}. ` +
        'Set EXCEL_FILE_PATH env var if your SharePoint sync location differs.'
    );
  }
}

async function loadWorkbook() {
  assertWorkbookExists();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(config.EXCEL_FILE_PATH);
  return workbook;
}

async function saveWorkbook(workbook) {
  await workbook.xlsx.writeFile(config.EXCEL_FILE_PATH);
}

function getSheetOrThrow(workbook, sheetName) {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in workbook. Storage structure may have changed.`);
  }
  return sheet;
}

// Reads all data rows (excluding header) from a sheet as plain objects,
// keyed by the header names in row 1. Skips fully-empty rows.
function readRowsAsObjects(sheet, columns) {
  const rows = [];
  const totalRows = sheet.rowCount;
  for (let r = 2; r <= totalRows; r++) {
    const row = sheet.getRow(r);
    const isEmpty = columns.every((_, idx) => {
      const cell = row.getCell(idx + 1);
      return cell.value === null || cell.value === undefined || cell.value === '';
    });
    if (isEmpty) continue;

    const obj = {};
    columns.forEach((colName, idx) => {
      obj[colName] = row.getCell(idx + 1).value;
    });
    obj.__rowNumber = r;
    rows.push(obj);
  }
  return rows;
}

function appendRow(sheet, columns, dataObj) {
  const values = columns.map((col) => dataObj[col] ?? null);
  sheet.addRow(values);
}

// Computes the next simple incrementing numeric ProjectID (1, 2, 3, ...).
function computeNextProjectId(existingProjects) {
  let maxId = 0;
  for (const p of existingProjects) {
    const idNum = Number(p.ProjectID);
    if (!Number.isNaN(idNum) && idNum > maxId) {
      maxId = idNum;
    }
  }
  return maxId + 1;
}

async function appendAuditLogEntry(workbook, { entityType, entityId, action, changedField, oldValue, newValue }) {
  const auditSheet = getSheetOrThrow(workbook, AUDIT_LOG_SHEET);
  const existingAudits = readRowsAsObjects(auditSheet, AUDIT_LOG_COLUMNS);

  let maxAuditId = 0;
  for (const a of existingAudits) {
    const idNum = Number(a.AuditID);
    if (!Number.isNaN(idNum) && idNum > maxAuditId) maxAuditId = idNum;
  }

  appendRow(auditSheet, AUDIT_LOG_COLUMNS, {
    AuditID: maxAuditId + 1,
    EntityType: entityType,
    EntityID: entityId,
    Action: action,
    ChangedField: changedField ?? '',
    OldValue: oldValue ?? '',
    NewValue: newValue ?? '',
    ChangeDate: new Date().toISOString(),
  });
}

/**
 * US-001 Create Project
 * Validates input, assigns the next simple incrementing ProjectID,
 * appends a new row to the Projects sheet, and logs the creation
 * to the append-only AuditLog sheet.
 */
async function createProject({ projectName, description, status, startDate, endDate }) {
  if (!projectName || !String(projectName).trim()) {
    const err = new Error('ProjectName is required.');
    err.statusCode = 400;
    throw err;
  }

  const workbook = await loadWorkbook();
  const projectsSheet = getSheetOrThrow(workbook, PROJECTS_SHEET);
  const existingProjects = readRowsAsObjects(projectsSheet, PROJECTS_COLUMNS);

  const duplicate = existingProjects.find(
    (p) => String(p.ProjectName).trim().toLowerCase() === String(projectName).trim().toLowerCase()
  );
  if (duplicate) {
    const err = new Error(`A project named "${projectName}" already exists (ProjectID ${duplicate.ProjectID}).`);
    err.statusCode = 409;
    throw err;
  }

  const projectId = computeNextProjectId(existingProjects);
  const now = new Date().toISOString();

  const newProject = {
    ProjectID: projectId,
    ProjectName: String(projectName).trim(),
    Description: description ?? '',
    Status: status && String(status).trim() ? status : 'Active',
    StartDate: startDate ?? '',
    EndDate: endDate ?? '',
    CreatedDate: now,
    UpdatedDate: now,
  };

  appendRow(projectsSheet, PROJECTS_COLUMNS, newProject);

  await appendAuditLogEntry(workbook, {
    entityType: 'Project',
    entityId: projectId,
    action: 'Create',
    changedField: '',
    oldValue: '',
    newValue: newProject.ProjectName,
  });

  await saveWorkbook(workbook);

  delete newProject.__rowNumber;
  return newProject;
}

module.exports = {
  createProject,
  // exported for future stories (US-002 Edit, US-003 View) to reuse safely
  loadWorkbook,
  saveWorkbook,
  getSheetOrThrow,
  readRowsAsObjects,
  PROJECTS_SHEET,
  PROJECTS_COLUMNS,
};
