/**
 * Projects sheet service.
 * Covers US-001 (Create Project), US-002 (Edit Project), and the read used
 * by US-003 (View Project) — see getProjectById() below and the controller
 * doc comment for how it's wired to US-003.
 *
 * Governance notes:
 * - Sheet/column names are never renamed.
 * - Column access is always by header name (see excelWorkbook.js), never
 *   fixed index.
 * - ProjectID and CreatedDate are immutable once set; UpdatedDate is
 *   system-managed and never accepted from the caller.
 */

const {
  loadWorkbook,
  saveWorkbook,
  getWorksheetOrThrow,
  getHeaderMap,
  readAllRows,
  findRowByColumnValue,
  rowToObject,
  appendRowByHeader,
  nextNumericId,
  withWorkbookLock,
} = require('./excelWorkbook');
const { appendAuditLogEntry } = require('./auditLogService');

const PROJECTS_SHEET = 'Projects';
const SETTINGS_SHEET = 'Settings';

// Only these columns may be changed by an edit.
const EDITABLE_PROJECT_COLUMNS = ['ProjectName', 'Description', 'Status', 'StartDate', 'EndDate'];

/**
 * US-001 Create Project.
 * Assigns the next simple incrementing ProjectID (1, 2, 3, ...), rejects
 * duplicate project names (case-insensitive), appends the row, and logs
 * the creation to AuditLog.
 */
async function createProject({ projectName, description, status, startDate, endDate }, filePath) {
  return withWorkbookLock(filePath, async () => {
    const workbook = await loadWorkbook(filePath);
    const sheet = getWorksheetOrThrow(workbook, PROJECTS_SHEET);
    const headerMap = getHeaderMap(sheet);
    const existingProjects = readAllRows(sheet, headerMap);

    const duplicate = existingProjects.find(
      (p) => String(p.ProjectName).trim().toLowerCase() === String(projectName).trim().toLowerCase()
    );
    if (duplicate) {
      const err = new Error(`A project named "${projectName}" already exists (ProjectID ${duplicate.ProjectID}).`);
      err.statusCode = 409;
      throw err;
    }

    const projectId = nextNumericId(existingProjects, 'ProjectID');
    const now = new Date().toISOString();

    // status is validated against the allowed list by the validator before this
    // is ever called (see validators/projects.validator.js), so by the time we
    // get here it's already a known-good value or absent.
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

    appendRowByHeader(sheet, headerMap, newProject);

    appendAuditLogEntry(workbook, {
      entityType: 'Project',
      entityId: projectId,
      action: 'Create',
      newValue: newProject.ProjectName,
    });

    await saveWorkbook(workbook, filePath);
    return newProject;
  });
}

/**
 * Fetches a single project by ProjectID.
 * Used by: US-003 View Project (via the controller's viewProject/getProject
 * handler), and to pre-populate the Edit Project (US-002) form.
 * Returns null if no matching project exists.
 */
async function getProjectById(projectId, filePath) {
  const workbook = await loadWorkbook(filePath);
  const sheet = getWorksheetOrThrow(workbook, PROJECTS_SHEET);
  const headerMap = getHeaderMap(sheet);

  const row = findRowByColumnValue(sheet, headerMap, 'ProjectID', projectId);
  if (!row) return null;

  return rowToObject(row, headerMap);
}

/**
 * US-002 Edit Project.
 * Updates only the editable fields present in `updates` for the matching
 * ProjectID, sets UpdatedDate, persists the workbook, and logs the change
 * to AuditLog. Returns the updated project, or null if no match exists.
 */
async function updateProject(projectId, updates, filePath) {
  return withWorkbookLock(filePath, async () => {
    const workbook = await loadWorkbook(filePath);
    const sheet = getWorksheetOrThrow(workbook, PROJECTS_SHEET);
    const headerMap = getHeaderMap(sheet);

    const row = findRowByColumnValue(sheet, headerMap, 'ProjectID', projectId);
    if (!row) return null;

    const before = rowToObject(row, headerMap);

    // Same duplicate-name rule as Create Project (US-001): a rename can't
    // collide with another existing project's name, case-insensitive.
    if (Object.prototype.hasOwnProperty.call(updates, 'ProjectName')) {
      const newName = String(updates.ProjectName).trim().toLowerCase();
      const existingProjects = readAllRows(sheet, headerMap);
      const duplicate = existingProjects.find(
        (p) =>
          String(p.ProjectID) !== String(projectId) && String(p.ProjectName).trim().toLowerCase() === newName
      );
      if (duplicate) {
        const err = new Error(
          `A project named "${updates.ProjectName}" already exists (ProjectID ${duplicate.ProjectID}).`
        );
        err.statusCode = 409;
        throw err;
      }
    }

    const changedFields = [];

    EDITABLE_PROJECT_COLUMNS.forEach((columnName) => {
      if (Object.prototype.hasOwnProperty.call(updates, columnName)) {
        const colIndex = headerMap[columnName];
        if (colIndex && updates[columnName] !== before[columnName]) {
          row.getCell(colIndex).value = updates[columnName];
          changedFields.push(columnName);
        }
      }
    });

    const updatedDateCol = headerMap['UpdatedDate'];
    const now = new Date().toISOString();
    if (updatedDateCol) {
      row.getCell(updatedDateCol).value = now;
    }
    row.commit();

    changedFields.forEach((field) => {
      appendAuditLogEntry(workbook, {
        entityType: 'Project',
        entityId: projectId,
        action: 'Update',
        changedField: field,
        oldValue: before[field] ?? '',
        newValue: updates[field] ?? '',
      });
    });

    await saveWorkbook(workbook, filePath);
    return rowToObject(row, headerMap);
  });
}

/**
 * Reads valid Status lookup values from the Settings sheet (SettingType = "Status").
 *
 * NOT currently used for Project status validation — in the real data template
 * the "Status" SettingType rows (Open / In Progress / Closed) are Bug statuses,
 * not Project statuses (Active / On Hold / Completed). Settings has no
 * Project-specific entry, so wiring this into Project validation would reject
 * every legitimate Project status, including the default a new project is
 * created with. See controllers/projects.controller.js, which validates
 * Project status against ALLOWED_STATUSES_FALLBACK instead.
 * Left in place (and still covered by its own tests) for whichever future
 * story adds an entity-specific Settings key (e.g. a "Module"/"Entity"
 * column, or a distinct SettingType like "ProjectStatus") — that's a
 * Settings-sheet/schema decision that needs approval per the Master AI
 * Coding Instructions, so it isn't made here.
 * Returns null if unavailable, so callers can fall back to a safe default.
 */
async function getAllowedStatuses(filePath) {
  const workbook = await loadWorkbook(filePath);
  const worksheet = workbook.getWorksheet(SETTINGS_SHEET);
  if (!worksheet) return null;

  const headerMap = getHeaderMap(worksheet);
  if (!headerMap['SettingType'] || !headerMap['Value']) return null;

  const rows = readAllRows(worksheet, headerMap);
  const statuses = rows
    .filter((r) => String(r.SettingType).trim() === 'Status')
    .map((r) => String(r.Value).trim())
    .filter(Boolean);

  return statuses.length > 0 ? statuses : null;
}

module.exports = {
  createProject,
  getProjectById,
  updateProject,
  getAllowedStatuses,
  PROJECTS_SHEET,
  SETTINGS_SHEET,
  EDITABLE_PROJECT_COLUMNS,
};
