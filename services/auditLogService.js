/**
 * AuditLog service — append-only, per TRD V2 storage rules.
 * Any resource service that creates/edits/deletes a record should call
 * appendAuditLogEntry() as part of the same workbook load/save cycle, so the
 * entity change and its audit record are saved together.
 *
 * Never edits or removes existing AuditLog rows — only ever appends.
 */

const {
  getWorksheetOrThrow,
  getHeaderMap,
  readAllRows,
  appendRowByHeader,
  nextNumericId,
} = require('./excelWorkbook');

const AUDIT_LOG_SHEET = 'AuditLog';

/**
 * @param {import('exceljs').Workbook} workbook - an already-loaded workbook (not yet saved)
 * @param {object} entry
 * @param {string} entry.entityType - e.g. 'Project'
 * @param {string|number} entry.entityId
 * @param {string} entry.action - e.g. 'Create', 'Update'
 * @param {string} [entry.changedField]
 * @param {string} [entry.oldValue]
 * @param {string} [entry.newValue]
 */
function appendAuditLogEntry(workbook, { entityType, entityId, action, changedField, oldValue, newValue }) {
  const sheet = getWorksheetOrThrow(workbook, AUDIT_LOG_SHEET);
  const headerMap = getHeaderMap(sheet);
  const existing = readAllRows(sheet, headerMap);
  const auditId = nextNumericId(existing, 'AuditID');

  appendRowByHeader(sheet, headerMap, {
    AuditID: auditId,
    EntityType: entityType,
    EntityID: entityId,
    Action: action,
    ChangedField: changedField ?? '',
    OldValue: oldValue ?? '',
    NewValue: newValue ?? '',
    ChangeDate: new Date().toISOString(),
  });
}

module.exports = { appendAuditLogEntry, AUDIT_LOG_SHEET };
