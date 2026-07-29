/**
 * Validation for the Projects resource.
 * validateCreateProjectPayload  -> US-001 Create Project
 * validateEditProjectPayload    -> US-002 Edit Project
 *
 * US-003 View Project needs no payload validation of its own — it's a
 * read-by-id, and "does this ProjectID exist" is already handled by
 * getProjectById() returning null -> the controller returning 404.
 */

// This is the actual source of truth for Project status right now, not just a
// "fallback" — the Settings sheet's "Status" SettingType rows are Bug statuses
// (Open / In Progress / Closed) in the real data template, not Project statuses,
// so callers (see projects.controller.js) pass this in explicitly rather than
// pulling from Settings. See projectsService.getAllowedStatuses() doc comment.
const ALLOWED_STATUSES_FALLBACK = ['Active', 'On Hold', 'Completed'];

const EDITABLE_FIELDS = ['ProjectName', 'Description', 'Status', 'StartDate', 'EndDate'];
const IMMUTABLE_FIELDS = ['ProjectID', 'CreatedDate', 'UpdatedDate'];

function isNonEmptyString(value) {
  return typeof value === 'string' && value.trim().length > 0;
}

function isValidDateValue(value) {
  if (value === undefined || value === null || value === '') return true; // dates are optional
  const parsed = new Date(value);
  return !isNaN(parsed.getTime());
}

function hasField(payload, field) {
  return Object.prototype.hasOwnProperty.call(payload, field);
}

function validateDateOrder(payload, errors) {
  if (hasField(payload, 'StartDate') && hasField(payload, 'EndDate') && payload.StartDate && payload.EndDate) {
    const start = new Date(payload.StartDate);
    const end = new Date(payload.EndDate);
    if (!isNaN(start.getTime()) && !isNaN(end.getTime()) && end < start) {
      errors.push('EndDate cannot be earlier than StartDate.');
    }
  }
}

/**
 * @param {*} payload - the request body for POST /api/projects
 * @param {string[]} [allowedStatuses] - valid Status values if `status` is provided;
 *   defaults to the same fallback list Edit Project uses, keeping Create and Edit
 *   consistent (see ALLOWED_STATUSES_FALLBACK doc note above about Settings).
 * @returns {{ valid: boolean, errors: string[] }}
 */
function validateCreateProjectPayload(payload, allowedStatuses = ALLOWED_STATUSES_FALLBACK) {
  const errors = [];

  if (payload === null || typeof payload !== 'object' || Array.isArray(payload)) {
    return { valid: false, errors: ['Request body must be a JSON object.'] };
  }

  if (!isNonEmptyString(payload.projectName)) {
    errors.push('projectName is required and must be a non-empty string.');
  }

  if (hasField(payload, 'description') && payload.description !== null && typeof payload.description !== 'string') {
    errors.push('description must be a string.');
  }

  // status is optional on Create (defaults to 'Active' in the service), but if the
  // caller does supply one it must be one of the allowed values — same rule Edit
  // already enforces on Status. Previously Create had no check here at all.
  if (hasField(payload, 'status') && payload.status !== null && payload.status !== '') {
    if (!isNonEmptyString(payload.status) || !allowedStatuses.includes(payload.status)) {
      errors.push(`status must be one of: ${allowedStatuses.join(', ')}.`);
    }
  }

  if (!isValidDateValue(payload.startDate)) errors.push('startDate must be a valid date.');
  if (!isValidDateValue(payload.endDate)) errors.push('endDate must be a valid date.');

  validateDateOrder({ StartDate: payload.startDate, EndDate: payload.endDate }, errors);

  return { valid: errors.length === 0, errors };
}

/**
 * @param {*} payload - the request body for PUT /api/projects/:id
 * @param {string[]} [allowedStatuses] - valid Status values, normally sourced from the Settings sheet
 * @returns {{ valid: boolean, errors: string[] }}
 */
function validateEditProjectPayload(payload, allowedStatuses = ALLOWED_STATUSES_FALLBACK) {
  const errors = [];

  if (payload === null || typeof payload !== 'object' || Array.isArray(payload)) {
    return { valid: false, errors: ['Request body must be a JSON object.'] };
  }

  IMMUTABLE_FIELDS.forEach((field) => {
    if (hasField(payload, field)) {
      errors.push(`Field "${field}" cannot be modified via Edit Project.`);
    }
  });

  if (hasField(payload, 'ProjectName') && !isNonEmptyString(payload.ProjectName)) {
    errors.push('ProjectName is required and must be a non-empty string.');
  }

  if (hasField(payload, 'Description') && payload.Description !== null && typeof payload.Description !== 'string') {
    errors.push('Description must be a string.');
  }

  if (hasField(payload, 'Status')) {
    if (!isNonEmptyString(payload.Status) || !allowedStatuses.includes(payload.Status)) {
      errors.push(`Status must be one of: ${allowedStatuses.join(', ')}.`);
    }
  }

  if (hasField(payload, 'StartDate') && !isValidDateValue(payload.StartDate)) {
    errors.push('StartDate must be a valid date.');
  }
  if (hasField(payload, 'EndDate') && !isValidDateValue(payload.EndDate)) {
    errors.push('EndDate must be a valid date.');
  }

  validateDateOrder(payload, errors);

  const hasAnyEditableField = EDITABLE_FIELDS.some((field) => hasField(payload, field));
  if (!hasAnyEditableField) {
    errors.push(`At least one editable field (${EDITABLE_FIELDS.join(', ')}) must be provided.`);
  }

  return { valid: errors.length === 0, errors };
}

module.exports = {
  validateCreateProjectPayload,
  validateEditProjectPayload,
  ALLOWED_STATUSES_FALLBACK,
  EDITABLE_FIELDS,
  IMMUTABLE_FIELDS,
};
