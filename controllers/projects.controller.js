const projectsService = require('../services/projectsService');
const {
  validateCreateProjectPayload,
  validateEditProjectPayload,
  ALLOWED_STATUSES_FALLBACK,
} = require('../validators/projects.validator');

/**
 * POST /api/projects
 * US-001 Create Project
 */
async function createProject(req, res) {
  const { valid, errors } = validateCreateProjectPayload(req.body, ALLOWED_STATUSES_FALLBACK);
  if (!valid) {
    return res.status(400).json({ errors });
  }

  try {
    const project = await projectsService.createProject(req.body);
    return res.status(201).json(project);
  } catch (err) {
    const statusCode = err.statusCode || 500;
    if (statusCode === 500) {
      // eslint-disable-next-line no-console
      console.error('Error creating project:', err);
    }
    return res.status(statusCode).json({ error: err.message });
  }
}

/**
 * GET /api/projects/:id
 * US-003 View Project.
 * Also doubles as the read used to pre-populate the Edit Project (US-002)
 * form before saving — same data, same shape, no need for a second endpoint.
 * Returns the project object directly (not wrapped), matching the response
 * shape used across the rest of this resource.
 */
async function getProject(req, res) {
  try {
    const { id } = req.params;
    const project = await projectsService.getProjectById(id);
    if (!project) {
      return res.status(404).json({ error: `Project with ProjectID "${id}" was not found.` });
    }
    return res.status(200).json(project);
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error('Error reading project:', err);
    return res.status(500).json({ error: 'Failed to read project.', details: err.message });
  }
}

/**
 * PUT /api/projects/:id
 * US-002 Edit Project
 */
async function editProject(req, res) {
  try {
    const { id } = req.params;

    // NOT sourced from projectsService.getAllowedStatuses() / the Settings sheet:
    // in the real data template, Settings' "Status" rows are Bug statuses (Open /
    // In Progress / Closed), not Project statuses, so pulling from there would
    // reject every legitimate Project status. See ALLOWED_STATUSES_FALLBACK's
    // doc comment in projects.validator.js.
    const allowedStatuses = ALLOWED_STATUSES_FALLBACK;

    const { valid, errors } = validateEditProjectPayload(req.body, allowedStatuses);
    if (!valid) {
      return res.status(400).json({ errors });
    }

    const existing = await projectsService.getProjectById(id);
    if (!existing) {
      return res.status(404).json({ error: `Project with ProjectID "${id}" was not found.` });
    }

    const updated = await projectsService.updateProject(id, req.body);
    return res.status(200).json(updated);
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error('Error updating project:', err);
    return res.status(500).json({ error: 'Failed to update project.', details: err.message });
  }
}

module.exports = { createProject, getProject, editProject };
