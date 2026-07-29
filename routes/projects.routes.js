const express = require('express');
const { createProject, getProject, editProject } = require('../controllers/projects.controller');

const router = express.Router();

// POST /api/projects        - US-001 Create Project
router.post('/', createProject);

// GET /api/projects/:id     - US-003 View Project (also used to pre-fill Edit Project)
router.get('/:id', getProject);

// PUT /api/projects/:id     - US-002 Edit Project
router.put('/:id', editProject);

module.exports = router;
