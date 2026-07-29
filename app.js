require('dotenv').config();

const express = require('express');
const cors = require('cors');
const config = require('./config');
const projectsRouter = require('./routes/projects.routes');

const app = express();

app.use(cors());
app.use(express.json());

app.use('/api/projects', projectsRouter);

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', excelFile: config.EXCEL_FILE_PATH });
});

// 404 handler — unmatched routes
app.use((req, res) => {
  res.status(404).json({ error: `Not found: ${req.method} ${req.originalUrl}` });
});

// Global error handler — catches anything not already handled by a route's try/catch
// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  // eslint-disable-next-line no-console
  console.error('Unhandled error:', err);
  res.status(err.statusCode || 500).json({ error: err.message || 'Internal server error' });
});

module.exports = app;
