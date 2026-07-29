// Central config for the QA Management Tool backend.
// Per TRD V2: storage is a SharePoint-hosted Excel workbook.
// Locally, this points at the synced copy of the workbook on disk.
// Override with the EXCEL_FILE_PATH environment variable if your
// SharePoint sync folder differs from the default project layout.

const path = require('path');

const EXCEL_FILE_PATH =
  process.env.EXCEL_FILE_PATH ||
  path.resolve(__dirname, '..', '..', '..', 'Data', 'QA_Management_Data_Template.xlsx');

const PORT = process.env.PORT || 3000;

module.exports = {
  EXCEL_FILE_PATH,
  PORT,
};
