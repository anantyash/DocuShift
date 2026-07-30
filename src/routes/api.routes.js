const express = require('express');
const router = express.Router();
const upload = require('../middleware/upload.middleware');
const controller = require('../controllers/conversion.controller');

// Batch convert Word files to PDF
router.post('/convert', upload.array('files', 20), (req, res) => controller.convertFiles(req, res));

// Merge converted PDFs
router.post('/merge', (req, res) => controller.mergeFiles(req, res));

// Create batch ZIP
router.post('/zip', (req, res) => controller.createZipBatch(req, res));

// Download file
router.get('/download/:fileId', (req, res) => controller.downloadFile(req, res));

// Stream PDF preview
router.get('/preview/:fileId', (req, res) => controller.previewFile(req, res));

// Health & System status
router.get('/health', (req, res) => controller.getHealthStats(req, res));

module.exports = router;
