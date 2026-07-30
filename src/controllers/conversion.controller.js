const path = require("path");
const fs = require("fs-extra");
const { v4: uuidv4 } = require("uuid");
const config = require("../config/config");
const wordConverterService = require("../services/wordConverter.service");
const pdfMergeService = require("../services/pdfMerge.service");
const zipArchiveService = require("../services/zipArchive.service");

/**
 * Safely decode Multer latin1 filename string to proper UTF-8 (Hindi, Devanagari, Chinese, Arabic, etc.)
 */
function fixUtf8Filename(str) {
  if (!str) return str;
  try {
    const decoded = Buffer.from(str, "latin1").toString("utf8");
    // If decoding produced valid characters, use it
    if (decoded && !decoded.includes("")) {
      return decoded;
    }
  } catch (e) {}
  return str;
}

class ConversionController {
  /**
   * Convert multiple Word documents (.docx/.doc) to PDF
   */
  async convertFiles(req, res) {
    try {
      if (!req.files || req.files.length === 0) {
        return res
          .status(400)
          .json({ success: false, error: "No Word files uploaded." });
      }

      const quality = req.body.quality || "Standard";
      const embedFonts = req.body.embedFonts !== "false";
      const batchId = uuidv4();

      await fs.ensureDir(config.OUTPUTS_DIR);

      const results = [];

      for (const file of req.files) {
        const fileId = path.parse(file.filename).name;
        const originalName = fixUtf8Filename(file.originalname);
        const baseName = path.parse(originalName).name;
        const pdfFilename = `${baseName}.pdf`;
        const internalPdfFilename = `${fileId}.pdf`;
        const outputPdfPath = path.join(
          config.OUTPUTS_DIR,
          internalPdfFilename,
        );

        try {
          const conversion = await wordConverterService.convertFile(
            file.path,
            outputPdfPath,
            {
              quality,
              embedFonts,
              clientId: req.body.clientId,
            },
          );

          results.push({
            jobId: conversion.jobId,
            fileId,
            originalName,
            baseName,
            pdfFilename,
            internalPdfName: internalPdfFilename,
            pageCount: conversion.pageCount,
            engine: conversion.engine,
            downloadUrl: `/api/download/${fileId}`,
            previewUrl: `/api/preview/${fileId}`,
            status: "success",
          });
        } catch (convErr) {
          console.error(`Failed converting ${originalName}:`, convErr.message);
          results.push({
            fileId,
            originalName,
            baseName,
            error: convErr.message,
            status: "error",
          });
        }
      }

      const successCount = results.filter((r) => r.status === "success").length;

      return res.json({
        success: successCount > 0,
        batchId,
        total: results.length,
        convertedCount: successCount,
        results,
      });
    } catch (err) {
      console.error("[ConversionController] Error:", err);
      return res.status(500).json({
        success: false,
        error: err.message || "Internal conversion error.",
      });
    }
  }

  /**
   * Merge multiple converted PDFs into a single PDF
   */
  async mergeFiles(req, res) {
    try {
      const { fileIds, customTitle } = req.body;
      if (!fileIds || !Array.isArray(fileIds) || fileIds.length === 0) {
        return res.status(400).json({
          success: false,
          error: "Select at least one converted document to merge.",
        });
      }

      const pdfPaths = fileIds
        .map((id) => path.join(config.OUTPUTS_DIR, `${id}.pdf`))
        .filter((p) => fs.existsSync(p));

      if (pdfPaths.length === 0) {
        return res.status(404).json({
          success: false,
          error: "Specified PDF files were not found.",
        });
      }

      const mergedFileId = `merged-${uuidv4()}`;
      const mergedPdfPath = path.join(
        config.OUTPUTS_DIR,
        `${mergedFileId}.pdf`,
      );

      const result = await pdfMergeService.mergePdfs(pdfPaths, mergedPdfPath);
      const downloadName =
        (customTitle ? customTitle.trim() : "Merged_Document") + ".pdf";

      return res.json({
        success: true,
        mergedFileId,
        downloadName,
        totalPages: result.totalPages,
        downloadUrl: `/api/download/${mergedFileId}?name=${encodeURIComponent(downloadName)}`,
        previewUrl: `/api/preview/${mergedFileId}`,
      });
    } catch (err) {
      console.error("[ConversionController] Merge error:", err);
      return res
        .status(500)
        .json({ success: false, error: err.message || "PDF merge failed." });
    }
  }

  /**
   * Create downloadable ZIP archive of batch PDFs
   */
  async createZipBatch(req, res) {
    try {
      const { files } = req.body; // Array of { fileId, displayName }
      if (!files || !Array.isArray(files) || files.length === 0) {
        return res.status(400).json({
          success: false,
          error: "No converted files provided for ZIP.",
        });
      }

      const zipId = `batch-${uuidv4()}`;
      const zipPath = path.join(config.OUTPUTS_DIR, `${zipId}.zip`);

      const fileList = files
        .map((f) => ({
          path: path.join(config.OUTPUTS_DIR, `${f.fileId}.pdf`),
          name: f.displayName || `${f.fileId}.pdf`,
        }))
        .filter((f) => fs.existsSync(f.path));

      if (fileList.length === 0) {
        return res.status(404).json({
          success: false,
          error: "No valid converted PDF files found for ZIP.",
        });
      }

      await zipArchiveService.createZip(fileList, zipPath);

      return res.json({
        success: true,
        zipId,
        downloadUrl: `/api/download/${zipId}?type=zip&name=${encodeURIComponent("Converted_Documents.zip")}`,
      });
    } catch (err) {
      console.error("[ConversionController] ZIP error:", err);
      return res.status(500).json({
        success: false,
        error: err.message || "ZIP archive creation failed.",
      });
    }
  }

  /**
   * Download a converted PDF or ZIP file with RFC 5987 UTF-8 filename header
   */
  async downloadFile(req, res) {
    try {
      const fileId = req.params.fileId;
      const isZip = req.query.type === "zip" || fileId.endsWith(".zip");
      const ext = isZip ? ".zip" : ".pdf";
      const cleanId = fileId.replace(/\.(pdf|zip)$/i, "");
      const filePath = path.join(config.OUTPUTS_DIR, `${cleanId}${ext}`);

      if (!(await fs.pathExists(filePath))) {
        return res.status(404).send("File not found or expired.");
      }

      const rawName =
        req.query.name ||
        (isZip ? "Converted_Documents.zip" : `${cleanId}.pdf`);
      const clientName = decodeURIComponent(rawName);
      const encodedFilename = encodeURIComponent(clientName);

      res.setHeader(
        "Content-Type",
        isZip ? "application/zip" : "application/pdf",
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${encodedFilename}"; filename*=UTF-8''${encodedFilename}`,
      );

      fs.createReadStream(filePath).pipe(res);
    } catch (err) {
      console.error("[ConversionController] Download error:", err);
      res.status(500).send("Error serving download.");
    }
  }

  /**
   * Stream PDF file inline for live in-browser preview
   */
  async previewFile(req, res) {
    try {
      const fileId = req.params.fileId.replace(/\.pdf$/i, "");
      const filePath = path.join(config.OUTPUTS_DIR, `${fileId}.pdf`);

      if (!(await fs.pathExists(filePath))) {
        return res.status(404).send("Preview file not found.");
      }

      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", "inline");
      fs.createReadStream(filePath).pipe(res);
    } catch (err) {
      res.status(500).send("Error loading preview.");
    }
  }

  /**
   * Healthcheck & engine status check
   */
  async getHealthStats(req, res) {
    const health = await wordConverterService.checkEngineHealth();
    const stats = wordConverterService.getStats();

    return res.json({
      success: true,
      engine: health,
      queue: stats,
    });
  }
}

module.exports = new ConversionController();
