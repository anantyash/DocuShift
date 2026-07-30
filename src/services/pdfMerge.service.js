const { PDFDocument } = require('pdf-lib');
const fs = require('fs-extra');

class PdfMergeService {
  /**
   * Merges multiple PDF files into one combined PDF document
   * @param {string[]} pdfPaths List of absolute paths to PDF files
   * @param {string} outputPath Destination path for merged PDF
   * @returns {Promise<{ mergedPath: string, totalPages: number }>}
   */
  async mergePdfs(pdfPaths, outputPath) {
    if (!pdfPaths || pdfPaths.length === 0) {
      throw new Error('No PDF files provided for merging.');
    }

    const mergedPdf = await PDFDocument.create();

    for (const filePath of pdfPaths) {
      if (!await fs.pathExists(filePath)) {
        console.warn(`[PdfMergeService] Skipping missing PDF file: ${filePath}`);
        continue;
      }

      const pdfBytes = await fs.readFile(filePath);
      const pdf = await PDFDocument.load(pdfBytes);
      const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());

      copiedPages.forEach((page) => mergedPdf.addPage(page));
    }

    const totalPages = mergedPdf.getPageCount();
    const mergedPdfBytes = await mergedPdf.save();
    await fs.writeFile(outputPath, mergedPdfBytes);

    return {
      mergedPath: outputPath,
      totalPages
    };
  }
}

module.exports = new PdfMergeService();
