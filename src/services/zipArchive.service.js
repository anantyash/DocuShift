const archiver = require('archiver');
const fs = require('fs-extra');

class ZipArchiveService {
  /**
   * Creates a ZIP archive containing multiple converted PDF files
   * @param {Array<{ path: string, name: string }>} files List of file descriptors
   * @param {string} outputPath Target ZIP file path
   * @returns {Promise<string>}
   */
  async createZip(files, outputPath) {
    return new Promise((resolve, reject) => {
      const output = fs.createWriteStream(outputPath);
      const archive = archiver('zip', {
        zlib: { level: 9 }
      });

      output.on('close', () => resolve(outputPath));
      archive.on('error', (err) => reject(err));

      archive.pipe(output);

      files.forEach((file) => {
        if (fs.existsSync(file.path)) {
          archive.file(file.path, {
            name: file.name,
            store: false
          });
        }
      });

      archive.finalize();
    });
  }
}

module.exports = new ZipArchiveService();
