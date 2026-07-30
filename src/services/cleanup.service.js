const fs = require('fs-extra');
const path = require('path');
const config = require('../config/config');

class CleanupService {
  constructor() {
    this.timer = null;
  }

  start() {
    // Run cleanup immediately, then schedule periodic runs
    this.cleanTempDirectories();
    const intervalMs = config.CLEANUP_INTERVAL_MINUTES * 60 * 1000;
    this.timer = setInterval(() => this.cleanTempDirectories(), intervalMs);
    console.log(`[CleanupService] Started automated cleanup task (Interval: ${config.CLEANUP_INTERVAL_MINUTES}m)`);
  }

  stop() {
    if (this.timer) {
      clearInterval(this.timer);
      this.timer = null;
    }
  }

  async cleanTempDirectories() {
    try {
      await fs.ensureDir(config.UPLOADS_DIR);
      await fs.ensureDir(config.OUTPUTS_DIR);

      const now = Date.now();
      const ttlMs = config.FILE_TTL_MINUTES * 60 * 1000;

      await this.cleanFolder(config.UPLOADS_DIR, now, ttlMs);
      await this.cleanFolder(config.OUTPUTS_DIR, now, ttlMs);
    } catch (err) {
      console.error('[CleanupService] Error cleaning temporary folders:', err.message);
    }
  }

  async cleanFolder(dirPath, now, ttlMs) {
    const files = await fs.readdir(dirPath);
    for (const file of files) {
      const filePath = path.join(dirPath, file);
      try {
        const stats = await fs.stat(filePath);
        if (now - stats.mtimeMs > ttlMs) {
          await fs.remove(filePath);
          console.log(`[CleanupService] Purged expired temp file: ${file}`);
        }
      } catch (err) {
        // File may be locked, skip
      }
    }
  }
}

module.exports = new CleanupService();
