const { execFile, exec } = require('child_process');
const path = require('path');
const fs = require('fs-extra');
const config = require('../config/config');

class WordConverterService {
  constructor() {
    this.queue = [];
    this.isProcessing = false;
    this.totalConverted = 0;
    this.totalErrors = 0;
  }

  /**
   * Healthcheck to test Word COM responsiveness
   */
  async checkEngineHealth() {
    return new Promise((resolve) => {
      const psCommand = 'powershell -NoProfile -Command "try { $w = New-Object -ComObject Word.Application; $v = $w.Version; $w.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($w) | Out-Null; Write-Output $v } catch { Write-Output \'ERROR\' }"';
      exec(psCommand, { timeout: 10000 }, (error, stdout) => {
        if (error || !stdout || stdout.includes('ERROR')) {
          resolve({ status: 'degraded', engine: 'Native COM (Unavailable/Busy)', version: null });
        } else {
          resolve({ status: 'healthy', engine: 'Microsoft Word Native Engine', version: stdout.trim() });
        }
      });
    });
  }

  /**
   * Enqueue a single file conversion job
   */
  async convertFile(inputPath, outputPath, options = {}) {
    return new Promise((resolve, reject) => {
      this.queue.push({
        inputPath,
        outputPath,
        options,
        resolve,
        reject
      });
      this.processQueue();
    });
  }

  /**
   * Sequential Queue Processor for Word COM stability
   */
  async processQueue() {
    if (this.isProcessing || this.queue.length === 0) return;

    this.isProcessing = true;
    const task = this.queue.shift();

    try {
      const result = await this.executePowerShellConversion(
        task.inputPath,
        task.outputPath,
        task.options
      );
      this.totalConverted++;
      task.resolve(result);
    } catch (err) {
      this.totalErrors++;
      task.reject(err);
    } finally {
      this.isProcessing = false;
      // Continue next queued conversion
      setImmediate(() => this.processQueue());
    }
  }

  /**
   * Execute convert_word.ps1 with timeout and resilience
   */
  executePowerShellConversion(inputPath, outputPath, options) {
    return new Promise((resolve, reject) => {
      const scriptPath = path.join(config.SCRIPTS_DIR, 'convert_word.ps1');
      const quality = options.quality || 'Standard';
      const embedFonts = options.embedFonts !== false ? '1' : '0';

      const args = [
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', scriptPath,
        '-InputPath', inputPath,
        '-OutputPath', outputPath,
        '-Quality', quality,
        '-EmbedFonts', embedFonts
      ];

      let isSettled = false;

      const child = execFile('powershell.exe', args, {
        timeout: config.WORD_COM_TIMEOUT_MS,
        maxBuffer: 1024 * 1024 * 10
      }, (error, stdout, stderr) => {
        if (isSettled) return;
        isSettled = true;

        if (error) {
          console.error('[WordConverterService] Execution error:', error.message, stderr);
          if (error.killed) {
            this.killStrayWordProcesses();
            return reject(new Error('Conversion timed out (file may be locked or password protected).'));
          }
          return reject(new Error(error.message || stderr || 'Conversion failed.'));
        }

        try {
          const lines = stdout.trim().split('\n');
          const jsonLine = lines.find(l => l.trim().startsWith('{') && l.trim().endsWith('}')) || lines[lines.length - 1];
          const parsed = JSON.parse(jsonLine);

          if (parsed.success) {
            resolve({
              pdfPath: parsed.pdfPath,
              pageCount: parsed.pageCount || 1,
              engine: 'Microsoft Word COM'
            });
          } else {
            reject(new Error(parsed.error || 'Conversion reported failure.'));
          }
        } catch (parseErr) {
          console.error('[WordConverterService] Failed to parse output:', stdout);
          reject(new Error('Failed to parse conversion output script result.'));
        }
      });
    });
  }

  killStrayWordProcesses() {
    exec('taskkill /f /im WINWORD.EXE', () => {
      console.log('[WordConverterService] Cleared orphan Word processes after timeout.');
    });
  }

  getStats() {
    return {
      queueLength: this.queue.length,
      isProcessing: this.isProcessing,
      totalConverted: this.totalConverted,
      totalErrors: this.totalErrors
    };
  }
}

module.exports = new WordConverterService();
