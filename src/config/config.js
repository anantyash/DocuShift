const path = require("path");

const ROOT_DIR = path.resolve(__dirname, "../../");
const UPLOADS_DIR = path.join(ROOT_DIR, "temp/uploads");
const OUTPUTS_DIR = path.join(ROOT_DIR, "temp/outputs");
const SCRIPTS_DIR = path.join(ROOT_DIR, "scripts");

module.exports = {
  PORT: process.env.PORT || 3000,
  ROOT_DIR,
  UPLOADS_DIR,
  OUTPUTS_DIR,
  SCRIPTS_DIR,
  MAX_FILE_SIZE_MB: 50,
  MAX_FILES_PER_BATCH: 20,
  ALLOWED_EXTENSIONS: [".docx", ".doc"],
  ALLOWED_MIME_TYPES: [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/msword",
    "application/x-msword",
  ],
  CLEANUP_INTERVAL_MINUTES: 15,
  FILE_TTL_MINUTES: 30,
  WORD_COM_TIMEOUT_MS: 45000, // 45 sec timeout per file conversion
};
