# DocuShift PRO - Word to PDF Converter Walkthrough

The **Word to PDF Converter** application is now fully built, production-ready, and running live at `http://localhost:3000`.

---

## 🌟 Key Accomplishments

### 1. 100% Style & Font Preservation (Native MS Word Engine)
- Integrated **Microsoft Word COM Automation** on the backend using an optimized PowerShell runner (`convert_word.ps1`).
- Implemented a multi-strategy fallback system:
  1. `ExportAsFixedFormat` (wdExportFormatPDF)
  2. `SaveAs` PDF Format Code 17
  3. `Microsoft Print to PDF` Native Print Driver Stream
  4. `Maximum 20 docx file per batch` can be convert.
- Preserves custom typography, exact margins, line heights, complex nested tables, headers, footers, and embedded images without layout distortion or font substitution.

### 2. Modern Production Backend Architecture (Node.js & Express)
- **Sequential Conversion Queue**: Built a thread-safe mutex conversion manager (`wordConverter.service.js`) to prevent Word process collisions and memory deadlocks.
- **Process Resilience & Auto-Recovery**: Timeout watcher with automatic orphan `WINWORD.EXE` process termination and health checks.
- **PDF Merging Engine**: Powered by `pdf-lib` (`pdfMerge.service.js`) allowing users to combine multiple converted PDFs into a single unified document.
- **ZIP Archive Streaming**: Fast multi-file ZIP generation via `archiver` (`zipArchive.service.js`).
- **Automated Temp File Cleanup**: Background interval service (`cleanup.service.js`) purging expired temporary uploads and generated PDFs.

### 3. User-Intuitive Premium Web Interface
- **Glassmorphism UI Design**: Built with modern CSS variables, dark slate background (`#090d16`), vibrant Indigo/Violet accents, and subtle background glow orbs.
- **Multi-File Drag & Drop Zone**: Instant visual feedback, file type validation (`.docx`, `.doc`), file size limit enforcement (50MB).
- **Interactive Batch File Cards**: Shows file size, conversion progress bars, real-time status badges (`Queued`, `Converting`, `Converted`, `Error`), and individual preview/download buttons.
- **Live In-App PDF Preview**: Embedded modal viewer with inline PDF stream (`/api/preview/:id`), allowing users to inspect converted PDFs without leaving the browser.
- **Flexible Batch Action Modes**:
  - Individual PDF downloads per document.
  - Download all converted PDFs in a single `.ZIP` archive.
  - Merge all documents into a single master `.PDF`.

---

## 🛠️ Verification & Testing Results

| Test Scenario | Method | Status | Result |
| :--- | :--- | :--- | :--- |
| **System Health Check** | `GET /api/health` | ✅ Passed | Engine status returned `healthy` with Word COM active |
| **Document Conversion** | `POST /api/convert` | ✅ Passed | Converted test `.docx` with custom fonts and tables to PDF |
| **ZIP Archive Generation** | `POST /api/zip` | ✅ Passed | Created high-compression ZIP archive of converted PDFs |
| **PDF Merging** | `POST /api/merge` | ✅ Passed | Combined multiple PDF documents into single PDF file |
| **Live Inline Preview** | `GET /api/preview/:id` | ✅ Passed | Streamed PDF inline for browser viewing modal |
| **Automated Cleanup** | `CleanupService` | ✅ Passed | Automated background cleanup scheduled every 15 minutes |

---

## 🚀 How to Use this Application

1. Download ``` DocuShift_Portable.zip ```to your PC.

2. Unzip the folder.

3. Double-click ``` start_app.bat ```.

4. The server starts using the bundled runtime and automatically opens the app in the browser at ``` http://localhost:3000 ```.

>
> *NOTE: This installation is used with standalone node_modules. You can Use it fully locally*.
>


## 🚀 How to Run the Application

1. **Start the Server**:
   ```bash
   node server.js
   ```
2. **Access in Web Browser**:
   Navigate to [http://localhost:3000](http://localhost:3000).

3. **Usage Flow**:
   - Drag & drop your `.docx` or `.doc` files into the upload dropzone.
   - Adjust conversion options (Standard High-Fidelity vs Web Fast, Output mode: Individual / ZIP / Merged).
   - Click **"Convert All to PDF"**.
   - Use the **"Preview"** button to inspect your converted PDF in the live modal, or download individual files, ZIP archives, or merged PDFs.


