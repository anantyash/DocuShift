/**
 * DocuShift PRO - Word to PDF Converter Application Client Logic
 */

document.addEventListener("DOMContentLoaded", () => {
  // DOM Elements
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const browseBtn = document.getElementById("browseBtn");

  const queueSection = document.getElementById("queueSection");
  const queueCount = document.getElementById("queueCount");
  const fileList = document.getElementById("fileList");
  const clearAllBtn = document.getElementById("clearAllBtn");

  const convertAllBtn = document.getElementById("convertAllBtn");
  const downloadZipBtn = document.getElementById("downloadZipBtn");
  const mergeDownloadBtn = document.getElementById("mergeDownloadBtn");
  const batchSummary = document.getElementById("batchSummary");

  const qualitySelect = document.getElementById("qualitySelect");
  const embedFontsToggle = document.getElementById("embedFontsToggle");

  const overallProgressContainer = document.getElementById(
    "overallProgressContainer",
  );
  const overallProgressText = document.getElementById("overallProgressText");
  const overallPercentText = document.getElementById("overallPercentText");
  const overallProgressFill = document.getElementById("overallProgressFill");

  const engineText = document.getElementById("engineText");

  // Modal Elements
  const previewModal = document.getElementById("previewModal");
  const modalFileName = document.getElementById("modalFileName");
  const pdfPreviewFrame = document.getElementById("pdfPreviewFrame");
  const modalDownloadBtn = document.getElementById("modalDownloadBtn");
  const modalNewTabBtn = document.getElementById("modalNewTabBtn");
  const closeModalBtn = document.getElementById("closeModalBtn");

  const toastContainer = document.getElementById("toastContainer");

  // Application State
  let filesQueue = []; // Array of { id, file, status, pageCount, resultData, errorMessage }
  const jobs = new Map();
  let socket = null;
  let clientId = null;
  let isConverting = false;

  // Initialize
  initializeWebSocket();
  checkSystemHealth();
  setupEventListeners();

  /**
   * System Health Check API
   */
  async function checkSystemHealth() {
    try {
      const res = await fetch("/api/health");
      const data = await res.json();
      if (data.success && data.engine) {
        if (data.engine.status === "healthy") {
          engineText.textContent = `MS Word Native Engine Active (${data.engine.version || "COM"})`;
        } else {
          engineText.textContent = "Native COM Engine Ready";
        }
      }
    } catch (err) {
      engineText.textContent = "Word Engine Ready";
    }
  }

  function initializeWebSocket() {
    const protocol = location.protocol === "https:" ? "wss:" : "ws:";
    socket = new WebSocket(`${protocol}//${location.host}`);

    socket.onopen = () => {
      console.log("WebSocket Connected");
    };

    socket.onmessage = (event) => {
      const message = JSON.parse(event.data);
      // console.log("WS:", message);

      switch (message.type) {
        case "connected":
          clientId = message.clientId;
          break;

        case "job":
          handleJobUpdate(message);
          break;

        case "pong":
          break;
      }
    };

    socket.onclose = () => {
      console.log("WebSocket Closed");
    };

    socket.onerror = (err) => {
      console.log(err);
    };
  }

  /**
   * Setup Event Listeners
   */
  function setupEventListeners() {
    browseBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      fileInput.click();
    });

    dropzone.addEventListener("click", (e) => {
      if (e.target !== browseBtn && !browseBtn.contains(e.target)) {
        fileInput.click();
      }
    });

    fileInput.addEventListener("change", (e) => {
      handleFilesSelected(Array.from(e.target.files));
      fileInput.value = "";
    });

    // Drag and Drop Events
    ["dragenter", "dragover"].forEach((eventName) => {
      dropzone.addEventListener(
        eventName,
        (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropzone.classList.add("dragover");
        },
        false,
      );
    });

    ["dragleave", "drop"].forEach((eventName) => {
      dropzone.addEventListener(
        eventName,
        (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropzone.classList.remove("dragover");
        },
        false,
      );
    });

    dropzone.addEventListener("drop", (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      handleFilesSelected(Array.from(files));
    });

    // Action Buttons
    clearAllBtn.addEventListener("click", clearQueue);
    convertAllBtn.addEventListener("click", startBatchConversion);
    downloadZipBtn.addEventListener("click", handleDownloadZip);
    mergeDownloadBtn.addEventListener("click", handleDownloadMergedPdf);

    // Modal Close
    closeModalBtn.addEventListener("click", closePreviewModal);
    previewModal.addEventListener("click", (e) => {
      if (e.target === previewModal) closePreviewModal();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && !previewModal.classList.contains("hidden")) {
        closePreviewModal();
      }
    });
  }

  /**
   * Handle File Selection
   */
  function handleFilesSelected(newFiles) {
    if (newFiles.length === 0) return;

    const validExtensions = [".docx", ".doc"];
    let addedCount = 0;

    newFiles.forEach((file) => {
      const ext = "." + file.name.split(".").pop().toLowerCase();
      if (!validExtensions.includes(ext)) {
        showToast(
          `Skipped "${file.name}": Only .docx and .doc formats supported.`,
          "error",
        );
        return;
      }

      if (file.size > 50 * 1024 * 1024) {
        showToast(
          `Skipped "${file.name}": File size exceeds 50MB limit.`,
          "error",
        );
        return;
      }

      const isDuplicate = filesQueue.some(
        (item) => item.file.name === file.name && item.file.size === file.size,
      );
      if (isDuplicate) {
        showToast(`"${file.name}" is already in the queue.`, "info");
        return;
      }

      const fileItem = {
        id: "file-" + Math.random().toString(36).substring(2, 9),
        file: file,
        status: "pending", // pending, converting, success, error
        progress: 0,
        resultData: null,
        errorMessage: null,
      };

      filesQueue.push(fileItem);
      addedCount++;
    });

    if (addedCount > 0) {
      showToast(
        `Added ${addedCount} document${addedCount > 1 ? "s" : ""} to queue.`,
        "success",
      );
      renderQueue();
    }
  }

  /**
   * Render Queue Items UI
   */
  function renderQueue() {
    if (filesQueue.length === 0) {
      queueSection.classList.add("hidden");
      return;
    }

    queueSection.classList.remove("hidden");
    queueCount.textContent = filesQueue.length;

    fileList.innerHTML = "";

    filesQueue.forEach((item) => {
      const fileCard = document.createElement("div");
      fileCard.className = `file-card ${item.status === "success" ? "file-card-success" : ""}`;
      fileCard.id = `card-${item.id}`;

      const formattedSize = formatBytes(item.file.size);
      const badgeHtml = getStatusBadgeHtml(item);

      fileCard.innerHTML = `
        <div class="file-card-main">
          <div class="file-info">
            <div class="file-type-icon ${item.status === "success" ? "icon-success" : ""}">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
              </svg>
            </div>
            <div class="file-details">
              <span class="file-name" title="${item.file.name}">${item.file.name}</span>
              <div class="file-meta">
                <span>${formattedSize}</span>
                ${item.pageCount ? `<span>• ${item.pageCount} page${item.pageCount > 1 ? "s" : ""}</span>` : ""}
                ${item.errorMessage ? `<span class="error-text" title="${item.errorMessage}">• Error: ${item.errorMessage}</span>` : ""}
              </div>
            </div>
          </div>

          <div class="file-actions">
            ${badgeHtml}
            ${
              item.status === "success" && item.resultData
                ? `
              <button class="btn btn-secondary btn-sm preview-btn" data-id="${item.id}" title="Preview PDF">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                  <circle cx="12" cy="12" r="3"></circle>
                </svg>
                <span>Preview</span>
              </button>
              
              <a href="${item.resultData.downloadUrl}?name=${encodeURIComponent(item.resultData.pdfFilename)}" class="btn btn-emerald btn-sm download-btn" download="${item.resultData.pdfFilename}" title="Download PDF">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                  <polyline points="7 10 12 15 17 10"></polyline>
                  <line x1="12" y1="15" x2="12" y2="3"></line>
                </svg>
                <span>Download PDF</span>
              </a>
            `
                : ""
            }

            ${
              item.status === "error" && !isConverting
                ? `
              <button class="btn btn-secondary btn-sm retry-btn" data-id="${item.id}">
                🔄 Retry
              </button>
            `
                : ""
            }

            ${
              !isConverting
                ? `
              <button class="btn btn-icon remove-btn" data-id="${item.id}" title="Remove file">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <line x1="18" y1="6" x2="6" y2="18"></line>
                  <line x1="6" y1="6" x2="18" y2="18"></line>
                </svg>
              </button>
            `
                : ""
            }
          </div>
        </div>
        ${
          item.status === "converting"
            ? `
          <div class="item-progress">
            <div class="item-progress-fill" style="width: ${item.progress}%"></div>
          </div>
        `
            : ""
        }
      `;

      fileList.appendChild(fileCard);
    });

    // Attach row event listeners
    document.querySelectorAll(".remove-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const id = e.currentTarget.getAttribute("data-id");
        removeFileFromQueue(id);
      });
    });

    document.querySelectorAll(".retry-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const id = e.currentTarget.getAttribute("data-id");
        const item = filesQueue.find((f) => f.id === id);
        if (item) {
          item.status = "pending";
          item.errorMessage = null;
          renderQueue();
        }
      });
    });

    document.querySelectorAll(".preview-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const id = e.currentTarget.getAttribute("data-id");
        const item = filesQueue.find((f) => f.id === id);
        if (item && item.resultData) {
          openPreviewModal(
            item.resultData.previewUrl,
            item.resultData.pdfFilename,
            item.resultData.downloadUrl,
          );
        }
      });
    });

    updateBatchSummary();
    updateActionButtonsVisibility();
  }

  /**
   * Badge HTML Generator
   */
  function getStatusBadgeHtml(item) {
    if (item.status === "pending") {
      return `<span class="status-badge badge-pending">⏳ Queued</span>`;
    }
    if (item.status === "converting") {
      return `<span class="status-badge badge-converting">⚙️ Converting...</span>`;
    }
    if (item.status === "success") {
      return `<span class="status-badge badge-success">✅ Converted</span>`;
    }
    if (item.status === "error") {
      return `<span class="status-badge badge-error" title="${item.errorMessage || ""}">❌ Failed</span>`;
    }
    return "";
  }

  /**
   * Update Summary and Action Buttons
   */
  function updateBatchSummary() {
    const successCount = filesQueue.filter(
      (f) => f.status === "success",
    ).length;
    const totalCount = filesQueue.length;
    if (successCount === 0) {
      batchSummary.innerHTML = `<span>${totalCount} document${totalCount > 1 ? "s" : ""} ready to convert</span>`;
    } else {
      batchSummary.innerHTML = `<span style="color: var(--emerald);">✅ ${successCount} of ${totalCount} converted to PDF</span>`;
    }
  }

  /**
   * Display Available Actions Simultaneously
   */
  function updateActionButtonsVisibility() {
    const successCount = filesQueue.filter(
      (f) => f.status === "success",
    ).length;
    const pendingCount = filesQueue.filter(
      (f) => f.status === "pending" || f.status === "error",
    ).length;

    if (isConverting) {
      convertAllBtn.classList.add("hidden");
      downloadZipBtn.classList.add("hidden");
      mergeDownloadBtn.classList.add("hidden");
      return;
    }

    // Convert Button visibility
    if (pendingCount > 0) {
      convertAllBtn.classList.remove("hidden");
      convertAllBtn.querySelector("span").textContent =
        successCount > 0 ? "Convert Remaining" : "Convert All to PDF";
    } else {
      convertAllBtn.classList.add("hidden");
    }

    // ZIP Button visibility (Show whenever 1 or more files are converted!)
    if (successCount > 0) {
      downloadZipBtn.classList.remove("hidden");
    } else {
      downloadZipBtn.classList.add("hidden");
    }

    // Merge Button visibility (Show whenever 1 or more files are converted!)
    if (successCount > 0) {
      mergeDownloadBtn.classList.remove("hidden");
    } else {
      mergeDownloadBtn.classList.add("hidden");
    }
  }

  /**
   * Remove Single File
   */
  function removeFileFromQueue(id) {
    filesQueue = filesQueue.filter((f) => f.id !== id);
    renderQueue();
  }

  /**
   * Clear Entire Queue
   */
  function clearQueue() {
    if (isConverting) return;
    filesQueue = [];
    renderQueue();
    overallProgressContainer.classList.add("hidden");
    showToast("Queue cleared.", "info");
  }

  /**
   * Start Batch Conversion
   */
  async function startBatchConversion() {
    const pendingItems = filesQueue.filter(
      (f) => f.status === "pending" || f.status === "error",
    );
    if (pendingItems.length === 0) {
      showToast("All documents are already converted!", "info");
      return;
    }

    isConverting = true;
    convertAllBtn.disabled = true;
    clearAllBtn.classList.add("hidden");

    overallProgressContainer.classList.remove("hidden");
    updateOverallProgress(0, pendingItems.length);

    const formData = new FormData();
    pendingItems.forEach((item) => {
      formData.append("files", item.file);
      item.status = "pending";
      item.progress = 0;
    });

    formData.append("quality", qualitySelect.value);
    formData.append("embedFonts", embedFontsToggle.checked ? "true" : "false");
    formData.append("clientId", clientId);

    renderQueue();

    try {
      const response = await fetch("/api/convert", {
        method: "POST",
        body: formData,
      });

      const data = await response.json();

      if (data.results && Array.isArray(data.results)) {
        data.results.forEach((res, index) => {
          // Robust matching by filename or array index
          let item = filesQueue.find((f) => f.file.name === res.originalName);
          if (!item && pendingItems[index]) {
            item = pendingItems[index];
          }

          if (item) {
            if (res.status === "success") {
              item.status = "success";
              item.progress = 100;
              item.pageCount = res.pageCount;
              item.resultData = res;
              jobs.set(res.jobId, item.id);
              item.errorMessage = null;
            } else {
              item.status = "error";
              item.errorMessage = res.error || "Conversion failed.";
            }
          }
        });

        const succeeded = data.results.filter(
          (r) => r.status === "success",
        ).length;
        updateOverallProgress(pendingItems.length, pendingItems.length);

        if (succeeded > 0) {
          showToast(
            `Converted ${succeeded} document${succeeded > 1 ? "s" : ""} successfully!`,
            "success",
          );
        } else {
          showToast("Conversion failed for uploaded documents.", "error");
        }
      } else {
        throw new Error(
          data.error || "Conversion server returned an invalid response.",
        );
      }
    } catch (err) {
      showToast(`Conversion error: ${err.message}`, "error");
      pendingItems.forEach((item) => {
        if (item.status === "converting") {
          item.status = "error";
          item.errorMessage = err.message;
        }
      });
    } finally {
      isConverting = false;
      convertAllBtn.disabled = false;
      clearAllBtn.classList.remove("hidden");
      renderQueue();
    }
  }

  function updateOverallProgress(completed, total) {
    const percent = Math.round((completed / total) * 100);
    overallProgressFill.style.width = `${percent}%`;
    overallProgressText.textContent = `Processing ${completed} of ${total} documents...`;
    overallPercentText.textContent = `${percent}%`;
  }

  /**
   * Handle Batch ZIP Download
   */
  async function handleDownloadZip() {
    const successItems = filesQueue.filter(
      (f) => f.status === "success" && f.resultData,
    );
    if (successItems.length === 0) return;

    try {
      showToast("Generating ZIP archive...", "info");
      const payload = {
        files: successItems.map((item) => ({
          fileId: item.resultData.fileId,
          displayName: item.resultData.pdfFilename,
        })),
      };

      const res = await fetch("/api/zip", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      const data = await res.json();
      if (data.success && data.downloadUrl) {
        window.location.href = data.downloadUrl;
        showToast("ZIP download started!", "success");
      } else {
        throw new Error(data.error || "Failed to create ZIP.");
      }
    } catch (err) {
      showToast(`ZIP Creation Error: ${err.message}`, "error");
    }
  }

  /**
   * Handle Merged PDF Download
   */
  async function handleDownloadMergedPdf() {
    const successItems = filesQueue.filter(
      (f) => f.status === "success" && f.resultData,
    );
    if (successItems.length === 0) return;

    try {
      showToast("Merging PDF documents...", "info");
      const payload = {
        fileIds: successItems.map((item) => item.resultData.fileId),
        customTitle: "Merged_Converted_Documents",
      };

      const res = await fetch("/api/merge", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      const data = await res.json();
      if (data.success && data.downloadUrl) {
        window.location.href = data.downloadUrl;
        showToast("Merged PDF download started!", "success");
      } else {
        throw new Error(data.error || "Failed to merge PDFs.");
      }
    } catch (err) {
      showToast(`Merge Error: ${err.message}`, "error");
    }
  }

  /**
   * Open In-App PDF Preview Modal
   */
  function openPreviewModal(previewUrl, filename, downloadUrl) {
    modalFileName.textContent = filename;
    pdfPreviewFrame.src = previewUrl;

    const fullDownloadUrl = `${downloadUrl}?name=${encodeURIComponent(filename)}`;
    modalDownloadBtn.href = fullDownloadUrl;
    modalDownloadBtn.setAttribute("download", filename);

    modalNewTabBtn.href = previewUrl;

    previewModal.classList.remove("hidden");
    document.body.style.overflow = "hidden";
  }

  function closePreviewModal() {
    previewModal.classList.add("hidden");
    pdfPreviewFrame.src = "";
    document.body.style.overflow = "";
  }

  /**
   * Toast Notification Manager
   */
  function showToast(message, type = "info") {
    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;

    let iconSvg = "";
    if (type === "success") {
      iconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#10b981" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>`;
    } else if (type === "error") {
      iconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#f43f5e" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>`;
    } else {
      iconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#6366f1" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>`;
    }

    toast.innerHTML = `${iconSvg}<span>${message}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = "0";
      toast.style.transform = "translateX(50px)";
      toast.style.transition = "all 0.3s ease";
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }

  /**
   * Helper: Format Bytes to Human Readable String
   */
  function formatBytes(bytes, decimals = 1) {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + " " + sizes[i];
  }

  function handleJobUpdate(job) {
    const item = filesQueue.find((f) => f.file.name === job.fileName);

    if (!item) return;

    switch (job.status) {
      case "queued":
        item.status = "pending";
        item.progress = 0;
        break;

      case "processing":
        item.status = "converting";
        item.progress = 50;
        item.message = job.message;
        break;

      case "completed":
        item.progress = 100;
        break;

      case "error":
        item.status = "error";
        item.errorMessage = job.message;
        break;
    }

    renderQueue();
  }
});
