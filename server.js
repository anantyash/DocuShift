const express = require("express");
const cors = require("cors");
const path = require("path");
const http = require("http");
const { exec } = require("child_process");
const config = require("./src/config/config");
const apiRoutes = require("./src/routes/api.routes");
const cleanupService = require("./src/services/cleanup.service");
const websocket = require("./src/websocket/websocket");

const app = express();
const server = http.createServer(app);

// Security & Cross-Origin Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static frontend files
app.use(express.static(path.join(__dirname, "public")));

// Mount API Routes
app.use("/api", apiRoutes);

// Single Page Application Fallback
app.get("*", (req, res) => {
  if (req.path.startsWith("/api")) {
    return res
      .status(404)
      .json({ success: false, error: "API Endpoint Not Found" });
  }
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// Global Error Handler
app.use((err, req, res, next) => {
  console.error("[ServerError]", err);
  res.status(500).json({
    success: false,
    error: err.message || "An unexpected server error occurred.",
  });
});

// Start cleanup service for temporary files
cleanupService.start();

websocket.initialize(server);

// Start HTTP Server
server.listen(config.PORT, () => {
  console.log(`=======================================================`);
  console.log(`🚀 Word to PDF Production Converter Server Ready`);
  console.log(`📡 URL: http://localhost:${config.PORT}`);
  console.log(`=======================================================`);

  // Open default browser automatically when starting
  if (
    process.env.NODE_ENV === "production" ||
    process.env.AUTO_OPEN === "true"
  ) {
    exec("start http://localhost:3000");
  }
});

// Graceful Shutdown
process.on("SIGTERM", gracefulShutdown);
process.on("SIGINT", gracefulShutdown);

function gracefulShutdown() {
  console.log("\n[Server] Shutting down gracefully...");
  cleanupService.stop();
  server.close(() => {
    console.log("[Server] Closed all connections.");
    console.log(`[WebSocket] Closing ${websocket.getClientCount()} clients...`);
    process.exit(0);
  });
}
