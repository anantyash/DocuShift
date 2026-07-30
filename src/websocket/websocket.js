const WebSocket = require("ws");
const crypto = require("crypto");
const events = require("../events/converter.events");

class WebSocketManager {
  constructor() {
    this.clients = new Map();
    this.wss = null;
  }

  initialize(server) {
    this.wss = new WebSocket.Server({
      server,
    });
    this.registerEventListeners();

    console.log("[WebSocket] Initializing...");

    this.wss.on("connection", (socket) => {
      const clientId = crypto.randomUUID();

      console.log(`[WebSocket] Client Connected : ${clientId}`);

      this.clients.set(clientId, socket);

      socket.send(
        JSON.stringify({
          type: "connected",
          clientId,
        }),
      );

      socket.on("message", (message) => {
        this.handleMessage(clientId, message);
      });

      socket.on("close", () => {
        console.log(`[WebSocket] Client Disconnected : ${clientId}`);

        this.clients.delete(clientId);
      });

      socket.on("error", (err) => {
        console.log("[WebSocket] Error", err.message);
      });
    });
  }

  registerEventListeners() {
    events.on("job-update", (job) => {
      this.send(job.clientId, {
        type: "job",
        jobId: job.jobId,
        fileId: job.fileId,
        status: job.status,
        message: job.message,
        queuePosition: job.queuePosition,
        downloadUrl: job.downloadUrl,
        previewUrl: job.previewUrl,
      });
    });
  }

  handleMessage(clientId, rawMessage) {
    try {
      const message = JSON.parse(rawMessage);

      switch (message.type) {
        case "ping":
          this.send(clientId, {
            type: "pong",
          });
          break;

        default:
          console.log("[WebSocket] Unknown message", message);
      }
    } catch (err) {
      console.log(err.message);
    }
  }

  send(clientId, payload) {
    const socket = this.clients.get(clientId);

    if (!socket) return false;
    if (socket.readyState !== WebSocket.OPEN) return false;

    socket.send(JSON.stringify(payload));
    return true;
  }

  broadcast(payload) {
    for (const socket of this.clients.values()) {
      if (socket.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify(payload));
      }
    }
  }

  getClient(clientId) {
    return this.clients.get(clientId);
  }

  getClientCount() {
    return this.clients.size;
  }
}

module.exports = new WebSocketManager();
