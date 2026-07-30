const EventEmitter = require("events");

class ConverterEvents extends EventEmitter {}

module.exports = new ConverterEvents();