/**
 * Simple logging utility with levels
 */

const LogLevel = {
  ERROR: 0,
  WARN: 1,
  INFO: 2,
  DEBUG: 3
};

class Logger {
  constructor(level = LogLevel.INFO) {
    this.level = level;
  }

  setLevel(level) {
    this.level = level;
  }

  error(...args) {
    if (this.level >= LogLevel.ERROR) {
      console.error('[ERROR]', ...args);
    }
  }

  warn(...args) {
    if (this.level >= LogLevel.WARN) {
      console.warn('[WARN]', ...args);
    }
  }

  info(...args) {
    if (this.level >= LogLevel.INFO) {
      console.log('[INFO]', ...args);
    }
  }

  debug(...args) {
    if (this.level >= LogLevel.DEBUG) {
      console.log('[DEBUG]', ...args);
    }
  }

  progress(message) {
    if (this.level >= LogLevel.INFO) {
      process.stdout.write(`\r${message}`);
    }
  }

  clearProgress() {
    if (this.level >= LogLevel.INFO) {
      process.stdout.write('\r\x1b[K');
    }
  }
}

// Global logger instance
const logger = new Logger();

module.exports = {
  Logger,
  LogLevel,
  logger
};
