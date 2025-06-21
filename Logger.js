/**
 * ===================================================================
 * ENHANCED LOGGER SERVICE
 * ===================================================================
 * Modern logging system with multiple outputs and structured data
 * Rewritten for Google Apps Script compatibility (no ES6 classes)
 */

// Logger levels
var LoggerLevels = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
  CRITICAL: 4
};

/**
 * Logger constructor function
 * @param {string} component - Component name
 */
function Logger(component) {
  this.component = component || 'System';
  this.startTime = new Date().getTime();
}

/**
 * Static method to create logger
 * @param {string} component - Component name
 * @return {Logger} Logger instance
 */
Logger.createLogger = function(component) {
  return new Logger(component);
};

/**
 * Log a message at the specified level
 * @param {string} level - Log level
 * @param {string} message - Message to log
 * @param {Object} context - Additional context
 */
Logger.prototype.log = function(level, message, context) {
  try {
    context = context || {};
    var config = getConfig('LOGGING') || {};
    var currentLevel = LoggerLevels[config.level || 'INFO'];
    
    if (LoggerLevels[level] < currentLevel) {
      return;
    }
    
    var entry = {
      timestamp: new Date().toISOString(),
      level: level,
      component: this.component,
      message: message,
      context: context,
      user: Session.getActiveUser().getEmail(),
      executionTime: new Date().getTime() - this.startTime
    };
    
    // Console logging - Google Apps Script doesn't support console.debug, warn, etc.
    console.log(JSON.stringify(entry, null, 2));
    
    // Persist critical errors
    if (LoggerLevels[level] >= LoggerLevels.ERROR) {
      this.persistToSheet(entry);
    }
  } catch (e) {
    // Fail silently to avoid breaking the application
    console.log('Logger error: ' + e.toString());
  }
};

/**
 * Persist log entry to sheet
 * @param {Object} entry - Log entry
 */
Logger.prototype.persistToSheet = function(entry) {
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var errorLog = ss.getSheetByName('Error Log');
    
    if (!errorLog) {
      errorLog = ss.insertSheet('Error Log');
      errorLog.getRange(1, 1, 1, 7).setValues([[
        'Timestamp', 'Level', 'Component', 'Message', 'Context', 'User', 'Execution Time'
      ]]);
      // Format header if function exists
      if (typeof formatHeaderRow === 'function') {
        formatHeaderRow(errorLog, 1, 7);
      }
    }
    
    errorLog.appendRow([
      entry.timestamp,
      entry.level,
      entry.component,
      entry.message,
      JSON.stringify(entry.context),
      entry.user,
      entry.executionTime + 'ms'
    ]);
  } catch (e) {
    console.log('Failed to persist log: ' + e.toString());
  }
};

// Add convenience methods
Logger.prototype.debug = function(message, context) {
  this.log('DEBUG', message, context);
};

Logger.prototype.info = function(message, context) {
  this.log('INFO', message, context);
};

Logger.prototype.warn = function(message, context) {
  this.log('WARN', message, context);
};

Logger.prototype.error = function(message, context) {
  this.log('ERROR', message, context);
};

Logger.prototype.critical = function(message, context) {
  this.log('CRITICAL', message, context);
};

/**
 * Start a timer for performance tracking
 * @param {string} operation - Operation name
 * @return {Object} Timer object with end method
 */
Logger.prototype.startTimer = function(operation) {
  var timerStart = new Date().getTime();
  var self = this;
  
  return {
    end: function() {
      var duration = new Date().getTime() - timerStart;
      self.debug(operation + ' completed in ' + duration + 'ms');
      return duration;
    }
  };
};

/**
 * Legacy support - Create logger using old AppLogger interface
 */
function AppLogger(componentName) {
  var logger = new Logger(componentName);
  
  // Map old methods to new
  this.debug = function(message, data) { logger.debug(message, data); };
  this.info = function(message, data) { logger.info(message, data); };
  this.warn = function(message, data) { logger.warn(message, data); };
  this.error = function(message, data) { logger.error(message, data); };
  this.critical = function(message, data) { logger.critical(message, data); };
  this.startTimer = function(operation) { return logger.startTimer(operation); };
}

/**
 * Create a logger instance
 * @param {string} component - Component name
 * @return {Logger} Logger instance
 */
function createLogger(component) {
  return new Logger(component);
}

/**
 * Global logger for general use
 */
var globalLogger = new Logger('System');