/**
 * ===================================================================
 * UNIFIED LOGGER - LOADS FIRST (UNDERSCORE PREFIX)
 * ===================================================================
 * Single source of truth for Logger implementation
 * Compatible with Google Apps Script (no ES6 classes)
 */

// Global logger configuration
var LOGGER_CONFIG = {
  LEVEL: 'INFO',
  LOG_TO_SHEET: true,
  SHEET_NAME: 'Error Log',
  MAX_SHEET_ROWS: 1000
};

// Logger levels
var LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
  CRITICAL: 4
};

/**
 * Create a logger instance - main factory function
 * @param {string} component - Component name
 * @return {Object} Logger instance
 */
function createLogger(component) {
  var logger = {
    component: component || 'System',
    startTime: new Date().getTime(),
    
    /**
     * Generic log method
     */
    log: function(level, message, context) {
      try {
        // Check log level
        var currentLevel = LOG_LEVELS[LOGGER_CONFIG.LEVEL] || LOG_LEVELS.INFO;
        if (LOG_LEVELS[level] < currentLevel) {
          return;
        }
        
        // Format log entry
        var timestamp = new Date().toISOString();
        var logEntry = '[' + timestamp + '] [' + level + '] [' + this.component + '] ' + message;
        
        // Log to console
        console.log(logEntry);
        if (context) {
          console.log('Context:', JSON.stringify(context));
        }
        
        // For errors, log to sheet
        if (LOGGER_CONFIG.LOG_TO_SHEET && (level === 'ERROR' || level === 'CRITICAL')) {
          this._logToSheet(timestamp, level, message, context);
        }
      } catch (e) {
        // Fail silently - logging should never break the app
        console.log('Logger error: ' + e.toString());
      }
    },
    
    /**
     * Log to Error Log sheet (private method)
     */
    _logToSheet: function(timestamp, level, message, context) {
      try {
        // Only proceed if we have getConfig function
        if (typeof getConfig !== 'function') return;
        
        var ssId = getConfig('DAILY_SUMMARY_SPREADSHEET_ID');
        if (!ssId) return;
        
        var ss = SpreadsheetApp.openById(ssId);
        var sheet = ss.getSheetByName(LOGGER_CONFIG.SHEET_NAME);
        
        // Create sheet if it doesn't exist
        if (!sheet) {
          sheet = ss.insertSheet(LOGGER_CONFIG.SHEET_NAME);
          sheet.getRange(1, 1, 1, 5).setValues([[
            'Timestamp', 'Level', 'Component', 'Message', 'Context'
          ]]);
          
          // Format header
          var headerRange = sheet.getRange(1, 1, 1, 5);
          headerRange.setFontWeight('bold');
          headerRange.setBackground('#f0f0f0');
        }
        
        // Add log entry
        sheet.appendRow([
          timestamp,
          level,
          this.component,
          message,
          context ? JSON.stringify(context) : ''
        ]);
        
        // Trim old entries if sheet is too large
        if (sheet.getLastRow() > LOGGER_CONFIG.MAX_SHEET_ROWS) {
          sheet.deleteRows(2, 100); // Delete 100 oldest entries
        }
      } catch (e) {
        // Fail silently
      }
    },
    
    // Convenience methods
    debug: function(message, context) {
      this.log('DEBUG', message, context);
    },
    
    info: function(message, context) {
      this.log('INFO', message, context);
    },
    
    warn: function(message, context) {
      this.log('WARN', message, context);
    },
    
    error: function(message, context) {
      this.log('ERROR', message, context);
    },
    
    critical: function(message, context) {
      this.log('CRITICAL', message, context);
    },
    
    /**
     * Start a timer for performance tracking
     */
    startTimer: function(operation) {
      var timerStart = new Date().getTime();
      var self = this;
      
      return {
        end: function() {
          var duration = new Date().getTime() - timerStart;
          self.debug(operation + ' completed in ' + duration + 'ms');
          return duration;
        }
      };
    }
  };
  
  return logger;
}

/**
 * Logger constructor for compatibility
 * @param {string} component - Component name
 */
function Logger(component) {
  return createLogger(component);
}

// Add static method for compatibility
Logger.createLogger = createLogger;

/**
 * AppLogger constructor for legacy support
 * @param {string} component - Component name
 */
function AppLogger(component) {
  return createLogger(component);
}

// Create global logger instance
var globalLogger = createLogger('System');

/**
 * Test logger functionality
 */
function testLogger() {
  try {
    var testLog = createLogger('LoggerTest');
    testLog.info('Testing logger functionality');
    testLog.debug('Debug message', { test: true });
    testLog.error('Test error', { error: 'This is a test' });
    
    SpreadsheetApp.getUi().alert(
      'Logger Test',
      'Logger is working correctly! Check the console and Error Log sheet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Logger Test Failed',
      'Error: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Clear error log sheet
 */
function clearErrorLog() {
  try {
    var ssId = getConfig('DAILY_SUMMARY_SPREADSHEET_ID');
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName(LOGGER_CONFIG.SHEET_NAME);
    
    if (sheet && sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
      SpreadsheetApp.getUi().alert('Error log cleared successfully.');
    } else {
      SpreadsheetApp.getUi().alert('Error log is already empty.');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error clearing log: ' + error.toString());
  }
}