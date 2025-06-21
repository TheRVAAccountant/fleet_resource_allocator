/**
 * ===================================================================
 * ENHANCED LOGGER SERVICE
 * ===================================================================
 * Modern logging system with multiple outputs and structured data
 */

class Logger {
  static levels = {
    DEBUG: 0,
    INFO: 1,
    WARN: 2,
    ERROR: 3,
    CRITICAL: 4
  };
  
  constructor(component) {
    this.component = component;
    this.startTime = new Date().getTime();
  }
  
  static createLogger(component) {
    return new Logger(component);
  }
  
  log(level, message, context = {}) {
    const config = getConfig('LOGGING') || {};
    const currentLevel = Logger.levels[config.level || 'INFO'];
    
    if (Logger.levels[level] < currentLevel) {
      return;
    }
    
    const entry = {
      timestamp: new Date().toISOString(),
      level,
      component: this.component,
      message,
      context,
      user: Session.getActiveUser().getEmail(),
      executionTime: new Date().getTime() - this.startTime
    };
    
    // Console logging
    console[level.toLowerCase()](JSON.stringify(entry, null, 2));
    
    // Persist critical errors
    if (Logger.levels[level] >= Logger.levels.ERROR) {
      this.persistToSheet(entry);
    }
  }
  
  persistToSheet(entry) {
    try {
      const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
      let errorLog = ss.getSheetByName('Error Log');
      
      if (!errorLog) {
        errorLog = ss.insertSheet('Error Log');
        errorLog.getRange(1, 1, 1, 7).setValues([[
          'Timestamp', 'Level', 'Component', 'Message', 'Context', 'User', 'Execution Time'
        ]]);
        formatHeaderRow(errorLog, 1, 7);
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
      console.error('Failed to persist log:', e);
    }
  }
  
  debug(message, context) { this.log('DEBUG', message, context); }
  info(message, context) { this.log('INFO', message, context); }
  warn(message, context) { this.log('WARN', message, context); }
  error(message, context) { this.log('ERROR', message, context); }
  critical(message, context) { this.log('CRITICAL', message, context); }
  
  startTimer(operation) {
    const timerStart = new Date().getTime();
    return {
      end: () => {
        const duration = new Date().getTime() - timerStart;
        this.debug(`${operation} completed in ${duration}ms`);
        return duration;
      }
    };
  }
}

/**
 * Legacy support - Create logger using old AppLogger interface
 */
function AppLogger(componentName) {
  const logger = Logger.createLogger(componentName);
  
  // Map old methods to new
  this.debug = (message, data) => logger.debug(message, data);
  this.info = (message, data) => logger.info(message, data);
  this.warn = (message, data) => logger.warn(message, data);
  this.error = (message, data) => logger.error(message, data);
  this.critical = (message, data) => logger.critical(message, data);
  this.startTimer = (operation) => logger.startTimer(operation);
}

/**
 * Create a logger instance (backward compatibility)
 * @param {string} component - Component name
 * @return {Logger} Logger instance
 */
function createLogger(component) {
  return Logger.createLogger(component);
}

/**
 * Global logger for general use
 */
var globalLogger = createLogger('System');