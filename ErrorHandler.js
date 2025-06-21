/**
 * ===================================================================
 * ERROR HANDLER SERVICE
 * ===================================================================
 * Centralized error handling with user-friendly messages
 */

/**
 * Custom error types for better error handling
 */
function ValidationError(message, field) {
  this.name = 'ValidationError';
  this.message = message;
  this.field = field;
  this.stack = (new Error()).stack;
}
ValidationError.prototype = Object.create(Error.prototype);
ValidationError.prototype.constructor = ValidationError;

function FileProcessingError(message, fileName) {
  this.name = 'FileProcessingError';
  this.message = message;
  this.fileName = fileName;
  this.stack = (new Error()).stack;
}
FileProcessingError.prototype = Object.create(Error.prototype);
FileProcessingError.prototype.constructor = FileProcessingError;

function AllocationError(message, details) {
  this.name = 'AllocationError';
  this.message = message;
  this.details = details;
  this.stack = (new Error()).stack;
}
AllocationError.prototype = Object.create(Error.prototype);
AllocationError.prototype.constructor = AllocationError;

/**
 * User-friendly error messages
 */
var ErrorMessages = {
  FILE_NOT_FOUND: 'The file could not be found. Please check if it exists and try again.',
  INVALID_FILE_FORMAT: 'The file format is not supported. Please upload an Excel (.xlsx) file.',
  MISSING_REQUIRED_SHEET: 'Required sheet "{sheet}" not found in the file.',
  MISSING_REQUIRED_COLUMN: 'Required column "{column}" not found in the sheet.',
  NO_VEHICLES_AVAILABLE: 'No operational vehicles are available for allocation.',
  ALLOCATION_FAILED: 'Vehicle allocation failed. Please check the data and try again.',
  PERMISSION_DENIED: 'You do not have permission to perform this action.',
  NETWORK_ERROR: 'Network error occurred. Please check your connection and try again.',
  UNKNOWN_ERROR: 'An unexpected error occurred. Please try again or contact support.'
};

/**
 * Error handler constructor
 */
function ErrorHandler() {
  this.logger = createLogger('ErrorHandler');
}
/**
 * Handle error and return user-friendly message
 * @param {Error} error - The error to handle
 * @param {Object} context - Additional context
 * @return {string} User-friendly error message
 */
ErrorHandler.prototype.handle = function(error, context) {
  context = context || {};
  // Log the full error
  this.logger.error(error.message, {
    stack: error.stack,
    context: context,
    errorType: error.constructor.name
  });
  
  // Return user-friendly message
  return this.getUserFriendlyMessage(error, context);
};

/**
 * Get user-friendly error message
 * @param {Error} error - The error
 * @param {Object} context - Additional context
 * @return {string} User-friendly message
 */
ErrorHandler.prototype.getUserFriendlyMessage = function(error, context) {
  if (error instanceof ValidationError) {
    return 'Validation Error: ' + error.message;
  }
  
  if (error instanceof FileProcessingError) {
    return 'File Processing Error: ' + error.message;
  }
  
  if (error instanceof AllocationError) {
    return 'Allocation Error: ' + error.message;
  }
  
  // Check for specific error patterns
  if (error.message.includes('not found')) {
    return ErrorMessages.FILE_NOT_FOUND;
  }
  
  if (error.message.includes('permission')) {
    return ErrorMessages.PERMISSION_DENIED;
  }
  
  if (error.message.includes('network') || error.message.includes('fetch')) {
    return ErrorMessages.NETWORK_ERROR;
  }
  
  // Default message
  return ErrorMessages.UNKNOWN_ERROR;
};

/**
 * Show error alert to user
 * @param {Error} error - The error
 * @param {Object} context - Additional context
 */
ErrorHandler.prototype.showAlert = function(error, context) {
  context = context || {};
  var message = this.handle(error, context);
  SpreadsheetApp.getUi().alert('Error', message, SpreadsheetApp.getUi().ButtonSet.OK);
};

/**
 * Create a safe wrapper for functions
 * @param {Function} fn - Function to wrap
 * @param {Object} context - Context for error handling
 * @return {Function} Wrapped function
 */
ErrorHandler.prototype.createSafeFunction = function(fn, context) {
  var self = this;
  context = context || {};
  return function() {
    try {
      return fn.apply(this, arguments);
    } catch (error) {
      var fullContext = {};
      for (var key in context) {
        fullContext[key] = context[key];
      }
      fullContext['function'] = fn.name;
      self.showAlert(error, fullContext);
      throw error; // Re-throw for debugging
    }
  };
};

/**
 * Global error handler instance
 */
var errorHandler = new ErrorHandler();

/**
 * Decorator for error handling
 * @param {Function} target - Function to decorate
 * @param {string} componentName - Component name for logging
 * @return {Function} Decorated function
 */
function withErrorHandling(target, componentName) {
  return function() {
    var logger = createLogger(componentName);
    var timer = logger.startTimer(target.name);
    var args = Array.prototype.slice.call(arguments);
    
    try {
      logger.debug('Starting ' + target.name, { args: args });
      var result = target.apply(this, args);
      timer.end();
      return result;
    } catch (error) {
      logger.error('Error in ' + target.name, {
        error: error.message,
        stack: error.stack,
        args: args
      });
      errorHandler.showAlert(error, { component: componentName, 'function': target.name });
      throw error;
    }
  };
}