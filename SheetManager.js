/**
 * ===================================================================
 * SHEET MANAGER SERVICE
 * ===================================================================
 * Centralized sheet operations with caching and batch updates
 * Rewritten for Google Apps Script compatibility (no ES6 classes)
 */

/**
 * SheetManager constructor
 * @param {string} spreadsheetId - Spreadsheet ID (optional)
 */
function SheetManager(spreadsheetId) {
  this.spreadsheetId = spreadsheetId || getConfig('DAILY_SUMMARY_SPREADSHEET_ID');
  this.logger = createLogger('SheetManager');
  this._spreadsheet = null;
  this._sheets = {};
  this._pendingUpdates = [];
}

/**
 * Get spreadsheet (lazy loading)
 * @return {Spreadsheet} Google Sheets spreadsheet object
 */
SheetManager.prototype.getSpreadsheet = function() {
  if (!this._spreadsheet) {
    var timer = this.logger.startTimer('openSpreadsheet');
    this._spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
    timer.end();
  }
  return this._spreadsheet;
};

/**
 * Get a sheet by name with caching
 * @param {string} sheetName - Name of the sheet
 * @return {SheetWrapper} Wrapped sheet object
 */
SheetManager.prototype.getSheet = function(sheetName) {
  if (this._sheets[sheetName]) {
    return this._sheets[sheetName];
  }
  
  var timer = this.logger.startTimer('getSheet: ' + sheetName);
  var sheet = this.getSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("Sheet '" + sheetName + "' not found in spreadsheet " + this.spreadsheetId);
  }
  
  var wrapper = new SheetWrapper(sheet, this);
  this._sheets[sheetName] = wrapper;
  timer.end();
  
  return wrapper;
};

/**
 * Create a new sheet
 * @param {string} sheetName - Name for the new sheet
 * @param {Object} options - Creation options
 * @return {SheetWrapper} Wrapped sheet object
 */
SheetManager.prototype.createSheet = function(sheetName, options) {
  options = options || {};
  var existingSheet = this.getSpreadsheet().getSheetByName(sheetName);
  
  if (existingSheet) {
    if (options.overwrite) {
      this.getSpreadsheet().deleteSheet(existingSheet);
    } else {
      return this.getSheet(sheetName);
    }
  }
  
  var timer = this.logger.startTimer('createSheet: ' + sheetName);
  var newSheet = this.getSpreadsheet().insertSheet(sheetName);
  
  if (options.headers && options.headers.length > 0) {
    newSheet.getRange(1, 1, 1, options.headers.length).setValues([options.headers]);
    if (typeof formatHeaderRow === 'function') {
      formatHeaderRow(newSheet, 1, options.headers.length);
    }
  }
  
  var wrapper = new SheetWrapper(newSheet, this);
  this._sheets[sheetName] = wrapper;
  timer.end();
  
  return wrapper;
};

/**
 * Queue a batch update
 * @param {Object} update - Update object
 */
SheetManager.prototype.queueUpdate = function(update) {
  this._pendingUpdates.push(update);
};

/**
 * Execute all pending updates
 */
SheetManager.prototype.flush = function() {
  if (this._pendingUpdates.length === 0) return;
  
  var timer = this.logger.startTimer('batchUpdate');
  
  try {
    // Group updates by sheet
    var updatesBySheet = {};
    
    for (var i = 0; i < this._pendingUpdates.length; i++) {
      var update = this._pendingUpdates[i];
      var sheetName = update.sheet.getName();
      
      if (!updatesBySheet[sheetName]) {
        updatesBySheet[sheetName] = [];
      }
      updatesBySheet[sheetName].push(update);
    }
    
    // Execute updates per sheet
    for (var sheetName in updatesBySheet) {
      if (updatesBySheet.hasOwnProperty(sheetName)) {
        this._executeBatchUpdate(updatesBySheet[sheetName]);
      }
    }
    
    this._pendingUpdates = [];
    timer.end();
    
  } catch (error) {
    this.logger.error('Batch update failed', { error: error.message });
    throw error;
  }
};

/**
 * Execute batch update for a single sheet
 * @param {Array} updates - Array of updates for one sheet
 * @private
 */
SheetManager.prototype._executeBatchUpdate = function(updates) {
  if (updates.length === 0) return;
  
  var sheet = updates[0].sheet;
  var rangeList = [];
  var values = [];
  
  for (var i = 0; i < updates.length; i++) {
    var update = updates[i];
    rangeList.push(update.range);
    values.push(update.values);
  }
  
  // Use batch update if available
  if (sheet.getRangeList) {
    var ranges = sheet.getRangeList(rangeList);
    ranges.setValues(values);
  } else {
    // Fallback to individual updates
    for (var j = 0; j < updates.length; j++) {
      sheet.getRange(updates[j].range).setValues(updates[j].values);
    }
  }
};

/**
 * SheetWrapper constructor
 * @param {Sheet} sheet - Google Sheets sheet object
 * @param {SheetManager} manager - Parent SheetManager
 */
function SheetWrapper(sheet, manager) {
  this._sheet = sheet;
  this.manager = manager;
  this._dataCache = null;
  this._lastRow = null;
  this._lastColumn = null;
}

/**
 * Get all data from the sheet
 * @param {boolean} useCache - Whether to use cached data
 * @return {Array} 2D array of sheet data
 */
SheetWrapper.prototype.getData = function(useCache) {
  if (useCache === undefined) useCache = true;
  
  if (useCache && this._dataCache) {
    return this._dataCache;
  }
  
  var timer = this.manager.logger.startTimer('getData: ' + this._sheet.getName());
  
  if (this.getLastRow() === 0) {
    this._dataCache = [];
    timer.end();
    return [];
  }
  
  var data = this._sheet.getRange(1, 1, this.getLastRow(), this.getLastColumn()).getValues();
  this._dataCache = data;
  timer.end();
  
  return data;
};

/**
 * Get last row with data
 * @return {number} Last row number
 */
SheetWrapper.prototype.getLastRow = function() {
  if (this._lastRow === null) {
    this._lastRow = this._sheet.getLastRow();
  }
  return this._lastRow;
};

/**
 * Get last column with data
 * @return {number} Last column number
 */
SheetWrapper.prototype.getLastColumn = function() {
  if (this._lastColumn === null) {
    this._lastColumn = this._sheet.getLastColumn();
  }
  return this._lastColumn;
};

/**
 * Clear the cache
 */
SheetWrapper.prototype.clearCache = function() {
  this._dataCache = null;
  this._lastRow = null;
  this._lastColumn = null;
};

/**
 * Append a row to the sheet
 * @param {Array} rowData - Data to append
 */
SheetWrapper.prototype.appendRow = function(rowData) {
  this._sheet.appendRow(rowData);
  this.clearCache();
};

/**
 * Set values in a range
 * @param {string} range - A1 notation range
 * @param {Array} values - 2D array of values
 * @param {boolean} batch - Whether to batch this update
 */
SheetWrapper.prototype.setValues = function(range, values, batch) {
  if (batch) {
    this.manager.queueUpdate({
      sheet: this._sheet,
      range: range,
      values: values
    });
  } else {
    this._sheet.getRange(range).setValues(values);
    this.clearCache();
  }
};

/**
 * Auto-resize columns
 * @param {number} startColumn - Starting column
 * @param {number} numColumns - Number of columns
 */
SheetWrapper.prototype.autoResizeColumns = function(startColumn, numColumns) {
  for (var i = 0; i < numColumns; i++) {
    this._sheet.autoResizeColumn(startColumn + i);
  }
};

// Expose the sheet property for compatibility
Object.defineProperty(SheetWrapper.prototype, 'sheet', {
  get: function() {
    return this._sheet;
  },
  enumerable: true
});