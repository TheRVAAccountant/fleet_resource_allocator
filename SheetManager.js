/**
 * ===================================================================
 * SHEET MANAGER SERVICE
 * ===================================================================
 * Centralized sheet operations with caching and batch updates
 */

class SheetManager {
  constructor(spreadsheetId) {
    this.spreadsheetId = spreadsheetId || getConfig('DAILY_SUMMARY_SPREADSHEET_ID');
    this.logger = Logger.createLogger('SheetManager');
    this._spreadsheet = null;
    this._sheets = new Map();
    this._pendingUpdates = [];
  }
  
  get spreadsheet() {
    if (!this._spreadsheet) {
      const timer = this.logger.startTimer('openSpreadsheet');
      this._spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
      timer.end();
    }
    return this._spreadsheet;
  }
  
  getSheet(sheetName) {
    if (this._sheets.has(sheetName)) {
      return this._sheets.get(sheetName);
    }
    
    const timer = this.logger.startTimer(`getSheet: ${sheetName}`);
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet '${sheetName}' not found in spreadsheet ${this.spreadsheetId}`);
    }
    
    const wrapper = new SheetWrapper(sheet, this);
    this._sheets.set(sheetName, wrapper);
    timer.end();
    
    return wrapper;
  }
  
  createSheet(sheetName, options = {}) {
    const existingSheet = this.spreadsheet.getSheetByName(sheetName);
    
    if (existingSheet) {
      if (options.overwrite) {
        this.spreadsheet.deleteSheet(existingSheet);
      } else {
        throw new Error(`Sheet '${sheetName}' already exists`);
      }
    }
    
    const newSheet = this.spreadsheet.insertSheet(sheetName);
    
    if (options.headers) {
      newSheet.getRange(1, 1, 1, options.headers.length)
        .setValues([options.headers])
        .setFontWeight('bold')
        .setBackground('#E8F0FE');
    }
    
    return new SheetWrapper(newSheet, this);
  }
  
  batchUpdate(updates) {
    this._pendingUpdates.push(...updates);
  }
  
  flushUpdates() {
    if (this._pendingUpdates.length === 0) return;
    
    const timer = this.logger.startTimer('batchUpdate');
    
    // Group updates by sheet
    const updatesBySheet = new Map();
    
    this._pendingUpdates.forEach(update => {
      if (!updatesBySheet.has(update.sheet)) {
        updatesBySheet.set(update.sheet, []);
      }
      updatesBySheet.get(update.sheet).push(update);
    });
    
    // Execute updates per sheet
    updatesBySheet.forEach((updates, sheet) => {
      const values = sheet.getDataRange().getValues();
      
      updates.forEach(update => {
        const { row, col, value } = update;
        if (values[row - 1]) {
          values[row - 1][col - 1] = value;
        }
      });
      
      sheet.getDataRange().setValues(values);
    });
    
    this._pendingUpdates = [];
    timer.end();
  }
  
  sheetExists(sheetName) {
    return this.spreadsheet.getSheetByName(sheetName) !== null;
  }
  
  getAllSheetNames() {
    return this.spreadsheet.getSheets().map(sheet => sheet.getName());
  }
}

class SheetWrapper {
  constructor(sheet, manager) {
    this.sheet = sheet;
    this.manager = manager;
    this.logger = Logger.createLogger(`Sheet:${sheet.getName()}`);
  }
  
  getName() {
    return this.sheet.getName();
  }
  
  getData(range) {
    const cacheKey = `sheet_data_${this.getName()}_${range || 'all'}`;
    
    return Cache.get(cacheKey, () => {
      const timer = this.logger.startTimer('getData');
      const data = range ? 
        this.sheet.getRange(range).getValues() : 
        this.sheet.getDataRange().getValues();
      timer.end();
      return data;
    }, { ttl: 60 });
  }
  
  setData(range, values, options = {}) {
    if (options.batch) {
      this.manager.batchUpdate([{
        sheet: this.sheet,
        range,
        values
      }]);
    } else {
      const timer = this.logger.startTimer('setData');
      this.sheet.getRange(range).setValues(values);
      timer.end();
    }
    
    // Invalidate cache
    Cache.invalidate(`sheet_data_${this.getName()}_*`);
  }
  
  appendRows(rows) {
    const timer = this.logger.startTimer('appendRows');
    
    rows.forEach(row => {
      this.sheet.appendRow(row);
    });
    
    timer.end();
    Cache.invalidate(`sheet_data_${this.getName()}_*`);
  }
  
  find(searchValue, columnIndex) {
    const data = this.getData();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][columnIndex] === searchValue) {
        return {
          row: i + 1,
          data: data[i]
        };
      }
    }
    
    return null;
  }
  
  getLastRow() {
    return this.sheet.getLastRow();
  }
  
  getLastColumn() {
    return this.sheet.getLastColumn();
  }
  
  getRange(a1Notation) {
    return this.sheet.getRange(a1Notation);
  }
  
  clear(options = {}) {
    if (options.contentsOnly) {
      this.sheet.clearContents();
    } else {
      this.sheet.clear();
    }
    Cache.invalidate(`sheet_data_${this.getName()}_*`);
  }
  
  sort(columnIndex, ascending = true) {
    this.sheet.sort(columnIndex, ascending);
    Cache.invalidate(`sheet_data_${this.getName()}_*`);
  }
  
  setColumnWidths(columnWidths) {
    Object.entries(columnWidths).forEach(([col, width]) => {
      this.sheet.setColumnWidth(parseInt(col), width);
    });
  }
  
  autoResizeColumns(startColumn, numColumns) {
    this.sheet.autoResizeColumns(startColumn, numColumns);
  }
}