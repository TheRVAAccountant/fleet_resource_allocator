/**
 * ===================================================================
 * DATA EXPORT SERVICE
 * ===================================================================
 * Handles data export functionality including CSV generation,
 * file compression, and Drive storage
 */

/**
 * Export all data from the system
 * Creates CSV files for all sheets and packages them with timestamp
 */
function exportAllData() {
  const logger = Logger.createLogger('DataExportService');
  const timer = logger.startTimer('exportAllData');
  
  try {
    UIHelpers.showLoading('Preparing data export...');
    
    const exportDate = new Date();
    const timestamp = Utilities.formatDate(exportDate, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const exportFolderName = `Fleet_Data_Export_${timestamp}`;
    
    // Create export folder in Drive
    const rootFolder = DriveApp.getRootFolder();
    const exportFolder = rootFolder.createFolder(exportFolderName);
    
    logger.info('Created export folder', { folderName: exportFolderName });
    
    // Get all sheets to export
    const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    const sheets = ss.getSheets();
    
    const exportSummary = {
      totalSheets: sheets.length,
      exportedSheets: [],
      errors: [],
      exportDate: formatDate(exportDate),
      exportTime: Utilities.formatDate(exportDate, Session.getScriptTimeZone(), 'HH:mm:ss')
    };
    
    // Track progress
    const progressTracker = new ProgressTracker('Exporting Data', sheets.length);
    
    sheets.forEach((sheet, index) => {
      try {
        const sheetName = sheet.getName();
        progressTracker.update(index + 1, `Exporting: ${sheetName}`);
        
        // Skip empty sheets
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        
        if (lastRow === 0 || lastCol === 0) {
          logger.debug(`Skipping empty sheet: ${sheetName}`);
          return;
        }
        
        // Get sheet data
        const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
        
        // Convert to CSV
        const csv = convertToCSV(data);
        
        // Create file in export folder
        const blob = Utilities.newBlob(csv, 'text/csv', `${sanitizeFilename(sheetName)}.csv`);
        exportFolder.createFile(blob);
        
        exportSummary.exportedSheets.push({
          name: sheetName,
          rows: lastRow,
          columns: lastCol,
          size: blob.getBytes().length
        });
        
      } catch (error) {
        logger.error(`Failed to export sheet: ${sheet.getName()}`, { error: error.message });
        exportSummary.errors.push({
          sheet: sheet.getName(),
          error: error.message
        });
      }
    });
    
    // Create export summary file
    const summaryContent = createExportSummary(exportSummary);
    const summaryBlob = Utilities.newBlob(summaryContent, 'text/plain', 'export_summary.txt');
    exportFolder.createFile(summaryBlob);
    
    // Create metadata file
    const metadata = {
      exportDate: exportDate.toISOString(),
      exportedBy: Session.getActiveUser().getEmail(),
      spreadsheetId: getConfig('DAILY_SUMMARY_SPREADSHEET_ID'),
      spreadsheetName: ss.getName(),
      totalSheets: exportSummary.totalSheets,
      exportedSheets: exportSummary.exportedSheets.length,
      errors: exportSummary.errors.length
    };
    
    const metadataBlob = Utilities.newBlob(
      JSON.stringify(metadata, null, 2), 
      'application/json', 
      'metadata.json'
    );
    exportFolder.createFile(metadataBlob);
    
    progressTracker.complete('Export completed successfully!');
    
    timer.end();
    
    logger.info('Data export completed', {
      folderName: exportFolderName,
      sheetsExported: exportSummary.exportedSheets.length,
      errors: exportSummary.errors.length
    });
    
    // Show success with folder link
    const folderUrl = exportFolder.getUrl();
    const html = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: 'Google Sans', Arial, sans-serif;">
        <h3 style="color: #34A853;">Export Completed Successfully!</h3>
        <p><strong>Export Folder:</strong> ${exportFolderName}</p>
        <p><strong>Sheets Exported:</strong> ${exportSummary.exportedSheets.length} of ${exportSummary.totalSheets}</p>
        ${exportSummary.errors.length > 0 ? 
          `<p style="color: #EA4335;"><strong>Errors:</strong> ${exportSummary.errors.length} sheets failed</p>` : 
          ''
        }
        <p style="margin-top: 20px;">
          <a href="${folderUrl}" target="_blank" style="
            background: #1a73e8;
            color: white;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 4px;
            display: inline-block;
          ">Open Export Folder</a>
        </p>
      </div>
    `)
    .setWidth(450)
    .setHeight(250);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Export Complete');
    
    // Record action
    SmartDefaults.recordAction('data_exported', {
      folderName: exportFolderName,
      sheetsCount: exportSummary.exportedSheets.length
    });
    
  } catch (error) {
    logger.error('Failed to export data', { error: error.message });
    google.script.host.close();
    UIHelpers.showError(error);
  }
}

/**
 * Convert 2D array to CSV format
 * @param {Array[]} data - 2D array of data
 * @return {string} CSV formatted string
 */
function convertToCSV(data) {
  return data.map(row => 
    row.map(cell => {
      // Handle different data types
      if (cell === null || cell === undefined) {
        return '';
      }
      
      // Convert dates to ISO format
      if (cell instanceof Date) {
        return cell.toISOString();
      }
      
      // Convert to string and escape quotes
      const cellStr = cell.toString();
      
      // If contains comma, newline, or quotes, wrap in quotes
      if (cellStr.includes(',') || cellStr.includes('\n') || cellStr.includes('"')) {
        return '"' + cellStr.replace(/"/g, '""') + '"';
      }
      
      return cellStr;
    }).join(',')
  ).join('\n');
}

/**
 * Sanitize filename for safe file creation
 * @param {string} filename - Original filename
 * @return {string} Sanitized filename
 */
function sanitizeFilename(filename) {
  return filename
    .replace(/[\/\\:*?"<>|]/g, '_') // Replace invalid characters
    .replace(/\s+/g, '_') // Replace spaces with underscores
    .replace(/_+/g, '_') // Remove multiple underscores
    .substring(0, 100); // Limit length
}

/**
 * Create export summary report
 * @param {Object} summary - Export summary data
 * @return {string} Formatted summary text
 */
function createExportSummary(summary) {
  let content = 'FLEET DATA EXPORT SUMMARY\n';
  content += '========================\n\n';
  content += `Export Date: ${summary.exportDate}\n`;
  content += `Export Time: ${summary.exportTime}\n`;
  content += `Exported By: ${Session.getActiveUser().getEmail()}\n\n`;
  
  content += 'EXPORT STATISTICS\n';
  content += '-----------------\n';
  content += `Total Sheets: ${summary.totalSheets}\n`;
  content += `Successfully Exported: ${summary.exportedSheets.length}\n`;
  content += `Failed: ${summary.errors.length}\n\n`;
  
  content += 'EXPORTED SHEETS\n';
  content += '---------------\n';
  summary.exportedSheets.forEach(sheet => {
    content += `${sheet.name}\n`;
    content += `  Rows: ${sheet.rows}\n`;
    content += `  Columns: ${sheet.columns}\n`;
    content += `  Size: ${formatFileSize(sheet.size)}\n\n`;
  });
  
  if (summary.errors.length > 0) {
    content += 'ERRORS\n';
    content += '------\n';
    summary.errors.forEach(error => {
      content += `${error.sheet}: ${error.error}\n`;
    });
  }
  
  return content;
}

/**
 * Format file size in human readable format
 * @param {number} bytes - File size in bytes
 * @return {string} Formatted file size
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Export specific date range data
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 */
function exportDateRangeData(startDate, endDate) {
  const logger = Logger.createLogger('DataExportService');
  
  try {
    UIHelpers.showLoading('Exporting date range data...');
    
    const manager = new SheetManager();
    const dailyDetails = manager.getSheet(getConfig('SHEETS.DAILY_DETAILS'));
    const data = dailyDetails.getData();
    
    // Filter data by date range
    const filteredData = [data[0]]; // Include headers
    
    data.slice(1).forEach(row => {
      const date = row[0];
      if (date instanceof Date && date >= startDate && date <= endDate) {
        filteredData.push(row);
      }
    });
    
    if (filteredData.length === 1) {
      UIHelpers.showError(new Error('No data found in the specified date range'));
      return;
    }
    
    // Create CSV
    const csv = convertToCSV(filteredData);
    const filename = `Fleet_Data_${formatDate(startDate)}_to_${formatDate(endDate)}.csv`;
    
    // Create file in Drive
    const blob = Utilities.newBlob(csv, 'text/csv', filename);
    const file = DriveApp.createFile(blob);
    
    google.script.host.close();
    
    // Show success with download link
    const fileUrl = file.getUrl();
    const html = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: 'Google Sans', Arial, sans-serif;">
        <h3 style="color: #34A853;">Export Completed!</h3>
        <p><strong>Date Range:</strong> ${formatDate(startDate)} to ${formatDate(endDate)}</p>
        <p><strong>Records Exported:</strong> ${filteredData.length - 1}</p>
        <p style="margin-top: 20px;">
          <a href="${fileUrl}" target="_blank" style="
            background: #1a73e8;
            color: white;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 4px;
            display: inline-block;
          ">Download CSV File</a>
        </p>
      </div>
    `)
    .setWidth(450)
    .setHeight(250);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Export Complete');
    
    logger.info('Date range export completed', {
      startDate: formatDate(startDate),
      endDate: formatDate(endDate),
      records: filteredData.length - 1
    });
    
  } catch (error) {
    logger.error('Failed to export date range data', { error: error.message });
    google.script.host.close();
    UIHelpers.showError(error);
  }
}

/**
 * Export current week's data
 */
function exportCurrentWeekData() {
  const today = new Date();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - today.getDay());
  startOfWeek.setHours(0, 0, 0, 0);
  
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  endOfWeek.setHours(23, 59, 59, 999);
  
  exportDateRangeData(startOfWeek, endOfWeek);
}