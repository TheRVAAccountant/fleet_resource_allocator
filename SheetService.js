/**
 * ===================================================================
 * SHEET SERVICE
 * ===================================================================
 * Handles all Google Sheets operations including reading, writing,
 * formatting, and sheet management.
 */

/**
 * Retrieves sheet data as a 2D array
 * @param {string} spreadsheetId - Spreadsheet ID
 * @param {string} sheetName - Name of the sheet
 * @return {Array[]} 2D array of sheet data
 */
function getSheetData(spreadsheetId, sheetName) {
  Logger.log("Retrieving data from spreadsheet ID: " + spreadsheetId + ", sheet: " + sheetName);
  
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("Sheet '" + sheetName + "' not found in spreadsheet: " + spreadsheetId);
  }
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow < 1 || lastCol < 1) {
    throw new Error("Sheet '" + sheetName + "' is empty.");
  }
  
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

/**
 * Creates or overwrites the Results sheet
 * @param {string} spreadsheetId - Target spreadsheet ID
 * @param {string} sheetName - Name for the results sheet
 * @param {Object[]} allocationResults - Allocation results data
 */
function createResultsSheet(spreadsheetId, sheetName, allocationResults) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var existing = ss.getSheetByName(sheetName);
  
  if (existing) {
    Logger.log("Results sheet '" + sheetName + "' exists. Deleting...");
    ss.deleteSheet(existing);
  }
  
  Logger.log("Creating new sheet: " + sheetName);
  var sheet = ss.insertSheet(sheetName);
  
  // Headers
  var headers = [
    "Route Code", "Service Type", "DSP", "Wave", "Staging Location",
    "Van ID", "Device Name", "Van Type", "Operational", "Associate Name", "Unique Identifier"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
  // Get today's date for unique identifier
  var today = formatDate(new Date());
  
  // Build rows with unique identifier
  var rows = allocationResults.map(function(r) {
    var uniqueIdentifier = today + "|" + r["Route Code"] + "|" + r["Associate Name"] + "|" + r["Van ID"];
    return [
      r["Route Code"],
      r["Service Type"],
      r["DSP"],
      r["Wave"],
      r["Staging Location"],
      r["Van ID"],
      r["Device Name"],
      r["Van Type"],
      r["Operational"],
      r["Associate Name"],
      uniqueIdentifier
    ];
  });
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Format header
  formatHeaderRow(sheet, 1, headers.length);
  sheet.autoResizeColumns(1, headers.length);
  
  Logger.log("Inserted " + rows.length + " rows into '" + sheetName + "'.");
}

/**
 * Creates or overwrites the Unassigned Vans sheet
 * @param {string} spreadsheetId - Target spreadsheet ID
 * @param {string} sheetName - Name for the unassigned sheet
 * @param {Array[]} data - 2D array of unassigned van data
 */
function createUnassignedSheetInDailySummary(spreadsheetId, sheetName, data) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var existing = ss.getSheetByName(sheetName);
  
  if (existing) {
    Logger.log("Sheet '" + sheetName + "' exists. Deleting...");
    ss.deleteSheet(existing);
  }
  
  Logger.log("Creating new sheet: " + sheetName);
  var sheet = ss.insertSheet(sheetName);
  
  if (data.length > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Format header if exists
    if (data.length > 0) {
      formatHeaderRow(sheet, 1, data[0].length);
      sheet.autoResizeColumns(1, data[0].length);
    }
  }
  
  Logger.log("Created unassigned sheet '" + sheetName + "'. Rows: " + data.length);
}

/**
 * Updates the Daily Details sheet with new allocation data
 * @param {string} spreadsheetId - Daily Summary spreadsheet ID
 * @param {string} resultsSheetName - Name of results sheet to read from
 */
function updateDailyDetails(spreadsheetId, resultsSheetName) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
  
  if (!dailyDetailsSheet) {
    throw new Error("Daily Details sheet not found.");
  }
  
  var resultsSheet = ss.getSheetByName(resultsSheetName);
  if (!resultsSheet) {
    throw new Error("Results sheet '" + resultsSheetName + "' not found.");
  }
  
  var resultsData = resultsSheet.getDataRange().getValues();
  if (resultsData.length < 2) {
    Logger.log("No rows in the Results sheet to process.");
    return;
  }
  
  var today = getTodayString();
  
  // Prepare new rows
  var newRows = [];
  var newUniqueIdentifiers = new Set();
  
  for (var i = 1; i < resultsData.length; i++) {
    var r = resultsData[i];
    var newRow = [today, r[0], r[9], "", r[5]];
    newRows.push(newRow);
    
    var uniqueIdentifier = r[10];
    if (uniqueIdentifier) {
      newUniqueIdentifiers.add(uniqueIdentifier);
    } else {
      Logger.log("Warning: Missing Unique Identifier in Results sheet for row " + (i + 1));
    }
  }
  
  // Check for duplicates
  var existingUniqueIdentifiers = getExistingUniqueIdentifiers(dailyDetailsSheet, today);
  
  var hasDuplicates = false;
  for (var uniqueIdentifier of newUniqueIdentifiers) {
    if (existingUniqueIdentifiers.has(uniqueIdentifier)) {
      hasDuplicates = true;
      Logger.log("Duplicate found, skipping row with Unique ID: " + uniqueIdentifier);
    }
  }
  
  if (hasDuplicates) {
    Logger.log("Duplicate rows detected. No new rows will be appended.");
    return;
  }
  
  // Append new rows
  var lastPopulatedRow = getLastPopulatedRowInColumns(dailyDetailsSheet);
  var startRow = lastPopulatedRow + 1;
  var numNewRows = newRows.length;
  
  var writeRange = dailyDetailsSheet.getRange(startRow, 1, numNewRows, 5);
  var uniqueIdWriteRange = dailyDetailsSheet.getRange(startRow, 22, numNewRows, 1);
  
  writeRange.clearDataValidations();
  writeRange.setValues(newRows);
  
  var uniqueIdValues = Array.from(newUniqueIdentifiers).map(function(id) { return [id]; });
  if (uniqueIdValues.length === numNewRows) {
    uniqueIdWriteRange.setValues(uniqueIdValues);
  } else {
    Logger.log("Warning: Mismatched Unique Identifier and New Rows count.");
  }
  
  // Center-align specific columns
  centerAlignDailyDetailsColumns(dailyDetailsSheet, startRow, numNewRows);
  
  Logger.log("Appended new rows in Daily Details at range: " + writeRange.getA1Notation());
  
  // Update data validation
  updateDailyDetailsValidation(dailyDetailsSheet, newRows, startRow, numNewRows);
}

/**
 * Retrieves existing unique identifiers for a given date
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Daily Details sheet
 * @param {string} dateStr - Date string to check
 * @return {Set} Set of existing unique identifiers
 */
function getExistingUniqueIdentifiers(sheet, dateStr) {
  const lastRow = getLastPopulatedRowInColumns(sheet);
  const uniqueIdentifiers = new Set();
  
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
    
    for (let i = 0; i < data.length; i++) {
      const existingDate = (data[i][0] instanceof Date)
        ? formatDate(data[i][0])
        : (data[i][0] ?? '').toString();
      
      if (existingDate === dateStr) {
        uniqueIdentifiers.add(data[i][21]);
      }
    }
  }
  
  return uniqueIdentifiers;
}

/**
 * Formats a header row with standard styling
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {number} row - Row number
 * @param {number} numColumns - Number of columns
 */
function formatHeaderRow(sheet, row, numColumns) {
  var headerRange = sheet.getRange(row, 1, 1, numColumns);
  headerRange.setFontWeight("bold")
    .setBackground("#4F81BD")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
}

/**
 * Center-aligns specific columns in Daily Details
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Daily Details sheet
 * @param {number} startRow - Starting row
 * @param {number} numRows - Number of rows to format
 */
function centerAlignDailyDetailsColumns(sheet, startRow, numRows) {
  var centerAlignRange = sheet.getRange(startRow, 1, numRows, 22);
  
  centerAlignRange.setHorizontalAlignments(
    Array(numRows).fill(
      [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22].map(function(col) {
        if ([1, 2, 3, 5, 22].includes(col)) {
          return "center";
        } else {
          return "left";
        }
      })
    )
  );
}

/**
 * Updates data validation for Daily Details
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Daily Details sheet
 * @param {Array[]} newRows - New rows being added
 * @param {number} startRow - Starting row
 * @param {number} numRows - Number of rows
 */
function updateDailyDetailsValidation(sheet, newRows, startRow, numRows) {
  var newNames = [];
  for (var i = 0; i < newRows.length; i++) {
    var assoc = newRows[i][2];
    if (assoc && newNames.indexOf(assoc) === -1) {
      newNames.push(assoc);
    }
  }
  
  var sampleCell = sheet.getRange("C2");
  var rule = sampleCell.getDataValidation();
  var allowedValues = (rule ? rule.getCriteriaValues()[0] : []);
  var unionAllowed = allowedValues.slice();
  
  for (var i = 0; i < newNames.length; i++) {
    if (unionAllowed.indexOf(newNames[i]) === -1) {
      unionAllowed.push(newNames[i]);
    }
  }
  
  if (unionAllowed.length > 0) {
    var newRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(unionAllowed, true)
      .build();
    sheet.getRange(startRow, 3, numRows, 1).setDataValidation(newRule);
  }
}