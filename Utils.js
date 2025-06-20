/**
 * ===================================================================
 * UTILITY FUNCTIONS
 * ===================================================================
 * Common utility functions used throughout the application for
 * data manipulation, formatting, and general operations.
 */

/**
 * Extracts the file ID from user input (URL or ID)
 * @param {string} input - User provided input
 * @return {string} Extracted file ID
 */
function extractFileId(input) {
  input = input.trim();
  
  // Check if it's already a file ID
  if (input.match(/^[a-zA-Z0-9-_]{30,}$/)) {
    return input;
  }
  
  // Try to extract from URL
  var match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    return match[1];
  }
  
  throw new Error("Could not parse file ID from the provided input:\n" + input);
}

/**
 * Converts a 2D array to an array of objects using first row as headers
 * @param {Array[]} data2D - 2D array with headers in first row
 * @return {Object[]} Array of objects
 */
function convertToObjectArray(data2D) {
  if (!data2D || data2D.length < 2) {
    return [];
  }
  
  var headers = data2D[0];
  var output = [];
  
  for (var row = 1; row < data2D.length; row++) {
    var rowObj = {};
    for (var col = 0; col < headers.length; col++) {
      var key = headers[col];
      rowObj[key] = data2D[row][col];
    }
    output.push(rowObj);
  }
  
  return output;
}

/**
 * Groups an array of objects by a specified field
 * @param {Object[]} arr - Array to group
 * @param {string} fieldName - Field to group by
 * @return {Object} Grouped object
 */
function groupBy(arr, fieldName) {
  var out = {};
  
  arr.forEach(function(obj) {
    var key = obj[fieldName];
    if (!out[key]) {
      out[key] = [];
    }
    out[key].push(obj);
  });
  
  return out;
}

/**
 * Verifies required columns exist in headers
 * @param {string[]} actualHeaders - Actual headers from sheet
 * @param {string[]} requiredHeaders - Required headers to check
 * @param {string} contextLabel - Context for error message
 * @throws {Error} If required columns are missing
 */
function verifyRequiredColumns(actualHeaders, requiredHeaders, contextLabel) {
  var missing = [];
  
  requiredHeaders.forEach(function(req) {
    if (actualHeaders.indexOf(req) === -1) {
      missing.push(req);
    }
  });
  
  if (missing.length > 0) {
    throw new Error("Missing required columns in " + contextLabel + ": " + missing.join(", "));
  }
}

/**
 * Maps service types to van types
 * @param {string} serviceType - Service type from route
 * @return {string|null} Corresponding van type or null
 */
function getVanType(serviceType) {
  var mapping = getConfig('VAN_TYPE_MAPPING');
  
  if (mapping[serviceType]) {
    return mapping[serviceType];
  } else if (serviceType && serviceType.indexOf("Nursery Route Level") !== -1) {
    return "Large";
  }
  
  return null;
}

/**
 * Gets the last populated row in columns A-E
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to check
 * @return {number} Last populated row number
 */
function getLastPopulatedRowInColumns(sheet) {
  var data = sheet.getRange("A:E").getValues();
  var lastPopulatedRow = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    for (var j = 0; j < row.length; j++) {
      if (row[j] && row[j].toString().trim() !== "") {
        lastPopulatedRow = i + 1;
        break;
      }
    }
  }
  
  return lastPopulatedRow;
}

/**
 * Formats date consistently across the application
 * @param {Date} date - Date to format
 * @return {string} Formatted date string
 */
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
}

/**
 * Gets current date string in standard format
 * @return {string} Today's date formatted
 */
function getTodayString() {
  return formatDate(new Date());
}

/**
 * Creates a timestamp string for file naming
 * @return {string} Formatted timestamp
 */
function getTimestampString() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM-dd-yy HH:mm:ss");
}