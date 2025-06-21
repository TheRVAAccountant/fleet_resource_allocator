/**
 * ===================================================================
 * UTILITY FUNCTIONS
 * ===================================================================
 * Common utility functions used throughout the application for
 * data manipulation, formatting, and general operations.
 */

/**
 * Extracts the file ID from user input (URL or ID)
 * @param {string} input - User provided input (can be full URL or just ID)
 * @return {string} Extracted file ID
 * @throws {ValidationError} If input format is invalid
 * @example
 * // Returns: "1abc123def456"
 * extractFileId("https://docs.google.com/spreadsheets/d/1abc123def456/edit")
 * 
 * // Returns: "1abc123def456" 
 * extractFileId("1abc123def456")
 */
function extractFileId(input) {
  if (!input || typeof input !== 'string') {
    throw new ValidationError('Invalid input: expected string', 'fileId');
  }
  
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
  
  throw new ValidationError("Could not parse file ID from the provided input: " + input, 'fileId');
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

/**
 * Normalize reporting time by removing suffixes like "(End of Day)"
 * @param {string} reportingTime - The reporting time string
 * @return {string} Normalized time string
 */
function normalizeReportingTime(reportingTime) {
  if (!reportingTime || typeof reportingTime !== 'string') {
    return '';
  }
  // Remove common suffixes and trim whitespace
  return reportingTime
    .replace(' (End of Day)', '')
    .replace('(End of Day)', '')
    .trim();
}

/**
 * Convert Date object to time string format (e.g., "1:40 PM")
 * @param {Date|string} timeValue - Time as Date object or string
 * @return {string} Formatted time string like "1:40 PM"
 */
function formatTimeString(timeValue) {
  // If already a string, check if it's in the right format
  if (typeof timeValue === 'string') {
    // Check if it already looks like a time string
    if (timeValue.match(/^\d{1,2}:\d{2}\s*(AM|PM)$/i)) {
      return timeValue;
    }
    // Try to parse it as a date
    timeValue = new Date(timeValue);
  }
  
  // If not a valid date, return empty string
  if (!(timeValue instanceof Date) || isNaN(timeValue)) {
    console.log('Invalid time value:', timeValue);
    return '';
  }
  
  // Extract hours and minutes
  var hours = timeValue.getHours();
  var minutes = timeValue.getMinutes();
  
  // Convert to 12-hour format
  var period = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12; // 0 should be 12
  
  // Format minutes with leading zero if needed
  var minutesStr = minutes < 10 ? '0' + minutes : minutes.toString();
  
  return hours + ':' + minutesStr + ' ' + period;
}

/**
 * Format time value as text to prevent Google Sheets date conversion
 * @param {string} timeValue - Time value (e.g., "18:30")
 * @return {string} Time value formatted as text in AM/PM format
 */
function formatTimeAsText(timeValue) {
  if (!timeValue || timeValue === '-') {
    return timeValue || '-';
  }
  
  // If it already starts with apostrophe, return as is
  if (typeof timeValue === 'string' && timeValue.startsWith("'")) {
    return timeValue;
  }
  
  // If it's already in AM/PM format, just add apostrophe
  if (typeof timeValue === 'string' && timeValue.match(/^\d{1,2}:\d{2}\s*(AM|PM)$/i)) {
    return "'" + timeValue;
  }
  
  // If it's a string that looks like 24-hour format (no AM/PM)
  if (typeof timeValue === 'string' && timeValue.match(/^\d{1,2}:\d{2}$/)) {
    // Convert 24-hour format to AM/PM
    var timeParts = timeValue.split(':');
    var hours = parseInt(timeParts[0]);
    var minutes = timeParts[1];
    
    var period = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // 0 should be 12
    
    return "'" + hours + ':' + minutes + ' ' + period;
  }
  
  // If it's a Date object, extract and format the time
  if (timeValue instanceof Date) {
    var hours = timeValue.getHours();
    var minutes = timeValue.getMinutes();
    
    var period = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // 0 should be 12
    
    var minutesStr = minutes < 10 ? '0' + minutes : minutes.toString();
    return "'" + hours + ':' + minutesStr + ' ' + period;
  }
  
  return String(timeValue);
}

/**
 * Convert 24-hour time format to 12-hour AM/PM format
 * @param {string} time24 - Time in 24-hour format (e.g., "18:30")
 * @return {string} Time in 12-hour AM/PM format (e.g., "6:30 PM")
 */
function convertTo12HourFormat(time24) {
  if (!time24 || typeof time24 !== 'string') {
    return time24 || '';
  }
  
  // If already in AM/PM format, return as is
  if (time24.match(/^\d{1,2}:\d{2}\s*(AM|PM)$/i)) {
    return time24;
  }
  
  // Parse 24-hour format
  var match = time24.match(/^(\d{1,2}):(\d{2})/);
  if (!match) {
    return time24;
  }
  
  var hours = parseInt(match[1]);
  var minutes = match[2];
  
  var period = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12; // 0 should be 12
  
  return hours + ':' + minutes + ' ' + period;
}

/**
 * Escape HTML special characters
 * @param {string} text - Text to escape
 * @return {string} Escaped text
 */
function escapeHtml(text) {
  if (!text) return '';
  
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Calculate the week number of the year for a given date
 * Week starts on Sunday (US convention)
 * @param {Date} date - Date to calculate week number for
 * @return {number} Week number (1-52/53)
 */
function getWeekNumber(date) {
  // Create a copy of the date to avoid modifying the original
  var d = new Date(date.getTime());
  
  // Set to nearest Thursday: current date + 4 - current day number
  // Make Sunday's day number 7
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  
  // Get first day of year
  var yearStart = new Date(d.getFullYear(), 0, 1);
  
  // Calculate full weeks to nearest Thursday
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  
  return weekNo;
}

/**
 * Get ISO week number (Monday as first day of week)
 * @param {Date} date - Date to calculate week number for
 * @return {number} ISO week number (1-52/53)
 */
function getISOWeekNumber(date) {
  var d = new Date(date.getTime());
  d.setHours(0, 0, 0, 0);
  
  // Thursday in current week decides the year
  d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
  
  // January 4 is always in week 1
  var week1 = new Date(d.getFullYear(), 0, 4);
  
  // Adjust to Thursday in week 1 and count number of weeks from date to week1
  return 1 + Math.round(((d.getTime() - week1.getTime()) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
}

/**
 * Get US week number (Sunday as first day of week, simpler calculation)
 * @param {Date} date - Date to calculate week number for
 * @return {number} Week number (1-53)
 */
function getUSWeekNumber(date) {
  var d = new Date(date.getTime());
  d.setHours(0, 0, 0, 0);
  
  // Get first day of the year
  var yearStart = new Date(d.getFullYear(), 0, 1);
  
  // Calculate days since start of year
  var daysSinceYearStart = Math.floor((d - yearStart) / 86400000);
  
  // Calculate week number (week 1 starts on Jan 1)
  var weekNumber = Math.ceil((daysSinceYearStart + yearStart.getDay() + 1) / 7);
  
  return weekNumber;
}

/**
 * Test week number calculations
 */
function testWeekNumberCalculations() {
  console.log('=== Testing Week Number Calculations ===');
  
  var testDates = [
    new Date(2025, 0, 1),   // January 1, 2025
    new Date(2025, 0, 5),   // January 5, 2025 (Sunday)
    new Date(2025, 0, 6),   // January 6, 2025 (Monday)
    new Date(2025, 5, 21),  // June 21, 2025
    new Date(2025, 11, 31), // December 31, 2025
    new Date(2024, 11, 30), // December 30, 2024
    new Date(2024, 11, 31), // December 31, 2024
    new Date(2024, 0, 1),   // January 1, 2024
  ];
  
  testDates.forEach(function(date) {
    var weekNum = getUSWeekNumber(date);
    var isoWeekNum = getISOWeekNumber(date);
    var standardWeekNum = getWeekNumber(date);
    
    console.log(
      formatDate(date) + 
      ' (' + ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][date.getDay()] + ')' +
      ': US Week=' + weekNum + 
      ', ISO Week=' + isoWeekNum + 
      ', Standard Week=' + standardWeekNum
    );
  });
  
  console.log('\n=== Week Number Test Complete ===');
  
  SpreadsheetApp.getUi().alert(
    'Week Number Test Complete',
    'Check the logs for detailed week number calculations.\n\n' +
    'Today (' + formatDate(new Date()) + ') is week ' + getUSWeekNumber(new Date()),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Format time with timezone for display
 * @param {string} timeString - Time in 24-hour format (e.g., "18:30")
 * @return {string} Formatted time with timezone (e.g., "6:30 PM EDT")
 */
function formatTimeWithTimezone(timeString) {
  if (!timeString) return '';
  
  // Convert to 12-hour format if needed
  var formattedTime = convertTo12HourFormat(timeString);
  
  // Get timezone abbreviation
  var timezone = getTimezoneAbbreviation();
  
  return formattedTime + ' ' + timezone;
}

/**
 * Get current timezone abbreviation (EDT/EST)
 * @return {string} Timezone abbreviation
 */
function getTimezoneAbbreviation() {
  var date = new Date();
  var timeZone = Session.getScriptTimeZone();
  
  // For Eastern Time, determine if it's EDT or EST based on date
  if (timeZone === 'America/New_York' || timeZone.includes('Eastern')) {
    // Check if we're in daylight saving time
    var jan = new Date(date.getFullYear(), 0, 1);
    var jul = new Date(date.getFullYear(), 6, 1);
    var stdOffset = Math.max(jan.getTimezoneOffset(), jul.getTimezoneOffset());
    var isDST = date.getTimezoneOffset() < stdOffset;
    
    return isDST ? 'EDT' : 'EST';
  }
  
  // For other timezones, try to extract abbreviation
  try {
    var formatter = Utilities.formatDate(date, timeZone, 'zzz');
    return formatter;
  } catch (e) {
    // Default to EDT if we can't determine
    return 'EDT';
  }
}

/**
 * Show information alert to user
 * @param {string} message - Message to display
 * @param {string} title - Alert title (optional)
 */
function showInfoAlert(message, title) {
  SpreadsheetApp.getUi().alert(
    title || 'Information',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Test week number column in Daily Details
 */
function testWeekNumberInDailyDetails() {
  console.log('=== Testing Week Number in Daily Details ===');
  
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
    
    if (!dailyDetailsSheet) {
      throw new Error('Daily Details sheet not found');
    }
    
    // Check the headers
    var headers = dailyDetailsSheet.getRange(1, 1, 1, 22).getValues()[0];
    console.log('Column U (index 20) header:', headers[20]);
    console.log('Column V (index 21) header:', headers[21]);
    
    // Check a few recent rows
    var lastRow = dailyDetailsSheet.getLastRow();
    if (lastRow > 1) {
      var recentData = dailyDetailsSheet.getRange(Math.max(2, lastRow - 4), 1, Math.min(5, lastRow - 1), 22).getValues();
      
      console.log('\nRecent rows:');
      recentData.forEach(function(row, index) {
        var date = row[0];
        var weekNumber = row[20]; // Column U
        var uniqueId = row[21]; // Column V
        
        console.log('Row ' + (lastRow - 4 + index) + ':');
        console.log('  Date:', date instanceof Date ? formatDate(date) : date);
        console.log('  Week Number (Col U):', weekNumber);
        console.log('  Unique ID (Col V):', uniqueId ? uniqueId.substring(0, 30) + '...' : 'empty');
      });
    }
    
    // Test what week number today would be
    var today = new Date();
    var todayWeekNum = getUSWeekNumber(today);
    console.log('\nToday (' + formatDate(today) + ') is week number: ' + todayWeekNum);
    
    SpreadsheetApp.getUi().alert(
      'Week Number Column Test',
      'Column U header: ' + headers[20] + '\n' +
      'Today\'s week number: ' + todayWeekNum + '\n\n' +
      'Check logs for more details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error testing week number column:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}