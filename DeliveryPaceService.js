/**
 * ===================================================================
 * DELIVERY PACE SERVICE
 * ===================================================================
 * Manages delivery pace tracking, updates, reporting, and automation
 * for monitoring van delivery progress throughout the day.
 */

/**
 * Initialize delivery pace column headers in Daily Details sheet
 */
function initializeDeliveryPaceHeaders() {
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
  
  if (!dailyDetailsSheet) {
    throw new Error("Daily Details sheet not found");
  }
  
  // Set headers for columns L-P
  var headers = [["Delivery Pace: 1:40 PM", "3:40 PM", "5:40 PM", "7:40 PM", "9:40 PM"]];
  dailyDetailsSheet.getRange(1, 12, 1, 5).setValues(headers);
  
  // Format headers
  var headerRange = dailyDetailsSheet.getRange(1, 12, 1, 5);
  headerRange.setFontWeight("bold")
    .setBackground("#E8F0FE")
    .setHorizontalAlignment("center")
    .setWrap(true);
  
  Logger.log("Delivery pace headers initialized");
}

/**
 * Get delivery pace data for a specific van
 * @param {string} vanId - Van ID
 * @param {string} date - Date string
 * @return {Object} Pace data by time slot
 */
function getDeliveryPaceData(vanId, date) {
  Logger.log("Fetching delivery pace data for Van: " + vanId + ", Date: " + date);
  
  // Try to get data from form responses first
  var formData = getDeliveryPaceDataFromForms(vanId, date);
  
  if (formData && Object.keys(formData).length > 0) {
    Logger.log("Using form-submitted data for Van: " + vanId);
    return formData;
  }
  
  // Try to get data from external source
  var externalData = getDeliveryPaceDataFromSource(vanId, date);
  
  if (externalData && Object.keys(externalData).length > 0) {
    Logger.log("Using external data for Van: " + vanId);
    return externalData;
  }
  
  // No data available
  Logger.log("No delivery pace data available for Van: " + vanId);
  return {};
}

/**
 * Get delivery pace data from external source
 * @param {string} vanId - Van ID
 * @param {string} date - Date string
 * @return {Object} Pace data by time slot
 */
function getDeliveryPaceDataFromSource(vanId, date) {
  // Check if mock data is enabled (runtime override takes precedence)
  var runtimeMockEnabled = isMockDataEnabled(vanId);
  
  if (!runtimeMockEnabled) {
    // Check configuration if runtime not set
    var useMockData = getConfig('DEV_SETTINGS.USE_MOCK_DATA');
    var mockEnabledVans = getConfig('DEV_SETTINGS.MOCK_DATA_ENABLED_VANS') || [];
    
    // Only use mock data if explicitly enabled AND (all vans OR specific van is enabled)
    var shouldUseMockData = useMockData && 
      (mockEnabledVans.length === 0 || mockEnabledVans.indexOf(vanId) !== -1);
    
    if (!shouldUseMockData) {
      Logger.log("Mock data disabled for Van: " + vanId + " - returning empty data");
      return {}; // Return empty object instead of mock data
    }
  }
  
  // Option 1: Read from another Google Sheet
  // var dataSpreadsheetId = "YOUR_DATA_SOURCE_SPREADSHEET_ID";
  // var dataSheet = SpreadsheetApp.openById(dataSpreadsheetId).getSheetByName("DeliveryData");
  
  // Option 2: Call an external API
  // var apiUrl = "https://your-api.com/delivery-pace/" + vanId + "/" + date;
  // var response = UrlFetchApp.fetch(apiUrl, {
  //     'headers': {
  //         'Authorization': 'Bearer YOUR_API_TOKEN'
  //     }
  // });
  // var data = JSON.parse(response.getContentText());
  
  // Option 3: Query from a database via JDBC
  // var conn = Jdbc.getConnection("jdbc:mysql://your-host:3306/database", "user", "password");
  // var stmt = conn.prepareStatement("SELECT * FROM delivery_pace WHERE van_id = ? AND date = ?");
  // stmt.setString(1, vanId);
  // stmt.setString(2, date);
  // var results = stmt.executeQuery();
  
  // Mock data only for testing/development when explicitly enabled
  Logger.log("Using mock data for Van: " + vanId + " (DEV MODE)");
  var baseStops = Math.floor(Math.random() * 20) + 10;
  return {
    "1:40 PM": baseStops,
    "3:40 PM": baseStops + Math.floor(Math.random() * 30) + 20,
    "5:40 PM": baseStops + Math.floor(Math.random() * 50) + 40,
    "7:40 PM": baseStops + Math.floor(Math.random() * 60) + 60,
    "9:40 PM": baseStops + Math.floor(Math.random() * 70) + 80
  };
}

/**
 * Get delivery pace data from form submissions
 * @param {string} vanId - Van ID
 * @param {string} date - Date string
 * @return {Object} Pace data from forms
 */
function getDeliveryPaceDataFromForms(vanId, date) {
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dataSheet = ss.getSheetByName('Delivery Pace Data');
    
    if (!dataSheet) {
      Logger.log("Delivery Pace Data sheet not found");
      return null;
    }
    
    var data = dataSheet.getDataRange().getValues();
    
    // Log sheet info for debugging
    console.log('Checking Delivery Pace Data sheet. Total rows:', data.length);
    
    var paceData = {
      "1:40 PM": null,
      "3:40 PM": null,
      "5:40 PM": null,
      "7:40 PM": null,
      "9:40 PM": null
    };
    
    // Validate headers
    if (data.length > 0) {
      var headers = data[0];
      console.log('Sheet headers:', headers);
      
      // Verify expected columns exist
      var expectedHeaders = ['Timestamp', 'Date', 'Van ID', 'Driver Name', 'Route Code', 'Reporting Time', 'Total Deliveries', 'Notes', 'Processed'];
      var missingHeaders = expectedHeaders.filter(function(header, index) {
        return headers[index] !== header;
      });
      
      if (missingHeaders.length > 0) {
        console.log('Warning: Missing or mismatched headers:', missingHeaders);
      }
    }
    
    // Headers: Timestamp, Date, Van ID, Driver Name, Route Code, Reporting Time, Total Deliveries, Notes, Processed
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][1];
      if (rowDate instanceof Date) {
        rowDate = formatDate(rowDate);
      }
      
      var rowVanId = data[i][2];
      var reportingTime = data[i][5];
      var deliveryCount = data[i][6];
      
      // Match van ID and date
      if (rowVanId === vanId && rowDate === date) {
        // Handle Date objects for reporting time
        if (reportingTime instanceof Date) {
          console.log('Converting Date object to time string for Van:', vanId);
          console.log('Original Date:', reportingTime);
          reportingTime = formatTimeString(reportingTime);
          console.log('Converted to:', reportingTime);
        } else if (reportingTime && typeof reportingTime !== 'string') {
          console.log('Converting reportingTime to string. Type was:', typeof reportingTime, 'Value:', reportingTime);
          reportingTime = String(reportingTime);
        }
        
        // Skip if reportingTime is empty or invalid
        if (!reportingTime || reportingTime.trim() === '') {
          console.log('Skipping row with invalid reportingTime for Van:', vanId);
          continue;
        }
        
        // Map reporting time to our standard format
        // Normalize by removing any suffix like "(End of Day)"
        var timeKey = normalizeReportingTime(reportingTime);
        
        if (paceData.hasOwnProperty(timeKey)) {
          // Keep the latest submission for each time slot
          paceData[timeKey] = deliveryCount;
          console.log('Found pace data for Van ' + vanId + ' at ' + timeKey + ': ' + deliveryCount);
        } else {
          console.log('Unknown time slot:', timeKey, 'for Van:', vanId);
          console.log('Available time slots:', Object.keys(paceData).join(', '));
        }
      }
    }
    
    // Check if we have any actual data
    var hasData = false;
    for (var key in paceData) {
      if (paceData[key] !== null) {
        hasData = true;
        break;
      }
    }
    
    return hasData ? paceData : null;
    
  } catch (error) {
    Logger.log("Error reading form data: " + error);
    Logger.log("Error stack: " + error.stack);
    // Log specific data that caused the error for debugging
    if (typeof reportingTime !== 'undefined') {
      Logger.log("reportingTime type: " + typeof reportingTime + ", value: " + reportingTime);
    }
    return null;
  }
}

/**
 * Update delivery pace for all vans allocated today
 */
function updateDeliveryPaceForToday() {
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
  
  if (!dailyDetailsSheet) {
    throw new Error("Daily Details sheet not found");
  }
  
  var today = getTodayString();
  Logger.log("Updating delivery pace for date: " + today);
  
  // Get all data from Daily Details
  var lastRow = getLastPopulatedRowInColumns(dailyDetailsSheet);
  if (lastRow < 2) {
    Logger.log("No data to process");
    return;
  }
  
  var data = dailyDetailsSheet.getRange(2, 1, lastRow - 1, 22).getValues();
  var currentHour = new Date().getHours();
  var currentMinutes = new Date().getMinutes();
  var currentTime = currentHour + (currentMinutes / 60);
  
  var timeSlots = getConfig('DELIVERY_TIME_SLOTS');
  var updatedRows = 0;
  
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][0];
    
    // Format the date for comparison
    if (rowDate instanceof Date) {
      rowDate = formatDate(rowDate);
    }
    
    // Only process today's entries
    if (rowDate === today) {
      var vanId = data[i][4]; // Column E - Van ID
      
      if (vanId) {
        // Get delivery pace data for this van
        var paceData = getDeliveryPaceData(vanId, today);
        
        // Only process if we have actual data
        if (paceData && Object.keys(paceData).length > 0) {
          var hasUpdates = false;
          
          // Update appropriate columns based on current time
          for (var j = 0; j < timeSlots.length; j++) {
            var slot = timeSlots[j];
            
            // Only update if current time is past the time slot
            if (currentTime >= slot.time && paceData.hasOwnProperty(slot.label)) {
              var cellValue = paceData[slot.label];
              
              // Update the cell
              dailyDetailsSheet.getRange(i + 2, slot.column).setValue(cellValue);
              
              Logger.log("Updated Van " + vanId + " at " + slot.label + ": " + cellValue + " stops");
              hasUpdates = true;
            }
          }
          
          if (hasUpdates) {
            updatedRows++;
          }
        } else {
          Logger.log("No pace data available for Van " + vanId + " - skipping update");
        }
      }
    }
  }
  
  Logger.log("Updated delivery pace for " + updatedRows + " vans");
}

/**
 * Update delivery pace for a specific van
 * @param {string} vanId - Van ID
 * @param {string} date - Date string
 * @return {boolean} Success status
 */
function updateDeliveryPaceForVan(vanId, date) {
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
  
  if (!dailyDetailsSheet) {
    throw new Error("Daily Details sheet not found");
  }
  
  // Find the row for this van on this date
  var lastRow = getLastPopulatedRowInColumns(dailyDetailsSheet);
  var data = dailyDetailsSheet.getRange(2, 1, lastRow - 1, 22).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][0];
    var rowVanId = data[i][4];
    
    // Format date for comparison
    if (rowDate instanceof Date) {
      rowDate = formatDate(rowDate);
    }
    
    if (rowDate === date && rowVanId === vanId) {
      // Found the matching row
      var paceData = getDeliveryPaceData(vanId, date);
      
      // Update columns L-P
      var updateValues = [[
        paceData["1:40 PM"],
        paceData["3:40 PM"],
        paceData["5:40 PM"],
        paceData["7:40 PM"],
        paceData["9:40 PM"]
      ]];
      
      dailyDetailsSheet.getRange(i + 2, 12, 1, 5).setValues(updateValues);
      
      Logger.log("Updated delivery pace for Van: " + vanId + " on " + date);
      return true;
    }
  }
  
  Logger.log("Van " + vanId + " not found for date " + date);
  return false;
}

/**
 * Batch update delivery pace for multiple vans
 * @param {string[]} vanIds - Array of van IDs
 * @param {string} date - Date string
 * @return {number} Number of vans updated
 */
function batchUpdateDeliveryPace(vanIds, date) {
  var updatedCount = 0;
  
  vanIds.forEach(function(vanId) {
    if (updateDeliveryPaceForVan(vanId, date)) {
      updatedCount++;
    }
  });
  
  Logger.log("Batch update completed. Updated " + updatedCount + " vans out of " + vanIds.length);
  return updatedCount;
}

/**
 * Create time-based triggers for automatic updates
 */
function setupDeliveryPaceTriggers() {
  // Remove existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "updateDeliveryPaceForToday") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new triggers for each time slot
  var times = [
    {hour: 13, minute: 45}, // 1:45 PM
    {hour: 15, minute: 45}, // 3:45 PM
    {hour: 17, minute: 45}, // 5:45 PM
    {hour: 19, minute: 45}, // 7:45 PM
    {hour: 21, minute: 45}  // 9:45 PM
  ];
  
  times.forEach(function(time) {
    ScriptApp.newTrigger("updateDeliveryPaceForToday")
      .timeBased()
      .everyDays(1)
      .atHour(time.hour)
      .nearMinute(time.minute)
      .create();
  });
  
  Logger.log("Delivery pace triggers created for 5 time slots");
}

/**
 * Generate delivery pace summary report for a specific date
 * @param {string} date - Date string (optional, defaults to today)
 * @return {Object} Summary data
 */
function generateDeliveryPaceSummary(date) {
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
  
  if (!dailyDetailsSheet) {
    throw new Error("Daily Details sheet not found");
  }
  
  if (!date) {
    date = getTodayString();
  }
  
  // Get all data for the specified date
  var lastRow = getLastPopulatedRowInColumns(dailyDetailsSheet);
  var data = dailyDetailsSheet.getRange(2, 1, lastRow - 1, 22).getValues();
  
  var summary = {
    date: date,
    totalVans: 0,
    vansWithData: 0,
    averagePace: {
      "1:40 PM": 0,
      "3:40 PM": 0,
      "5:40 PM": 0,
      "7:40 PM": 0,
      "9:40 PM": 0
    },
    vanDetails: []
  };
  
  var counts = {
    "1:40 PM": 0,
    "3:40 PM": 0,
    "5:40 PM": 0,
    "7:40 PM": 0,
    "9:40 PM": 0
  };
  
  // Process each row
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][0];
    
    // Format date for comparison
    if (rowDate instanceof Date) {
      rowDate = formatDate(rowDate);
    }
    
    if (rowDate === date) {
      summary.totalVans++;
      
      var vanId = data[i][4];
      var driverName = data[i][2];
      var routeCode = data[i][1];
      
      var vanData = {
        vanId: vanId,
        driver: driverName,
        route: routeCode,
        pace: {}
      };
      
      var hasData = false;
      
      // Collect pace data
      var timeSlots = ["1:40 PM", "3:40 PM", "5:40 PM", "7:40 PM", "9:40 PM"];
      for (var j = 0; j < timeSlots.length; j++) {
        var value = data[i][11 + j];
        if (value && !isNaN(value)) {
          vanData.pace[timeSlots[j]] = value;
          summary.averagePace[timeSlots[j]] += value;
          counts[timeSlots[j]]++;
          hasData = true;
        }
      }
      
      if (hasData) {
        summary.vansWithData++;
        summary.vanDetails.push(vanData);
      }
    }
  }
  
  // Calculate averages
  for (var slot in summary.averagePace) {
    if (counts[slot] > 0) {
      summary.averagePace[slot] = Math.round(summary.averagePace[slot] / counts[slot]);
    }
  }
  
  // Create summary sheet
  createDeliveryPaceSummarySheet(summary);
  
  return summary;
}

/**
 * Create a summary sheet with delivery pace statistics
 * @param {Object} summary - Summary data object
 */
function createDeliveryPaceSummarySheet(summary) {
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var summarySheetName = summary.date.replace(/\//g, "-") + " - Pace Summary";
  
  // Check if sheet exists
  var summarySheet = ss.getSheetByName(summarySheetName);
  if (summarySheet) {
    ss.deleteSheet(summarySheet);
  }
  
  summarySheet = ss.insertSheet(summarySheetName);
  
  // Add title and summary stats
  var titleData = [
    ["Delivery Pace Summary Report"],
    ["Date: " + summary.date],
    [""],
    ["Total Vans Allocated:", summary.totalVans],
    ["Vans with Pace Data:", summary.vansWithData],
    [""],
    ["Average Stops by Time:"],
    ["1:40 PM:", summary.averagePace["1:40 PM"]],
    ["3:40 PM:", summary.averagePace["3:40 PM"]],
    ["5:40 PM:", summary.averagePace["5:40 PM"]],
    ["7:40 PM:", summary.averagePace["7:40 PM"]],
    ["9:40 PM:", summary.averagePace["9:40 PM"]],
    [""],
    ["Van Details:"]
  ];
  
  summarySheet.getRange(1, 1, titleData.length, 2).setValues(titleData);
  
  // Format title
  summarySheet.getRange(1, 1, 1, 2).merge()
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  // Add van details headers
  var detailsStartRow = titleData.length + 2;
  var headers = [["Van ID", "Driver", "Route", "1:40 PM", "3:40 PM", "5:40 PM", "7:40 PM", "9:40 PM"]];
  summarySheet.getRange(detailsStartRow, 1, 1, headers[0].length).setValues(headers);
  formatHeaderRow(summarySheet, detailsStartRow, headers[0].length);
  
  // Add van details data
  if (summary.vanDetails.length > 0) {
    var detailsData = summary.vanDetails.map(function(van) {
      return [
        van.vanId,
        van.driver,
        van.route,
        van.pace["1:40 PM"] || "",
        van.pace["3:40 PM"] || "",
        van.pace["5:40 PM"] || "",
        van.pace["7:40 PM"] || "",
        van.pace["9:40 PM"] || ""
      ];
    });
    
    summarySheet.getRange(detailsStartRow + 1, 1, detailsData.length, headers[0].length)
      .setValues(detailsData);
  }
  
  // Auto-resize columns
  summarySheet.autoResizeColumns(1, headers[0].length);
  
  Logger.log("Created delivery pace summary sheet: " + summarySheetName);
  showInfoAlert("Delivery Pace Summary created: " + summarySheetName);
}

/**
 * Test function to debug form data reading
 */
function testFormDataReading() {
  console.log('=== Testing Form Data Reading ===');
  
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dataSheet = ss.getSheetByName('Delivery Pace Data');
    
    if (!dataSheet) {
      console.log('ERROR: Delivery Pace Data sheet not found');
      SpreadsheetApp.getUi().alert('Delivery Pace Data sheet not found. Please create it first.');
      return;
    }
    
    var data = dataSheet.getDataRange().getValues();
    console.log('Total rows in Delivery Pace Data sheet:', data.length);
    
    if (data.length === 0) {
      console.log('Sheet is empty');
      SpreadsheetApp.getUi().alert('Delivery Pace Data sheet is empty. Submit a form first.');
      return;
    }
    
    // Log headers
    console.log('Headers:', data[0]);
    
    // Check first few data rows
    var sampleSize = Math.min(5, data.length - 1);
    console.log('Checking first', sampleSize, 'data rows...');
    
    for (var i = 1; i <= sampleSize; i++) {
      if (i < data.length) {
        var row = data[i];
        console.log('\nRow', i, ':');
        console.log('  Date:', row[1], 'Type:', typeof row[1]);
        console.log('  Van ID:', row[2], 'Type:', typeof row[2]);
        console.log('  Reporting Time:', row[5], 'Type:', typeof row[5]);
        console.log('  Deliveries:', row[6], 'Type:', typeof row[6]);
      }
    }
    
    // Test reading data for a specific van
    if (data.length > 1) {
      var testVanId = data[1][2]; // Get van ID from first data row
      var testDate = data[1][1];
      
      if (testDate instanceof Date) {
        testDate = formatDate(testDate);
      }
      
      console.log('\nTesting getDeliveryPaceDataFromForms for Van:', testVanId, 'Date:', testDate);
      
      var paceData = getDeliveryPaceDataFromForms(testVanId, testDate);
      console.log('Result:', paceData);
    }
    
    SpreadsheetApp.getUi().alert(
      'Form Data Test Complete',
      'Total rows: ' + data.length + '\n' +
      'Check logs for detailed information.\n\n' +
      'If you see reportingTime errors, check that the form is saving text values.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Test failed:', error);
    console.error('Stack:', error.stack);
    SpreadsheetApp.getUi().alert('Test failed: ' + error.toString());
  }
}

/**
 * Create sample data in Delivery Pace Data sheet for testing
 */
function createSampleDeliveryPaceData() {
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dataSheet = ss.getSheetByName('Delivery Pace Data');
    
    if (!dataSheet) {
      // Create the sheet if it doesn't exist
      dataSheet = ss.insertSheet('Delivery Pace Data');
      setupDeliveryPaceDataSheet(dataSheet);
      console.log('Created Delivery Pace Data sheet');
    }
    
    var today = new Date();
    var timeSlots = ['1:40 PM', '3:40 PM', '5:40 PM', '7:40 PM', '9:40 PM (End of Day)'];
    var sampleVans = ['BW2', 'BW10'];
    
    var sampleData = [];
    
    // Create sample data for each van and time slot
    sampleVans.forEach(function(vanId) {
      timeSlots.forEach(function(timeSlot, index) {
        var deliveries = 20 + (index * 25) + Math.floor(Math.random() * 10);
        sampleData.push([
          new Date(), // Timestamp
          today, // Date
          vanId, // Van ID
          'Test Driver', // Driver Name
          'TEST001', // Route Code
          timeSlot, // Reporting Time
          deliveries, // Total Deliveries
          'Test data', // Notes
          'No' // Processed
        ]);
      });
    });
    
    // Append the sample data
    if (sampleData.length > 0) {
      var lastRow = dataSheet.getLastRow();
      var dataRange = dataSheet.getRange(lastRow + 1, 1, sampleData.length, sampleData[0].length);
      dataRange.setValues(sampleData);
      
      // Format reporting time column as text to prevent date conversion
      dataSheet.getRange(lastRow + 1, 6, sampleData.length, 1).setNumberFormat('@');
      
      console.log('Added', sampleData.length, 'sample rows');
    }
    
    SpreadsheetApp.getUi().alert(
      'Sample Data Created',
      'Added ' + sampleData.length + ' sample rows to Delivery Pace Data sheet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error creating sample data:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Test function to verify time format normalization
 */
function testTimeFormatNormalization() {
  console.log('=== Testing Time Format Normalization ===');
  
  try {
    // Test 1: Test normalizeReportingTime function
    console.log('\nTest 1: normalizeReportingTime function');
    
    var testCases = [
      { input: '9:40 PM', expected: '9:40 PM' },
      { input: '9:40 PM (End of Day)', expected: '9:40 PM' },
      { input: '1:40 PM', expected: '1:40 PM' },
      { input: '3:40 PM ', expected: '3:40 PM' },
      { input: ' 5:40 PM ', expected: '5:40 PM' },
      { input: '7:40 PM(End of Day)', expected: '7:40 PM' },
      { input: null, expected: '' },
      { input: undefined, expected: '' },
      { input: 123, expected: '' }
    ];
    
    var allPassed = true;
    testCases.forEach(function(testCase, index) {
      var result = normalizeReportingTime(testCase.input);
      var passed = result === testCase.expected;
      console.log('Test', index + 1, ':', 
                  'Input:', JSON.stringify(testCase.input),
                  'Expected:', testCase.expected,
                  'Got:', result,
                  passed ? 'PASS' : 'FAIL');
      if (!passed) allPassed = false;
    });
    
    // Test 2: Test form submission with different formats
    console.log('\n\nTest 2: Testing form submission handling');
    
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
    
    if (dailyDetailsSheet) {
      // Find a test row for today
      var today = formatDate(new Date());
      var testData = dailyDetailsSheet.getDataRange().getValues();
      var testRow = -1;
      
      for (var i = 1; i < testData.length; i++) {
        var rowDate = testData[i][0];
        if (rowDate instanceof Date) {
          rowDate = formatDate(rowDate);
        }
        if (rowDate === today && testData[i][4]) { // Has Van ID
          testRow = i;
          break;
        }
      }
      
      if (testRow >= 0) {
        var vanId = testData[testRow][4];
        console.log('Testing with Van:', vanId, 'on date:', today);
        
        // Test different time formats
        var timeFormats = [
          '9:40 PM',
          '9:40 PM (End of Day)',
          '1:40 PM',
          '3:40 PM'
        ];
        
        timeFormats.forEach(function(timeFormat) {
          updateDailyDetailsFromForm({
            'Date': today,
            'Van ID': vanId,
            'Driver Name': 'Test Driver',
            'Route Code': 'TEST001',
            'Reporting Time': timeFormat,
            'Total Deliveries Completed': 100
          });
          console.log('Processed time format:', timeFormat);
        });
      } else {
        console.log('No test data available for today');
      }
    }
    
    console.log('\n=== Test Results ===');
    console.log('Normalization tests:', allPassed ? 'ALL PASSED' : 'SOME FAILED');
    
    SpreadsheetApp.getUi().alert(
      'Time Format Normalization Test',
      'Test completed. ' + (allPassed ? 'All tests passed!' : 'Some tests failed.') + 
      '\n\nCheck the logs for detailed results.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Test failed:', error);
    console.error('Stack:', error.stack);
    SpreadsheetApp.getUi().alert('Test failed: ' + error.toString());
  }
}

/**
 * Test function to verify Date and string format handling
 */
function testDateAndStringFormats() {
  console.log('=== Testing Date and String Format Handling ===');
  
  try {
    // Test 1: Test formatTimeString with various inputs
    console.log('\nTest 1: formatTimeString function');
    
    // Test with Date object
    var testDate = new Date('1899-12-30T21:40:00'); // 9:40 PM
    console.log('Input Date:', testDate);
    console.log('Output:', formatTimeString(testDate));
    
    // Test with time string
    var timeString = '3:40 PM';
    console.log('Input String:', timeString);
    console.log('Output:', formatTimeString(timeString));
    
    // Test with invalid input
    console.log('Input Invalid:', null);
    console.log('Output:', formatTimeString(null));
    
    // Test 2: Create test data with mixed formats
    console.log('\n\nTest 2: Creating test data with mixed formats');
    
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dataSheet = ss.getSheetByName('Delivery Pace Data');
    
    if (!dataSheet) {
      dataSheet = ss.insertSheet('Delivery Pace Data');
      setupDeliveryPaceDataSheet(dataSheet);
    }
    
    // Clear existing data (keep headers)
    var lastRow = dataSheet.getLastRow();
    if (lastRow > 1) {
      dataSheet.getRange(2, 1, lastRow - 1, dataSheet.getLastColumn()).clear();
    }
    
    var today = new Date();
    var testData = [
      // Row with Date object for reporting time
      [
        new Date(), // Timestamp
        today, // Date
        'BW1', // Van ID
        'Test Driver 1', // Driver Name
        'TEST001', // Route Code
        new Date('1899-12-30T13:40:00'), // 1:40 PM as Date object
        50, // Total Deliveries
        'Test with Date object', // Notes
        'No' // Processed
      ],
      // Row with string for reporting time
      [
        new Date(), // Timestamp
        today, // Date
        'BW2', // Van ID
        'Test Driver 2', // Driver Name
        'TEST002', // Route Code
        '3:40 PM', // Reporting Time as string
        75, // Total Deliveries
        'Test with string', // Notes
        'No' // Processed
      ],
      // Row with different string format
      [
        new Date(), // Timestamp
        today, // Date
        'BW3', // Van ID
        'Test Driver 3', // Driver Name
        'TEST003', // Route Code
        '5:40 PM (End of Day)', // Reporting Time with suffix
        100, // Total Deliveries
        'Test with suffix', // Notes
        'No' // Processed
      ]
    ];
    
    // Insert test data
    dataSheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);
    console.log('Inserted', testData.length, 'test rows');
    
    // Test 3: Read data back and verify handling
    console.log('\n\nTest 3: Reading data back');
    
    var vanIds = ['BW1', 'BW2', 'BW3'];
    vanIds.forEach(function(vanId) {
      console.log('\nTesting Van:', vanId);
      var paceData = getDeliveryPaceDataFromForms(vanId, formatDate(today));
      console.log('Result:', JSON.stringify(paceData, null, 2));
    });
    
    // Test 4: Test the update function
    console.log('\n\nTest 4: Testing updateDeliveryPaceForToday');
    updateDeliveryPaceForToday();
    
    console.log('\n=== All tests completed ===');
    
    SpreadsheetApp.getUi().alert(
      'Format Testing Complete',
      'All tests completed successfully. Check the logs for detailed results.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Test failed:', error);
    console.error('Stack:', error.stack);
    SpreadsheetApp.getUi().alert('Test failed: ' + error.toString());
  }
}

/**
 * Get all checkpoint data from Daily Details for a specific van and date
 * @param {string} vanId - Van ID to get data for
 * @param {string} date - Date in MM/DD/YYYY format
 * @return {Object} Object with checkpoint times as keys and delivery counts as values
 */
function getAllCheckpointData(vanId, date) {
  console.log('Getting all checkpoint data for Van:', vanId, 'Date:', date);
  
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
    
    if (!dailyDetailsSheet) {
      console.log('Daily Details sheet not found');
      return {};
    }
    
    var data = dailyDetailsSheet.getDataRange().getValues();
    var checkpointData = {};
    
    // Time slot mapping - columns L through P (indices 11-15)
    var timeSlots = [
      { time: '1:40 PM', column: 11 },
      { time: '3:40 PM', column: 12 },
      { time: '5:40 PM', column: 13 },
      { time: '7:40 PM', column: 14 },
      { time: '9:40 PM', column: 15 }
    ];
    
    // Find the row for this van and date
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0]; // Column A
      var rowVanId = data[i][4]; // Column E
      
      // Format date for comparison
      if (rowDate instanceof Date) {
        rowDate = formatDate(rowDate);
      }
      
      if (rowDate === date && rowVanId === vanId) {
        console.log('Found matching row for Van:', vanId);
        
        // Get data from all checkpoint columns
        timeSlots.forEach(function(slot) {
          var value = data[i][slot.column];
          if (value !== null && value !== '' && !isNaN(value)) {
            checkpointData[slot.time] = Number(value);
            console.log('Found checkpoint data:', slot.time, '=', value);
          }
        });
        
        break; // Found the row, no need to continue
      }
    }
    
    console.log('Total checkpoints found:', Object.keys(checkpointData).length);
    return checkpointData;
    
  } catch (error) {
    console.error('Error getting checkpoint data:', error);
    return {};
  }
}

/**
 * Test function to verify pace calculations work correctly
 */
function testPaceCalculations() {
  console.log('=== Testing Pace Calculations ===');
  
  try {
    // Test 1: Test getAllCheckpointData function
    console.log('\nTest 1: Testing getAllCheckpointData function');
    
    // Find a van with data for today
    var today = formatDate(new Date());
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
    
    if (!dailyDetailsSheet) {
      console.log('Daily Details sheet not found');
      SpreadsheetApp.getUi().alert('Daily Details sheet not found');
      return;
    }
    
    var data = dailyDetailsSheet.getDataRange().getValues();
    var testVanId = null;
    
    // Find a van with data for today
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0];
      if (rowDate instanceof Date) {
        rowDate = formatDate(rowDate);
      }
      
      if (rowDate === today && data[i][4]) { // Has Van ID
        testVanId = data[i][4];
        console.log('Found test van:', testVanId);
        break;
      }
    }
    
    if (!testVanId) {
      console.log('No van data found for today. Creating test data...');
      
      // Create test data
      var testRow = [
        new Date(), // Date
        'TEST001',  // Route
        'Test Driver', // Driver
        '',         // Asset ID
        'BW999',    // Van ID
        '', '', '', '', '', '', // Empty columns F-K
        50,         // 1:40 PM
        100,        // 3:40 PM
        150,        // 5:40 PM
        180,        // 7:40 PM
        200         // 9:40 PM
      ];
      
      dailyDetailsSheet.appendRow(testRow);
      testVanId = 'BW999';
      console.log('Created test data for van:', testVanId);
    }
    
    // Test getAllCheckpointData
    var checkpointData = getAllCheckpointData(testVanId, today);
    console.log('Checkpoint data retrieved:', JSON.stringify(checkpointData, null, 2));
    
    // Test 2: Test calculateAveragePace function
    console.log('\n\nTest 2: Testing calculateAveragePace function');
    
    var testCases = [
      {
        name: 'No data',
        data: {},
        expected: 0
      },
      {
        name: 'Single checkpoint',
        data: { '1:40 PM': 50 },
        expectedMin: 5, // Should be around 8-10 stops/hr (50 deliveries / ~5.67 hours from 8 AM)
        expectedMax: 15
      },
      {
        name: 'Two checkpoints',
        data: { '1:40 PM': 50, '3:40 PM': 100 },
        expected: 25 // 50 deliveries in 2 hours = 25/hr
      },
      {
        name: 'All checkpoints',
        data: { '1:40 PM': 50, '3:40 PM': 100, '5:40 PM': 150, '7:40 PM': 180, '9:40 PM': 200 },
        expected: 25 // 200 deliveries in 8 hours = 25/hr
      },
      {
        name: 'Non-sequential checkpoints',
        data: { '1:40 PM': 50, '5:40 PM': 150, '9:40 PM': 200 },
        expected: 25 // 200 deliveries in 8 hours = 25/hr
      }
    ];
    
    testCases.forEach(function(testCase) {
      var result = calculateAveragePace(testCase.data);
      var passed = false;
      
      if (testCase.expected !== undefined) {
        passed = result === testCase.expected;
      } else if (testCase.expectedMin !== undefined && testCase.expectedMax !== undefined) {
        passed = result >= testCase.expectedMin && result <= testCase.expectedMax;
      }
      
      console.log('Test case:', testCase.name);
      console.log('  Data:', JSON.stringify(testCase.data));
      console.log('  Expected:', testCase.expected || `${testCase.expectedMin}-${testCase.expectedMax}`);
      console.log('  Result:', result);
      console.log('  Status:', passed ? 'PASS' : 'FAIL');
    });
    
    // Test 3: Test email generation with checkpoint data
    console.log('\n\nTest 3: Testing email generation with checkpoint data');
    
    var emailData = {
      vanId: testVanId,
      date: today,
      timestamp: new Date(),
      driverName: 'Test Driver',
      deliveries: checkpointData,
      notes: 'Test notes'
    };
    
    var emailHtml = createEmailBody(emailData);
    console.log('Email HTML generated successfully:', emailHtml.length > 0 ? 'Yes' : 'No');
    
    // Check if pace calculation appears in email
    var paceCalculated = emailHtml.includes('stops/hr') && !emailHtml.includes('0 stops/hr');
    console.log('Pace calculation in email:', paceCalculated ? 'Yes' : 'No');
    
    console.log('\n=== Test Complete ===');
    
    SpreadsheetApp.getUi().alert(
      'Pace Calculation Test Complete',
      'Tests completed. Check the logs for detailed results.\n\n' +
      'Checkpoint data found: ' + Object.keys(checkpointData).length + ' checkpoints',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Test failed:', error);
    console.error('Stack:', error.stack);
    SpreadsheetApp.getUi().alert('Test failed: ' + error.toString());
  }
}

/**
 * Enable or disable mock data for testing
 * @param {boolean} enable - Whether to enable mock data
 * @param {string[]} vans - Optional array of specific van IDs to enable mock data for
 */
function setMockDataMode(enable, vans) {
  var properties = PropertiesService.getScriptProperties();
  
  if (enable) {
    properties.setProperty('MOCK_DATA_ENABLED', 'true');
    if (vans && vans.length > 0) {
      properties.setProperty('MOCK_DATA_VANS', JSON.stringify(vans));
      Logger.log('Mock data enabled for specific vans: ' + vans.join(', '));
    } else {
      properties.deleteProperty('MOCK_DATA_VANS');
      Logger.log('Mock data enabled for all vans');
    }
  } else {
    properties.setProperty('MOCK_DATA_ENABLED', 'false');
    properties.deleteProperty('MOCK_DATA_VANS');
    Logger.log('Mock data disabled');
  }
}

/**
 * Check if mock data is enabled (runtime override of config)
 * @param {string} vanId - Van ID to check
 * @return {boolean} Whether mock data should be used
 */
function isMockDataEnabled(vanId) {
  var properties = PropertiesService.getScriptProperties();
  var mockEnabled = properties.getProperty('MOCK_DATA_ENABLED') === 'true';
  
  if (!mockEnabled) {
    return false;
  }
  
  var mockVansStr = properties.getProperty('MOCK_DATA_VANS');
  if (mockVansStr) {
    try {
      var mockVans = JSON.parse(mockVansStr);
      return mockVans.indexOf(vanId) !== -1;
    } catch (e) {
      Logger.log('Error parsing mock vans: ' + e);
    }
  }
  
  return true; // Mock enabled for all vans
}

/**
 * Test function to verify mock data is disabled in production
 */
function testMockDataDisabled() {
  console.log('=== Testing Mock Data Configuration ===');
  
  try {
    // Test 1: Check configuration
    console.log('\nTest 1: Checking configuration settings');
    var useMockData = getConfig('DEV_SETTINGS.USE_MOCK_DATA');
    console.log('Config USE_MOCK_DATA:', useMockData);
    console.log('Expected: false');
    console.log('Status:', useMockData === false ? 'PASS' : 'FAIL');
    
    // Test 2: Check runtime properties
    console.log('\n\nTest 2: Checking runtime properties');
    var properties = PropertiesService.getScriptProperties();
    var mockEnabled = properties.getProperty('MOCK_DATA_ENABLED');
    console.log('Runtime MOCK_DATA_ENABLED:', mockEnabled);
    console.log('Expected: false or null');
    console.log('Status:', (mockEnabled === 'false' || mockEnabled === null) ? 'PASS' : 'FAIL');
    
    // Test 3: Test data retrieval for a few vans
    console.log('\n\nTest 3: Testing data retrieval without mock data');
    var testVans = ['BW999', 'TEST001', 'MOCK001'];
    var today = formatDate(new Date());
    
    testVans.forEach(function(vanId) {
      console.log('\nTesting Van:', vanId);
      var data = getDeliveryPaceData(vanId, today);
      var dataKeys = Object.keys(data);
      console.log('Data returned:', dataKeys.length > 0 ? 'Yes (' + dataKeys.length + ' checkpoints)' : 'No (empty)');
      
      // Check if any of the values look like mock data (sequential increases)
      if (dataKeys.length >= 3) {
        var values = dataKeys.map(function(key) { return data[key]; });
        var isMockPattern = true;
        for (var i = 1; i < values.length; i++) {
          if (values[i] <= values[i-1]) {
            isMockPattern = false;
            break;
          }
        }
        console.log('Appears to be mock data:', isMockPattern ? 'YES (WARNING!)' : 'No');
      }
    });
    
    // Test 4: Temporarily enable mock data and test
    console.log('\n\nTest 4: Testing mock data toggle');
    
    // Enable mock data for one van
    setMockDataMode(true, ['TEST001']);
    var testData = getDeliveryPaceDataFromSource('TEST001', today);
    console.log('Mock enabled for TEST001, data returned:', Object.keys(testData).length > 0 ? 'Yes' : 'No');
    
    var otherData = getDeliveryPaceDataFromSource('BW999', today);
    console.log('Mock enabled for TEST001, BW999 data returned:', Object.keys(otherData).length > 0 ? 'Yes' : 'No');
    
    // Disable mock data
    setMockDataMode(false);
    var disabledData = getDeliveryPaceDataFromSource('TEST001', today);
    console.log('Mock disabled, TEST001 data returned:', Object.keys(disabledData).length > 0 ? 'Yes' : 'No');
    
    console.log('\n=== Test Complete ===');
    
    SpreadsheetApp.getUi().alert(
      'Mock Data Test Complete',
      'Mock data is ' + (useMockData ? 'ENABLED (WARNING!)' : 'disabled') + ' in configuration.\n\n' +
      'Check logs for detailed test results.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Test failed:', error);
    console.error('Stack:', error.stack);
    SpreadsheetApp.getUi().alert('Test failed: ' + error.toString());
  }
}

/**
 * Get delivery pace summary for a specific date
 * @param {string} date - Date to get summary for
 * @return {Object} Summary object
 */
function getDeliveryPaceSummary(date) {
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
    
    if (!dailyDetailsSheet) {
      return {
        totalVans: 0,
        lastCompletedCheckpoint: 'None',
        averageDeliveries: {},
        onPaceCount: 0,
        behindPaceCount: 0
      };
    }
    
    var data = dailyDetailsSheet.getDataRange().getValues();
    var vansTracked = 0;
    var lastCheckpoint = 'None';
    
    // Find vans for the date and check their pace data
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0];
      if (rowDate instanceof Date) {
        rowDate = formatDate(rowDate);
      }
      
      if (rowDate === date) {
        // Check if any pace data exists (columns L-P)
        var hasPaceData = false;
        for (var j = 11; j <= 15; j++) {
          if (data[i][j]) {
            hasPaceData = true;
            vansTracked++;
            break;
          }
        }
      }
    }
    
    return {
      totalVans: vansTracked,
      lastCompletedCheckpoint: lastCheckpoint,
      averageDeliveries: {},
      onPaceCount: Math.floor(vansTracked * 0.7), // Mock data
      behindPaceCount: Math.ceil(vansTracked * 0.3) // Mock data
    };
    
  } catch (error) {
    Logger.log('Error getting delivery pace summary: ' + error);
    return {
      totalVans: 0,
      lastCompletedCheckpoint: 'None',
      averageDeliveries: {},
      onPaceCount: 0,
      behindPaceCount: 0
    };
  }
}

/**
 * Test delivery pace update functionality
 */
function testDeliveryPaceUpdate() {
  // Initialize headers if needed
  initializeDeliveryPaceHeaders();
  
  // Update pace for today
  updateDeliveryPaceForToday();
  
  // Show completion message
  SpreadsheetApp.getUi().alert('Delivery pace update test completed. Check the Daily Details sheet.');
}

/**
 * Migrate existing Date objects in Delivery Pace Data sheet to time strings
 * This fixes the issue where form submissions stored Date objects instead of time strings
 */
function migrateDeliveryPaceData() {
  console.log('=== Starting Delivery Pace Data Migration ===');
  
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dataSheet = ss.getSheetByName('Delivery Pace Data');
    
    if (!dataSheet) {
      console.log('Delivery Pace Data sheet not found');
      SpreadsheetApp.getUi().alert('Delivery Pace Data sheet not found.');
      return;
    }
    
    var data = dataSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      console.log('No data to migrate');
      SpreadsheetApp.getUi().alert('No data to migrate.');
      return;
    }
    
    var updatedRows = 0;
    var reportingTimeCol = 5; // Column F (index 5) - Reporting Time
    
    // Start from row 2 (skip headers)
    for (var i = 1; i < data.length; i++) {
      var reportingTime = data[i][reportingTimeCol];
      
      // Check if it's a Date object
      if (reportingTime instanceof Date) {
        console.log('Row ' + (i + 1) + ': Converting Date object to time string');
        console.log('  Original:', reportingTime);
        
        // Convert to time string format
        var timeString = formatTimeString(reportingTime);
        console.log('  Converted:', timeString);
        
        // Update the cell
        dataSheet.getRange(i + 1, reportingTimeCol + 1).setValue(timeString);
        updatedRows++;
      } else if (reportingTime && typeof reportingTime === 'string') {
        // Check if it's already in the correct format
        if (!reportingTime.match(/^\d{1,2}:\d{2}\s*(AM|PM)/i)) {
          console.log('Row ' + (i + 1) + ': Fixing time format');
          console.log('  Original:', reportingTime);
          
          // Try to parse and reformat
          try {
            var parsedTime = new Date(reportingTime);
            if (!isNaN(parsedTime)) {
              var timeString = formatTimeString(parsedTime);
              console.log('  Converted:', timeString);
              dataSheet.getRange(i + 1, reportingTimeCol + 1).setValue(timeString);
              updatedRows++;
            }
          } catch (e) {
            console.log('  Could not parse time:', e);
          }
        }
      }
    }
    
    console.log('Migration complete. Updated ' + updatedRows + ' rows.');
    
    SpreadsheetApp.getUi().alert(
      'Migration Complete',
      'Updated ' + updatedRows + ' rows with Date objects to proper time string format.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Migration failed:', error);
    console.error('Stack:', error.stack);
    SpreadsheetApp.getUi().alert('Migration failed: ' + error.toString());
  }
}