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
  
  // Fallback to other data sources or mock data
  Logger.log("No form data found, using mock data for Van: " + vanId);
  return getDeliveryPaceDataFromSource(vanId, date);
}

/**
 * Get delivery pace data from external source
 * @param {string} vanId - Van ID
 * @param {string} date - Date string
 * @return {Object} Pace data by time slot
 */
function getDeliveryPaceDataFromSource(vanId, date) {
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
  
  // For demonstration, returning simulated progressive data
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
    var paceData = {
      "1:40 PM": null,
      "3:40 PM": null,
      "5:40 PM": null,
      "7:40 PM": null,
      "9:40 PM": null
    };
    
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
        // Map reporting time to our standard format
        var timeKey = reportingTime.replace(' (End of Day)', '');
        
        if (paceData.hasOwnProperty(timeKey)) {
          // Keep the latest submission for each time slot
          paceData[timeKey] = deliveryCount;
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
        
        // Update appropriate columns based on current time
        for (var j = 0; j < timeSlots.length; j++) {
          var slot = timeSlots[j];
          
          // Only update if current time is past the time slot
          if (currentTime >= slot.time) {
            var cellValue = paceData[slot.label];
            
            // Update the cell
            dailyDetailsSheet.getRange(i + 2, slot.column).setValue(cellValue);
            
            Logger.log("Updated Van " + vanId + " at " + slot.label + ": " + cellValue + " stops");
          }
        }
        
        updatedRows++;
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