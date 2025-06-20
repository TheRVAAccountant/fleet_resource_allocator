/**
 * ===================================================================
 * ALLOCATION SERVICE
 * ===================================================================
 * Handles vehicle allocation logic, route matching, and results
 * generation for the fleet resource allocation system.
 */

/**
 * Main allocation function that orchestrates the entire process
 * @param {string} dayOfOpsId - Day of Ops spreadsheet ID
 * @param {string} dailyRoutesId - Daily Routes spreadsheet ID
 */
function mainAllocation(dayOfOpsId, dailyRoutesId) {
  Logger.log("------------------------------");
  Logger.log("Starting mainAllocation...");
  
  Logger.log("User-supplied DayOfOps ID: " + dayOfOpsId);
  Logger.log("User-supplied DailyRoutes ID: " + dailyRoutesId);
  
  // Load data from spreadsheets
  var dayOfOpsData = getSheetData(dayOfOpsId, getConfig('SHEETS.DAY_OF_OPS_SOLUTION'));
  Logger.log("Loaded DayOfOps data. Rows: " + (dayOfOpsData.length - 1));
  
  var vehicleStatusData = getSheetData(
    getConfig('DAILY_SUMMARY_SPREADSHEET_ID'), 
    getConfig('SHEETS.VEHICLE_STATUS')
  );
  Logger.log("Loaded Fleet (VehicleStatus) data. Rows: " + (vehicleStatusData.length - 1));
  
  var routesData = getSheetData(dailyRoutesId, getConfig('SHEETS.DAILY_ROUTES'));
  Logger.log("Loaded DailyRoutes data. Rows: " + (routesData.length - 1));
  
  // Verify required columns
  verifyRequiredColumns(
    dayOfOpsData[0],
    getConfig('REQUIRED_COLUMNS.DAY_OF_OPS'),
    "DayOfOps (" + getConfig('SHEETS.DAY_OF_OPS_SOLUTION') + ")"
  );
  verifyRequiredColumns(
    vehicleStatusData[0],
    getConfig('REQUIRED_COLUMNS.VEHICLE_STATUS'),
    "VehicleStatus (" + getConfig('SHEETS.VEHICLE_STATUS') + ")"
  );
  verifyRequiredColumns(
    routesData[0],
    getConfig('REQUIRED_COLUMNS.DAILY_ROUTES'),
    "DailyRoutes (" + getConfig('SHEETS.DAILY_ROUTES') + ")"
  );
  
  // Convert to object arrays
  var dayOfOpsObj = convertToObjectArray(dayOfOpsData);
  var fleetObj = convertToObjectArray(vehicleStatusData);
  var dailyRoutesObj = convertToObjectArray(routesData);
  
  // Filter for target DSP
  Logger.log("Filtering Day of Ops to DSP == '" + getConfig('TARGET_DSP') + "'...");
  dayOfOpsObj = dayOfOpsObj.filter(function(r) { 
    return r["DSP"] === getConfig('TARGET_DSP'); 
  });
  Logger.log("Remaining " + getConfig('TARGET_DSP') + " routes: " + dayOfOpsObj.length);
  
  // Allocate vehicles
  Logger.log("Allocating vehicles to routes...");
  var allocationOutcome = allocateVehiclesToRoutes(dayOfOpsObj, fleetObj);
  var allocationResults = allocationOutcome.allocationResults;
  var assignedVanIds = allocationOutcome.assignedVanIds;
  Logger.log("Completed allocation. Routes allocated: " + allocationResults.length);
  
  // Create results sheets
  var now = new Date();
  var dateStamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM-dd-yy");
  var resultsSheetName = dateStamp + " - Results";
  
  Logger.log("Creating/Overwriting sheet: " + resultsSheetName + " in Daily Summary...");
  createResultsSheet(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'), resultsSheetName, allocationResults);
  Logger.log("Results sheet created successfully.");
  
  // Update with associate names
  Logger.log("Updating Associate Names from the daily routes...");
  updateResultsWithDailyRoutes(
    getConfig('DAILY_SUMMARY_SPREADSHEET_ID'), 
    resultsSheetName, 
    dailyRoutesObj
  );
  Logger.log("Associate Names updated successfully.");
  
  // Update Daily Details
  Logger.log("Updating Daily Details sheet with new allocation data...");
  updateDailyDetails(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'), resultsSheetName);
  Logger.log("Daily Details sheet updated with new allocation data.");
  
  // Create unassigned vans sheet
  Logger.log("Determining unassigned vans from 'Vehicle Status'...");
  var unassigned2D = extractUnassignedRows(vehicleStatusData, assignedVanIds);
  
  var unassignedSheetName = dateStamp + " - Available & Unassigned Vans";
  Logger.log("Creating new sheet '" + unassignedSheetName + "' in Daily Summary...");
  createUnassignedSheetInDailySummary(
    getConfig('DAILY_SUMMARY_SPREADSHEET_ID'), 
    unassignedSheetName, 
    unassigned2D
  );
  Logger.log("Unassigned sheet created in Daily Summary.");
  
  // Create Route Assignments file
  Logger.log("Creating separate 'Route Assignments' file...");
  var routeAssignmentsFileId = createRouteAssignmentsFile(
    resultsSheetName,
    unassignedSheetName,
    getConfig('DAILY_SUMMARY_SPREADSHEET_ID'),
    unassigned2D,
    getConfig('ROUTE_ASSIGNMENTS_FOLDER_ID')
  );
  Logger.log("Route Assignments file created. ID: " + routeAssignmentsFileId);
  
  Logger.log("mainAllocation completed successfully.");
  showAllocationCompleteAlert(resultsSheetName, unassignedSheetName, routeAssignmentsFileId);
  
  Logger.log("------------------------------");
}

/**
 * Allocates vehicles to routes based on type matching
 * @param {Object[]} dayOfOpsObj - Route objects
 * @param {Object[]} fleetObj - Fleet vehicle objects
 * @return {Object} Allocation results and assigned van IDs
 */
function allocateVehiclesToRoutes(dayOfOpsObj, fleetObj) {
  Logger.log("Filtering fleet for operational vehicles (Opnal? Y/N == 'Y')...");
  var operationalFleet = fleetObj.filter(function(f) {
    return f["Opnal?\nY/N"] === "Y";
  });
  Logger.log("Operational fleet count: " + operationalFleet.length);
  
  var fleetGroups = groupBy(operationalFleet, "Type");
  var allocationResults = [];
  var assignedVanIds = [];
  
  dayOfOpsObj.forEach(function(route) {
    var serviceType = route["Service Type"];
    var requiredVanType = getVanType(serviceType);
    
    if (!requiredVanType) {
      Logger.log("No matching van type for Service Type: " + serviceType +
        " (Route: " + route["Route Code"] + "). Skipping.");
      return;
    }
    
    if (!fleetGroups[requiredVanType] || fleetGroups[requiredVanType].length === 0) {
      Logger.log("No available vehicle of type '" + requiredVanType +
        "' for route: " + route["Route Code"]);
      return;
    }
    
    var assignedVehicle = fleetGroups[requiredVanType].shift();
    assignedVanIds.push(assignedVehicle["Van ID"]);
    
    var result = {
      "Route Code": route["Route Code"],
      "Service Type": serviceType,
      "DSP": route["DSP"],
      "Wave": route["Wave"],
      "Staging Location": route["Staging Location"],
      "Van ID": assignedVehicle["Van ID"],
      "Device Name": assignedVehicle["Van ID"],
      "Van Type": assignedVehicle["Type"],
      "Operational": assignedVehicle["Opnal?\nY/N"],
      "Associate Name": ""
    };
    
    allocationResults.push(result);
    Logger.log("Assigned Van '" + assignedVehicle["Van ID"] +
      "' (type: " + assignedVehicle["Type"] +
      ") to route '" + route["Route Code"] + "'");
  });
  
  return {
    allocationResults: allocationResults,
    assignedVanIds: assignedVanIds
  };
}

/**
 * Updates results sheet with associate names from daily routes
 * @param {string} spreadsheetId - Daily Summary spreadsheet ID
 * @param {string} resultsSheetName - Results sheet name
 * @param {Object[]} dailyRoutesObj - Daily routes data
 */
function updateResultsWithDailyRoutes(spreadsheetId, resultsSheetName, dailyRoutesObj) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(resultsSheetName);
  
  if (!sheet) {
    throw new Error("Results sheet '" + resultsSheetName + "' not found in daily summary!");
  }
  
  // Create route to driver mapping
  var routeToDriver = {};
  dailyRoutesObj.forEach(function(r) {
    var rc = r["Route code"];
    routeToDriver[rc] = r["Driver name"] || "N/A";
  });
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No data in Results sheet; skipping driver assignment.");
    return;
  }
  
  // Get and update data
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 11);
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; i++) {
    var routeCode = data[i][0];
    var associateName = routeToDriver[routeCode] || "N/A";
    data[i][9] = associateName;
    
    // Rebuild unique identifier
    var uniqueIdentifier = getTodayString() + "|" +
      data[i][0] + "|" +
      associateName + "|" +
      data[i][5];
    data[i][10] = uniqueIdentifier;
    Logger.log("Updated Unique Identifier for route: " + routeCode);
  }
  
  dataRange.setValues(data);
  Logger.log("Updated results sheet '" + resultsSheetName + "' with Associate Names and Unique Identifiers.");
}

/**
 * Identifies unassigned operational vans
 * @param {Array[]} vehicleStatusData - Vehicle status data
 * @param {string[]} assignedVanIds - IDs of assigned vans
 * @return {Array[]} 2D array of unassigned van data
 */
function extractUnassignedRows(vehicleStatusData, assignedVanIds) {
  var headerRow = vehicleStatusData[0];
  var colCount = 15;
  var vanIdCol = headerRow.indexOf("Van ID");
  var opCol = headerRow.indexOf("Opnal?\nY/N");
  
  if (vanIdCol === -1 || opCol === -1) {
    throw new Error("Could not find 'Van ID' or 'Opnal?\nY/N' in Vehicle Status header.");
  }
  
  var out = [];
  out.push(headerRow.slice(0, colCount));
  
  var assignedSet = new Set(assignedVanIds);
  
  for (var r = 1; r < vehicleStatusData.length; r++) {
    var row = vehicleStatusData[r];
    if (!row || row.length < colCount) continue;
    
    var opValue = row[opCol];
    var thisVanId = row[vanIdCol];
    
    if (opValue === "Y" && !assignedSet.has(thisVanId)) {
      out.push(row.slice(0, colCount));
    }
  }
  
  Logger.log("Unassigned operational vans count (excluding header): " + (out.length - 1));
  return out;
}

/**
 * Creates the Route Assignments file
 * @param {string} resultsSheetName - Results sheet name
 * @param {string} unassignedSheetName - Unassigned sheet name
 * @param {string} dailySummarySpreadsheetId - Daily Summary ID
 * @param {Array[]} unassigned2D - Unassigned data
 * @param {string} folderId - Folder ID for the new file
 * @return {string} File ID of created spreadsheet
 */
function createRouteAssignmentsFile(resultsSheetName, unassignedSheetName, 
                                   dailySummarySpreadsheetId, unassigned2D, folderId) {
  var ss = SpreadsheetApp.openById(dailySummarySpreadsheetId);
  var resultsSheet = ss.getSheetByName(resultsSheetName);
  
  if (!resultsSheet) {
    throw new Error("No results sheet named '" + resultsSheetName + "'");
  }
  
  var lastRow = resultsSheet.getLastRow();
  var lastCol = resultsSheet.getLastColumn();
  
  if (lastRow < 1 || lastCol < 1) {
    throw new Error("Results sheet is empty.");
  }
  
  // Column reordering
  var columnOrder = [4, 10, 6, 7, 1, 5, 2, 8, 9, 3];
  var displayData = resultsSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var reorderedData = [];
  
  for (var r = 0; r < displayData.length; r++) {
    var newRow = [];
    for (var c = 0; c < columnOrder.length; c++) {
      var oldColIndex = columnOrder[c] - 1;
      if (oldColIndex < displayData[r].length) {
        newRow.push(displayData[r][oldColIndex]);
      } else {
        newRow.push('');
        Logger.log("Warning: Column index out of bounds for row " + r + ", column " + c);
      }
    }
    reorderedData.push(newRow);
  }
  
  // Create new spreadsheet
  var fileName = getTimestampString() + " - Route Assignments";
  Logger.log("Creating new spreadsheet: " + fileName);
  var newSS = SpreadsheetApp.create(fileName);
  var sheet = newSS.getActiveSheet();
  sheet.setName("Route Assignments");
  
  // Set data and format
  sheet.getRange(1, 1, reorderedData.length, columnOrder.length).setValues(reorderedData);
  sheet.getRange(1, 1, reorderedData.length, columnOrder.length).setNumberFormat("@");
  
  if (reorderedData.length > 0) {
    formatHeaderRow(sheet, 1, columnOrder.length);
    sheet.autoResizeColumns(1, columnOrder.length);
  }
  
  // Add unassigned vans sheet
  var unassignedSheet = newSS.insertSheet(unassignedSheetName + " - Summary Data", 2);
  if (unassigned2D.length > 0) {
    unassignedSheet
      .getRange(1, 1, unassigned2D.length, unassigned2D[0].length)
      .setValues(unassigned2D)
      .setNumberFormat("@");
    
    formatHeaderRow(unassignedSheet, 1, unassigned2D[0].length);
    unassignedSheet.autoResizeColumns(1, unassigned2D[0].length);
  }
  
  // Move to folder
  var folder = DriveApp.getFolderById(folderId);
  folder.addFile(DriveApp.getFileById(newSS.getId()));
  
  var newFileId = newSS.getId();
  Logger.log("Created 'Route Assignments' spreadsheet with ID: " + newFileId +
    " in folder: " + folderId);
  
  return newFileId;
}