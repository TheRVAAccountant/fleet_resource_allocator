/**
 * ===================================================================
 * DIAGNOSTIC MAIN - TROUBLESHOOTING VERSION
 * ===================================================================
 * This file helps diagnose why menus aren't appearing
 */

var SCRIPT_VERSION = "2.1.0-diagnostic";

/**
 * Diagnostic onOpen that logs every step
 */
function onOpenDiagnostic() {
  var ui = SpreadsheetApp.getUi();
  var log = [];
  
  try {
    log.push("1. Starting onOpen at " + new Date().toString());
    
    // Test 1: Basic menu
    try {
      ui.createMenu("Diagnostic Menu 1")
        .addItem("View Log", "showDiagnosticResults")
        .addToUi();
      log.push("2. SUCCESS: Created Diagnostic Menu 1");
    } catch (e) {
      log.push("2. FAILED: Diagnostic Menu 1 - " + e.toString());
    }
    
    // Test 2: Menu with multiple items
    try {
      ui.createMenu("Diagnostic Menu 2")
        .addItem("Item 1", "diagnosticFunction1")
        .addItem("Item 2", "diagnosticFunction2")
        .addSeparator()
        .addItem("Item 3", "diagnosticFunction3")
        .addToUi();
      log.push("3. SUCCESS: Created Diagnostic Menu 2 with multiple items");
    } catch (e) {
      log.push("3. FAILED: Diagnostic Menu 2 - " + e.toString());
    }
    
    // Test 3: Menu with submenu
    try {
      ui.createMenu("Diagnostic Menu 3")
        .addItem("Direct Item", "diagnosticFunction1")
        .addSubMenu(ui.createMenu("Submenu")
          .addItem("Sub Item 1", "diagnosticFunction1")
          .addItem("Sub Item 2", "diagnosticFunction2"))
        .addToUi();
      log.push("4. SUCCESS: Created Diagnostic Menu 3 with submenu");
    } catch (e) {
      log.push("4. FAILED: Diagnostic Menu 3 with submenu - " + e.toString());
    }
    
    // Test 4: Check if specific functions exist
    var functionsToCheck = [
      "showUploadDialog",
      "initializeDeliveryPaceHeaders",
      "updateDeliveryPaceForToday",
      "generateTodaysSummary"
    ];
    
    functionsToCheck.forEach(function(funcName) {
      if (typeof this[funcName] === 'function') {
        log.push("5. Function exists: " + funcName);
      } else {
        log.push("5. Function MISSING: " + funcName);
      }
    });
    
    // Test 5: Try the actual menu structure
    try {
      ui.createMenu("Fleet Resource Allocator")
        .addItem("Upload Files for Allocation", "showUploadDialog")
        .addToUi();
      log.push("6. SUCCESS: Created Fleet Resource Allocator menu");
    } catch (e) {
      log.push("6. FAILED: Fleet Resource Allocator - " + e.toString());
    }
    
  } catch (error) {
    log.push("CRITICAL ERROR in onOpen: " + error.toString());
  }
  
  // Store the log
  PropertiesService.getScriptProperties().setProperty('diagnosticLog', JSON.stringify(log));
  
  // Also try to show immediate feedback
  try {
    ui.createMenu("Diagnostic Results")
      .addItem("View Results (" + log.length + " items)", "showDiagnosticResults")
      .addToUi();
  } catch (e) {
    // Even this failed!
  }
}

/**
 * Show diagnostic results
 */
function showDiagnosticResults() {
  var logJson = PropertiesService.getScriptProperties().getProperty('diagnosticLog');
  var log = logJson ? JSON.parse(logJson) : ["No diagnostic data found"];
  
  var message = "Diagnostic Results:\n\n" + log.join("\n");
  
  // Show in alert (limited to ~1000 chars)
  if (message.length > 1000) {
    message = message.substring(0, 1000) + "\n...(truncated)";
  }
  
  SpreadsheetApp.getUi().alert("Menu Diagnostic Results", message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Diagnostic functions for menu testing
 */
function diagnosticFunction1() {
  SpreadsheetApp.getUi().alert("Diagnostic", "Function 1 called successfully", SpreadsheetApp.getUi().ButtonSet.OK);
}

function diagnosticFunction2() {
  SpreadsheetApp.getUi().alert("Diagnostic", "Function 2 called successfully", SpreadsheetApp.getUi().ButtonSet.OK);
}

function diagnosticFunction3() {
  SpreadsheetApp.getUi().alert("Diagnostic", "Function 3 called successfully", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Nuclear option - create menus with all functions defined inline
 */
function createSelfContainedMenus() {
  var ui = SpreadsheetApp.getUi();
  
  // Remove any existing menus
  try {
    ["Vehicle Assignment Tool", "Fleet Resource Allocator", "Fleet Tools", "Diagnostic Menu 1", "Diagnostic Menu 2", "Diagnostic Menu 3"].forEach(function(menuName) {
      try {
        SpreadsheetApp.getActiveSpreadsheet().removeMenu(menuName);
      } catch (e) {
        // Menu might not exist
      }
    });
  } catch (e) {
    // Continue anyway
  }
  
  // Create new menus with inline functions
  ui.createMenu("Fleet Operations")
    .addItem("Upload Files", "inlineUpload")
    .addItem("Vehicle Status", "inlineVehicleStatus")
    .addItem("Daily Details", "inlineDailyDetails")
    .addToUi();
    
  ui.createMenu("Quick Actions")
    .addItem("Initialize Headers", "inlineInitHeaders")
    .addItem("Update Delivery Pace", "inlineUpdatePace")
    .addItem("Generate Report", "inlineGenerateReport")
    .addToUi();
    
  SpreadsheetApp.getUi().alert("Success", "Self-contained menus created. Check for 'Fleet Operations' and 'Quick Actions' menus.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// Inline functions that don't depend on any other files
function inlineUpload() {
  SpreadsheetApp.getUi().alert("Upload", "Upload functionality will be available soon", SpreadsheetApp.getUi().ButtonSet.OK);
}

function inlineVehicleStatus() {
  var ss = SpreadsheetApp.openById("1fgwW9tcozBqiB6zrpg7jzactFMkzpRXCcmPs0eUsaqI");
  var sheet = ss.getSheetByName("Vehicle Status");
  if (sheet) {
    SpreadsheetApp.getUi().alert("Vehicle Status", "Found " + sheet.getLastRow() + " vehicles in status sheet", SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("Error", "Vehicle Status sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function inlineDailyDetails() {
  var ss = SpreadsheetApp.openById("1fgwW9tcozBqiB6zrpg7jzactFMkzpRXCcmPs0eUsaqI");
  var sheet = ss.getSheetByName("Daily Details");
  if (sheet) {
    SpreadsheetApp.getUi().alert("Daily Details", "Found " + sheet.getLastRow() + " rows in daily details", SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("Error", "Daily Details sheet not found", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function inlineInitHeaders() {
  SpreadsheetApp.getUi().alert("Initialize Headers", "Headers initialization will be available soon", SpreadsheetApp.getUi().ButtonSet.OK);
}

function inlineUpdatePace() {
  SpreadsheetApp.getUi().alert("Update Pace", "Delivery pace update will be available soon", SpreadsheetApp.getUi().ButtonSet.OK);
}

function inlineGenerateReport() {
  SpreadsheetApp.getUi().alert("Generate Report", "Report generation will be available soon", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Check script version and environment
 */
function checkEnvironment() {
  var info = [];
  
  info.push("Script Version: " + SCRIPT_VERSION);
  info.push("Spreadsheet ID: " + SpreadsheetApp.getActiveSpreadsheet().getId());
  info.push("Script ID: " + ScriptApp.getScriptId());
  info.push("Time Zone: " + Session.getScriptTimeZone());
  info.push("User Email: " + Session.getActiveUser().getEmail());
  
  // Check if key functions exist
  var functions = ["onOpen", "showUploadDialog", "getConfig", "mainAllocation"];
  functions.forEach(function(fn) {
    info.push(fn + ": " + (typeof this[fn] === 'function' ? "EXISTS" : "MISSING"));
  });
  
  SpreadsheetApp.getUi().alert("Environment Check", info.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
}