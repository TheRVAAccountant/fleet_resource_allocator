/**
 * ===================================================================
 * MENU TEST AND DIAGNOSTICS
 * ===================================================================
 * Functions to test and diagnose menu creation issues
 */

/**
 * Simple test to create a basic menu
 * Run this from the Script Editor to test basic functionality
 */
function testBasicMenu() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Test Menu")
      .addItem("Test Item 1", "testFunction1")
      .addItem("Test Item 2", "testFunction2")
      .addToUi();
    
    ui.alert("Success", "Test menu created successfully!", ui.ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error", "Failed to create test menu: " + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Test function 1
 */
function testFunction1() {
  SpreadsheetApp.getUi().alert("Test 1", "Test function 1 executed successfully!", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Test function 2
 */
function testFunction2() {
  SpreadsheetApp.getUi().alert("Test 2", "Test function 2 executed successfully!", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Check all required files are present
 */
function checkRequiredFiles() {
  var requiredFiles = [
    'Main.js',
    'Config.js',
    'AllocationService.js',
    'DeliveryPaceService.js',
    'SheetService.js',
    'UIService.js',
    'Utils.js',
    'Logger.js',
    'ErrorHandler.js',
    'SheetManager.js'
  ];
  
  var missingFiles = [];
  var message = "File Check Results:\n\n";
  
  // In Google Apps Script, we can't directly check files
  // But we can check if key functions from each file exist
  var fileChecks = {
    'Main.js': 'onOpen',
    'Config.js': 'getConfig',
    'AllocationService.js': 'mainAllocation',
    'DeliveryPaceService.js': 'initializeDeliveryPaceHeaders',
    'SheetService.js': 'getVehicleStatusData',
    'UIService.js': 'showUploadDialog',
    'Utils.js': 'formatDate',
    'Logger.js': 'createLogger',
    'ErrorHandler.js': 'ErrorHandler',
    'SheetManager.js': 'SheetManager'
  };
  
  for (var file in fileChecks) {
    var funcName = fileChecks[file];
    if (typeof this[funcName] === 'function') {
      message += "✓ " + file + " - OK (found " + funcName + ")\n";
    } else {
      message += "✗ " + file + " - MISSING (couldn't find " + funcName + ")\n";
      missingFiles.push(file);
    }
  }
  
  message += "\n" + missingFiles.length + " files may be missing or not loaded properly.";
  
  SpreadsheetApp.getUi().alert("File Check", message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Force reload of onOpen
 */
function reloadMenus() {
  try {
    // Call onOpen directly
    onOpen();
    SpreadsheetApp.getUi().alert("Success", "Menus reloaded. Please check if they appear now.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error", "Failed to reload menus: " + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Create a minimal working menu
 */
function createMinimalMenu() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Create the absolute minimum menu
    ui.createMenu("Fleet Tools")
      .addItem("Upload Files", "showUploadDialog")
      .addItem("Check Functions", "checkRequiredFiles")
      .addItem("Test Basic Menu", "testBasicMenu")
      .addToUi();
      
    ui.alert("Success", "Minimal menu created. Look for 'Fleet Tools' menu.", ui.ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error", "Even minimal menu failed: " + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Debug function to check global scope
 */
function debugGlobalScope() {
  var globalFunctions = [];
  var count = 0;
  
  // Check for common functions
  var checkFunctions = [
    'onOpen', 'showUploadDialog', 'getConfig', 'mainAllocation',
    'initializeDeliveryPaceHeaders', 'updateDeliveryPaceForToday',
    'formatDate', 'createLogger', 'SheetManager', 'ErrorHandler'
  ];
  
  var message = "Global Scope Check:\n\n";
  
  checkFunctions.forEach(function(funcName) {
    if (typeof this[funcName] !== 'undefined') {
      var type = typeof this[funcName];
      message += funcName + ": " + type + "\n";
      if (type === 'function') count++;
    } else {
      message += funcName + ": UNDEFINED\n";
    }
  });
  
  message += "\n" + count + " functions found in global scope.";
  
  SpreadsheetApp.getUi().alert("Debug Info", message, SpreadsheetApp.getUi().ButtonSet.OK);
}