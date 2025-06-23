/**
 * ===================================================================
 * MAIN ENTRY POINTS AND MENU SETUP
 * ===================================================================
 * Handles application initialization, menu creation, and primary
 * entry points for user interactions.
 */

/**
 * Called automatically when the spreadsheet is opened.
 * Sets up the application menus.
 */
function onOpen() {
  try {
    setupAllMenus();
  } catch (error) {
    // Fallback to basic menu
    SpreadsheetApp.getUi()
      .createMenu("Fleet System")
      .addItem("Initialize System", "initializeSystem")
      .addItem("Show Error", "showLastError")
      .addToUi();
    
    // Store error for debugging
    PropertiesService.getScriptProperties().setProperty('lastMenuError', error.toString());
  }
}

/**
 * Set up all application menus
 */
function setupAllMenus() {
  var ui = SpreadsheetApp.getUi();
  
  // Main menu - Fleet Resource Allocator
  ui.createMenu('Fleet Resource Allocator')
    .addItem('Upload Files for Allocation', 'showUploadDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('Quick Actions')
      .addItem('Check Vehicle Status', 'showVehicleStatus')
      .addItem('View Daily Details', 'showDailyDetails')
      .addItem('Generate Report', 'quickReport'))
    .addToUi();
  
  // Delivery Pace menu
  ui.createMenu('Delivery Pace')
    .addItem('Initialize Headers', 'initializeDeliveryPaceHeaders')
    .addItem('Update Today\'s Pace', 'updateDeliveryPaceForToday')
    .addItem('Generate Today\'s Summary', 'generateTodaysSummary')
    .addSeparator()
    .addItem('Update Specific Van', 'showUpdateVanDialog')
    .addItem('Setup Auto-Update Triggers', 'setupDeliveryPaceTriggers')
    .addItem('Test Update', 'testDeliveryPaceUpdate')
    .addSeparator()
    .addSubMenu(ui.createMenu('Form Management')
      .addItem('Create/Update Collection Form', 'createDeliveryForm')
      .addItem('Get Form Link & QR Code', 'showFormInfo')
      .addItem('Setup Form Trigger', 'setupFormSubmitTrigger')
      .addItem('Test Van Filtering', 'testFormVanFiltering'))
    .addToUi();
  
  // Reports menu
  ui.createMenu('Reports')
    .addItem('Vehicle Utilization Report', 'generateVehicleUtilizationReport')
    .addItem('Driver Performance Report', 'generateDriverPerformanceReport')
    .addItem('Weekly Summary Report', 'generateWeeklySummaryReport')
    .addSeparator()
    .addItem('Analytics Dashboard', 'showAnalyticsDashboard')
    .addItem('Export All Data', 'exportAllData')
    .addToUi();
  
  // Forms menu
  ui.createMenu('Forms')
    .addItem('Delivery Pace Form', 'showDeliveryPaceForm')
    .addItem('RTS Form', 'showRTSForm')
    .addSeparator()
    .addItem('Form Management', 'showFormManagement')
    .addToUi();
  
  // Administration menu
  ui.createMenu('Administration')
    .addItem('Dashboard', 'showDashboard')
    .addItem('View Error Log', 'showErrorLog')
    .addItem('Clear Error Log', 'clearErrorLog')
    .addSeparator()
    .addItem('Test Logger', 'testLogger')
    .addItem('Run Diagnostics', 'runDiagnostics')
    .addItem('Initialize All Functions', 'initializeAllFunctions')
    .addSeparator()
    .addItem('User Guide', 'showUserGuide')
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Initialize system - fallback function
 */
function initializeSystem() {
  try {
    setupAllMenus();
    SpreadsheetApp.getUi().alert(
      'System Initialized',
      'All menus have been created. Please refresh the spreadsheet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Initialization Error',
      'Error: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Show last menu creation error
 */
function showLastError() {
  var error = PropertiesService.getScriptProperties().getProperty('lastMenuError') || 'No error recorded';
  SpreadsheetApp.getUi().alert('Last Menu Error', error, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Run diagnostics to understand why menus aren't appearing
 */
function runDiagnostics() {
  var ui = SpreadsheetApp.getUi();
  var report = [];
  
  report.push("=== MENU CREATION DIAGNOSTICS ===");
  report.push("Time: " + new Date().toString());
  report.push("");
  
  // Test 1: Check if we can create a test menu
  try {
    ui.createMenu("Test Menu")
      .addItem("Test Item", "showUploadDialog")
      .addToUi();
    report.push("✓ Can create menus");
  } catch (e) {
    report.push("✗ Cannot create menus: " + e.toString());
  }
  
  // Test 2: Check key functions
  var functions = [
    "showUploadDialog",
    "initializeDeliveryPaceHeaders", 
    "updateDeliveryPaceForToday",
    "getConfig",
    "mainAllocation"
  ];
  
  report.push("");
  report.push("Function Check:");
  functions.forEach(function(fn) {
    if (typeof this[fn] === 'function') {
      report.push("✓ " + fn + " exists");
    } else {
      report.push("✗ " + fn + " is missing");
    }
  });
  
  // Test 3: Check script properties
  report.push("");
  report.push("Script Properties:");
  try {
    var props = PropertiesService.getScriptProperties().getKeys();
    report.push("Properties found: " + props.length);
  } catch (e) {
    report.push("Cannot access properties: " + e.toString());
  }
  
  // Show results
  ui.alert("Diagnostic Results", report.join("\n"), ui.ButtonSet.OK);
}

/**
 * Force create all menus - bypass onOpen
 */
function forceCreateAllMenus() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Create each menu individually
    ui.createMenu("Vehicle Assignment")
      .addItem("Upload Files", "showUploadDialog")
      .addToUi();
    
    ui.createMenu("Delivery Pace")
      .addItem("Initialize Headers", "initializeDeliveryPaceHeaders")
      .addItem("Update Today's Pace", "updateDeliveryPaceForToday")
      .addItem("Generate Summary", "generateTodaysSummary")
      .addToUi();
    
    ui.createMenu("Reports")
      .addItem("Vehicle Utilization", "generateVehicleUtilizationReport")
      .addItem("Driver Performance", "generateDriverPerformanceReport")
      .addItem("Export All Data", "exportAllData")
      .addToUi();
    
    ui.createMenu("Administration")
      .addItem("View Error Log", "showErrorLog")
      .addItem("Form Management", "showFormManagement")
      .addToUi();
    
    ui.alert("Success", "All menus have been created. You may need to refresh the spreadsheet to see them all.", ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert("Error", "Failed to create menus: " + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Alternative simple menu structure - one menu at a time
 */
function createBasicMenus() {
  var ui = SpreadsheetApp.getUi();
  
  // Create each menu separately to isolate failures
  try {
    ui.createMenu("Vehicle Assignment")
      .addItem("Upload Files", "showUploadDialog")
      .addToUi();
  } catch (e) {
    console.log("Failed to create Vehicle Assignment menu: " + e.toString());
  }
  
  try {
    ui.createMenu("Delivery Pace")
      .addItem("Initialize Headers", "initializeDeliveryPaceHeaders")
      .addItem("Update Today", "updateDeliveryPaceForToday")
      .addToUi();
  } catch (e) {
    console.log("Failed to create Delivery Pace menu: " + e.toString());
  }
  
  try {
    ui.createMenu("Reports")
      .addItem("Vehicle Report", "generateVehicleUtilizationReport")
      .addItem("Export Data", "exportAllData")
      .addToUi();
  } catch (e) {
    console.log("Failed to create Reports menu: " + e.toString());
  }
  
  ui.alert("Basic menus created. Check for any that might be missing.");
}

/**
 * Initialize all functions that might be missing
 * This ensures all menu items have corresponding functions
 */
function initializeAllFunctions() {
  // First, let's make sure all the required functions exist
  var missingFunctions = [];
  var requiredFunctions = [
    'showUploadDialog',
    'initializeDeliveryPaceHeaders',
    'updateDeliveryPaceForToday',
    'generateTodaysSummary',
    'showUpdateVanDialog',
    'setupDeliveryPaceTriggers',
    'testDeliveryPaceUpdate',
    'generateVehicleUtilizationReport',
    'generateDriverPerformanceReport',
    'generateWeeklySummaryReport',
    'showAnalyticsDashboard',
    'exportAllData',
    'showDeliveryPaceForm',
    'showRTSForm',
    'showFormManagement',
    'showErrorLog',
    'showUserGuide',
    'showAbout'
  ];
  
  // Check which functions are missing
  for (var i = 0; i < requiredFunctions.length; i++) {
    var funcName = requiredFunctions[i];
    if (typeof this[funcName] !== 'function') {
      missingFunctions.push(funcName);
    }
  }
  
  if (missingFunctions.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Missing Functions Found',
      'The following functions are missing:\n\n' + missingFunctions.join('\n') + 
      '\n\nThese need to be defined for the menus to work properly.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      'All Functions Present',
      'All required functions are defined. Try running createBasicMenus() to create the menus.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Test function to verify the script is loaded
 */
function testScriptLoaded() {
  SpreadsheetApp.getUi().alert('Script Loaded', 'The Main.js script is loaded and working.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Main allocation entry point - called from UI
 * @param {string} dayOfOpsId - File ID for Day of Ops spreadsheet
 * @param {string} dailyRoutesId - File ID for Daily Routes spreadsheet
 */
function runAllocation(dayOfOpsId, dailyRoutesId) {
  try {
    mainAllocation(dayOfOpsId, dailyRoutesId);
  } catch (err) {
    Logger.log("Error in runAllocation: " + err);
    SpreadsheetApp.getUi().alert("Error during allocation: " + err);
  }
}

/**
 * Generate delivery pace summary for today - menu handler
 */
function generateTodaysSummary() {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  generateDeliveryPaceSummary(today);
}

/**
 * Test delivery pace updates - menu handler
 */
function testDeliveryPaceUpdate() {
  // Initialize headers if needed
  initializeDeliveryPaceHeaders();
  
  // Update pace for today
  updateDeliveryPaceForToday();
  
  // Show completion message
  SpreadsheetApp.getUi().alert("Delivery pace update completed. Check the Daily Details sheet.");
}

/**
 * Create or update the delivery pace collection form
 */
function createDeliveryForm() {
  try {
    var formUrl = createDeliveryPaceForm();
    SpreadsheetApp.getUi().alert(
      "Form Created Successfully!",
      "Delivery Pace Collection Form has been created/updated.\n\n" +
      "Form URL: " + formUrl + "\n\n" +
      "Share this link with drivers or use 'Get Form Link & QR Code' to generate a QR code.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error creating form: " + error.toString());
  }
}

/**
 * Show form information and QR code
 */
function showFormInfo() {
  try {
    var info = generateFormQRCode();
    
    var html = HtmlService.createHtmlOutput(
      '<div style="padding: 20px; text-align: center;">' +
      '<h3>Delivery Pace Collection Form</h3>' +
      '<p><strong>Form URL:</strong><br>' +
      '<a href="' + info.formUrl + '" target="_blank">' + info.formUrl + '</a></p>' +
      
      '<p><strong>QR Code:</strong><br>' +
      '<img src="' + info.qrCodeUrl + '" alt="QR Code" style="margin: 10px auto;">' +
      '</p>' +
      
      '<p style="font-size: 12px; color: #666;">' +
      'Drivers can scan this QR code with their mobile devices<br>' +
      'to quickly access the delivery pace reporting form.' +
      '</p>' +
      '</div>'
    )
    .setWidth(400)
    .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Form Information');
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error showing form info: " + error.toString());
  }
}

/**
 * Show dialog to update specific van
 */
function showUpdateVanDialog() {
  var html = HtmlService.createHtmlOutput(
    '<div style="padding: 20px;">' +
    '<p>Enter Van ID to update delivery pace:</p>' +
    '<input type="text" id="vanId" placeholder="e.g., BW1" style="width: 200px; padding: 5px;">' +
    '<br><br>' +
    '<button onclick="updateVan()" style="padding: 5px 15px;">Update</button>' +
    '<button onclick="google.script.host.close()" style="padding: 5px 15px; margin-left: 10px;">Cancel</button>' +
    '</div>' +
    '<script>' +
    'function updateVan() {' +
    '  var vanId = document.getElementById("vanId").value;' +
    '  if (vanId) {' +
    '    google.script.run' +
    '      .withSuccessHandler(function() { google.script.host.close(); })' +
    '      .withFailureHandler(function(error) { alert("Error: " + error); })' +
    '      .updateDeliveryPaceForVan(vanId, new Date());' +
    '  } else {' +
    '    alert("Please enter a Van ID");' +
    '  }' +
    '}' +
    '</script>'
  )
  .setWidth(300)
  .setHeight(150);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Van Delivery Pace');
}

// ===================================================================
// STUB FUNCTIONS FOR MENU ITEMS
// ===================================================================
// These ensure all menu items have corresponding functions

/**
 * Show dashboard interface
 */
function showDashboard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('Dashboard')
      .setWidth(800)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Fleet Operations Dashboard');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Dashboard is not available. Error: ' + error.toString());
  }
}

/**
 * Show delivery pace form
 */
function showDeliveryPaceForm() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('DeliveryPaceForm')
      .setWidth(600)
      .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Delivery Pace Report');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Delivery Pace form is not available. Error: ' + error.toString());
  }
}

/**
 * Show RTS (Return to Station) form
 */
function showRTSForm() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('RTSForm')
      .setWidth(600)
      .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'End of Day RTS Report');
  } catch (error) {
    SpreadsheetApp.getUi().alert('RTS form is not available. Error: ' + error.toString());
  }
}

/**
 * Show error log
 */
function showErrorLog() {
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var errorSheet = ss.getSheetByName('Error Log');
    
    var errorLog = 'No errors logged.';
    
    if (errorSheet && errorSheet.getLastRow() > 1) {
      var data = errorSheet.getRange(2, 1, Math.min(20, errorSheet.getLastRow() - 1), 5).getValues();
      errorLog = data.map(function(row) {
        return row[0] + ' | ' + row[1] + ' | ' + row[2];
      }).join('\n');
    }
    
    var html = HtmlService.createHtmlOutput(
      '<div style="padding: 20px; font-family: monospace;">' +
      '<h3>Recent Errors</h3>' +
      '<pre style="white-space: pre-wrap; word-wrap: break-word;">' + errorLog + '</pre>' +
      '</div>'
    )
    .setWidth(700)
    .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Error Log');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error log not available: ' + error.toString());
  }
}

/**
 * Generate vehicle utilization report
 */
function generateVehicleUtilizationReport() {
  try {
    SpreadsheetApp.getUi().alert(
      'Vehicle Utilization Report',
      'This feature will analyze vehicle usage patterns.\n\nComing soon!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Show analytics dashboard
 */
function showAnalyticsDashboard() {
  try {
    if (typeof generateAnalyticsDashboard === 'function') {
      generateAnalyticsDashboard();
    } else {
      SpreadsheetApp.getUi().alert(
        'Analytics Dashboard',
        'Analytics dashboard feature coming soon!',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Show user guide
 */
function showUserGuide() {
  var html = HtmlService.createHtmlOutput(
    '<div style="padding: 20px; font-family: Arial, sans-serif;">' +
    '<h2>Fleet Resource Allocator User Guide</h2>' +
    '<h3>Getting Started</h3>' +
    '<ol>' +
    '<li><strong>Upload Files:</strong> Use "Fleet Resource Allocator" → "Upload Files for Allocation"</li>' +
    '<li><strong>Select Files:</strong> Choose Day of Ops and Daily Routes Excel files</li>' +
    '<li><strong>Wait for Processing:</strong> The system will allocate vehicles automatically</li>' +
    '<li><strong>Review Results:</strong> Check the Results sheet for allocations</li>' +
    '</ol>' +
    '<h3>Features</h3>' +
    '<ul>' +
    '<li><strong>Delivery Pace:</strong> Track delivery progress throughout the day</li>' +
    '<li><strong>Reports:</strong> Generate utilization and performance reports</li>' +
    '<li><strong>Forms:</strong> Submit delivery pace and RTS reports</li>' +
    '</ul>' +
    '</div>'
  )
  .setWidth(600)
  .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'User Guide');
}

/**
 * Show about dialog
 */
function showAbout() {
  var html = HtmlService.createHtmlOutput(
    '<div style="padding: 20px; text-align: center; font-family: Arial, sans-serif;">' +
    '<h2>Fleet Resource Allocator</h2>' +
    '<p>Version 2.0</p>' +
    '<p>Automated vehicle assignment system for delivery operations</p>' +
    '<hr>' +
    '<p><strong>Features:</strong></p>' +
    '<ul style="text-align: left; display: inline-block;">' +
    '<li>Automated vehicle allocation</li>' +
    '<li>Delivery pace tracking</li>' +
    '<li>Performance analytics</li>' +
    '<li>Comprehensive reporting</li>' +
    '</ul>' +
    '<hr>' +
    '<p>© 2025 Fleet Operations</p>' +
    '</div>'
  )
  .setWidth(400)
  .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'About');
}

/**
 * Generate driver performance report
 * Stub function that calls the actual implementation if available
 */
function generateDriverPerformanceReport() {
  try {
    // Try to call the actual function from ReportService
    if (typeof ReportService !== 'undefined' && typeof ReportService.generateDriverPerformanceReport === 'function') {
      ReportService.generateDriverPerformanceReport();
    } else {
      SpreadsheetApp.getUi().alert(
        'Driver Performance Report',
        'This feature analyzes driver performance metrics.\n\nComing soon!',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error generating report: ' + error.toString());
  }
}

/**
 * Generate weekly summary report
 * Stub function that calls the actual implementation if available
 */
function generateWeeklySummaryReport() {
  try {
    // Try to call the actual function from ReportService
    if (typeof ReportService !== 'undefined' && typeof ReportService.generateWeeklySummaryReport === 'function') {
      ReportService.generateWeeklySummaryReport();
    } else {
      SpreadsheetApp.getUi().alert(
        'Weekly Summary Report',
        'This feature generates a comprehensive weekly summary.\n\nComing soon!',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error generating report: ' + error.toString());
  }
}

/**
 * Export all data
 * Stub function that calls the actual implementation if available
 */
function exportAllData() {
  try {
    // Try to call the actual function from DataExportService
    if (typeof DataExportService !== 'undefined' && typeof DataExportService.exportAllData === 'function') {
      DataExportService.exportAllData();
    } else {
      SpreadsheetApp.getUi().alert(
        'Export All Data',
        'This feature exports all data to CSV files.\n\nComing soon!',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error exporting data: ' + error.toString());
  }
}

/**
 * Show vehicle status - quick view of operational vehicles
 */
function showVehicleStatus() {
  try {
    var logger = createLogger('Main');
    logger.info('Showing vehicle status');
    
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var vehicleSheet = ss.getSheetByName('Vehicle Status');
    
    if (!vehicleSheet) {
      SpreadsheetApp.getUi().alert('Vehicle Status sheet not found.');
      return;
    }
    
    var data = vehicleSheet.getDataRange().getValues();
    var operational = 0;
    var nonOperational = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === 'Y') operational++;
      else if (data[i][2] === 'N') nonOperational++;
    }
    
    SpreadsheetApp.getUi().alert(
      'Vehicle Status',
      'Total Vehicles: ' + (data.length - 1) + '\n' +
      'Operational: ' + operational + '\n' +
      'Non-Operational: ' + nonOperational,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Show daily details - quick view of today's allocations
 */
function showDailyDetails() {
  try {
    var logger = createLogger('Main');
    logger.info('Showing daily details');
    
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dailySheet = ss.getSheetByName('Daily Details');
    
    if (!dailySheet) {
      SpreadsheetApp.getUi().alert('Daily Details sheet not found.');
      return;
    }
    
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
    var data = dailySheet.getDataRange().getValues();
    var todayCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      var dateStr = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'MM/dd/yyyy');
      if (dateStr === today) todayCount++;
    }
    
    SpreadsheetApp.getUi().alert(
      'Daily Details',
      'Date: ' + today + '\n' +
      'Routes Allocated Today: ' + todayCount + '\n' +
      'Total Historical Records: ' + (data.length - 1),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Generate quick report
 */
function quickReport() {
  try {
    var logger = createLogger('Main');
    logger.info('Generating quick report');
    
    SpreadsheetApp.getUi().alert(
      'Quick Report',
      'This feature will generate a quick summary report.\n\nComing soon!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Setup form submit trigger
 */
function setupFormSubmitTrigger() {
  try {
    SpreadsheetApp.getUi().alert(
      'Form Trigger Setup',
      'This will set up automatic processing of form submissions.\n\nComing soon!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Test form van filtering
 */
function testFormVanFiltering() {
  try {
    SpreadsheetApp.getUi().alert(
      'Van Filtering Test',
      'This will test the van filtering functionality in forms.\n\nComing soon!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Show form management interface
 * Stub function that calls the actual implementation if available
 */
function showFormManagement() {
  try {
    // Try to call the actual function from FormManagementService
    if (typeof FormManagementService !== 'undefined' && typeof FormManagementService.showFormManagement === 'function') {
      FormManagementService.showFormManagement();
    } else {
      SpreadsheetApp.getUi().alert(
        'Form Management',
        'This feature allows you to manage all system forms.\n\nComing soon!',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error showing form management: ' + error.toString());
  }
}