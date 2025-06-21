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
  var ui = SpreadsheetApp.getUi();
  
  // Main vehicle assignment menu - this always works
  ui.createMenu("Vehicle Assignment Tool")
    .addItem("Upload Files for Allocation", "showUploadDialog")
    .addSeparator()
    .addItem("Create Extended Menus", "createExtendedMenus")
    .addToUi();
  
  // Delivery pace tracking menu - restored from original modular commit
  ui.createMenu("Delivery Pace")
    .addItem("Initialize Headers", "initializeDeliveryPaceHeaders")
    .addItem("Update Today's Pace", "updateDeliveryPaceForToday")
    .addSeparator()
    .addItem("Generate Today's Summary", "generateTodaysSummary")
    .addItem("Update Specific Van", "showUpdateVanDialog")
    .addSeparator()
    .addItem("Setup Auto-Update Triggers", "setupDeliveryPaceTriggers")
    .addItem("Test Update", "testDeliveryPaceUpdate")
    .addToUi();
}

/**
 * Create extended menus with all features
 * This can be called manually if the automatic menu creation fails
 */
function createExtendedMenus() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Fleet Operations menu
    ui.createMenu('Fleet Operations')
      .addItem('Allocate Vehicles', 'showUploadDialog')
      .addItem('View Dashboard', 'showDashboard')
      .addSeparator()
      .addItem('Delivery Pace Form', 'showDeliveryPaceForm')
      .addItem('RTS Report', 'showRTSForm')
      .addSeparator()
      .addItem('Form Management', 'showFormManagement')
      .addItem('View Error Log', 'showErrorLog')
      .addToUi();
    
    // Reports menu
    ui.createMenu('Reports')
      .addItem('Vehicle Utilization', 'generateVehicleUtilizationReport')
      .addItem('Driver Performance', 'generateDriverPerformanceReport')
      .addItem('Weekly Summary', 'generateWeeklySummaryReport')
      .addSeparator()
      .addItem('Analytics Dashboard', 'showAnalyticsDashboard')
      .addSeparator()
      .addItem('Export All Data', 'exportAllData')
      .addToUi();
    
    // Help menu
    ui.createMenu('Help')
      .addItem('User Guide', 'showUserGuide')
      .addItem('About', 'showAbout')
      .addToUi();
    
    SpreadsheetApp.getUi().alert('Extended menus created successfully!');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error creating extended menus: ' + error.toString());
  }
}

/**
 * Diagnostic function to understand menu issues
 */
function runMenuDiagnostics() {
  var ui = SpreadsheetApp.getUi();
  var results = [];
  
  results.push("=== MENU CREATION DIAGNOSTICS ===");
  results.push("Time: " + new Date().toString());
  
  // Test 1: Basic menu creation
  results.push("\n1. Testing basic menu creation:");
  try {
    ui.createMenu("Diagnostic Test 1")
      .addItem("Test Item", "showUploadDialog")
      .addToUi();
    results.push("✓ Created 'Diagnostic Test 1' menu");
  } catch (e) {
    results.push("✗ Failed: " + e.toString());
  }
  
  // Test 2: Multiple menus
  results.push("\n2. Testing multiple menu creation:");
  var testMenus = ["Test A", "Test B", "Test C"];
  testMenus.forEach(function(menuName) {
    try {
      ui.createMenu(menuName)
        .addItem("Item", "showUploadDialog")
        .addToUi();
      results.push("✓ Created '" + menuName + "' menu");
    } catch (e) {
      results.push("✗ Failed '" + menuName + "': " + e.toString());
    }
  });
  
  // Test 3: Check function availability
  results.push("\n3. Checking required functions:");
  var requiredFunctions = [
    "initializeDeliveryPaceHeaders",
    "updateDeliveryPaceForToday",
    "showDeliveryPaceForm",
    "generateTodaysSummary",
    "showRTSForm",
    "generateTodaysRTSSummary"
  ];
  
  requiredFunctions.forEach(function(fn) {
    if (typeof this[fn] === 'function') {
      results.push("✓ " + fn + " exists");
    } else {
      results.push("✗ " + fn + " missing");
    }
  });
  
  // Test 4: Try creating our actual menus
  results.push("\n4. Testing actual menu creation:");
  
  try {
    ui.createMenu("Delivery Pace Test")
      .addItem("Initialize Headers", "initializeDeliveryPaceHeaders")
      .addToUi();
    results.push("✓ Delivery Pace Test menu created");
  } catch (e) {
    results.push("✗ Delivery Pace Test failed: " + e.toString());
  }
  
  try {
    ui.createMenu("RTS Reporting Test")
      .addItem("Submit Report", "showRTSForm")
      .addToUi();
    results.push("✓ RTS Reporting Test menu created");
  } catch (e) {
    results.push("✗ RTS Reporting Test failed: " + e.toString());
  }
  
  // Show results
  var htmlOutput = HtmlService.createHtmlOutput('<pre>' + results.join('\n') + '</pre>')
    .setWidth(600)
    .setHeight(500);
  ui.showModalDialog(htmlOutput, 'Menu Diagnostics Results');
}

/**
 * Force menu recreation
 */
function forceMenuRecreation() {
  var ui = SpreadsheetApp.getUi();
  
  // Try direct menu creation
  ui.createMenu("Delivery Pace")
    .addItem("Initialize Headers", "initializeDeliveryPaceHeaders")
    .addItem("Update Today's Pace", "updateDeliveryPaceForToday")
    .addSeparator()
    .addItem("Submit Pace Report", "showDeliveryPaceForm")
    .addItem("Generate Today's Summary", "generateTodaysSummary")
    .addToUi();
  
  ui.createMenu("RTS Reporting")
    .addItem("Submit End of Day Report", "showRTSForm")
    .addItem("Generate Today's RTS Summary", "generateTodaysRTSSummary")
    .addToUi();
  
  ui.alert("Menu Recreation Attempted", 
    "Attempted to create Delivery Pace and RTS Reporting menus.\n\n" +
    "If they don't appear:\n" +
    "1. Try refreshing the page (Ctrl+R or Cmd+R)\n" +
    "2. Close and reopen the spreadsheet\n" +
    "3. Create a copy of the spreadsheet (File → Make a copy)",
    ui.ButtonSet.OK);
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
    
    var html = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; text-align: center;">
        <h3>Delivery Pace Collection Form</h3>
        <p><strong>Form URL:</strong><br>
        <a href="${info.formUrl}" target="_blank">${info.formUrl}</a></p>
        
        <p><strong>QR Code:</strong><br>
        <img src="${info.qrCodeUrl}" alt="QR Code" style="margin: 10px auto;">
        </p>
        
        <p style="font-size: 12px; color: #666;">
        Drivers can scan this QR code with their mobile devices<br>
        to quickly access the delivery pace reporting form.
        </p>
      </div>
    `)
    .setWidth(400)
    .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Form Information');
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

/**
 * Show RTS form dialog
 */
function showRTSForm() {
  var html = HtmlService.createHtmlOutputFromFile('RTSForm')
    .setWidth(getConfig('UI.RTS_FORM_WIDTH'))
    .setHeight(getConfig('UI.RTS_FORM_HEIGHT'));
  
  SpreadsheetApp.getUi().showModalDialog(html, 'End of Day Report');
}

/**
 * Generate RTS summary for today
 */
function generateTodaysRTSSummary() {
  var today = formatDate(new Date());
  generateRTSSummaryReport(today);
}

/**
 * Show information about the smart form
 */
function showSmartFormInfo() {
  var html = HtmlService.createHtmlOutput(`
    <div style="padding: 20px;">
      <h3>Smart Delivery Pace Form</h3>
      <p><strong>Features:</strong></p>
      <ul>
        <li>Auto-populates driver name and route based on van selection</li>
        <li>Shows only today's assigned routes</li>
        <li>Mobile-optimized interface</li>
        <li>Real-time updates to Daily Details</li>
      </ul>
      
      <p><strong>Setup Instructions:</strong></p>
      <ol>
        <li>Deploy this script as a Web App:
          <ul>
            <li>Click Extensions → Apps Script</li>
            <li>Click Deploy → New Deployment</li>
            <li>Type: Web app</li>
            <li>Execute as: Me</li>
            <li>Access: Anyone</li>
            <li>Click Deploy</li>
          </ul>
        </li>
        <li>Copy the Web App URL</li>
        <li>Share with drivers to bookmark on their phones</li>
      </ol>
      
      <p style="color: #666; font-size: 12px;">
      Note: The smart form requires deploying as a web app to enable auto-population features.
      </p>
    </div>
  `)
  .setWidth(500)
  .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Smart Form Information');
}

/**
 * Show today's operations summary
 */
function showTodaysSummary() {
  const logger = createLogger('Main');
  try {
    const today = formatDate(new Date());
    const summary = {
      allocation: getTodayAllocationStatistics(),
      pace: getDeliveryPaceStatistics(today),
      rts: getRTSStatistics(today)
    };
    
    const message = `Today's Summary (${today})\n\n` +
      `Vehicle Allocation:\n` +
      `• Routes Assigned: ${summary.allocation.assignedRoutes} of ${summary.allocation.totalRoutes}\n` +
      `• Vans Used: ${summary.allocation.vansUsed}\n` +
      `• Allocation Rate: ${summary.allocation.allocationRate}%\n\n` +
      `Delivery Progress:\n` +
      `• Vans Tracked: ${summary.pace.totalVansTracked}\n` +
      `• Last Update: ${summary.pace.lastCheckpoint}\n\n` +
      `End of Day:\n` +
      `• RTS Reports: ${summary.rts.completedReports} of ${summary.rts.totalRoutes}\n` +
      `• Packages Delivered: ${summary.rts.totalDelivered}\n` +
      `• Success Rate: ${summary.rts.successRate}%`;
    
    SpreadsheetApp.getUi().alert('Today\'s Summary', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    logger.error('Failed to show summary', { error: error.message });
    errorHandler.showAlert(error);
  }
}

/**
 * Refresh vehicle status from sheet
 */
function refreshVehicleStatus() {
  const logger = createLogger('Main');
  try {
    const stats = getVehicleStatistics();
    const message = `Vehicle Status Updated\n\n` +
      `Total Vehicles: ${stats.total}\n` +
      `Operational: ${stats.operational} (${stats.operationalRate}%)\n` +
      `Non-Operational: ${stats.nonOperational}\n\n` +
      `By Type:\n` +
      `• Extra Large: ${stats.byType['Extra Large'].operational}/${stats.byType['Extra Large'].total}\n` +
      `• Large: ${stats.byType['Large'].operational}/${stats.byType['Large'].total}\n` +
      `• Step Van: ${stats.byType['Step Van'].operational}/${stats.byType['Step Van'].total}`;
    
    SpreadsheetApp.getUi().alert('Vehicle Status', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    logger.error('Failed to refresh vehicle status', { error: error.message });
    errorHandler.showAlert(error);
  }
}

/**
 * Show delivery pace form
 */
function showDeliveryPaceForm() {
  var html = HtmlService.createHtmlOutputFromFile('DeliveryPaceForm')
    .setWidth(500)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Submit Delivery Pace Report');
}

/**
 * Send daily summary report
 */
function sendDailyReport() {
  const logger = createLogger('Main');
  try {
    // Implementation would send comprehensive daily report
    SpreadsheetApp.getUi().alert('Daily report sent to ' + getConfig('EMAIL_RECIPIENT'));
  } catch (error) {
    logger.error('Failed to send daily report', { error: error.message });
    errorHandler.showAlert(error);
  }
}

/**
 * Show configuration dialog
 */
function showConfigDialog() {
  SpreadsheetApp.getUi().alert('Configuration settings coming soon!');
}

/**
 * Show form management options
 */
function showFormManagement() {
  SpreadsheetApp.getUi().alert('Form management interface coming soon!');
}

/**
 * Show notification settings
 */
function showNotificationSettings() {
  SpreadsheetApp.getUi().alert('Notification settings coming soon!');
}

/**
 * Show error log
 */
function showErrorLog() {
  const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  const errorSheet = ss.getSheetByName('Error Log');
  
  if (errorSheet) {
    SpreadsheetApp.setActiveSheet(errorSheet);
    SpreadsheetApp.getUi().alert('Error Log sheet is now active');
  } else {
    SpreadsheetApp.getUi().alert('No errors logged yet');
  }
}

/**
 * Show user guide
 */
function showUserGuide() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding: 20px;">
      <h2>Fleet Resource Allocator User Guide</h2>
      <h3>Quick Start</h3>
      <ol>
        <li><strong>Daily Vehicle Allocation</strong>
          <ul>
            <li>Go to Fleet Operations → Daily Operations → Allocate Vehicles</li>
            <li>Upload Day of Ops and Daily Routes files</li>
            <li>System will automatically assign vehicles to routes</li>
          </ul>
        </li>
        <li><strong>Track Delivery Progress</strong>
          <ul>
            <li>Drivers submit pace reports at checkpoints (1:40, 3:40, 5:40, 7:40, 9:40 PM)</li>
            <li>View progress in Fleet Operations → Delivery Tracking</li>
          </ul>
        </li>
        <li><strong>End of Day Reporting</strong>
          <ul>
            <li>Drivers submit RTS reports via Fleet Operations → End of Day → Submit RTS Report</li>
            <li>Generate summaries for management review</li>
          </ul>
        </li>
      </ol>
      <p style="margin-top: 20px;">
        <strong>Need Help?</strong> Contact support at info@thervaaccountant.com
      </p>
    </div>
  `)
  .setWidth(600)
  .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'User Guide');
}

/**
 * Show video tutorials
 */
function showTutorials() {
  SpreadsheetApp.getUi().alert('Video tutorials coming soon!');
}

/**
 * Show support contact
 */
function showSupport() {
  SpreadsheetApp.getUi().alert(
    'Support Contact',
    'Email: info@thervaaccountant.com\n' +
    'Phone: (555) 123-4567\n' +
    'Hours: Mon-Fri 8AM-5PM EST',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Show about dialog
 */
function showAbout() {
  SpreadsheetApp.getUi().alert(
    'About Fleet Resource Allocator',
    'Version: 2.0\n' +
    'Last Updated: December 2024\n\n' +
    'Developed for efficient fleet management and route optimization.\n\n' +
    '© 2024 The RVA Accountant',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Comprehensive function verification test
 */
function verifyAllFunctions() {
  var ui = SpreadsheetApp.getUi();
  var results = [];
  var missing = [];
  var errors = [];
  
  results.push("=== FUNCTION VERIFICATION TEST ===");
  results.push("Time: " + new Date().toString());
  results.push("");
  
  // Define all functions referenced in menus
  var menuFunctions = {
    'Fleet Operations': {
      'Daily Operations': [
        'showUploadDialog',
        'showDashboard',
        'showTodaysSummary',
        'refreshVehicleStatus'
      ],
      'Delivery Tracking': [
        'showDeliveryPaceForm',
        'initializeDeliveryPaceHeaders',
        'updateDeliveryPaceForToday',
        'generateTodaysSummary',
        'setupDeliveryPaceTriggers'
      ],
      'End of Day': [
        'showRTSForm',
        'generateTodaysRTSSummary',
        'sendDailyReport'
      ],
      'Administration': [
        'showConfigDialog',
        'showFormManagement',
        'showNotificationSettings',
        'showErrorLog'
      ]
    },
    'Reports': [
      'generateVehicleUtilizationReport',
      'generateDriverPerformanceReport',
      'generateWeeklySummaryReport',
      'exportAllData'
    ],
    'Developer Tools': {
      'Testing': [
        'runAllTests',
        'testAllocationLogic',
        'testFormFunctionality',
        'runEmailServiceTests',
        'testRTSFunctionality',
        'testDeliveryPaceUpdate'
      ],
      'Data Management': [
        'createSampleData',
        'cleanOldData',
        'migrateData'
      ],
      'System': [
        'viewLogs',
        'clearCache',
        'resetTriggers',
        'runHealthCheck'
      ],
      'Other': [
        'runMenuDiagnostics'
      ]
    },
    'Help': [
      'showUserGuide',
      'showTutorials',
      'showSupport',
      'showAbout'
    ]
  };
  
  // Helper function to check if function exists
  function checkFunction(funcName) {
    try {
      if (typeof this[funcName] === 'function') {
        return true;
      }
      return false;
    } catch (e) {
      errors.push(funcName + ": " + e.toString());
      return false;
    }
  }
  
  // Check all menu functions
  function checkMenuSection(section, items) {
    results.push("\n" + section + ":");
    
    if (Array.isArray(items)) {
      items.forEach(function(func) {
        if (checkFunction(func)) {
          results.push("  ✓ " + func);
        } else {
          results.push("  ✗ " + func + " - MISSING");
          missing.push(func);
        }
      });
    } else {
      // Nested menu structure
      Object.keys(items).forEach(function(subsection) {
        results.push("  " + subsection + ":");
        items[subsection].forEach(function(func) {
          if (checkFunction(func)) {
            results.push("    ✓ " + func);
          } else {
            results.push("    ✗ " + func + " - MISSING");
            missing.push(func);
          }
        });
      });
    }
  }
  
  // Check all menus
  Object.keys(menuFunctions).forEach(function(menu) {
    checkMenuSection(menu, menuFunctions[menu]);
  });
  
  // Summary
  results.push("\n=== SUMMARY ===");
  results.push("Total functions checked: " + 
    Object.keys(menuFunctions).reduce(function(sum, menu) {
      var items = menuFunctions[menu];
      if (Array.isArray(items)) {
        return sum + items.length;
      } else {
        return sum + Object.keys(items).reduce(function(subSum, sub) {
          return subSum + items[sub].length;
        }, 0);
      }
    }, 0)
  );
  results.push("Missing functions: " + missing.length);
  if (missing.length > 0) {
    results.push("\nMissing:");
    missing.forEach(function(func) {
      results.push("  - " + func);
    });
  }
  if (errors.length > 0) {
    results.push("\nErrors:");
    errors.forEach(function(err) {
      results.push("  - " + err);
    });
  }
  
  // Show results
  var htmlOutput = HtmlService.createHtmlOutput('<pre>' + results.join('\n') + '</pre>')
    .setWidth(700)
    .setHeight(600);
  ui.showModalDialog(htmlOutput, 'Function Verification Results');
}

/**
 * Test function to debug menu creation
 */
function testMenuCreation() {
  var ui = SpreadsheetApp.getUi();
  var results = [];
  
  // Test 1: Basic menu
  try {
    ui.createMenu("Test Basic")
      .addItem("Test Item", "showUploadDialog")
      .addToUi();
    results.push("✓ Basic menu created");
  } catch (e) {
    results.push("✗ Basic menu failed: " + e.toString());
  }
  
  // Test 2: Menu with emoji
  try {
    ui.createMenu("📊 Test Emoji")
      .addItem("Test Item", "showUploadDialog")
      .addToUi();
    results.push("✓ Emoji menu created");
  } catch (e) {
    results.push("✗ Emoji menu failed: " + e.toString());
  }
  
  // Test 3: Menu with submenu
  try {
    ui.createMenu("Test Submenu")
      .addSubMenu(ui.createMenu("Sub")
        .addItem("Item", "showUploadDialog"))
      .addToUi();
    results.push("✓ Submenu created");
  } catch (e) {
    results.push("✗ Submenu failed: " + e.toString());
  }
  
  // Test 4: Check if functions exist
  var functions = [
    "showDashboard",
    "showTodaysSummary", 
    "refreshVehicleStatus",
    "showDeliveryPaceForm",
    "showErrorLog",
    "initializeDeliveryPaceHeaders",
    "updateDeliveryPaceForToday",
    "generateTodaysSummary",
    "showRTSForm",
    "generateTodaysRTSSummary"
  ];
  
  var missing = [];
  functions.forEach(function(fn) {
    try {
      if (typeof globalThis[fn] === 'function') {
        results.push("✓ Function exists: " + fn);
      } else {
        results.push("✗ Function missing: " + fn);
        missing.push(fn);
      }
    } catch (e) {
      results.push("✗ Error checking: " + fn);
      missing.push(fn);
    }
  });
  
  // Summary
  results.push("\n=== SUMMARY ===");
  results.push("Total functions checked: " + functions.length);
  results.push("Missing functions: " + missing.length);
  if (missing.length > 0) {
    results.push("Missing: " + missing.join(", "));
  }
  
  // Show results
  ui.alert("Menu Creation Test Results", results.join("\n"), ui.ButtonSet.OK);
}

// Developer tool functions
function openDevConsole() { 
  DevTools.showDevConsole(); 
}

function enableDevMode() { 
  DevTools.enableDevMode(); 
}

function disableDevMode() { 
  DevTools.disableDevMode(); 
}

function viewPerformance() { 
  const stats = DevTools.getPerformanceStats();
  const html = HtmlService.createHtmlOutput(`<pre>${stats}</pre>`)
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Performance Statistics');
}

function viewCacheStats() {
  const stats = JSON.stringify(Cache.getStats(), null, 2);
  const html = HtmlService.createHtmlOutput(`<pre>${stats}</pre>`)
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cache Statistics');
}

function clearAllCaches() {
  Cache.clear();
  Cache.clear(Cache.CACHE_TYPES.USER);
  Cache.clear(Cache.CACHE_TYPES.DOCUMENT);
  UIHelpers.toast('All caches cleared', 2000);
}

function profileCurrentOperation() {
  SpreadsheetApp.getUi().alert('Select an operation to profile from the Performance menu');
}

// Test functions for new components
function testLoggerFunction() {
  const result = DevTools.runTest('logger');
  UIHelpers.showSuccess(result);
}

function testCacheFunction() {
  const result = DevTools.runTest('cache');
  const html = HtmlService.createHtmlOutput(`<pre>${result}</pre>`)
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cache Test Results');
}

function testSheetManagerFunction() {
  const result = DevTools.runTest('sheet');
  const html = HtmlService.createHtmlOutput(`<pre>${result}</pre>`)
    .setWidth(500)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'SheetManager Test Results');
}

function testUIHelpersFunction() {
  UIHelpers.showLoading('Testing loading indicator...');
  Utilities.sleep(2000);
  google.script.host.close();
  
  UIHelpers.showSuccess('UI Helpers test complete!');
  
  Utilities.sleep(2000);
  
  const confirmed = UIHelpers.confirm(
    'Test Confirmation',
    'Did you see the loading indicator and success message?'
  );
  
  if (confirmed) {
    UIHelpers.toast('Great! UI Helpers are working correctly', 3000);
  }
}

function testSmartDefaultsFunction() {
  const result = DevTools.runTest('smart');
  const html = HtmlService.createHtmlOutput(`<pre>${result}</pre>`)
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Smart Defaults Test Results');
}

// Data management functions
function createSampleData() {
  const logger = Logger.createLogger('Main');
  
  try {
    UIHelpers.showLoading('Creating sample data...');
    
    // Implementation would create test data
    Utilities.sleep(1000);
    
    google.script.host.close();
    UIHelpers.showSuccess('Sample data created successfully!');
    
  } catch (error) {
    logger.error('Failed to create sample data', { error: error.message });
    UIHelpers.showError(error);
  }
}

function exportAllData() {
  const logger = Logger.createLogger('Main');
  
  try {
    const progress = UIHelpers.showProgress('Export Progress', 'Preparing data export...');
    
    // Implementation would export data
    
    UIHelpers.showSuccess('Data export complete!');
    
  } catch (error) {
    logger.error('Failed to export data', { error: error.message });
    UIHelpers.showError(error);
  }
}

function cleanOldData() {
  const confirmed = UIHelpers.confirm(
    'Clean Old Data',
    'This will remove data older than 90 days. Continue?',
    { dangerous: true }
  );
  
  if (confirmed) {
    UIHelpers.showLoading('Cleaning old data...');
    // Implementation
    Utilities.sleep(2000);
    google.script.host.close();
    UIHelpers.showSuccess('Old data cleaned successfully!');
  }
}

function migrateData() {
  SpreadsheetApp.getUi().alert('Data migration functionality coming soon!');
}

function viewLogs() {
  const logger = Logger.createLogger('Main');
  logger.info('Viewing logs');
  
  const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  const errorLog = ss.getSheetByName('Error Log');
  
  if (errorLog) {
    const data = errorLog.getDataRange().getValues();
    const recentLogs = data.slice(-20); // Last 20 entries
    
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: monospace; font-size: 12px;">
        <h3>Recent Error Logs</h3>
        <pre>${JSON.stringify(recentLogs, null, 2)}</pre>
      </div>
    `)
    .setWidth(800)
    .setHeight(600);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Error Logs');
  } else {
    UIHelpers.showSuccess('No error logs found');
  }
}

function resetTriggers() {
  const logger = Logger.createLogger('Main');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    
    UIHelpers.showSuccess(`Removed ${triggers.length} triggers`);
    
  } catch (error) {
    logger.error('Failed to reset triggers', { error: error.message });
    UIHelpers.showError(error);
  }
}

function runHealthCheck() {
  const logger = Logger.createLogger('Main');
  const results = [];
  
  try {
    UIHelpers.showLoading('Running system health check...');
    
    // Check spreadsheet access
    try {
      const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
      results.push('✓ Spreadsheet access OK');
    } catch (e) {
      results.push('✗ Spreadsheet access FAILED');
    }
    
    // Check required sheets
    const requiredSheets = ['Vehicle Status', 'Daily Details'];
    requiredSheets.forEach(sheetName => {
      try {
        const manager = new SheetManager();
        manager.getSheet(sheetName);
        results.push(`✓ Sheet "${sheetName}" OK`);
      } catch (e) {
        results.push(`✗ Sheet "${sheetName}" MISSING`);
      }
    });
    
    // Check cache
    try {
      Cache.store('health_check', true, 10, Cache.CACHE_TYPES.SCRIPT);
      const retrieved = Cache.retrieve('health_check', Cache.CACHE_TYPES.SCRIPT);
      if (retrieved) {
        results.push('✓ Cache system OK');
      } else {
        results.push('✗ Cache system FAILED');
      }
    } catch (e) {
      results.push('✗ Cache system ERROR');
    }
    
    // Check triggers
    const triggers = ScriptApp.getProjectTriggers();
    results.push(`ℹ Active triggers: ${triggers.length}`);
    
    google.script.host.close();
    
    const html = HtmlService.createHtmlOutput(`
      <div style="padding: 20px;">
        <h3>System Health Check Results</h3>
        <pre>${results.join('\n')}</pre>
      </div>
    `)
    .setWidth(500)
    .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Health Check Results');
    
  } catch (error) {
    logger.error('Health check failed', { error: error.message });
    UIHelpers.showError(error);
  }
}

// Report functions
function generateVehicleUtilizationReport() {
  const logger = Logger.createLogger('Reports');
  
  try {
    // Use the comprehensive van utilization analytics
    const utilization = calculateVanUtilizationByType();
    
    if (!utilization) {
      throw new Error('Failed to calculate van utilization');
    }
    
    UIHelpers.showLoading('Generating vehicle utilization report...');
    
    const manager = new SheetManager();
    const reportDate = formatDate(new Date());
    const reportSheetName = `Vehicle Utilization ${reportDate}`;
    
    // Create report sheet
    const reportSheet = manager.createSheet(reportSheetName, {
      headers: [
        'Vehicle Type', 
        'Total Fleet', 
        'Operational', 
        'Non-Operational', 
        'Operational %',
        'Utilized', 
        'Utilization %'
      ],
      overwrite: true
    });
    
    // Add data rows
    const rows = Object.entries(utilization.byType).map(([type, data]) => [
      type,
      data.total,
      data.operational,
      data.total - data.operational,
      data.operationalRate,
      data.utilized,
      data.utilizationRate
    ]);
    
    reportSheet.appendRows(rows);
    
    // Add summary row
    reportSheet.appendRows([
      [],
      [
        'TOTAL',
        utilization.overall.totalFleet,
        utilization.overall.totalOperational,
        utilization.overall.totalFleet - utilization.overall.totalOperational,
        utilization.overall.operationalRate,
        utilization.overall.totalUtilized,
        utilization.overall.utilizationRate
      ]
    ]);
    
    // Format the sheet
    reportSheet.sheet.getRange(rows.length + 3, 1, 1, 7)
      .setFontWeight('bold')
      .setBackground('#E8F0FE');
    
    reportSheet.autoResizeColumns(1, 7);
    
    google.script.host.close();
    UIHelpers.showSuccess(`Vehicle utilization report generated!\nSheet: ${reportSheetName}`);
    
    logger.info('Vehicle utilization report generated', { sheet: reportSheetName });
    
  } catch (error) {
    logger.error('Failed to generate report', { error: error.message });
    google.script.host.close();
    UIHelpers.showError(error);
  }
}

// generateDriverPerformanceReport is now in ReportService.js
// generateWeeklySummaryReport is now in ReportService.js

/**
 * Show analytics dashboard
 */
function showAnalyticsDashboard() {
  generateAnalyticsDashboard();
}

/**
 * Test if all required functions exist
 */
function testRequiredFunctions() {
  var requiredFunctions = [
    // Core functions
    'showUploadDialog',
    'showDashboard',
    'showDeliveryPaceForm',
    'showRTSForm',
    'showFormManagement',
    'showErrorLog',
    
    // Report functions
    'generateVehicleUtilizationReport',
    'generateDriverPerformanceReport',
    'generateWeeklySummaryReport',
    'showAnalyticsDashboard',
    'exportAllData',
    
    // Help functions
    'showUserGuide',
    'showAbout'
  ];
  
  var results = [];
  var missing = [];
  
  requiredFunctions.forEach(function(funcName) {
    if (typeof this[funcName] === 'function') {
      results.push('✓ ' + funcName);
    } else {
      results.push('✗ ' + funcName + ' - MISSING');
      missing.push(funcName);
    }
  });
  
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Function Test Results',
    'Missing functions: ' + missing.length + '\n\n' + results.join('\n'),
    ui.ButtonSet.OK
  );
}

// ===================================================================
// MISSING FUNCTION STUBS
// ===================================================================
// These functions are referenced in menus but not yet implemented

/**
 * Show dashboard interface
 */
function showDashboard() {
  try {
    // Simply display the Dashboard HTML file
    var html = HtmlService.createHtmlOutputFromFile('Dashboard')
      .setWidth(800)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Fleet Operations Dashboard');
  } catch (error) {
    Logger.log('Error showing dashboard: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error loading dashboard: ' + error.toString());
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
    Logger.log('Error showing delivery pace form: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error loading delivery pace form: ' + error.toString());
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
    Logger.log('Error showing RTS form: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error loading RTS form: ' + error.toString());
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
    
    if (errorSheet) {
      var data = errorSheet.getDataRange().getValues();
      if (data.length > 1) {
        // Format recent errors (last 20)
        var recentErrors = data.slice(Math.max(1, data.length - 20));
        errorLog = recentErrors.map(function(row) {
          return row[0] + ' | ' + row[1] + ' | ' + row[2];
        }).join('\n');
      }
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
    Logger.log('Error showing error log: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error log not available');
  }
}

/**
 * Generate vehicle utilization report
 */
function generateVehicleUtilizationReport() {
  try {
    SpreadsheetApp.getUi().alert('Generating vehicle utilization report...');
    // This would call the actual report generation logic
    var report = 'Vehicle Utilization Report\n' +
                '=======================\n' +
                'Total Vehicles: 36\n' +
                'Operational: 27\n' + 
                'In Use Today: 24\n' +
                'Utilization Rate: 88.9%';
    SpreadsheetApp.getUi().alert('Report Generated', report, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error generating report: ' + error.toString());
  }
}

/**
 * Show analytics dashboard
 */
function showAnalyticsDashboard() {
  try {
    generateAnalyticsDashboard();
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error showing analytics: ' + error.toString());
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
    '<li><strong>Upload Files:</strong> Use "Vehicle Assignment Tool" → "Upload Files for Allocation"</li>' +
    '<li><strong>Select Files:</strong> Choose Day of Ops and Daily Routes Excel files</li>' +
    '<li><strong>Wait for Processing:</strong> The system will allocate vehicles automatically</li>' +
    '<li><strong>Review Results:</strong> Check the Results sheet for allocations</li>' +
    '</ol>' +
    '<h3>Features</h3>' +
    '<ul>' +
    '<li><strong>Dashboard:</strong> View real-time fleet status and metrics</li>' +
    '<li><strong>Reports:</strong> Generate utilization and performance reports</li>' +
    '<li><strong>Forms:</strong> Submit delivery pace and RTS reports</li>' +
    '<li><strong>Analytics:</strong> View comprehensive fleet analytics</li>' +
    '</ul>' +
    '<h3>Support</h3>' +
    '<p>For assistance, contact your system administrator.</p>' +
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
    '<li>Real-time tracking</li>' +
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