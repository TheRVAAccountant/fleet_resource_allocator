/**
 * ===================================================================
 * MENU SYSTEM TEST - COMPREHENSIVE VERIFICATION
 * ===================================================================
 * Tests all aspects of the menu system and identifies issues
 */

/**
 * Run comprehensive menu system test
 */
function testMenuSystem() {
  var ui = SpreadsheetApp.getUi();
  var results = [];
  
  results.push("=== MENU SYSTEM TEST REPORT ===");
  results.push("Time: " + new Date().toString());
  results.push("");
  
  // Test 1: Logger functionality
  results.push("1. LOGGER TEST:");
  try {
    var testLog = createLogger('MenuTest');
    testLog.info('Testing logger');
    results.push("✓ Logger created successfully");
    results.push("✓ createLogger function works");
  } catch (e) {
    results.push("✗ Logger error: " + e.toString());
  }
  
  // Test 2: Config access
  results.push("");
  results.push("2. CONFIG TEST:");
  try {
    var configValue = getConfig('DAILY_SUMMARY_SPREADSHEET_ID');
    if (configValue) {
      results.push("✓ Config access works");
      results.push("  Spreadsheet ID: " + configValue);
    } else {
      results.push("✗ Config returns null/undefined");
    }
  } catch (e) {
    results.push("✗ Config error: " + e.toString());
  }
  
  // Test 3: Menu creation capability
  results.push("");
  results.push("3. MENU CREATION TEST:");
  try {
    var testMenu = ui.createMenu('Test Menu');
    testMenu.addItem('Test Item', 'testMenuSystem');
    testMenu.addToUi();
    results.push("✓ Can create menus");
  } catch (e) {
    results.push("✗ Menu creation error: " + e.toString());
  }
  
  // Test 4: Check all required functions
  results.push("");
  results.push("4. FUNCTION AVAILABILITY:");
  var requiredFunctions = [
    'onOpen',
    'setupAllMenus',
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
    'clearErrorLog',
    'testLogger',
    'runDiagnostics',
    'showUserGuide',
    'showAbout',
    'createDeliveryForm',
    'showFormInfo',
    'showVehicleStatus',
    'showDailyDetails',
    'quickReport',
    'showDashboard'
  ];
  
  var missingFunctions = [];
  var presentFunctions = [];
  
  requiredFunctions.forEach(function(funcName) {
    if (typeof this[funcName] === 'function') {
      presentFunctions.push(funcName);
    } else {
      missingFunctions.push(funcName);
    }
  });
  
  results.push("Functions found: " + presentFunctions.length + "/" + requiredFunctions.length);
  if (missingFunctions.length > 0) {
    results.push("Missing functions:");
    missingFunctions.forEach(function(fn) {
      results.push("  ✗ " + fn);
    });
  } else {
    results.push("✓ All required functions are present");
  }
  
  // Test 5: Try to run setupAllMenus
  results.push("");
  results.push("5. MENU SETUP TEST:");
  try {
    setupAllMenus();
    results.push("✓ setupAllMenus executed without errors");
  } catch (e) {
    results.push("✗ setupAllMenus error: " + e.toString());
  }
  
  // Test 6: Check for common issues
  results.push("");
  results.push("6. COMMON ISSUES CHECK:");
  
  // Check for ES6 syntax issues
  try {
    var loggerSource = Logger.toString();
    if (loggerSource.indexOf('class ') > -1) {
      results.push("⚠ Found ES6 class syntax in Logger");
    } else {
      results.push("✓ No ES6 class syntax detected");
    }
  } catch (e) {
    results.push("  Could not check for ES6 syntax");
  }
  
  // Test 7: File loading order
  results.push("");
  results.push("7. FILE LOADING:");
  results.push("  Files with underscore prefix load first");
  results.push("  Current logger source: " + (typeof createLogger));
  
  // Show results
  var html = HtmlService.createHtmlOutput(
    '<div style="padding: 20px; font-family: monospace; font-size: 12px;">' +
    '<pre>' + results.join('\n') + '</pre>' +
    '<hr>' +
    '<p><strong>Recommendations:</strong></p>' +
    '<ol>' +
    '<li>If Logger errors persist, refresh the script editor</li>' +
    '<li>If menus don\'t appear, run setupAllMenus() manually</li>' +
    '<li>Check the Error Log sheet for runtime errors</li>' +
    '<li>Ensure the Daily Summary Spreadsheet ID is correct</li>' +
    '</ol>' +
    '</div>'
  )
  .setWidth(700)
  .setHeight(600);
  
  ui.showModalDialog(html, 'Menu System Test Report');
}

/**
 * Force refresh all menus
 */
function forceRefreshMenus() {
  try {
    // Clear any cached menu errors
    PropertiesService.getScriptProperties().deleteProperty('lastMenuError');
    
    // Run setup
    setupAllMenus();
    
    SpreadsheetApp.getUi().alert(
      'Menu Refresh Complete',
      'All menus have been recreated. You may need to refresh the spreadsheet to see changes.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Menu Refresh Error',
      'Error: ' + error.toString() + '\n\nTry running testMenuSystem() for diagnostics.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Quick fix for common issues
 */
function quickFix() {
  var fixes = [];
  
  try {
    // Fix 1: Ensure logger is available
    if (typeof createLogger !== 'function') {
      fixes.push('✗ createLogger not found - Logger system may be broken');
    } else {
      fixes.push('✓ Logger system OK');
    }
    
    // Fix 2: Check config
    if (typeof getConfig !== 'function') {
      fixes.push('✗ getConfig not found - Config system may be broken');
    } else {
      fixes.push('✓ Config system OK');
    }
    
    // Fix 3: Try to create menus
    try {
      setupAllMenus();
      fixes.push('✓ Menus created successfully');
    } catch (e) {
      fixes.push('✗ Menu creation failed: ' + e.toString());
    }
    
    SpreadsheetApp.getUi().alert(
      'Quick Fix Results',
      fixes.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Quick Fix Error',
      'Critical error: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}