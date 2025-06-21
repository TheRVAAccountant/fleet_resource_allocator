/**
 * ===================================================================
 * MENU DEBUG HELPER
 * ===================================================================
 * Helps diagnose menu creation issues
 */

/**
 * Simple onOpen function that creates menus one by one
 * This helps identify which menu is causing the issue
 */
function onOpenDebug() {
  var ui = SpreadsheetApp.getUi();
  var results = [];
  
  // Test creating each menu individually
  try {
    ui.createMenu("Test 1 - Basic")
      .addItem("Upload Files", "showUploadDialog")
      .addToUi();
    results.push("✓ Basic menu created");
  } catch (e) {
    results.push("✗ Basic menu failed: " + e.toString());
  }
  
  try {
    ui.createMenu('Test 2 - Fleet Ops')
      .addItem('Allocate Vehicles', 'showUploadDialog')
      .addItem('View Dashboard', 'showDashboard')
      .addToUi();
    results.push("✓ Fleet Ops menu created");
  } catch (e) {
    results.push("✗ Fleet Ops menu failed: " + e.toString());
  }
  
  try {
    ui.createMenu('Test 3 - Reports')
      .addItem('Vehicle Utilization', 'generateVehicleUtilizationReport')
      .addItem('Driver Performance', 'generateDriverPerformanceReport')
      .addToUi();
    results.push("✓ Reports menu created");
  } catch (e) {
    results.push("✗ Reports menu failed: " + e.toString());
  }
  
  try {
    ui.createMenu('Test 4 - With Submenu')
      .addSubMenu(ui.createMenu('Submenu')
        .addItem('Item 1', 'showUploadDialog'))
      .addToUi();
    results.push("✓ Submenu created");
  } catch (e) {
    results.push("✗ Submenu failed: " + e.toString());
  }
  
  // Show results
  ui.alert('Menu Debug Results', results.join('\n'), ui.ButtonSet.OK);
}

/**
 * Simplified onOpen that creates all menus without complex logic
 */
function onOpenSimple() {
  var ui = SpreadsheetApp.getUi();
  
  // Vehicle Assignment Tool
  ui.createMenu("Vehicle Assignment Tool")
    .addItem("Upload Files for Allocation", "showUploadDialog")
    .addSeparator()
    .addItem("View Dashboard", "showDashboard")
    .addItem("Run Menu Diagnostics", "runMenuDiagnostics")
    .addToUi();
  
  // Fleet Operations
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
  
  // Reports
  ui.createMenu('Reports')
    .addItem('Vehicle Utilization', 'generateVehicleUtilizationReport')
    .addItem('Driver Performance', 'generateDriverPerformanceReport')
    .addItem('Weekly Summary', 'generateWeeklySummaryReport')
    .addSeparator()
    .addItem('Analytics Dashboard', 'showAnalyticsDashboard')
    .addSeparator()
    .addItem('Export All Data', 'exportAllData')
    .addToUi();
  
  // Help
  ui.createMenu('Help')
    .addItem('User Guide', 'showUserGuide')
    .addItem('About', 'showAbout')
    .addToUi();
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

/**
 * Replace the current onOpen with a simple version
 */
function useSimpleMenus() {
  // This would need to be manually copied to replace onOpen
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Simple Menu Code',
    'Copy the onOpenSimple function and rename it to onOpen to use simplified menus.',
    ui.ButtonSet.OK
  );
}