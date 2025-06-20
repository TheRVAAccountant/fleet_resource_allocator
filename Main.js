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
  
  // Main vehicle assignment menu
  ui.createMenu("Vehicle Assignment Tool")
    .addItem("Upload Files for Allocation", "showUploadDialog")
    .addToUi();
  
  // Delivery pace tracking menu
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