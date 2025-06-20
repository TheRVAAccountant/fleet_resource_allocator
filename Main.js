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
    .addSubMenu(ui.createMenu("Form Management")
      .addItem("Create Smart Form (Recommended)", "showSmartFormInfo")
      .addItem("Create Basic Google Form", "createDeliveryForm")
      .addSeparator()
      .addItem("Get Form Link & QR Code", "showFormInfo")
      .addItem("Setup Form Trigger", "setupFormTrigger")
      .addItem("Test Van Filtering", "testVanFiltering"))
    .addSeparator()
    .addItem("Setup Auto-Update Triggers", "setupDeliveryPaceTriggers")
    .addItem("Test Update", "testDeliveryPaceUpdate")
    .addItem("Test Email Notification", "testDeliveryPaceEmail")
    .addItem("Debug Email Test", "debugTestDeliveryPaceEmail")
    .addItem("Run Email Tests", "runEmailServiceTests")
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