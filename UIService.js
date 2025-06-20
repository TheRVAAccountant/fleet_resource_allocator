/**
 * ===================================================================
 * UI SERVICE
 * ===================================================================
 * Handles all user interface operations including dialogs, forms,
 * and user interactions.
 */

/**
 * Shows the file upload dialog
 */
function showUploadDialog() {
  var html = HtmlService.createHtmlOutputFromFile('UploadDialog')
    .setWidth(getConfig('UI.UPLOAD_DIALOG_WIDTH'))
    .setHeight(getConfig('UI.UPLOAD_DIALOG_HEIGHT'));
  SpreadsheetApp.getUi().showModalDialog(html, 'Vehicle Allocation Tool');
}

/**
 * Shows dialog to update specific van
 */
function showUpdateVanDialog() {
  var today = getTodayString();
  var html = HtmlService.createHtmlOutput(`
    <div style="padding: 20px;">
      <h3>Update Delivery Pace for Specific Van</h3>
      <form id="updateForm">
        <label for="vanId">Van ID:</label><br>
        <input type="text" id="vanId" required style="margin: 10px 0; padding: 5px;"><br>
        
        <label for="date">Date (MM/DD/YYYY):</label><br>
        <input type="text" id="date" required style="margin: 10px 0; padding: 5px;" 
               value="${today}"><br>
        
        <input type="submit" value="Update Van" style="padding: 10px 20px; margin-top: 10px;">
      </form>
      <div id="status" style="margin-top: 20px; color: blue;"></div>
    </div>
    <script>
      document.getElementById('updateForm').addEventListener('submit', function(e) {
        e.preventDefault();
        var vanId = document.getElementById('vanId').value;
        var date = document.getElementById('date').value;
        document.getElementById('status').innerText = 'Updating...';
        
        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('status').innerText = result ? 
              'Successfully updated van ' + vanId : 
              'Van not found for the specified date';
          })
          .withFailureHandler(function(error) {
            document.getElementById('status').innerText = 'Error: ' + error.message;
          })
          .updateDeliveryPaceForVan(vanId, date);
      });
    </script>
  `)
  .setWidth(getConfig('UI.UPDATE_VAN_DIALOG_WIDTH'))
  .setHeight(getConfig('UI.UPDATE_VAN_DIALOG_HEIGHT'));
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Van Delivery Pace');
}

/**
 * Uploads and converts an XLSX file to a Google Sheet
 * @param {string} fileData - Base64 encoded file data
 * @param {string} fileName - Name of the file
 * @return {string} File ID of created Google Sheet
 */
function uploadAndConvertXLSX(fileData, fileName) {
  try {
    if (typeof Drive === 'undefined' || !Drive.Files || typeof Drive.Files.create !== 'function') {
      throw new Error("Advanced Drive Service is not enabled or not configured for v3. " +
                      "Please enable it via Resources > Advanced Google Services.");
    }
    
    var base64Data = fileData.split(',')[1];
    var contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, fileName);
    
    var resource = {
      name: fileName,
      mimeType: "application/vnd.google-apps.spreadsheet"
    };
    
    var file = Drive.Files.create(resource, blob);
    return file.id;
  } catch (err) {
    throw new Error("Error uploading file " + fileName + ": " + err);
  }
}

/**
 * Shows a completion alert with allocation results
 * @param {string} resultsSheetName - Name of results sheet
 * @param {string} unassignedSheetName - Name of unassigned sheet
 * @param {string} routeAssignmentsFileId - File ID of route assignments
 */
function showAllocationCompleteAlert(resultsSheetName, unassignedSheetName, routeAssignmentsFileId) {
  var uiMsg = "Allocation completed!\n\n" +
    "Results sheet created: " + resultsSheetName + " (in Daily Summary)\n" +
    "Available & Unassigned Vans sheet: " + unassignedSheetName + "\n" +
    "Route Assignments file ID: " + routeAssignmentsFileId + "\n\n" +
    "Check Logs in Apps Script > Executions for details.";
  
  SpreadsheetApp.getUi().alert(uiMsg);
}

/**
 * Shows an error alert
 * @param {string} message - Error message to display
 */
function showErrorAlert(message) {
  SpreadsheetApp.getUi().alert("Error: " + message);
}

/**
 * Shows an information alert
 * @param {string} message - Information message to display
 */
function showInfoAlert(message) {
  SpreadsheetApp.getUi().alert(message);
}