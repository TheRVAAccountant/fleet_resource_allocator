/**
 * ===================================================================
 * FORM SERVICE
 * ===================================================================
 * Manages Google Forms integration for collecting delivery pace data
 * from drivers throughout the day.
 */

/**
 * Creates or updates the delivery pace collection form
 * @return {string} URL of the created/updated form
 */
function createDeliveryPaceForm() {
  try {
    // Check if form already exists
    var existingFormId = PropertiesService.getScriptProperties().getProperty('DELIVERY_PACE_FORM_ID');
    var form;
    
    if (existingFormId) {
      try {
        form = FormApp.openById(existingFormId);
        // Clear existing items to rebuild
        var items = form.getItems();
        for (var i = items.length - 1; i >= 0; i--) {
          form.deleteItem(items[i]);
        }
      } catch (e) {
        // Form doesn't exist anymore, create new one
        form = FormApp.create('Delivery Pace Report');
      }
    } else {
      form = FormApp.create('Delivery Pace Report');
    }
    
    // Configure form settings
    form.setDescription('Report your delivery progress at each checkpoint throughout the day.')
      .setConfirmationMessage('Thank you! Your delivery pace has been recorded.')
      .setAllowResponseEdits(true)
      .setLimitOneResponsePerUser(false)
      .setProgressBar(true);
    
    // Add form header
    form.addSectionHeaderItem()
      .setTitle('Delivery Progress Report')
      .setHelpText('Please submit your delivery count at each checkpoint time.');
    
    // Date field (auto-populated but editable)
    form.addDateItem()
      .setTitle('Date')
      .setRequired(true)
      .setHelpText('Leave as today unless reporting for a different date');
    
    // Van ID dropdown - populated from Vehicle Status
    var vanIdItem = form.addListItem()
      .setTitle('Van ID')
      .setRequired(true)
      .setHelpText('Select your assigned van');
    
    // Get van IDs from Vehicle Status
    var vanChoices = getActiveVanChoices();
    vanIdItem.setChoiceValues(vanChoices);
    
    // Driver name - will be auto-populated when van is selected
    var driverNameItem = form.addTextItem()
      .setTitle('Driver Name')
      .setRequired(true)
      .setHelpText('Auto-populated based on van selection (editable if needed)');
    
    // Route code dropdown - populated from today's assignments
    var routeCodeItem = form.addListItem()
      .setTitle('Route Code')
      .setRequired(true)
      .setHelpText('Auto-populated based on van selection');
    
    // Get today's route assignments
    var routeChoices = getTodayRouteChoices();
    if (routeChoices.length > 0) {
      routeCodeItem.setChoiceValues(routeChoices);
    } else {
      // Fallback to text input if no routes found
      form.deleteItem(routeCodeItem);
      form.addTextItem()
        .setTitle('Route Code')
        .setRequired(true)
        .setHelpText('Enter your route code manually');
    }
    
    // Time checkpoint dropdown
    var timeItem = form.addListItem()
      .setTitle('Reporting Time')
      .setRequired(true)
      .setHelpText('Select the checkpoint you are reporting for');
    
    timeItem.setChoices([
      timeItem.createChoice('1:40 PM'),
      timeItem.createChoice('3:40 PM'),
      timeItem.createChoice('5:40 PM'),
      timeItem.createChoice('7:40 PM'),
      timeItem.createChoice('9:40 PM (End of Day)')
    ]);
    
    // Delivery count
    form.addTextItem()
      .setTitle('Total Deliveries Completed')
      .setRequired(true)
      .setValidation(FormApp.createTextValidation()
        .setHelpText('Please enter a number')
        .requireNumber()
        .build())
      .setHelpText('Enter the TOTAL number of deliveries completed so far today');
    
    // Optional notes
    form.addParagraphTextItem()
      .setTitle('Notes (Optional)')
      .setHelpText('Any issues or comments about your route');
    
    // Set up form response destination
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    
    // Create or get Delivery Pace Data sheet
    var dataSheet = ss.getSheetByName('Delivery Pace Data');
    if (!dataSheet) {
      dataSheet = ss.insertSheet('Delivery Pace Data');
      setupDeliveryPaceDataSheet(dataSheet);
    }
    
    // Link form to sheet
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Save form ID
    PropertiesService.getScriptProperties().setProperty('DELIVERY_PACE_FORM_ID', form.getId());
    
    // Create shortened URL for easy sharing
    var formUrl = form.getPublishedUrl();
    var shortUrl = form.shortenFormUrl(formUrl);
    
    Logger.log('Delivery Pace Form created/updated: ' + shortUrl);
    
    return shortUrl;
    
  } catch (error) {
    Logger.log('Error creating form: ' + error);
    throw error;
  }
}

/**
 * Get list of active vans from Vehicle Status
 * @return {string[]} Array of van IDs
 */
function getActiveVanChoices() {
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var vehicleSheet = ss.getSheetByName(getConfig('SHEETS.VEHICLE_STATUS'));
  
  if (!vehicleSheet) {
    // Return default list if sheet not found
    return ['BW1', 'BW2', 'BW3', 'BW4', 'BW5'];
  }
  
  var data = vehicleSheet.getDataRange().getValues();
  var vanIds = [];
  
  // Find Van ID and Operational columns
  var headers = data[0];
  var vanIdCol = headers.indexOf('Van ID');
  var opCol = headers.indexOf('Opnal?\nY/N');
  
  if (vanIdCol === -1) {
    return ['BW1', 'BW2', 'BW3', 'BW4', 'BW5'];
  }
  
  // Collect operational van IDs
  for (var i = 1; i < data.length; i++) {
    if (opCol === -1 || data[i][opCol] === 'Y') {
      var vanId = data[i][vanIdCol];
      if (vanId && vanId.toString().trim() !== '') {
        vanIds.push(vanId.toString());
      }
    }
  }
  
  return vanIds.length > 0 ? vanIds : ['BW1', 'BW2', 'BW3', 'BW4', 'BW5'];
}

/**
 * Get list of routes assigned today from Daily Details
 * @return {string[]} Array of route codes
 */
function getTodayRouteChoices() {
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
    
    if (!dailyDetailsSheet) {
      return [];
    }
    
    var today = formatDate(new Date());
    var data = dailyDetailsSheet.getDataRange().getValues();
    var routeCodes = [];
    var uniqueRoutes = new Set();
    
    // Find today's routes
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0];
      if (rowDate instanceof Date) {
        rowDate = formatDate(rowDate);
      }
      
      if (rowDate === today) {
        var routeCode = data[i][1]; // Column B
        if (routeCode && !uniqueRoutes.has(routeCode)) {
          uniqueRoutes.add(routeCode);
          routeCodes.push(routeCode.toString());
        }
      }
    }
    
    return routeCodes.sort();
    
  } catch (error) {
    Logger.log('Error getting today\'s routes: ' + error);
    return [];
  }
}

/**
 * Get van-to-route-to-driver mapping for today
 * @return {Object} Mapping object
 */
function getTodayAssignments() {
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
    
    if (!dailyDetailsSheet) {
      return {};
    }
    
    var today = formatDate(new Date());
    var data = dailyDetailsSheet.getDataRange().getValues();
    var assignments = {};
    
    // Build mapping: vanId -> {route, driver}
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][0];
      if (rowDate instanceof Date) {
        rowDate = formatDate(rowDate);
      }
      
      if (rowDate === today) {
        var routeCode = data[i][1]; // Column B
        var driverName = data[i][2]; // Column C
        var vanId = data[i][4]; // Column E
        
        if (vanId) {
          assignments[vanId] = {
            route: routeCode || '',
            driver: driverName || ''
          };
        }
      }
    }
    
    return assignments;
    
  } catch (error) {
    Logger.log('Error getting today\'s assignments: ' + error);
    return {};
  }
}

/**
 * Set up the Delivery Pace Data sheet structure
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to set up
 */
function setupDeliveryPaceDataSheet(sheet) {
  // Set headers
  var headers = [
    'Timestamp',
    'Date',
    'Van ID',
    'Driver Name',
    'Route Code',
    'Reporting Time',
    'Total Deliveries',
    'Notes',
    'Processed'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
    .setBackground('#E8F0FE')
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 100); // Date
  sheet.setColumnWidth(3, 80);  // Van ID
  sheet.setColumnWidth(4, 150); // Driver Name
  sheet.setColumnWidth(5, 100); // Route Code
  sheet.setColumnWidth(6, 120); // Reporting Time
  sheet.setColumnWidth(7, 120); // Total Deliveries
  sheet.setColumnWidth(8, 200); // Notes
  sheet.setColumnWidth(9, 80);  // Processed
  
  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Process form responses and update Daily Details
 * Called by form submit trigger
 */
function onDeliveryPaceFormSubmit(e) {
  try {
    var response = e.response;
    var itemResponses = response.getItemResponses();
    
    // Extract form data
    var formData = {};
    for (var i = 0; i < itemResponses.length; i++) {
      var itemResponse = itemResponses[i];
      formData[itemResponse.getItem().getTitle()] = itemResponse.getResponse();
    }
    
    // Update the Daily Details sheet
    updateDailyDetailsFromForm(formData);
    
    Logger.log('Processed delivery pace form submission for Van: ' + formData['Van ID']);
    
  } catch (error) {
    Logger.log('Error processing form submission: ' + error);
  }
}

/**
 * Update Daily Details with form submission data
 * @param {Object} formData - Data from form submission
 */
function updateDailyDetailsFromForm(formData) {
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
  
  if (!dailyDetailsSheet) {
    throw new Error('Daily Details sheet not found');
  }
  
  // Format date for comparison
  var formDate = formData['Date'];
  if (formDate instanceof Date) {
    formDate = formatDate(formDate);
  }
  
  var vanId = formData['Van ID'];
  var reportingTime = formData['Reporting Time'];
  var deliveryCount = formData['Total Deliveries Completed'];
  
  // Find the matching row
  var data = dailyDetailsSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0];
    if (rowDate instanceof Date) {
      rowDate = formatDate(rowDate);
    }
    
    var rowVanId = data[i][4]; // Column E
    
    if (rowDate === formDate && rowVanId === vanId) {
      // Found the matching row
      var updateColumn;
      
      // Map reporting time to column
      switch (reportingTime) {
        case '1:40 PM':
          updateColumn = 12; // Column L
          break;
        case '3:40 PM':
          updateColumn = 13; // Column M
          break;
        case '5:40 PM':
          updateColumn = 14; // Column N
          break;
        case '7:40 PM':
          updateColumn = 15; // Column O
          break;
        case '9:40 PM (End of Day)':
          updateColumn = 16; // Column P
          break;
        default:
          Logger.log('Unknown reporting time: ' + reportingTime);
          return;
      }
      
      // Update the cell
      dailyDetailsSheet.getRange(i + 1, updateColumn).setValue(deliveryCount);
      
      Logger.log('Updated Van ' + vanId + ' at ' + reportingTime + ': ' + deliveryCount + ' deliveries');
      
      // Mark form response as processed
      markFormResponseProcessed(formData);
      
      return;
    }
  }
  
  Logger.log('No matching row found for Van ' + vanId + ' on ' + formDate);
}

/**
 * Mark a form response as processed
 * @param {Object} formData - Form data to mark as processed
 */
function markFormResponseProcessed(formData) {
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var dataSheet = ss.getSheetByName('Delivery Pace Data');
  
  if (!dataSheet) {
    return;
  }
  
  var data = dataSheet.getDataRange().getValues();
  
  // Find the most recent matching unprocessed entry
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][2] === formData['Van ID'] && 
        data[i][5] === formData['Reporting Time'] &&
        data[i][8] !== 'Yes') {
      
      // Mark as processed
      dataSheet.getRange(i + 1, 9).setValue('Yes');
      break;
    }
  }
}

/**
 * Set up form submit trigger
 */
function setupFormTrigger() {
  var formId = PropertiesService.getScriptProperties().getProperty('DELIVERY_PACE_FORM_ID');
  
  if (!formId) {
    throw new Error('No form ID found. Please create the form first.');
  }
  
  // Remove existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onDeliveryPaceFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger
  ScriptApp.newTrigger('onDeliveryPaceFormSubmit')
    .forForm(formId)
    .onFormSubmit()
    .create();
  
  Logger.log('Form submit trigger created');
}

/**
 * Generate QR code for form URL (for printing/posting)
 * @return {Object} Form URL and QR code URL
 */
function generateFormQRCode() {
  var formId = PropertiesService.getScriptProperties().getProperty('DELIVERY_PACE_FORM_ID');
  
  if (!formId) {
    throw new Error('No form ID found. Please create the form first.');
  }
  
  var form = FormApp.openById(formId);
  var formUrl = form.getPublishedUrl();
  var shortUrl = form.shortenFormUrl(formUrl);
  
  // Generate QR code using quickchart.io API (more reliable than deprecated Google Charts)
  var qrCodeUrl = 'https://quickchart.io/qr?text=' + 
                  encodeURIComponent(shortUrl) + 
                  '&size=300&dark=1a73e8&light=ffffff&margin=2';
  
  Logger.log('Form URL: ' + shortUrl);
  Logger.log('QR Code URL: ' + qrCodeUrl);
  
  return {
    formUrl: shortUrl,
    qrCodeUrl: qrCodeUrl
  };
}

/**
 * Send form link to drivers via email
 * @param {string[]} emailAddresses - Array of driver email addresses
 */
function sendFormToDrivers(emailAddresses) {
  var formId = PropertiesService.getScriptProperties().getProperty('DELIVERY_PACE_FORM_ID');
  
  if (!formId) {
    throw new Error('No form ID found. Please create the form first.');
  }
  
  var form = FormApp.openById(formId);
  var formUrl = form.getPublishedUrl();
  var shortUrl = form.shortenFormUrl(formUrl);
  
  var subject = 'Delivery Pace Reporting Form';
  var body = `Hello Driver,

Please use this form to report your delivery progress at each checkpoint throughout the day:

${shortUrl}

Reporting times:
- 1:40 PM
- 3:40 PM
- 5:40 PM
- 7:40 PM
- 9:40 PM (End of Day)

Please bookmark this link on your mobile device for easy access.

Thank you for your timely reporting!`;
  
  emailAddresses.forEach(function(email) {
    if (email && email.includes('@')) {
      MailApp.sendEmail(email, subject, body);
    }
  });
  
  Logger.log('Form links sent to ' + emailAddresses.length + ' drivers');
}

/**
 * Create HTML-based delivery pace form with auto-population
 * @return {string} URL of the web app
 */
function createSmartDeliveryForm() {
  // Deploy as web app to get URL
  var url = ScriptApp.getService().getUrl();
  
  if (!url) {
    throw new Error('Please deploy this script as a web app first');
  }
  
  return url;
}

/**
 * Serve the HTML form
 * @return {HtmlOutput} HTML form page
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('DeliveryPaceForm')
    .setTitle('Delivery Pace Report')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Get form data for dropdowns
 * @return {Object} Form data including vans, routes, and assignments
 */
function getFormData() {
  return {
    vans: getActiveVanChoices(),
    routes: getTodayRouteChoices(),
    assignments: getTodayAssignments()
  };
}

/**
 * Process form submission from HTML form
 * @param {Object} formData - Form data object
 * @return {Object} Result object
 */
function submitDeliveryPaceReport(formData) {
  try {
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    
    // Get or create Delivery Pace Data sheet
    var dataSheet = ss.getSheetByName('Delivery Pace Data');
    if (!dataSheet) {
      dataSheet = ss.insertSheet('Delivery Pace Data');
      setupDeliveryPaceDataSheet(dataSheet);
    }
    
    // Add timestamp
    var timestamp = new Date();
    
    // Prepare row data
    var rowData = [
      timestamp,
      formData.date,
      formData.vanId,
      formData.driverName,
      formData.routeCode,
      formData.reportingTime,
      formData.deliveryCount,
      formData.notes || '',
      'No' // Processed flag
    ];
    
    // Append to sheet
    dataSheet.appendRow(rowData);
    
    // Update Daily Details immediately
    updateDailyDetailsFromForm({
      'Date': formData.date,
      'Van ID': formData.vanId,
      'Driver Name': formData.driverName,
      'Route Code': formData.routeCode,
      'Reporting Time': formData.reportingTime,
      'Total Deliveries Completed': formData.deliveryCount
    });
    
    Logger.log('Delivery pace report submitted for Van: ' + formData.vanId);
    
    return {
      success: true,
      message: 'Report submitted successfully'
    };
    
  } catch (error) {
    Logger.log('Error submitting delivery pace report: ' + error);
    throw error;
  }
}