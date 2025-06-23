/**
 * ===================================================================
 * SIMPLE MENU - SELF-CONTAINED MINIMAL IMPLEMENTATION
 * ===================================================================
 * This file starts with underscore to load first alphabetically.
 * It contains a minimal, self-contained menu system that doesn't
 * depend on any other files.
 */

/**
 * Super simple onOpen - if even this doesn't work, there's a deeper issue
 */
function onOpenSimple() {
  SpreadsheetApp.getUi()
    .createMenu("Fleet System")
    .addItem("Upload", "simpleUpload")
    .addItem("Status", "simpleStatus")
    .addToUi();
}

/**
 * Simple upload function
 */
function simpleUpload() {
  SpreadsheetApp.getUi().alert("Upload", "Upload feature placeholder", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Simple status function
 */
function simpleStatus() {
  SpreadsheetApp.getUi().alert("Status", "System is running", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Install menus manually - run this from Script Editor
 */
function installMenusManually() {
  // First, let's check what's preventing menu creation
  var issues = [];
  
  // Check 1: Can we access UI?
  try {
    var ui = SpreadsheetApp.getUi();
    issues.push("✓ Can access UI");
  } catch (e) {
    issues.push("✗ Cannot access UI: " + e.toString());
    SpreadsheetApp.getUi().alert("Critical Error", "Cannot access UI service", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Check 2: Can we create a simple menu?
  try {
    SpreadsheetApp.getUi()
      .createMenu("Test")
      .addItem("Item", "simpleStatus")
      .addToUi();
    issues.push("✓ Can create simple menu");
  } catch (e) {
    issues.push("✗ Cannot create menu: " + e.toString());
  }
  
  // Check 3: What about the onOpen function?
  if (typeof onOpen === 'function') {
    issues.push("✓ onOpen function exists");
    
    // Try to see what's in onOpen
    var onOpenString = onOpen.toString();
    if (onOpenString.indexOf("Vehicle Assignment Tool") > -1) {
      issues.push("✓ onOpen contains expected menu code");
    } else {
      issues.push("⚠ onOpen might not have the right code");
    }
  } else {
    issues.push("✗ onOpen function is missing!");
  }
  
  // Show diagnostic results
  SpreadsheetApp.getUi().alert(
    "Installation Diagnostic",
    issues.join("\n") + "\n\nNow attempting to create menus...",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  // Now try to create all menus
  var ui = SpreadsheetApp.getUi();
  var menusCreated = [];
  
  // Menu 1: Basic Operations
  try {
    ui.createMenu("Fleet Operations")
      .addItem("Upload Files", "showUploadDialog")
      .addItem("Vehicle Status", "checkVehicleStatus") 
      .addItem("Daily Details", "checkDailyDetails")
      .addToUi();
    menusCreated.push("Fleet Operations");
  } catch (e) {
    // Continue
  }
  
  // Menu 2: Delivery Pace
  try {
    ui.createMenu("Delivery Pace")
      .addItem("Initialize", "initDeliveryPace")
      .addItem("Update", "updateDeliveryPace")
      .addToUi();
    menusCreated.push("Delivery Pace");
  } catch (e) {
    // Continue
  }
  
  // Menu 3: Reports
  try {
    ui.createMenu("Reports")
      .addItem("Generate Report", "generateSimpleReport")
      .addToUi();
    menusCreated.push("Reports");
  } catch (e) {
    // Continue
  }
  
  // Show results
  if (menusCreated.length > 0) {
    SpreadsheetApp.getUi().alert(
      "Success",
      "Created " + menusCreated.length + " menus:\n" + menusCreated.join(", ") + 
      "\n\nPlease refresh the spreadsheet to see them.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "Failed",
      "Could not create any menus. There may be a permissions issue.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// Minimal placeholder functions
function checkVehicleStatus() {
  SpreadsheetApp.getUi().alert("Vehicle Status", "Checking vehicle status...", SpreadsheetApp.getUi().ButtonSet.OK);
}

function checkDailyDetails() {
  SpreadsheetApp.getUi().alert("Daily Details", "Checking daily details...", SpreadsheetApp.getUi().ButtonSet.OK);
}

function initDeliveryPace() {
  SpreadsheetApp.getUi().alert("Initialize", "Initializing delivery pace...", SpreadsheetApp.getUi().ButtonSet.OK);
}

function updateDeliveryPace() {
  SpreadsheetApp.getUi().alert("Update", "Updating delivery pace...", SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateSimpleReport() {
  SpreadsheetApp.getUi().alert("Report", "Generating report...", SpreadsheetApp.getUi().ButtonSet.OK);
}