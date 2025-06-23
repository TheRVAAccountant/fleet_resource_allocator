# Menu Creation Fix Plan

## Step 1: Verify Script Version
First, we need to confirm which version of the code is actually running.

### Add Version Tracking
Add this to the top of Main.js:
```javascript
var SCRIPT_VERSION = "2.1.0"; // Update this when making changes

function getScriptVersion() {
  SpreadsheetApp.getUi().alert("Script Version", "Current version: " + SCRIPT_VERSION, SpreadsheetApp.getUi().ButtonSet.OK);
}
```

## Step 2: Create Diagnostic onOpen
Replace the current onOpen with a diagnostic version that shows exactly where it fails:

```javascript
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var steps = [];
  
  try {
    steps.push("Starting onOpen");
    
    // Step 1: Create basic menu
    ui.createMenu("Vehicle Assignment Tool")
      .addItem("Upload Files for Allocation", "showUploadDialog")
      .addItem("Show Diagnostic Log", "showDiagnosticLog")
      .addToUi();
    steps.push("Created Vehicle Assignment Tool menu");
    
    // Store diagnostic info
    PropertiesService.getScriptProperties().setProperty('menuDiagnostic', JSON.stringify(steps));
    
    // Step 2: Try to create second menu
    ui.createMenu("Test Menu 2")
      .addItem("Test Item", "showUploadDialog")
      .addToUi();
    steps.push("Created Test Menu 2");
    
  } catch (error) {
    steps.push("Error: " + error.toString());
    PropertiesService.getScriptProperties().setProperty('menuDiagnostic', JSON.stringify(steps));
  }
}

function showDiagnosticLog() {
  var log = PropertiesService.getScriptProperties().getProperty('menuDiagnostic');
  SpreadsheetApp.getUi().alert("Diagnostic Log", log || "No diagnostic data", SpreadsheetApp.getUi().ButtonSet.OK);
}
```

## Step 3: Nuclear Option - Single File Approach
Create a new file called `MenuOnly.js` with EVERYTHING needed for menus in one place:

```javascript
// MenuOnly.js - Self-contained menu system
function onOpenMenuOnly() {
  var ui = SpreadsheetApp.getUi();
  
  // Create menus with inline functions
  ui.createMenu("Fleet Operations v2")
    .addItem("Upload Files", "menuUploadFiles")
    .addItem("Initialize Headers", "menuInitHeaders")
    .addItem("Vehicle Report", "menuVehicleReport")
    .addToUi();
}

function menuUploadFiles() {
  SpreadsheetApp.getUi().alert("Upload Files", "Feature coming soon", SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuInitHeaders() {
  SpreadsheetApp.getUi().alert("Initialize Headers", "Feature coming soon", SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuVehicleReport() {
  SpreadsheetApp.getUi().alert("Vehicle Report", "Feature coming soon", SpreadsheetApp.getUi().ButtonSet.OK);
}
```

## Step 4: Force Script Reload
1. In Script Editor, make a small change (add a space)
2. Save (Ctrl+S)
3. Click "Deploy" > "Test deployments"
4. Create a new test deployment
5. This forces Google to reload all scripts

## Step 5: Check Script Load Order
The issue might be file load order. Google Apps Script loads files alphabetically. Check if critical functions are in files that load after Main.js.

### Solution: Create `_Init.js` (underscore makes it load first)
```javascript
// _Init.js - Loads before all other files
// Define critical functions here that menus need
function showUploadDialog() {
  SpreadsheetApp.getUi().alert("Upload Dialog", "Placeholder", SpreadsheetApp.getUi().ButtonSet.OK);
}
```

## Step 6: Authorization Check
The script might be failing due to authorization issues.

```javascript
function checkAuthorization() {
  try {
    // Try to access all services
    SpreadsheetApp.getActiveSpreadsheet();
    DriveApp.getRootFolder();
    FormApp.create('temp').getId();
    
    SpreadsheetApp.getUi().alert("Authorization OK", "All services authorized", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("Authorization Error", error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
```

## Step 7: Manual Override
If all else fails, create a custom function that users can run manually:

```javascript
function installMenus() {
  // Clear any cached menus
  SpreadsheetApp.getActiveSpreadsheet().removeMenu("Vehicle Assignment Tool");
  
  // Wait a moment
  Utilities.sleep(1000);
  
  // Create all menus
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Fleet Tools")
    .addItem("Upload", "showUploadDialog")
    .addToUi();
    
  ui.createMenu("Reports")
    .addItem("Vehicle Report", "generateVehicleUtilizationReport")
    .addToUi();
    
  SpreadsheetApp.getUi().alert("Success", "Menus installed. Please refresh the spreadsheet.", SpreadsheetApp.getUi().ButtonSet.OK);
}
```

## Immediate Action Plan

1. **First**: Run this in Script Editor to check what's happening:
```javascript
function debugOnOpen() {
  try {
    onOpen();
    SpreadsheetApp.getUi().alert("onOpen executed without errors");
  } catch (error) {
    SpreadsheetApp.getUi().alert("onOpen failed: " + error.toString());
  }
}
```

2. **Second**: Check if it's a simple typo or syntax error:
```javascript
function validateMenuCode() {
  var ui = SpreadsheetApp.getUi();
  
  // Test 1: Can we create any menu?
  try {
    ui.createMenu("Test1").addItem("Item", "showUploadDialog").addToUi();
    console.log("Test 1 passed");
  } catch (e) {
    console.log("Test 1 failed: " + e);
  }
  
  // Test 2: Can we create submenus?
  try {
    ui.createMenu("Test2")
      .addSubMenu(ui.createMenu("Sub")
        .addItem("Item", "showUploadDialog"))
      .addToUi();
    console.log("Test 2 passed");
  } catch (e) {
    console.log("Test 2 failed: " + e);
  }
}
```

3. **Third**: Nuclear option - bypass onOpen entirely:
- Create a Google Sheets custom menu using a different trigger
- Use installable triggers instead of simple triggers
- Create a sidebar with all functions instead of menus