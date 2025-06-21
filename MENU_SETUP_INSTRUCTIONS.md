# Menu Setup Instructions

Since the menus are not appearing automatically, you need to set them up manually through the Google Apps Script Editor.

## Step-by-Step Instructions

### 1. Open the Script Editor
- In your Google Sheets, go to **Extensions** → **Apps Script**
- This will open the Google Apps Script editor in a new tab

### 2. Find the Setup Function
- In the script editor, look for the file list on the left side
- Click on **Main.js** (it might be called Main.gs)
- Use Ctrl+F (or Cmd+F on Mac) to search for `setupMenus`

### 3. Run the Setup Function
- Click anywhere inside the `setupMenus` function
- At the top of the editor, you'll see a dropdown that says "Select function"
- Make sure it shows **setupMenus**
- Click the **Run** button (looks like a play button ▶️)

### 4. Grant Permissions (First Time Only)
- If this is your first time running the script, you'll see a permission dialog
- Click **Review Permissions**
- Choose your Google account
- Click **Advanced** (at the bottom left)
- Click **Go to Fleet Resource Allocator (unsafe)**
- Click **Allow**

### 5. Wait for Success Message
- The script will run and show a success message
- You should see "Menu Initialization Complete" 

### 6. Return to Your Spreadsheet
- Go back to your Google Sheets tab
- You should now see all menus:
  - Vehicle Assignment Tool
  - Delivery Pace
  - Fleet Operations
  - Reports
  - Help

## Alternative Method: Run from Script Editor Console

If the above doesn't work, try this:

1. In the Script Editor, look for the "Execution Log" at the bottom
2. If you don't see it, go to **View** → **Logs**
3. In the code editor, add this temporary function at the bottom of Main.js:

```javascript
function testMenuCreation() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Testing menu creation...');
  
  try {
    ui.createMenu("Test Menu")
      .addItem("Test Item", "showUploadDialog")
      .addToUi();
    ui.alert('Test menu created! Check if it appears.');
  } catch (error) {
    ui.alert('Error: ' + error.toString());
  }
}
```

4. Run this function to see what error messages appear

## If Menus Still Don't Appear

### Check for Errors
1. In Script Editor, go to **View** → **Executions**
2. Look for any failed executions
3. Click on them to see error details

### Common Issues
- **Missing Functions**: The error log might show "ReferenceError: functionName is not defined"
- **Permission Issues**: Make sure you've authorized the script
- **File Loading**: Some .js files might not be loading properly

### Manual Menu Creation
As a last resort, you can create a simplified menu directly in the Script Editor:

1. Replace the entire `onOpen()` function in Main.js with:

```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Fleet Tools")
    .addItem("Upload Files", "showUploadDialog")
    .addItem("Initialize Headers", "initializeDeliveryPaceHeaders")
    .addItem("Driver Report", "generateDriverPerformanceReport") 
    .addItem("Export Data", "exportAllData")
    .addToUi();
}
```

2. Save the file (Ctrl+S or Cmd+S)
3. Refresh your spreadsheet

## Contact Support

If none of these methods work:
1. Take a screenshot of any error messages
2. Check the Execution log in Apps Script
3. Note which step failed
4. Contact your system administrator with this information