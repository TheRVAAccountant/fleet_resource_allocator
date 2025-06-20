# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
This is a Google Apps Script project for fleet resource allocation that automates vehicle assignment to delivery routes. The system processes Excel files to match available vehicles with routes based on service types and operational status.

## Architecture & Key Components

### Modular File Structure
The codebase is organized into separate modules following Google Apps Script best practices:

- **Config.js**: Configuration constants and external resource IDs
- **Main.js**: Entry points, menu setup, and initialization
- **AllocationService.js**: Vehicle allocation business logic
- **DeliveryPaceService.js**: Delivery pace tracking functionality
- **FormService.js**: Google Forms integration for pace data collection
- **SheetService.js**: Google Sheets operations and data management
- **UIService.js**: User interface handlers and dialogs
- **Utils.js**: Utility functions and helpers
- **UploadDialog.html**: HTML interface for file upload functionality
- **appsscript.json**: Google Apps Script manifest configuration

### Code Organization Benefits
- **Separation of Concerns**: Each service handles a specific domain
- **Improved Readability**: Related functions grouped together
- **Easier Maintenance**: Changes isolated to relevant modules
- **Better Testing**: Individual modules can be tested independently
- **Configuration Management**: All settings centralized in Config.js

### Key Functionality
1. **File Upload System**: Processes XLSX files (Day of Ops & Daily Routes) via web dialog
2. **Vehicle Allocation**: Matches routes to available vehicles based on service type requirements
3. **Results Generation**: Creates detailed allocation reports and unassigned vehicle lists
4. **Daily Details Tracking**: Updates cumulative tracking sheet with duplicate prevention
5. **Route Assignments Export**: Generates separate files for distribution

### External Dependencies
- **Daily Summary Spreadsheet ID**: `1fgwW9tcozBqiB6zrpg7jzactFMkzpRXCcmPs0eUsaqI` (Updated 6-19-25)
- **Route Assignments Folder ID**: `1_WxSEO5uw47pkuSzfMlrQTTU67Jafb2z`
- **Google Drive API v3**: Required for file operations

## Development Commands

### Google Apps Script Commands
```bash
# This is a Google Apps Script project - use the Apps Script editor for development
# Access via: https://script.google.com/
# Or through Google Sheets: Extensions > Apps Script

# To deploy:
# 1. Save all files in Apps Script editor
# 2. Click Deploy > New Deployment
# 3. Select type: Web app or Add-on
# 4. Configure settings and click Deploy

# To test:
# 1. Click Run > Select function to test
# 2. Use Test as add-on for UI testing
# 3. Check Executions for logs
```

### Local Development Setup (if using clasp)
```bash
# Install clasp globally
npm install -g @google/clasp

# Login to Google
clasp login

# Clone existing project (replace with your script ID)
clasp clone <scriptId>

# Push changes
clasp push

# Open in Apps Script editor
clasp open
```

## Key Business Logic

### Vehicle Type Mapping
The system maps service types to vehicle types:
- `"Standard Parcel - Extra Large Van - US"` → `"Extra Large"`
- `"Standard Parcel - Large Van"` → `"Large"`
- `"Standard Parcel Step Van - US"` → `"Step Van"`
- Routes containing `"Nursery Route Level"` → `"Large"`

### Data Flow
1. User uploads Day of Ops (contains routes) and Daily Routes (contains driver assignments)
2. System filters for DSP = "BWAY" routes only
3. Allocates operational vehicles (Opnal Y/N = "Y") to routes
4. Creates Results sheet with allocations
5. Updates Daily Details with unique identifier preventing duplicates
6. Generates Route Assignments file with reordered columns

### Unique Identifier Format
`MM/DD/YYYY|Route Code|Associate Name|Van ID`

Used for duplicate prevention in Daily Details tracking.

## Important Sheet Names & Structures

### Input Sheets
- **Day of Ops**: Sheet name "Solution" - requires columns: Route Code, Service Type, DSP, Wave, Staging Location
- **Daily Routes**: Sheet name "Routes" - requires columns: Route code, Driver name
- **Vehicle Status**: Sheet name "Vehicle Status" - requires columns: Van ID, Type, Opnal?\nY/N

### Output Sheets
- **Results**: Named as "MM-DD-YY - Results" with 11 columns including Unique Identifier
- **Unassigned Vans**: Named as "MM-DD-YY - Available & Unassigned Vans"
- **Daily Details**: Cumulative tracking sheet with columns A-E for core data, column V for Unique ID

## Common Tasks

### Running the Allocation
1. Open the spreadsheet containing this script
2. Click "Vehicle Assignment Tool" menu > "Upload Files for Allocation"
3. Select Day of Ops XLSX and Daily Routes XLSX files
4. Wait for processing to complete
5. Check newly created sheets in Daily Summary spreadsheet

### Debugging Issues
- Check Apps Script Executions log for detailed error messages
- Verify sheet names match expected values (case-sensitive)
- Ensure required columns exist in input files
- Confirm Drive API v3 is enabled in Advanced Services
- Validate spreadsheet IDs and folder permissions

### Modifying Vehicle Type Mappings
Edit the `VAN_TYPE_MAPPING` object in Config.js or update the `getVanType()` function in Utils.js to add new service type mappings.

### Adjusting Column Order in Route Assignments
Modify the `columnOrder` array in the `createRouteAssignmentsFile()` function in AllocationService.js to change output column arrangement.

### Adding New Configuration
All configuration values should be added to the CONFIG object in Config.js and accessed using the `getConfig()` helper function.

## Error Handling Patterns
- File upload errors are caught and displayed to user
- Missing required columns throw descriptive errors
- Empty sheets are detected and reported
- Duplicate entries are logged but don't halt processing
- COM-style explicit resource cleanup for spreadsheet operations

## Daily Summary Spreadsheet Structure

Based on analysis of the Daily Summary Log 2025 spreadsheet, the Google Apps Script interacts with the following structure:

### Sheet Organization
The spreadsheet contains 250+ sheets including:
- **Daily sheets**: Named by date (e.g., "06-18-25", "7-1-23") - historical allocations
- **Results sheets**: Named as "MM-DD-YY - Results" - contains allocation outputs
- **Unassigned sheets**: Named as "MM-DD-YY - Available & Unassign" - lists unallocated vehicles
- **Core reference sheets**: Vehicle Status, Daily Details, Validation, etc.

### Vehicle Status Sheet
- **Purpose**: Master list of all fleet vehicles
- **Key columns** (15 total):
  - Van ID: Unique identifier (e.g., "BW1", "BW2")
  - Type: Vehicle category (Extra Large: 24, Large: 7, Step Van: 5)
  - Opnal?\nY/N: Operational status (Y=27 operational, N=9 non-operational)
  - Additional fields: Year, Make, Model, License info, VIN, Issues, Grounding dates

### Daily Details Sheet  
- **Purpose**: Cumulative log of all route assignments
- **Structure**: 22 columns, 1588+ rows of historical data
- **Core columns (A-E)**:
  - A: Date (allocation date)
  - B: Route # (route code)
  - C: Name (driver/associate name)
  - D: Asset ID (typically empty)
  - E: Van ID (assigned vehicle)
- **Column V (index 21)**: Unique Identifier for duplicate prevention
- **Unique ID Format**: `MM/DD/YYYY|Route Code|Associate Name|Van ID`
  - Example: `02/11/2025|CX9|Marquis Thomas|BW2`

### Results Sheet Structure
Created by the script with 11 columns:
1. Route Code
2. Service Type
3. DSP (Delivery Service Partner)
4. Wave
5. Staging Location
6. Van ID
7. Device Name (same as Van ID)
8. Van Type
9. Operational (Y/N status)
10. Associate Name
11. Unique Identifier

### Data Flow Integration
1. Script reads Vehicle Status to get available operational vehicles
2. Filters and allocates vehicles based on type matching
3. Creates new Results sheet with allocation data
4. Appends allocation data to Daily Details sheet (checking for duplicates via Unique ID)
5. Creates Available & Unassigned sheet for remaining vehicles
6. Historical sheets preserve daily allocation records

## Delivery Pace Tracking System

### Overview
The enhanced Google Apps Script now includes a comprehensive delivery pace tracking system that monitors van delivery progress throughout the day at 2-hour intervals.

### Time-Based Tracking Columns (L-P)
- **Column L**: 1:40 PM checkpoint
- **Column M**: 3:40 PM checkpoint  
- **Column N**: 5:40 PM checkpoint
- **Column O**: 7:40 PM checkpoint
- **Column P**: 9:40 PM checkpoint

Each column tracks the cumulative number of stops completed by that time.

### Key Functions

#### Core Tracking Functions (DeliveryPaceService.js)
- `initializeDeliveryPaceHeaders()`: Sets up column headers for time-based tracking
- `updateDeliveryPaceForToday()`: Updates pace data for all vans allocated today
- `updateDeliveryPaceForVan(vanId, date)`: Updates pace for a specific van
- `getDeliveryPaceData(vanId, date)`: Fetches pace data (currently mock data, ready for integration)

#### Automation Functions (DeliveryPaceService.js)
- `setupDeliveryPaceTriggers()`: Creates time-based triggers to run updates automatically at:
  - 1:45 PM, 3:45 PM, 5:45 PM, 7:45 PM, 9:45 PM
- `batchUpdateDeliveryPace(vanIds, date)`: Updates multiple vans in one operation

#### Reporting Functions (DeliveryPaceService.js)
- `generateDeliveryPaceSummary(date)`: Creates comprehensive summary report
- `createDeliveryPaceSummarySheet(summary)`: Generates formatted summary sheet with:
  - Total vans allocated
  - Average stops by time period
  - Individual van performance details

### Data Collection via Google Forms

The system now includes a comprehensive form-based data collection system for drivers to report their delivery progress:

#### Form Features
- **Mobile-Friendly**: Optimized for smartphone use
- **Pre-populated Fields**: Van IDs loaded from Vehicle Status
- **Time Checkpoints**: All 5 daily reporting times
- **Validation**: Ensures numeric delivery counts
- **Optional Notes**: Drivers can report issues

#### Form Management (FormService.js)
- `createDeliveryPaceForm()`: Creates/updates the collection form
- `onDeliveryPaceFormSubmit()`: Processes form submissions automatically
- `generateFormQRCode()`: Creates QR code for easy mobile access
- `sendFormToDrivers()`: Email form links to driver list

#### Data Flow
1. **Driver submits form** at checkpoint time (1:40 PM, 3:40 PM, etc.)
2. **Form response saved** to "Delivery Pace Data" sheet
3. **Trigger processes submission** and updates Daily Details
4. **getDeliveryPaceData()** reads from form responses first
5. **Summary reports** aggregate all submitted data

#### Setting Up Forms
1. Go to `Delivery Pace` → `Form Management` → `Create/Update Collection Form`
2. System creates form linked to Daily Summary spreadsheet
3. Get form URL or QR code from `Get Form Link & QR Code`
4. Set up automatic processing with `Setup Form Trigger`
5. Share form link with drivers (bookmark on phones)

### Alternative Integration Points
If forms don't meet your needs, the system still supports:

1. **Direct API Integration**
2. **Database Connections**
3. **Other Google Sheets**
4. **CSV Imports**

### Menu Structure
The Delivery Pace menu provides:
- Initialize Headers
- Update Today's Pace
- Generate Today's Summary
- Update Specific Van (with dialog)
- Setup Auto-Update Triggers
- Test Update

### Usage Workflow
1. Run "Initialize Headers" once to set up column headers
2. Use "Setup Auto-Update Triggers" to enable automatic updates
3. Manual updates available via "Update Today's Pace"
4. Generate summaries with "Generate Today's Summary"
5. Update individual vans through the dialog interface

## Smart Form Van Filtering Enhancement

The delivery pace form has been enhanced to intelligently filter the Van ID dropdown to show only vans that are assigned for the current day. This reduces errors and improves the user experience.

### Key Features
- **Filtered Dropdown**: Shows only vans with assignments for today
- **Auto-Selection**: If only one van is assigned, it's automatically selected
- **Visual Feedback**: Displays "Showing only vans assigned today" message
- **Fallback Logic**: Shows all operational vans if no assignments exist yet
- **Configurable**: Can be enabled/disabled via configuration

### Configuration
In Config.js:
```javascript
FORM_SETTINGS: {
  FILTER_VANS_BY_ASSIGNMENT: true  // Set to false to show all operational vans
}
```

### How It Works
1. When the form loads, it queries Daily Details for today's assignments
2. Extracts unique van IDs that have been assigned routes
3. Populates the dropdown with only these vans
4. If no vans are assigned, falls back to showing all operational vans

### Testing
To test the filtering:
1. Go to "Delivery Pace" → "Form Management" → "Test Van Filtering"
2. The test will show:
   - Number of vans assigned today
   - Number shown in filtered dropdown
   - Total operational vans available

### Benefits
- **Reduced Errors**: Drivers can't accidentally select unassigned vans
- **Faster Selection**: Fewer options to choose from
- **Better UX**: Clear indication of what vans are available
- **Smart Defaults**: Auto-selects when only one option exists