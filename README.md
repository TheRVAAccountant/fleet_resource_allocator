# Fleet Resource Allocator

## Overview
The Fleet Resource Allocator is a Google Apps Script application that automates vehicle assignment to delivery routes. It processes Excel files to match available vehicles with routes based on service types and operational status.

## Recent Updates (2025-06-21)

### ✅ Completed Enhancements
1. **Report Functions**
   - Driver Performance Report - Analyzes individual driver metrics
   - Weekly Summary Report - Provides comprehensive weekly overview

2. **Data Export Features**  
   - Export All Data - Exports all sheets to CSV format in Google Drive
   - Export Current Week - Quick export of current week's data

3. **Form Management UI**
   - Central interface to manage all system forms
   - Enable/disable forms, configure settings, view URLs

4. **Comprehensive Analytics**
   - Daily allocation success rate
   - Average delivery completion time
   - Van utilization by type
   - Weekly trend analysis
   - Analytics dashboard

### 🐛 Fixed Issues
- Menu creation issues have been resolved
- All menus should now appear correctly: Vehicle Assignment Tool, Fleet Operations, Reports, and Help

## How to Use

### Basic Vehicle Allocation
1. Open the Google Sheets containing this script
2. Go to **Vehicle Assignment Tool** → **Upload Files for Allocation**
3. Select your Day of Ops and Daily Routes Excel files
4. Wait for processing to complete
5. Check the Results sheet in the Daily Summary spreadsheet

### Access New Features

#### View Reports
- **Reports** → **Driver Performance** - Analyze driver metrics
- **Reports** → **Weekly Summary** - Get weekly overview
- **Reports** → **Analytics Dashboard** - View all analytics

#### Export Data
- **Reports** → **Export All Data** - Export everything to Drive
- **Reports** → **Export Current Week** - Quick weekly export

#### Manage Forms
- **Fleet Operations** → **Form Management** - Configure all forms

#### Access Dashboard
- **Fleet Operations** → **View Dashboard** - Real-time fleet status

## Menu Structure

### Vehicle Assignment Tool
- Upload Files for Allocation
- Run Menu Diagnostics
- Test Required Functions

### Fleet Operations
- Allocate Vehicles
- View Dashboard
- Delivery Pace Form
- RTS Report
- Form Management
- View Error Log

### Reports
- Vehicle Utilization
- Driver Performance
- Weekly Summary
- Analytics Dashboard
- Export All Data

### Help
- User Guide
- About

## Troubleshooting

### Menus Not Appearing
1. Refresh the spreadsheet (Ctrl+R or Cmd+R)
2. Close and reopen the spreadsheet
3. Use **Vehicle Assignment Tool** → **Test Required Functions** to check for missing functions

### Allocation Errors
1. Check that your Excel files have the correct sheet names:
   - Day of Ops needs sheet named "Solution"
   - Daily Routes needs sheet named "Routes"
2. Verify required columns exist
3. Check **Fleet Operations** → **View Error Log** for details

### Report Generation Issues
1. Ensure Daily Details sheet has data
2. Check date formats are consistent
3. View Error Log for specific error messages

## Key Files

### Core Services
- `AllocationService.js` - Vehicle allocation logic
- `DashboardService.js` - Dashboard functionality
- `ReportService.js` - Report generation
- `AnalyticsService.js` - Analytics engine
- `FormManagementService.js` - Form management

### UI Components
- `Dashboard.html` - Dashboard interface
- `DeliveryPaceForm.html` - Delivery pace reporting
- `RTSForm.html` - Return to station reporting
- `UploadDialog.html` - File upload interface

### Infrastructure
- `Config.js` - Configuration settings
- `Logger.js` - Logging system
- `ErrorHandler.js` - Error handling
- `SheetManager.js` - Sheet operations
- `Cache.js` - Performance caching

## Configuration

Key settings in `Config.js`:
- `DAILY_SUMMARY_SPREADSHEET_ID` - Main data spreadsheet
- `ROUTE_ASSIGNMENTS_FOLDER_ID` - Output folder
- `VAN_TYPE_MAPPING` - Service type to van type mapping

## Support

For issues or questions:
1. Check the Error Log first
2. Run Menu Diagnostics
3. Test Required Functions
4. Contact your system administrator

## Version History

- **v2.0** (2025-06-21) - Added comprehensive reporting, analytics, and form management
- **v1.5** - Added delivery pace tracking and RTS reporting
- **v1.0** - Initial release with basic allocation functionality