/**
 * ===================================================================
 * CONFIGURATION AND CONSTANTS
 * ===================================================================
 * Central location for all configuration values, constants, and 
 * external resource IDs used throughout the application.
 */

/**
 * External Resource IDs
 */
var CONFIG = {
  // Spreadsheet IDs
  DAILY_SUMMARY_SPREADSHEET_ID: "1fgwW9tcozBqiB6zrpg7jzactFMkzpRXCcmPs0eUsaqI",
  ROUTE_ASSIGNMENTS_FOLDER_ID: "1_WxSEO5uw47pkuSzfMlrQTTU67Jafb2z",
  
  // Sheet Names
  SHEETS: {
    VEHICLE_STATUS: "Vehicle Status",
    DAILY_DETAILS: "Daily Details",
    DAY_OF_OPS_SOLUTION: "Solution",
    DAILY_ROUTES: "Routes"
  },
  
  // Column Headers
  REQUIRED_COLUMNS: {
    DAY_OF_OPS: ["Route Code", "Service Type", "DSP", "Wave", "Staging Location"],
    VEHICLE_STATUS: ["Van ID", "Type", "Opnal?\nY/N"],
    DAILY_ROUTES: ["Route code", "Driver name"]
  },
  
  // Time Slots for Delivery Pace
  DELIVERY_TIME_SLOTS: [
    {time: 13.67, column: 12, label: "1:40 PM"},
    {time: 15.67, column: 13, label: "3:40 PM"},
    {time: 17.67, column: 14, label: "5:40 PM"},
    {time: 19.67, column: 15, label: "7:40 PM"},
    {time: 21.67, column: 16, label: "9:40 PM"}
  ],
  
  // Daily Details Fields - Column indices (0-based)
  DAILY_DETAILS_FIELDS: {
    DATE: 0,             // Column A (index 0)
    ROUTE_NUMBER: 1,     // Column B (index 1)
    NAME: 2,             // Column C (index 2)
    ASSET_ID: 3,         // Column D (index 3)
    VAN_ID: 4,           // Column E (index 4)
    WEEK_NUMBER: 20,     // Column U (index 20) - Fixed from incorrect index 5
    UNIQUE_ID: 21        // Column V (index 21)
  },
  
  // RTS (Return to Station) Fields - Column indices (0-based)
  RTS_FIELDS: {
    RTS_TIME: 16,        // Column Q (index 16)
    PKG_DELIVERED: 17,   // Column R (index 17)
    PKG_RETURNED: 18,    // Column S (index 18)
    ROUTE_NOTES: 19      // Column T (index 19)
  },
  
  // Van Type Mappings
  VAN_TYPE_MAPPING: {
    "Standard Parcel - Extra Large Van - US": "Extra Large",
    "Standard Parcel - Large Van": "Large",
    "Standard Parcel Step Van - US": "Step Van"
  },
  
  // DSP Filter
  TARGET_DSP: "BWAY",
  
  // UI Settings
  UI: {
    UPLOAD_DIALOG_WIDTH: 550,
    UPLOAD_DIALOG_HEIGHT: 650,
    UPDATE_VAN_DIALOG_WIDTH: 400,
    UPDATE_VAN_DIALOG_HEIGHT: 300,
    RTS_FORM_WIDTH: 600,
    RTS_FORM_HEIGHT: 800
  },
  
  // Email Settings
  EMAIL_RECIPIENT: "info@thervaaccountant.com",
  
  // Form Settings
  FORM_SETTINGS: {
    FILTER_VANS_BY_ASSIGNMENT: true  // If true, only show vans assigned today in forms
  },
  
  // Development Settings
  DEV_SETTINGS: {
    USE_MOCK_DATA: false,  // Set to true only for testing/development
    MOCK_DATA_ENABLED_VANS: []  // Specific vans to use mock data for (empty = none)
  },
  
  // Logging Settings
  LOGGING: {
    level: 'INFO',  // Options: DEBUG, INFO, WARN, ERROR, CRITICAL
    persistToSheet: true,  // Log errors to Error Log sheet
    externalEndpoint: null  // Optional external logging service URL
  },
  
  // Cache Settings
  CACHE: {
    defaultTTL: 300,  // Default cache time-to-live in seconds (5 minutes)
    maxSize: 100,  // Maximum number of cache entries
    enabled: true  // Global cache enable/disable
  },
  
  // UI Settings Extended
  UI_SETTINGS: {
    showAnimations: true,
    autoCloseDialogs: true,
    autoCloseDuration: 2000,  // milliseconds
    progressUpdateInterval: 100,  // Update progress every N items
    theme: 'light'  // 'light' or 'dark'
  },
  
  // Performance Settings
  PERFORMANCE: {
    batchSize: 100,  // Number of items to process in batch operations
    maxConcurrent: 10,  // Maximum concurrent operations
    timeout: 300000  // Operation timeout in milliseconds (5 minutes)
  },
  
  // Feature Flags
  FEATURES: {
    DASHBOARD_ENABLED: true,
    AUTO_EMAIL_NOTIFICATIONS: true,
    PACE_TRACKING_ENABLED: true,
    RTS_TRACKING_ENABLED: true,
    SMART_DEFAULTS_ENABLED: true,
    DEV_TOOLS_ENABLED: true,
    PROGRESS_TRACKING_ENABLED: true,
    ADVANCED_CACHING_ENABLED: true
  }
};

/**
 * Get configuration value
 * @param {string} key - Dot notation path to config value
 * @return {*} Configuration value
 */
function getConfig(key) {
  var keys = key.split('.');
  var value = CONFIG;
  
  for (var i = 0; i < keys.length; i++) {
    value = value[keys[i]];
    if (value === undefined) {
      throw new Error("Configuration key not found: " + key);
    }
  }
  
  return value;
}