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
    UPDATE_VAN_DIALOG_HEIGHT: 300
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