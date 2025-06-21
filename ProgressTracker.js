/**
 * ===================================================================
 * PROGRESS TRACKER
 * ===================================================================
 * Track and display progress for long-running operations
 * Rewritten for Google Apps Script compatibility (no ES6 classes)
 */

/**
 * ProgressTracker constructor
 * @param {string} title - Title for the progress dialog
 * @param {number} total - Total number of items to process
 * @param {Object} options - Configuration options
 */
function ProgressTracker(title, total, options) {
  this.title = title;
  this.total = total;
  this.current = 0;
  this.startTime = new Date().getTime();
  
  // Merge options with defaults
  this.options = {
    showETA: true,
    showPercentage: true,
    showCurrent: true,
    updateInterval: 1 // Update UI every N items
  };
  
  // Apply user options
  if (options) {
    for (var key in options) {
      if (options.hasOwnProperty(key)) {
        this.options[key] = options[key];
      }
    }
  }
  
  this.dialog = null;
  this.createDialog();
}

/**
 * Create the progress dialog
 */
ProgressTracker.prototype.createDialog = function() {
  var html = HtmlService.createHtmlOutput(
    '<div id="progressContainer" style="padding: 20px; font-family: \'Google Sans\', Arial, sans-serif;">' +
    '<h3 style="margin: 0 0 20px 0; color: #202124;">' + this.title + '</h3>' +
    
    '<div style="background: #e0e0e0; height: 8px; border-radius: 4px; overflow: hidden;">' +
    '<div id="progressBar" style="' +
    'background: #1a73e8;' +
    'height: 100%;' +
    'width: 0%;' +
    'transition: width 0.3s ease;' +
    '"></div>' +
    '</div>' +
    
    '<div style="display: flex; justify-content: space-between; margin-top: 12px; font-size: 14px; color: #5f6368;">' +
    '<span id="progressText">0 / ' + this.total + '</span>' +
    '<span id="progressPercent">0%</span>' +
    '</div>' +
    
    '<div id="progressETA" style="margin-top: 8px; font-size: 12px; color: #5f6368;">' +
    'Calculating ETA...' +
    '</div>' +
    
    '<div id="currentItem" style="margin-top: 16px; font-size: 12px; color: #5f6368; font-style: italic;">' +
    'Starting...' +
    '</div>' +
    '</div>' +
    
    '<script>' +
    'function updateProgress(data) {' +
    '  document.getElementById("progressBar").style.width = data.percentage + "%";' +
    '  document.getElementById("progressText").textContent = data.current + " / " + data.total;' +
    '  document.getElementById("progressPercent").textContent = data.percentage + "%";' +
    '  if (data.eta) {' +
    '    document.getElementById("progressETA").textContent = "ETA: " + data.eta;' +
    '  }' +
    '  if (data.currentItem) {' +
    '    document.getElementById("currentItem").textContent = "Processing: " + data.currentItem;' +
    '  }' +
    '}' +
    '</script>'
  )
  .setWidth(400)
  .setHeight(250);
  
  this.dialog = SpreadsheetApp.getUi().showModalDialog(html, this.title);
};

/**
 * Update the progress
 * @param {number} increment - Number to increment by (default 1)
 * @param {string} currentItem - Current item being processed
 */
ProgressTracker.prototype.update = function(increment, currentItem) {
  if (increment === undefined) increment = 1;
  
  this.current += increment;
  
  // Only update UI based on updateInterval
  if (this.current % this.options.updateInterval !== 0 && this.current !== this.total) {
    return;
  }
  
  var percentage = Math.round((this.current / this.total) * 100);
  var elapsed = new Date().getTime() - this.startTime;
  var estimatedTotal = (elapsed / this.current) * this.total;
  var remaining = estimatedTotal - elapsed;
  
  var data = {
    current: this.current,
    total: this.total,
    percentage: percentage
  };
  
  if (this.options.showETA && this.current > 0) {
    data.eta = this.formatTime(remaining);
  }
  
  if (currentItem && this.options.showCurrent) {
    data.currentItem = currentItem;
  }
  
  // Update the dialog
  try {
    google.script.run.updateProgress(data);
  } catch (e) {
    // Fail silently - progress tracking shouldn't break the main operation
  }
};

/**
 * Format time in milliseconds to human readable format
 * @param {number} ms - Milliseconds
 * @return {string} Formatted time
 */
ProgressTracker.prototype.formatTime = function(ms) {
  if (ms < 1000) return 'less than a second';
  
  var seconds = Math.floor(ms / 1000);
  var minutes = Math.floor(seconds / 60);
  var hours = Math.floor(minutes / 60);
  
  if (hours > 0) {
    return hours + 'h ' + (minutes % 60) + 'm';
  } else if (minutes > 0) {
    return minutes + 'm ' + (seconds % 60) + 's';
  } else {
    return seconds + 's';
  }
};

/**
 * Close the progress dialog
 */
ProgressTracker.prototype.close = function() {
  try {
    google.script.host.close();
  } catch (e) {
    // Fail silently
  }
};

/**
 * Set the progress to a specific value
 * @param {number} value - Progress value
 * @param {string} currentItem - Current item being processed
 */
ProgressTracker.prototype.setProgress = function(value, currentItem) {
  this.current = value;
  this.update(0, currentItem);
};

/**
 * Mark as complete
 */
ProgressTracker.prototype.complete = function() {
  this.current = this.total;
  this.update(0, 'Complete!');
  
  // Auto-close after 1 second
  var self = this;
  Utilities.sleep(1000);
  self.close();
};

/**
 * Static factory method for compatibility
 * @param {string} title - Title for the progress dialog
 * @param {number} total - Total number of items
 * @param {Object} options - Options
 * @return {ProgressTracker} New progress tracker instance
 */
ProgressTracker.create = function(title, total, options) {
  return new ProgressTracker(title, total, options);
};