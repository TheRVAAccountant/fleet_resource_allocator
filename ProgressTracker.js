/**
 * ===================================================================
 * PROGRESS TRACKER
 * ===================================================================
 * Track and display progress for long-running operations
 */

class ProgressTracker {
  constructor(title, total, options = {}) {
    this.title = title;
    this.total = total;
    this.current = 0;
    this.startTime = new Date().getTime();
    this.options = {
      showETA: true,
      showPercentage: true,
      showCurrent: true,
      updateInterval: 1, // Update UI every N items
      ...options
    };
    
    this.dialog = null;
    this.createDialog();
  }
  
  createDialog() {
    const html = HtmlService.createHtmlOutput(`
      <div id="progressContainer" style="padding: 20px; font-family: 'Google Sans', Arial, sans-serif;">
        <h3 style="margin: 0 0 20px 0; color: #202124;">${this.title}</h3>
        
        <div style="background: #e0e0e0; height: 8px; border-radius: 4px; overflow: hidden;">
          <div id="progressBar" style="
            background: #1a73e8;
            height: 100%;
            width: 0%;
            transition: width 0.3s ease;
          "></div>
        </div>
        
        <div style="display: flex; justify-content: space-between; margin-top: 12px; font-size: 14px; color: #5f6368;">
          <span id="progressText">0 / ${this.total}</span>
          <span id="progressPercent">0%</span>
        </div>
        
        <div id="progressETA" style="margin-top: 8px; font-size: 12px; color: #5f6368;">
          Calculating ETA...
        </div>
        
        <div id="currentItem" style="margin-top: 16px; font-size: 12px; color: #5f6368; font-style: italic;">
          Starting...
        </div>
      </div>
      
      <script>
        function updateProgress(data) {
          document.getElementById('progressBar').style.width = data.percentage + '%';
          document.getElementById('progressText').textContent = data.current + ' / ' + data.total;
          document.getElementById('progressPercent').textContent = data.percentage + '%';
          
          if (data.eta) {
            document.getElementById('progressETA').textContent = 'ETA: ' + data.eta;
          }
          
          if (data.currentItem) {
            document.getElementById('currentItem').textContent = data.currentItem;
          }
        }
        
        // Listen for updates from server
        window.addEventListener('message', function(event) {
          if (event.data && event.data.type === 'progress') {
            updateProgress(event.data.data);
          }
        });
      </script>
    `).setWidth(400).setHeight(250);
    
    this.dialog = html;
    SpreadsheetApp.getUi().showModalDialog(html, ' ');
  }
  
  update(current, currentItem = null) {
    this.current = current;
    
    if (current % this.options.updateInterval === 0 || current === this.total) {
      const percentage = Math.round((current / this.total) * 100);
      const elapsed = new Date().getTime() - this.startTime;
      const eta = this.calculateETA(elapsed, current, this.total);
      
      const updateData = {
        current,
        total: this.total,
        percentage,
        eta: eta ? this.formatTime(eta) : null,
        currentItem
      };
      
      // Store progress in cache for client to retrieve
      CacheService.getScriptCache().put('progress_' + this.title, JSON.stringify(updateData), 60);
      
      // For immediate UI updates in same-script context
      if (this.dialog) {
        try {
          // This would need to be implemented differently in production
          // as we can't directly update an open dialog
          Logger.createLogger('ProgressTracker').debug('Progress update', updateData);
        } catch (e) {
          // Silent fail for UI updates
        }
      }
    }
  }
  
  calculateETA(elapsed, current, total) {
    if (current === 0) return null;
    
    const rate = current / elapsed;
    const remaining = total - current;
    return remaining / rate;
  }
  
  formatTime(milliseconds) {
    const seconds = Math.floor(milliseconds / 1000);
    const minutes = Math.floor(seconds / 60);
    const hours = Math.floor(minutes / 60);
    
    if (hours > 0) {
      return `${hours}h ${minutes % 60}m`;
    } else if (minutes > 0) {
      return `${minutes}m ${seconds % 60}s`;
    } else {
      return `${seconds}s`;
    }
  }
  
  complete(message = 'Complete!') {
    this.update(this.total, message);
    
    // Close dialog and show success
    Utilities.sleep(1000); // Give time for final update
    google.script.host.close();
    UIHelpers.showSuccess(message);
  }
  
  error(error) {
    google.script.host.close();
    UIHelpers.showError(error);
  }
}

/**
 * Helper function to track progress in server-side operations
 */
function trackProgress(title, items, processFn, options = {}) {
  const tracker = new ProgressTracker(title, items.length, options);
  const results = [];
  
  try {
    items.forEach((item, index) => {
      const result = processFn(item, index);
      results.push(result);
      tracker.update(index + 1, `Processing: ${item.toString().substring(0, 50)}...`);
    });
    
    tracker.complete(`Successfully processed ${items.length} items`);
    return results;
  } catch (error) {
    tracker.error(error);
    throw error;
  }
}

/**
 * Example usage function
 */
function exampleProgressTracking() {
  const items = Array.from({length: 100}, (_, i) => `Item ${i + 1}`);
  
  trackProgress('Processing Items', items, (item, index) => {
    // Simulate work
    Utilities.sleep(100);
    return `Processed: ${item}`;
  });
}