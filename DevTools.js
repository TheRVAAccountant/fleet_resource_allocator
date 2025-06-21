/**
 * ===================================================================
 * DEVELOPMENT TOOLS
 * ===================================================================
 * Tools for development, debugging, and monitoring
 */

class DevTools {
  static startTime = new Date().getTime();
  
  static isDevMode() {
    return Session.getActiveUser().getEmail().includes('@test.com') ||
           PropertiesService.getScriptProperties().getProperty('DEV_MODE') === 'true';
  }
  
  static enableDevMode() {
    PropertiesService.getScriptProperties().setProperty('DEV_MODE', 'true');
    this.showDevConsole('Development mode enabled');
  }
  
  static disableDevMode() {
    PropertiesService.getScriptProperties().deleteProperty('DEV_MODE');
    UIHelpers.toast('Development mode disabled');
  }
  
  static showDevConsole(initialMessage = '') {
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: 'Consolas', 'Monaco', monospace; font-size: 12px;">
        <div style="background: #1e1e1e; color: #d4d4d4; padding: 10px; border-bottom: 1px solid #333;">
          <strong>Developer Console</strong>
          <button onclick="clearConsole()" style="float: right; background: #333; color: white; border: none; padding: 2px 10px; cursor: pointer;">Clear</button>
        </div>
        <div id="console" style="background: #252526; color: #d4d4d4; padding: 10px; height: 400px; overflow-y: auto;">
          <div>${initialMessage}</div>
        </div>
        <div style="background: #1e1e1e; padding: 10px; border-top: 1px solid #333;">
          <input type="text" id="command" style="width: 100%; background: #3c3c3c; color: white; border: 1px solid #555; padding: 5px;" placeholder="Enter command..." onkeypress="handleCommand(event)">
        </div>
      </div>
      
      <script>
        function log(message, type = 'info') {
          const console = document.getElementById('console');
          const timestamp = new Date().toISOString().substr(11, 8);
          const color = {
            info: '#d4d4d4',
            success: '#4ec9b0',
            warning: '#dcdcaa',
            error: '#f48771'
          }[type] || '#d4d4d4';
          
          console.innerHTML += '<div style="color: ' + color + '">[' + timestamp + '] ' + message + '</div>';
          console.scrollTop = console.scrollHeight;
        }
        
        function clearConsole() {
          document.getElementById('console').innerHTML = '';
        }
        
        function handleCommand(event) {
          if (event.key === 'Enter') {
            const command = event.target.value;
            log('> ' + command, 'info');
            
            // Execute command
            google.script.run
              .withSuccessHandler(result => log(result, 'success'))
              .withFailureHandler(error => log(error.message, 'error'))
              .executeDevCommand(command);
            
            event.target.value = '';
          }
        }
        
        // Auto-refresh logs
        setInterval(() => {
          google.script.run
            .withSuccessHandler(logs => {
              if (logs && logs.length > 0) {
                logs.forEach(logEntry => log(logEntry.message, logEntry.type));
              }
            })
            .getDevLogs();
        }, 2000);
      </script>
    `).setWidth(600).setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Developer Console');
  }
  
  static executeDevCommand(command) {
    const parts = command.split(' ');
    const cmd = parts[0];
    const args = parts.slice(1);
    
    switch (cmd) {
      case 'clear':
        Cache.clear();
        return 'All caches cleared';
        
      case 'cache':
        return JSON.stringify(Cache.getStats(), null, 2);
        
      case 'props':
        const props = PropertiesService.getScriptProperties().getProperties();
        return JSON.stringify(props, null, 2);
        
      case 'user':
        return JSON.stringify({
          email: Session.getActiveUser().getEmail(),
          timezone: Session.getScriptTimeZone(),
          locale: Session.getActiveUserLocale()
        }, null, 2);
        
      case 'test':
        return this.runTest(args[0]);
        
      case 'perf':
        return this.getPerformanceStats();
        
      case 'sheets':
        return this.listSheets();
        
      case 'config':
        return JSON.stringify(getConfig(), null, 2);
        
      case 'history':
        return JSON.stringify(SmartDefaults.getHistory().slice(-10), null, 2);
        
      case 'help':
        return `Available commands:
  clear - Clear all caches
  cache - Show cache statistics
  props - Show script properties
  user - Show current user info
  test [name] - Run specific test
  perf - Show performance stats
  sheets - List all sheets
  config - Show configuration
  history - Show recent action history
  help - Show this help`;
        
      default:
        return `Unknown command: ${cmd}. Type 'help' for available commands.`;
    }
  }
  
  static getPerformanceStats() {
    const stats = {
      executionTime: new Date().getTime() - this.startTime,
      quotaUsed: {
        urlFetch: UrlFetchApp.getQuota ? UrlFetchApp.getQuota() : 'N/A',
        triggers: ScriptApp.getProjectTriggers().length
      },
      cache: Cache.getStats(),
      memory: {
        properties: {
          script: Object.keys(PropertiesService.getScriptProperties().getProperties()).length,
          user: Object.keys(PropertiesService.getUserProperties().getProperties()).length
        }
      }
    };
    
    return JSON.stringify(stats, null, 2);
  }
  
  static listSheets() {
    const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    const sheets = ss.getSheets().map(sheet => ({
      name: sheet.getName(),
      rows: sheet.getMaxRows(),
      columns: sheet.getMaxColumns(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn()
    }));
    
    return JSON.stringify(sheets, null, 2);
  }
  
  static runTest(testName) {
    try {
      switch (testName) {
        case 'logger':
          return this.testLogger();
        case 'cache':
          return this.testCache();
        case 'sheet':
          return this.testSheetManager();
        case 'ui':
          return this.testUIHelpers();
        case 'smart':
          return this.testSmartDefaults();
        default:
          return 'Unknown test. Available: logger, cache, sheet, ui, smart';
      }
    } catch (error) {
      return `Test failed: ${error.message}`;
    }
  }
  
  static testLogger() {
    const logger = Logger.createLogger('TestComponent');
    const timer = logger.startTimer('testOperation');
    
    logger.debug('Debug message');
    logger.info('Info message');
    logger.warn('Warning message');
    logger.error('Error message', { test: true });
    
    Utilities.sleep(100);
    timer.end();
    
    return 'Logger test complete. Check console and Error Log sheet.';
  }
  
  static testCache() {
    const key = 'test_key';
    const value = { test: true, timestamp: new Date().toISOString() };
    
    // Test cache miss
    const miss = Cache.get(key, () => value);
    
    // Test cache hit
    const hit = Cache.get(key, () => ({ should: 'not be called' }));
    
    // Clear and verify
    Cache.clear();
    const afterClear = Cache.get(key, () => ({ cleared: true }));
    
    return JSON.stringify({
      miss,
      hit,
      afterClear,
      stats: Cache.getStats()
    }, null, 2);
  }
  
  static testSheetManager() {
    const manager = new SheetManager();
    const vehicleSheet = manager.getSheet(getConfig('SHEETS.VEHICLE_STATUS'));
    
    // Test data retrieval with caching
    const timer1 = new Date().getTime();
    const data1 = vehicleSheet.getData();
    const time1 = new Date().getTime() - timer1;
    
    const timer2 = new Date().getTime();
    const data2 = vehicleSheet.getData(); // Should be cached
    const time2 = new Date().getTime() - timer2;
    
    return JSON.stringify({
      sheetName: vehicleSheet.getName(),
      rows: data1.length,
      firstLoadTime: time1 + 'ms',
      cachedLoadTime: time2 + 'ms',
      cacheImprovement: ((time1 - time2) / time1 * 100).toFixed(2) + '%'
    }, null, 2);
  }
  
  static testUIHelpers() {
    UIHelpers.toast('Testing toast notification');
    return 'UI test complete. You should see a toast notification.';
  }
  
  static testSmartDefaults() {
    // Record test action
    SmartDefaults.recordAction('test_action', {
      timestamp: new Date().toISOString(),
      test: true
    });
    
    const suggestion = SmartDefaults.suggestNextAction();
    const history = SmartDefaults.getHistory().slice(-5);
    
    return JSON.stringify({
      suggestion,
      recentHistory: history
    }, null, 2);
  }
  
  static profile(fn, label) {
    return function(...args) {
      const startTime = new Date().getTime();
      const startQuota = UrlFetchApp.getQuota ? UrlFetchApp.getQuota() : null;
      
      try {
        const result = fn.apply(this, args);
        
        const endTime = new Date().getTime();
        const endQuota = UrlFetchApp.getQuota ? UrlFetchApp.getQuota() : null;
        
        const profile = {
          label,
          duration: endTime - startTime,
          apiCalls: startQuota && endQuota ? startQuota - endQuota : 'N/A',
          timestamp: new Date().toISOString()
        };
        
        Logger.createLogger('Performance').info('Profile', profile);
        
        return result;
      } catch (error) {
        Logger.createLogger('Performance').error('Profile error', { label, error: error.message });
        throw error;
      }
    };
  }
  
  static getDevLogs() {
    // This would return any pending log entries for the console
    // In a real implementation, you might store these in cache
    return [];
  }
}

/**
 * Add developer menu to UI
 */
function addDeveloperMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🔧 Developer')
    .addItem('Open Console', 'openDevConsole')
    .addItem('Enable Dev Mode', 'enableDevMode')
    .addItem('Disable Dev Mode', 'disableDevMode')
    .addSeparator()
    .addItem('View Performance', 'viewPerformance')
    .addItem('View Cache Stats', 'viewCacheStats')
    .addItem('Clear All Caches', 'clearAllCaches')
    .addSeparator()
    .addItem('Run Tests', 'runAllTests')
    .addToUi();
}

// Helper functions for menu
function openDevConsole() { DevTools.showDevConsole(); }
function enableDevMode() { DevTools.enableDevMode(); }
function disableDevMode() { DevTools.disableDevMode(); }
function viewPerformance() { 
  UIHelpers.showSuccess(DevTools.getPerformanceStats());
}
function viewCacheStats() {
  UIHelpers.showSuccess(JSON.stringify(Cache.getStats(), null, 2));
}
function clearAllCaches() {
  Cache.clear();
  UIHelpers.toast('All caches cleared');
}
function runAllTests() {
  const tests = ['logger', 'cache', 'sheet', 'ui', 'smart'];
  const results = {};
  
  tests.forEach(test => {
    try {
      results[test] = DevTools.runTest(test);
    } catch (e) {
      results[test] = `Failed: ${e.message}`;
    }
  });
  
  const html = HtmlService.createHtmlOutput(`<pre>${JSON.stringify(results, null, 2)}</pre>`)
    .setWidth(600)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Test Results');
}