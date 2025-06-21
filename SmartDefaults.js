/**
 * ===================================================================
 * SMART DEFAULTS SERVICE
 * ===================================================================
 * Intelligent defaults and predictions based on usage patterns
 */

class SmartDefaults {
  static PREFERENCES_KEY = 'user_preferences';
  static HISTORY_KEY = 'user_history';
  
  static savePreference(key, value) {
    const preferences = this.getPreferences();
    preferences[key] = {
      value,
      updated: new Date().toISOString(),
      count: (preferences[key]?.count || 0) + 1
    };
    
    PropertiesService.getUserProperties()
      .setProperty(this.PREFERENCES_KEY, JSON.stringify(preferences));
  }
  
  static getPreferences() {
    const stored = PropertiesService.getUserProperties()
      .getProperty(this.PREFERENCES_KEY);
    return stored ? JSON.parse(stored) : {};
  }
  
  static getDefault(key, fallback = null) {
    const preferences = this.getPreferences();
    return preferences[key]?.value || fallback;
  }
  
  static recordAction(action, data) {
    const history = this.getHistory();
    
    history.push({
      action,
      data,
      timestamp: new Date().toISOString(),
      user: Session.getActiveUser().getEmail()
    });
    
    // Keep only last 100 actions
    if (history.length > 100) {
      history.shift();
    }
    
    PropertiesService.getScriptProperties()
      .setProperty(this.HISTORY_KEY, JSON.stringify(history));
  }
  
  static getHistory() {
    const stored = PropertiesService.getScriptProperties()
      .getProperty(this.HISTORY_KEY);
    return stored ? JSON.parse(stored) : [];
  }
  
  static predictVanAssignment(route) {
    const history = this.getHistory()
      .filter(h => h.action === 'van_assignment')
      .filter(h => h.data.routeType === route.serviceType);
    
    if (history.length === 0) {
      return null;
    }
    
    // Find most common van type for this route type
    const vanTypeCounts = {};
    history.forEach(h => {
      const vanType = h.data.vanType;
      vanTypeCounts[vanType] = (vanTypeCounts[vanType] || 0) + 1;
    });
    
    const mostCommon = Object.entries(vanTypeCounts)
      .sort((a, b) => b[1] - a[1])[0];
    
    return mostCommon ? mostCommon[0] : null;
  }
  
  static suggestNextAction() {
    const history = this.getHistory().slice(-10); // Last 10 actions
    const now = new Date();
    const hour = now.getHours();
    
    // Time-based suggestions
    if (hour >= 6 && hour < 9) {
      return {
        action: 'allocateVehicles',
        message: 'Good morning! Ready to allocate vehicles for today?',
        confidence: 0.9
      };
    } else if (hour >= 13 && hour < 14) {
      return {
        action: 'checkDeliveryPace',
        message: 'Time for the 1:40 PM delivery pace check',
        confidence: 0.8
      };
    } else if (hour >= 19) {
      return {
        action: 'generateRTSSummary',
        message: 'Generate end of day RTS summary?',
        confidence: 0.7
      };
    }
    
    // Pattern-based suggestions
    const lastAction = history[history.length - 1];
    if (lastAction?.action === 'uploadDayOfOps') {
      return {
        action: 'uploadDailyRoutes',
        message: 'Upload Daily Routes file next?',
        confidence: 0.95
      };
    }
    
    return null;
  }
  
  static getFrequentlyUsedVans(limit = 10) {
    const history = this.getHistory()
      .filter(h => h.action === 'van_assignment');
    
    const vanCounts = {};
    history.forEach(h => {
      const vanId = h.data.vanId;
      vanCounts[vanId] = (vanCounts[vanId] || 0) + 1;
    });
    
    return Object.entries(vanCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, limit)
      .map(([vanId, count]) => ({ vanId, count }));
  }
  
  static getPeakOperationTimes() {
    const history = this.getHistory();
    const hourCounts = {};
    
    history.forEach(h => {
      const hour = new Date(h.timestamp).getHours();
      hourCounts[hour] = (hourCounts[hour] || 0) + 1;
    });
    
    return Object.entries(hourCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([hour, count]) => ({
        hour: parseInt(hour),
        count,
        label: `${hour}:00`
      }));
  }
  
  static clearHistory() {
    PropertiesService.getScriptProperties()
      .deleteProperty(this.HISTORY_KEY);
    Logger.createLogger('SmartDefaults').info('History cleared');
  }
  
  static clearPreferences() {
    PropertiesService.getUserProperties()
      .deleteProperty(this.PREFERENCES_KEY);
    Logger.createLogger('SmartDefaults').info('Preferences cleared');
  }
}

/**
 * Apply smart defaults to forms and operations
 */
function applySmartDefaults() {
  const suggestion = SmartDefaults.suggestNextAction();
  
  if (suggestion && suggestion.confidence > 0.7) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Smart Suggestion',
      suggestion.message,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      switch (suggestion.action) {
        case 'allocateVehicles':
          showUploadDialog();
          break;
        case 'checkDeliveryPace':
          updateDeliveryPaceForToday();
          break;
        case 'generateRTSSummary':
          generateRTSSummary();
          break;
        case 'uploadDailyRoutes':
          // This would need custom handling
          showUploadDialog();
          break;
      }
    }
  }
}

/**
 * Example function to demonstrate smart defaults
 */
function testSmartDefaults() {
  // Record some sample actions
  SmartDefaults.recordAction('van_assignment', {
    routeType: 'Standard Parcel - Large Van',
    vanType: 'Large',
    vanId: 'BW5'
  });
  
  SmartDefaults.recordAction('van_assignment', {
    routeType: 'Standard Parcel - Large Van',
    vanType: 'Large',
    vanId: 'BW7'
  });
  
  // Save a preference
  SmartDefaults.savePreference('defaultWave', 'Wave 1');
  
  // Get suggestions
  const suggestion = SmartDefaults.suggestNextAction();
  const frequentVans = SmartDefaults.getFrequentlyUsedVans();
  const peakTimes = SmartDefaults.getPeakOperationTimes();
  
  console.log('Suggestion:', suggestion);
  console.log('Frequent Vans:', frequentVans);
  console.log('Peak Times:', peakTimes);
  
  UIHelpers.showSuccess('Smart Defaults test complete! Check logs for results.');
}