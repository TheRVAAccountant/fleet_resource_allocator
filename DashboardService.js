/**
 * ===================================================================
 * DASHBOARD SERVICE
 * ===================================================================
 * Provides at-a-glance insights and visualizations
 */

/**
 * Create or update the dashboard
 */
function showDashboard() {
  const logger = createLogger('DashboardService');
  
  try {
    logger.info('Generating dashboard');
    
    // Get dashboard data
    const dashboardData = generateDashboardData();
    
    // Create HTML for dashboard
    const htmlTemplate = HtmlService.createTemplateFromFile('Dashboard');
    htmlTemplate.data = dashboardData;
    
    const html = htmlTemplate.evaluate()
      .setWidth(800)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Fleet Operations Dashboard');
    
  } catch (error) {
    logger.error('Failed to show dashboard', { error: error.message });
    errorHandler.showAlert(error, { component: 'Dashboard' });
  }
}

/**
 * Generate dashboard data
 * @return {Object} Dashboard data
 */
function generateDashboardData() {
  const logger = createLogger('DashboardService');
  const timer = logger.startTimer('generateDashboardData');
  
  try {
    const today = formatDate(new Date());
    const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    
    // Get vehicle status data
    const vehicleStats = getVehicleStatistics();
    
    // Get today's allocation stats
    const allocationStats = getTodayAllocationStatistics();
    
    // Get delivery pace stats
    const paceStats = getDeliveryPaceStatistics(today);
    
    // Get RTS stats
    const rtsStats = getRTSStatistics(today);
    
    // Get recent activity
    const recentActivity = getRecentActivity();
    
    const dashboardData = {
      date: today,
      lastUpdate: new Date().toLocaleTimeString(),
      vehicleStats: vehicleStats,
      allocationStats: allocationStats,
      paceStats: paceStats,
      rtsStats: rtsStats,
      recentActivity: recentActivity,
      alerts: generateAlerts(vehicleStats, allocationStats, paceStats, rtsStats)
    };
    
    timer.end();
    return dashboardData;
    
  } catch (error) {
    logger.error('Failed to generate dashboard data', { error: error.message });
    throw error;
  }
}

/**
 * Get vehicle statistics
 * @return {Object} Vehicle statistics
 */
function getVehicleStatistics() {
  const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  const vehicleSheet = ss.getSheetByName(getConfig('SHEETS.VEHICLE_STATUS'));
  
  if (!vehicleSheet) {
    throw new Error('Vehicle Status sheet not found');
  }
  
  const data = vehicleSheet.getDataRange().getValues();
  const stats = {
    total: 0,
    operational: 0,
    nonOperational: 0,
    byType: {
      'Extra Large': { total: 0, operational: 0 },
      'Large': { total: 0, operational: 0 },
      'Step Van': { total: 0, operational: 0 }
    }
  };
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const vanType = data[i][1];
    const operational = data[i][2] === 'Y';
    
    stats.total++;
    if (operational) {
      stats.operational++;
    } else {
      stats.nonOperational++;
    }
    
    if (stats.byType[vanType]) {
      stats.byType[vanType].total++;
      if (operational) {
        stats.byType[vanType].operational++;
      }
    }
  }
  
  stats.operationalRate = ((stats.operational / stats.total) * 100).toFixed(1);
  
  return stats;
}

/**
 * Get today's allocation statistics
 * @return {Object} Allocation statistics
 */
function getTodayAllocationStatistics() {
  const today = formatDate(new Date());
  const assignments = getTodayAssignments();
  
  const stats = {
    totalRoutes: 0,
    assignedRoutes: 0,
    unassignedRoutes: 0,
    vansUsed: Object.keys(assignments).length,
    allocationRate: 0
  };
  
  // Get today's results sheet to count routes
  const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  const resultsSheetName = today.replace(/\//g, '-') + ' - Results';
  const resultsSheet = ss.getSheetByName(resultsSheetName);
  
  if (resultsSheet) {
    const data = resultsSheet.getDataRange().getValues();
    stats.totalRoutes = data.length - 1; // Subtract header
    stats.assignedRoutes = stats.vansUsed;
    stats.unassignedRoutes = stats.totalRoutes - stats.assignedRoutes;
    stats.allocationRate = stats.totalRoutes > 0 ? 
      ((stats.assignedRoutes / stats.totalRoutes) * 100).toFixed(1) : 0;
  }
  
  return stats;
}

/**
 * Get delivery pace statistics
 * @param {string} date - Date to get stats for
 * @return {Object} Pace statistics
 */
function getDeliveryPaceStatistics(date) {
  const summary = getDeliveryPaceSummary(date);
  
  return {
    totalVansTracked: summary.totalVans || 0,
    lastCheckpoint: summary.lastCompletedCheckpoint || 'None',
    averageDeliveries: summary.averageDeliveries || {},
    onPaceCount: summary.onPaceCount || 0,
    behindPaceCount: summary.behindPaceCount || 0
  };
}

/**
 * Get RTS statistics
 * @param {string} date - Date to get stats for
 * @return {Object} RTS statistics
 */
function getRTSStatistics(date) {
  const summary = getRTSSummary(date);
  
  return {
    totalRoutes: summary.totalRoutes,
    completedReports: summary.completedReports,
    completionRate: summary.completionRate,
    totalDelivered: summary.totalDelivered,
    totalReturned: summary.totalReturned,
    successRate: summary.overallSuccessRate
  };
}

/**
 * Get recent activity
 * @return {Array} Recent activity items
 */
function getRecentActivity() {
  const activities = [];
  const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  
  // Check for recent sheets created
  const sheets = ss.getSheets();
  const today = new Date();
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    // Look for recently created results sheets
    if (sheetName.includes('Results') || sheetName.includes('RTS Summary')) {
      const dateMatch = sheetName.match(/(\d{2}-\d{2}-\d{2})/);
      if (dateMatch) {
        activities.push({
          type: sheetName.includes('Results') ? 'allocation' : 'rts',
          description: `${sheetName} created`,
          timestamp: dateMatch[1]
        });
      }
    }
  });
  
  // Sort by most recent first
  activities.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
  
  // Return only last 5 activities
  return activities.slice(0, 5);
}

/**
 * Generate alerts based on statistics
 * @param {Object} vehicleStats - Vehicle statistics
 * @param {Object} allocationStats - Allocation statistics
 * @param {Object} paceStats - Pace statistics
 * @param {Object} rtsStats - RTS statistics
 * @return {Array} Alert messages
 */
function generateAlerts(vehicleStats, allocationStats, paceStats, rtsStats) {
  const alerts = [];
  
  // Vehicle alerts
  if (vehicleStats.operationalRate < 80) {
    alerts.push({
      type: 'warning',
      message: `Only ${vehicleStats.operationalRate}% of vehicles are operational`
    });
  }
  
  // Allocation alerts
  if (allocationStats.unassignedRoutes > 0) {
    alerts.push({
      type: 'info',
      message: `${allocationStats.unassignedRoutes} routes are unassigned`
    });
  }
  
  // RTS alerts
  if (rtsStats.completionRate < 50 && rtsStats.totalRoutes > 0) {
    alerts.push({
      type: 'warning',
      message: `Only ${rtsStats.completionRate}% of routes have submitted RTS reports`
    });
  }
  
  // Pace alerts
  if (paceStats.behindPaceCount > paceStats.onPaceCount) {
    alerts.push({
      type: 'warning',
      message: `${paceStats.behindPaceCount} vans are behind delivery pace`
    });
  }
  
  return alerts;
}

/**
 * Get today's allocation statistics (simplified version)
 * @return {Object} Allocation statistics
 */
function getTodayAllocationStatistics() {
  try {
    const today = formatDate(new Date());
    const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    
    // Check for today's results sheet
    const resultsSheetName = today.replace(/\//g, '-') + ' - Results';
    const resultsSheet = ss.getSheetByName(resultsSheetName);
    
    const stats = {
      totalRoutes: 0,
      assignedRoutes: 0,
      unassignedRoutes: 0,
      vansUsed: 0,
      allocationRate: 0
    };
    
    if (resultsSheet) {
      const data = resultsSheet.getDataRange().getValues();
      if (data.length > 1) {
        stats.totalRoutes = data.length - 1; // Subtract header
        stats.assignedRoutes = data.length - 1; // All routes in results are assigned
        stats.vansUsed = stats.assignedRoutes; // One van per route
        stats.allocationRate = '100'; // If results exist, allocation is complete
      }
    }
    
    return stats;
  } catch (error) {
    console.error('Error getting allocation statistics:', error);
    return {
      totalRoutes: 0,
      assignedRoutes: 0,
      unassignedRoutes: 0,
      vansUsed: 0,
      allocationRate: 0
    };
  }
}

/**
 * Export dashboard data to sheet
 */
function exportDashboardToSheet() {
  const logger = createLogger('DashboardService');
  
  try {
    const dashboardData = generateDashboardData();
    const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    
    let dashboardSheet = ss.getSheetByName('Dashboard Export');
    if (dashboardSheet) {
      ss.deleteSheet(dashboardSheet);
    }
    
    dashboardSheet = ss.insertSheet('Dashboard Export');
    
    // Add data to sheet
    const rows = [
      ['Fleet Operations Dashboard Export'],
      ['Generated:', new Date()],
      [''],
      ['Vehicle Statistics'],
      ['Total Vehicles:', dashboardData.vehicleStats.total],
      ['Operational:', dashboardData.vehicleStats.operational],
      ['Non-Operational:', dashboardData.vehicleStats.nonOperational],
      ['Operational Rate:', dashboardData.vehicleStats.operationalRate + '%'],
      [''],
      ['Allocation Statistics'],
      ['Total Routes:', dashboardData.allocationStats.totalRoutes],
      ['Assigned Routes:', dashboardData.allocationStats.assignedRoutes],
      ['Vans Used:', dashboardData.allocationStats.vansUsed],
      ['Allocation Rate:', dashboardData.allocationStats.allocationRate + '%'],
      [''],
      ['RTS Statistics'],
      ['Completed Reports:', dashboardData.rtsStats.completedReports],
      ['Completion Rate:', dashboardData.rtsStats.completionRate + '%'],
      ['Total Delivered:', dashboardData.rtsStats.totalDelivered],
      ['Success Rate:', dashboardData.rtsStats.successRate + '%']
    ];
    
    dashboardSheet.getRange(1, 1, rows.length, 2).setValues(rows);
    
    // Format the sheet
    dashboardSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
    dashboardSheet.getRange(4, 1).setFontWeight('bold');
    dashboardSheet.getRange(10, 1).setFontWeight('bold');
    dashboardSheet.getRange(16, 1).setFontWeight('bold');
    
    dashboardSheet.autoResizeColumns(1, 2);
    
    logger.info('Dashboard exported to sheet');
    showInfoAlert('Dashboard exported to "Dashboard Export" sheet');
    
  } catch (error) {
    logger.error('Failed to export dashboard', { error: error.message });
    errorHandler.showAlert(error, { component: 'Dashboard Export' });
  }
}