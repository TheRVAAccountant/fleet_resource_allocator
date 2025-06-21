/**
 * ===================================================================
 * ANALYTICS SERVICE
 * ===================================================================
 * Provides comprehensive analytics and insights for fleet operations
 */

/**
 * Calculate daily allocation success rate
 * @param {Date} date - Date to analyze (optional, defaults to today)
 * @return {Object} Allocation metrics
 */
function calculateDailyAllocationRate(date = new Date()) {
  const logger = Logger.createLogger('AnalyticsService');
  
  try {
    const dateStr = formatDate(date);
    const manager = new SheetManager();
    
    // Check for results sheet
    const resultsSheetName = dateStr.replace(/\//g, '-') + ' - Results';
    const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    const resultsSheet = ss.getSheetByName(resultsSheetName);
    
    const metrics = {
      date: dateStr,
      totalRoutes: 0,
      assignedRoutes: 0,
      unassignedRoutes: 0,
      allocationRate: 0,
      byServiceType: {},
      byWave: {},
      byStagingLocation: {}
    };
    
    if (!resultsSheet) {
      logger.info('No allocation data for date', { date: dateStr });
      return metrics;
    }
    
    const data = resultsSheet.getDataRange().getValues();
    const headers = data[0];
    const records = data.slice(1);
    
    metrics.totalRoutes = records.length;
    
    // Analyze each route
    records.forEach(row => {
      const serviceType = row[1];
      const wave = row[3];
      const stagingLocation = row[4];
      const vanId = row[5];
      
      // Count assigned routes
      if (vanId) {
        metrics.assignedRoutes++;
      } else {
        metrics.unassignedRoutes++;
      }
      
      // Group by service type
      if (!metrics.byServiceType[serviceType]) {
        metrics.byServiceType[serviceType] = {
          total: 0,
          assigned: 0,
          unassigned: 0
        };
      }
      metrics.byServiceType[serviceType].total++;
      if (vanId) {
        metrics.byServiceType[serviceType].assigned++;
      } else {
        metrics.byServiceType[serviceType].unassigned++;
      }
      
      // Group by wave
      if (!metrics.byWave[wave]) {
        metrics.byWave[wave] = {
          total: 0,
          assigned: 0
        };
      }
      metrics.byWave[wave].total++;
      if (vanId) {
        metrics.byWave[wave].assigned++;
      }
      
      // Group by staging location
      if (!metrics.byStagingLocation[stagingLocation]) {
        metrics.byStagingLocation[stagingLocation] = {
          total: 0,
          assigned: 0
        };
      }
      metrics.byStagingLocation[stagingLocation].total++;
      if (vanId) {
        metrics.byStagingLocation[stagingLocation].assigned++;
      }
    });
    
    // Calculate rates
    metrics.allocationRate = metrics.totalRoutes > 0
      ? ((metrics.assignedRoutes / metrics.totalRoutes) * 100).toFixed(1)
      : 0;
    
    // Calculate rates for groups
    Object.values(metrics.byServiceType).forEach(group => {
      group.rate = group.total > 0
        ? ((group.assigned / group.total) * 100).toFixed(1)
        : 0;
    });
    
    Object.values(metrics.byWave).forEach(group => {
      group.rate = group.total > 0
        ? ((group.assigned / group.total) * 100).toFixed(1)
        : 0;
    });
    
    Object.values(metrics.byStagingLocation).forEach(group => {
      group.rate = group.total > 0
        ? ((group.assigned / group.total) * 100).toFixed(1)
        : 0;
    });
    
    return metrics;
    
  } catch (error) {
    logger.error('Failed to calculate allocation rate', { error: error.message });
    return null;
  }
}

/**
 * Calculate average delivery completion time
 * Analyzes delivery pace data across all routes
 * @return {Object} Delivery time analytics
 */
function calculateAverageDeliveryCompletionTime() {
  const logger = Logger.createLogger('AnalyticsService');
  
  try {
    const manager = new SheetManager();
    const dailyDetails = manager.getSheet(getConfig('SHEETS.DAILY_DETAILS'));
    const data = dailyDetails.getData();
    
    const timeSlots = getConfig('DELIVERY_TIME_SLOTS');
    const analytics = {
      overallMetrics: {},
      byTimeSlot: {},
      byDayOfWeek: {
        'Sunday': {},
        'Monday': {},
        'Tuesday': {},
        'Wednesday': {},
        'Thursday': {},
        'Friday': {},
        'Saturday': {}
      }
    };
    
    // Initialize time slot analytics
    timeSlots.forEach(slot => {
      analytics.byTimeSlot[slot.label] = {
        totalStops: 0,
        routeCount: 0,
        averageStops: 0,
        completionRate: 0
      };
    });
    
    // Process each row
    let totalRoutes = 0;
    let routesWithPaceData = 0;
    
    data.slice(1).forEach(row => {
      const date = row[0];
      if (!(date instanceof Date)) return;
      
      const dayOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][date.getDay()];
      totalRoutes++;
      
      let hasPaceData = false;
      let lastCompletedSlot = null;
      let totalStopsForRoute = 0;
      
      // Check each time slot
      timeSlots.forEach((slot, index) => {
        const stops = parseInt(row[slot.column - 1]) || 0;
        
        if (stops > 0) {
          hasPaceData = true;
          lastCompletedSlot = slot;
          totalStopsForRoute = stops; // Latest number is cumulative
          
          analytics.byTimeSlot[slot.label].totalStops += stops;
          analytics.byTimeSlot[slot.label].routeCount++;
          
          // Day of week analytics
          if (!analytics.byDayOfWeek[dayOfWeek][slot.label]) {
            analytics.byDayOfWeek[dayOfWeek][slot.label] = {
              totalStops: 0,
              routeCount: 0
            };
          }
          analytics.byDayOfWeek[dayOfWeek][slot.label].totalStops += stops;
          analytics.byDayOfWeek[dayOfWeek][slot.label].routeCount++;
        }
      });
      
      if (hasPaceData) {
        routesWithPaceData++;
      }
    });
    
    // Calculate averages and rates
    timeSlots.forEach(slot => {
      const slotData = analytics.byTimeSlot[slot.label];
      if (slotData.routeCount > 0) {
        slotData.averageStops = (slotData.totalStops / slotData.routeCount).toFixed(1);
        slotData.completionRate = ((slotData.routeCount / totalRoutes) * 100).toFixed(1);
      }
    });
    
    // Calculate day of week averages
    Object.keys(analytics.byDayOfWeek).forEach(day => {
      const dayData = analytics.byDayOfWeek[day];
      Object.keys(dayData).forEach(slotLabel => {
        const slot = dayData[slotLabel];
        if (slot.routeCount > 0) {
          slot.averageStops = (slot.totalStops / slot.routeCount).toFixed(1);
        }
      });
    });
    
    // Overall metrics
    analytics.overallMetrics = {
      totalRoutesAnalyzed: totalRoutes,
      routesWithPaceData: routesWithPaceData,
      paceDataCoverage: totalRoutes > 0 
        ? ((routesWithPaceData / totalRoutes) * 100).toFixed(1) + '%'
        : '0%'
    };
    
    // Find peak completion time
    let peakTime = null;
    let peakCount = 0;
    
    Object.entries(analytics.byTimeSlot).forEach(([time, data]) => {
      if (data.routeCount > peakCount) {
        peakCount = data.routeCount;
        peakTime = time;
      }
    });
    
    analytics.overallMetrics.peakCompletionTime = peakTime;
    analytics.overallMetrics.peakCompletionCount = peakCount;
    
    return analytics;
    
  } catch (error) {
    logger.error('Failed to calculate delivery completion time', { error: error.message });
    return null;
  }
}

/**
 * Calculate van utilization by type
 * @param {Date} startDate - Start date for analysis
 * @param {Date} endDate - End date for analysis
 * @return {Object} Van utilization metrics
 */
function calculateVanUtilizationByType(startDate = null, endDate = null) {
  const logger = Logger.createLogger('AnalyticsService');
  
  try {
    const manager = new SheetManager();
    const vehicleStatus = manager.getSheet(getConfig('SHEETS.VEHICLE_STATUS'));
    const dailyDetails = manager.getSheet(getConfig('SHEETS.DAILY_DETAILS'));
    
    // Get vehicle data
    const vehicleData = vehicleStatus.getData();
    const vehicles = {};
    const utilization = {
      'Extra Large': { total: 0, operational: 0, utilized: new Set(), utilizationDays: {} },
      'Large': { total: 0, operational: 0, utilized: new Set(), utilizationDays: {} },
      'Step Van': { total: 0, operational: 0, utilized: new Set(), utilizationDays: {} }
    };
    
    // Build vehicle lookup
    vehicleData.slice(1).forEach(row => {
      const vanId = row[0];
      const type = row[1];
      const operational = row[2] === 'Y';
      
      vehicles[vanId] = { type, operational };
      
      if (utilization[type]) {
        utilization[type].total++;
        if (operational) {
          utilization[type].operational++;
        }
      }
    });
    
    // Analyze daily details
    const detailsData = dailyDetails.getData();
    
    detailsData.slice(1).forEach(row => {
      const date = row[0];
      const vanId = row[4];
      
      if (!vanId) return;
      
      // Apply date filter if provided
      if (startDate && date < startDate) return;
      if (endDate && date > endDate) return;
      
      const vehicle = vehicles[vanId];
      if (!vehicle) return;
      
      const dateStr = date instanceof Date ? formatDate(date) : date;
      
      if (utilization[vehicle.type]) {
        utilization[vehicle.type].utilized.add(vanId);
        
        // Track daily utilization
        if (!utilization[vehicle.type].utilizationDays[dateStr]) {
          utilization[vehicle.type].utilizationDays[dateStr] = new Set();
        }
        utilization[vehicle.type].utilizationDays[dateStr].add(vanId);
      }
    });
    
    // Calculate metrics
    const metrics = {
      dateRange: {
        start: startDate ? formatDate(startDate) : 'All time',
        end: endDate ? formatDate(endDate) : 'All time'
      },
      byType: {}
    };
    
    let totalFleet = 0;
    let totalOperational = 0;
    let totalUtilized = 0;
    
    Object.entries(utilization).forEach(([type, data]) => {
      const utilizedCount = data.utilized.size;
      
      metrics.byType[type] = {
        total: data.total,
        operational: data.operational,
        utilized: utilizedCount,
        operationalRate: data.total > 0 
          ? ((data.operational / data.total) * 100).toFixed(1) + '%' 
          : '0%',
        utilizationRate: data.operational > 0 
          ? ((utilizedCount / data.operational) * 100).toFixed(1) + '%' 
          : '0%',
        averageDailyUtilization: 0
      };
      
      // Calculate average daily utilization
      const utilizationDays = Object.keys(data.utilizationDays).length;
      if (utilizationDays > 0) {
        const totalDailyUtilization = Object.values(data.utilizationDays)
          .reduce((sum, daySet) => sum + daySet.size, 0);
        metrics.byType[type].averageDailyUtilization = 
          (totalDailyUtilization / utilizationDays).toFixed(1);
      }
      
      totalFleet += data.total;
      totalOperational += data.operational;
      totalUtilized += utilizedCount;
    });
    
    // Overall metrics
    metrics.overall = {
      totalFleet,
      totalOperational,
      totalUtilized,
      operationalRate: totalFleet > 0 
        ? ((totalOperational / totalFleet) * 100).toFixed(1) + '%' 
        : '0%',
      utilizationRate: totalOperational > 0 
        ? ((totalUtilized / totalOperational) * 100).toFixed(1) + '%' 
        : '0%'
    };
    
    return metrics;
    
  } catch (error) {
    logger.error('Failed to calculate van utilization', { error: error.message });
    return null;
  }
}

/**
 * Generate weekly trend analysis
 * @param {number} weeks - Number of weeks to analyze (default 4)
 * @return {Object} Weekly trend data
 */
function generateWeeklyTrendAnalysis(weeks = 4) {
  const logger = Logger.createLogger('AnalyticsService');
  
  try {
    const trends = {
      weeks: [],
      metrics: {
        routes: [],
        deliveries: [],
        returns: [],
        deliveryRate: [],
        vanUtilization: []
      }
    };
    
    const today = new Date();
    const manager = new SheetManager();
    const dailyDetails = manager.getSheet(getConfig('SHEETS.DAILY_DETAILS'));
    const data = dailyDetails.getData();
    
    // Analyze each week
    for (let w = weeks - 1; w >= 0; w--) {
      const weekStart = new Date(today);
      weekStart.setDate(today.getDate() - (w * 7) - today.getDay());
      weekStart.setHours(0, 0, 0, 0);
      
      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 6);
      weekEnd.setHours(23, 59, 59, 999);
      
      const weekNum = getWeekNumber(weekStart);
      trends.weeks.push(`Week ${weekNum}`);
      
      // Filter data for this week
      const weekData = data.slice(1).filter(row => {
        const date = row[0];
        return date instanceof Date && date >= weekStart && date <= weekEnd;
      });
      
      // Calculate metrics
      const weekMetrics = {
        routes: weekData.length,
        deliveries: 0,
        returns: 0,
        vansUsed: new Set()
      };
      
      weekData.forEach(row => {
        const delivered = parseInt(row[17]) || 0;
        const returned = parseInt(row[18]) || 0;
        const vanId = row[4];
        
        weekMetrics.deliveries += delivered;
        weekMetrics.returns += returned;
        if (vanId) weekMetrics.vansUsed.add(vanId);
      });
      
      const totalPackages = weekMetrics.deliveries + weekMetrics.returns;
      const deliveryRate = totalPackages > 0 
        ? ((weekMetrics.deliveries / totalPackages) * 100).toFixed(1)
        : 0;
      
      trends.metrics.routes.push(weekMetrics.routes);
      trends.metrics.deliveries.push(weekMetrics.deliveries);
      trends.metrics.returns.push(weekMetrics.returns);
      trends.metrics.deliveryRate.push(parseFloat(deliveryRate));
      trends.metrics.vanUtilization.push(weekMetrics.vansUsed.size);
    }
    
    // Calculate trend directions
    trends.trendDirection = {};
    Object.entries(trends.metrics).forEach(([metric, values]) => {
      if (values.length >= 2) {
        const recent = values[values.length - 1];
        const previous = values[values.length - 2];
        const change = recent - previous;
        const percentChange = previous !== 0 
          ? ((change / previous) * 100).toFixed(1)
          : 0;
        
        trends.trendDirection[metric] = {
          direction: change > 0 ? 'up' : change < 0 ? 'down' : 'stable',
          change: change,
          percentChange: percentChange + '%'
        };
      }
    });
    
    return trends;
    
  } catch (error) {
    logger.error('Failed to generate weekly trends', { error: error.message });
    return null;
  }
}

/**
 * Generate comprehensive analytics dashboard data
 * @return {Object} Complete analytics data
 */
function generateAnalyticsDashboard() {
  const logger = Logger.createLogger('AnalyticsService');
  const timer = logger.startTimer('generateAnalyticsDashboard');
  
  try {
    UIHelpers.showLoading('Generating analytics dashboard...');
    
    const dashboard = {
      generated: new Date().toISOString(),
      daily: {
        today: calculateDailyAllocationRate(new Date()),
        yesterday: calculateDailyAllocationRate((() => {
          const d = new Date();
          d.setDate(d.getDate() - 1);
          return d;
        })())
      },
      deliveryTimes: calculateAverageDeliveryCompletionTime(),
      vanUtilization: calculateVanUtilizationByType(),
      weeklyTrends: generateWeeklyTrendAnalysis(4),
      summary: {}
    };
    
    // Calculate summary metrics
    dashboard.summary = {
      currentAllocationRate: dashboard.daily.today.allocationRate + '%',
      fleetUtilization: dashboard.vanUtilization.overall.utilizationRate,
      peakDeliveryTime: dashboard.deliveryTimes.overallMetrics.peakCompletionTime,
      weeklyTrend: dashboard.weeklyTrends.trendDirection.routes?.direction || 'stable'
    };
    
    timer.end();
    google.script.host.close();
    
    // Create analytics sheet
    const manager = new SheetManager();
    const sheetName = `Analytics ${formatDate(new Date())}`;
    const analyticsSheet = manager.createSheet(sheetName, {
      headers: ['Fleet Analytics Dashboard'],
      overwrite: true
    });
    
    // Format and add data
    const reportData = formatAnalyticsReport(dashboard);
    analyticsSheet.sheet.getRange(1, 1, reportData.length, reportData[0].length)
      .setValues(reportData);
    
    // Format the sheet
    analyticsSheet.sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
    analyticsSheet.autoResizeColumns(1, 5);
    
    logger.info('Analytics dashboard generated', { sheet: sheetName });
    
    UIHelpers.showSuccess(`Analytics dashboard generated!\nSheet: ${sheetName}`);
    
    return dashboard;
    
  } catch (error) {
    logger.error('Failed to generate analytics dashboard', { error: error.message });
    google.script.host.close();
    UIHelpers.showError(error);
    return null;
  }
}

/**
 * Format analytics data for sheet output
 * @param {Object} dashboard - Dashboard data
 * @return {Array[]} Formatted 2D array
 */
function formatAnalyticsReport(dashboard) {
  const data = [
    ['Fleet Analytics Dashboard'],
    [`Generated: ${new Date().toLocaleString()}`],
    [],
    ['DAILY ALLOCATION METRICS'],
    ['Metric', 'Today', 'Yesterday', 'Change'],
    ['Total Routes', dashboard.daily.today.totalRoutes, dashboard.daily.yesterday.totalRoutes, 
      dashboard.daily.today.totalRoutes - dashboard.daily.yesterday.totalRoutes],
    ['Assigned Routes', dashboard.daily.today.assignedRoutes, dashboard.daily.yesterday.assignedRoutes,
      dashboard.daily.today.assignedRoutes - dashboard.daily.yesterday.assignedRoutes],
    ['Allocation Rate', dashboard.daily.today.allocationRate + '%', dashboard.daily.yesterday.allocationRate + '%',
      (parseFloat(dashboard.daily.today.allocationRate) - parseFloat(dashboard.daily.yesterday.allocationRate)).toFixed(1) + '%'],
    [],
    ['VAN UTILIZATION BY TYPE'],
    ['Type', 'Total', 'Operational', 'Utilized', 'Utilization Rate'],
  ];
  
  Object.entries(dashboard.vanUtilization.byType).forEach(([type, metrics]) => {
    data.push([type, metrics.total, metrics.operational, metrics.utilized, metrics.utilizationRate]);
  });
  
  data.push([]);
  data.push(['DELIVERY TIME ANALYSIS']);
  data.push(['Time Checkpoint', 'Routes Reporting', 'Average Stops', 'Completion Rate']);
  
  Object.entries(dashboard.deliveryTimes.byTimeSlot).forEach(([time, metrics]) => {
    data.push([time, metrics.routeCount, metrics.averageStops, metrics.completionRate + '%']);
  });
  
  data.push([]);
  data.push(['WEEKLY TRENDS']);
  data.push(['Week', 'Routes', 'Deliveries', 'Returns', 'Delivery Rate', 'Vans Used']);
  
  dashboard.weeklyTrends.weeks.forEach((week, index) => {
    data.push([
      week,
      dashboard.weeklyTrends.metrics.routes[index],
      dashboard.weeklyTrends.metrics.deliveries[index],
      dashboard.weeklyTrends.metrics.returns[index],
      dashboard.weeklyTrends.metrics.deliveryRate[index] + '%',
      dashboard.weeklyTrends.metrics.vanUtilization[index]
    ]);
  });
  
  return data;
}