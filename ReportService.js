/**
 * ===================================================================
 * REPORT SERVICE
 * ===================================================================
 * Comprehensive reporting functionality for fleet operations including
 * driver performance, vehicle utilization, and operational analytics
 */

/**
 * Generate driver performance report
 * Aggregates driver data from Daily Details and calculates key metrics
 */
function generateDriverPerformanceReport() {
  const logger = Logger.createLogger('ReportService');
  const timer = logger.startTimer('generateDriverPerformanceReport');
  
  try {
    UIHelpers.showLoading('Generating driver performance report...');
    
    const manager = new SheetManager();
    const dailyDetails = manager.getSheet(getConfig('SHEETS.DAILY_DETAILS'));
    const data = dailyDetails.getData();
    
    // Skip header row
    const records = data.slice(1);
    
    // Aggregate by driver
    const driverMetrics = {};
    
    records.forEach(row => {
      const date = row[0];
      const routeCode = row[1];
      const driverName = row[2];
      const vanId = row[4];
      const rtsTime = row[16]; // Column Q
      const packagesDelivered = row[17] || 0; // Column R
      const packagesReturned = row[18] || 0; // Column S
      
      if (!driverName || driverName === 'N/A') return;
      
      if (!driverMetrics[driverName]) {
        driverMetrics[driverName] = {
          name: driverName,
          routesCompleted: 0,
          totalPackagesDelivered: 0,
          totalPackagesReturned: 0,
          deliveryRate: 0,
          rtsTimes: [],
          daysWorked: new Set(),
          vansUsed: new Set()
        };
      }
      
      const driver = driverMetrics[driverName];
      driver.routesCompleted++;
      driver.totalPackagesDelivered += parseInt(packagesDelivered) || 0;
      driver.totalPackagesReturned += parseInt(packagesReturned) || 0;
      
      if (date instanceof Date) {
        driver.daysWorked.add(formatDate(date));
      }
      
      if (vanId) {
        driver.vansUsed.add(vanId);
      }
      
      if (rtsTime) {
        driver.rtsTimes.push(rtsTime);
      }
    });
    
    // Calculate additional metrics
    Object.values(driverMetrics).forEach(driver => {
      const totalPackages = driver.totalPackagesDelivered + driver.totalPackagesReturned;
      driver.deliveryRate = totalPackages > 0 
        ? ((driver.totalPackagesDelivered / totalPackages) * 100).toFixed(1)
        : 0;
      
      // Calculate average RTS time
      if (driver.rtsTimes.length > 0) {
        const validTimes = driver.rtsTimes
          .filter(time => time && time !== '-')
          .map(time => {
            // Parse time and convert to minutes from start of day
            const timeParts = time.toString().match(/(\d+):(\d+)\s*(AM|PM)/i);
            if (timeParts) {
              let hours = parseInt(timeParts[1]);
              const minutes = parseInt(timeParts[2]);
              const period = timeParts[3];
              
              if (period.toUpperCase() === 'PM' && hours !== 12) hours += 12;
              if (period.toUpperCase() === 'AM' && hours === 12) hours = 0;
              
              return hours * 60 + minutes;
            }
            return null;
          })
          .filter(time => time !== null);
        
        if (validTimes.length > 0) {
          const avgMinutes = validTimes.reduce((a, b) => a + b, 0) / validTimes.length;
          const avgHours = Math.floor(avgMinutes / 60);
          const avgMins = Math.floor(avgMinutes % 60);
          driver.averageRtsTime = `${avgHours}:${avgMins.toString().padStart(2, '0')}`;
        } else {
          driver.averageRtsTime = 'N/A';
        }
      } else {
        driver.averageRtsTime = 'N/A';
      }
      
      driver.daysWorkedCount = driver.daysWorked.size;
      driver.vansUsedCount = driver.vansUsed.size;
    });
    
    // Sort by routes completed
    const sortedDrivers = Object.values(driverMetrics)
      .sort((a, b) => b.routesCompleted - a.routesCompleted);
    
    // Create report sheet
    const reportDate = formatDate(new Date());
    const reportSheetName = `Driver Performance ${reportDate}`;
    
    const reportSheet = manager.createSheet(reportSheetName, {
      headers: [
        'Driver Name',
        'Routes Completed',
        'Days Worked',
        'Packages Delivered',
        'Packages Returned',
        'Delivery Rate %',
        'Average RTS Time',
        'Different Vans Used'
      ],
      overwrite: true
    });
    
    // Add data rows
    const rows = sortedDrivers.map(driver => [
      driver.name,
      driver.routesCompleted,
      driver.daysWorkedCount,
      driver.totalPackagesDelivered,
      driver.totalPackagesReturned,
      driver.deliveryRate + '%',
      driver.averageRtsTime,
      driver.vansUsedCount
    ]);
    
    reportSheet.appendRows(rows);
    
    // Add summary statistics
    const totalRoutes = sortedDrivers.reduce((sum, d) => sum + d.routesCompleted, 0);
    const totalDelivered = sortedDrivers.reduce((sum, d) => sum + d.totalPackagesDelivered, 0);
    const totalReturned = sortedDrivers.reduce((sum, d) => sum + d.totalPackagesReturned, 0);
    const overallDeliveryRate = totalDelivered + totalReturned > 0
      ? ((totalDelivered / (totalDelivered + totalReturned)) * 100).toFixed(1)
      : 0;
    
    reportSheet.appendRows([
      [], // Empty row
      ['SUMMARY', '', '', '', '', '', '', ''],
      ['Total Drivers:', sortedDrivers.length, '', '', '', '', '', ''],
      ['Total Routes:', totalRoutes, '', '', '', '', '', ''],
      ['Total Delivered:', totalDelivered, '', '', '', '', '', ''],
      ['Total Returned:', totalReturned, '', '', '', '', '', ''],
      ['Overall Delivery Rate:', overallDeliveryRate + '%', '', '', '', '', '', '']
    ]);
    
    // Format the sheet
    reportSheet.autoResizeColumns(1, 8);
    
    timer.end();
    google.script.host.close();
    
    logger.info('Driver performance report generated', {
      sheet: reportSheetName,
      drivers: sortedDrivers.length,
      totalRoutes: totalRoutes
    });
    
    UIHelpers.showSuccess(`Driver performance report generated!\nSheet: ${reportSheetName}`);
    
    // Record action for smart defaults
    SmartDefaults.recordAction('report_generated', {
      type: 'driver_performance',
      driversCount: sortedDrivers.length,
      totalRoutes: totalRoutes
    });
    
  } catch (error) {
    logger.error('Failed to generate driver performance report', { error: error.message });
    google.script.host.close();
    UIHelpers.showError(error);
  }
}

/**
 * Generate weekly summary report
 * Provides comprehensive weekly metrics and trends
 */
function generateWeeklySummaryReport() {
  const logger = Logger.createLogger('ReportService');
  const timer = logger.startTimer('generateWeeklySummaryReport');
  
  try {
    UIHelpers.showLoading('Generating weekly summary report...');
    
    const manager = new SheetManager();
    const dailyDetails = manager.getSheet(getConfig('SHEETS.DAILY_DETAILS'));
    const vehicleStatus = manager.getSheet(getConfig('SHEETS.VEHICLE_STATUS'));
    
    // Get current week date range
    const today = new Date();
    const currentWeek = getWeekNumber(today);
    const currentYear = today.getFullYear();
    
    // Get start of week (Sunday)
    const startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay());
    startOfWeek.setHours(0, 0, 0, 0);
    
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(startOfWeek.getDate() + 6);
    endOfWeek.setHours(23, 59, 59, 999);
    
    logger.info('Generating report for week', {
      weekNumber: currentWeek,
      startDate: formatDate(startOfWeek),
      endDate: formatDate(endOfWeek)
    });
    
    // Get data
    const dailyData = dailyDetails.getData();
    const vehicleData = vehicleStatus.getData();
    
    // Filter for current week
    const weekRecords = dailyData.slice(1).filter(row => {
      const date = row[0];
      if (date instanceof Date) {
        return date >= startOfWeek && date <= endOfWeek;
      }
      return false;
    });
    
    // Daily metrics
    const dailyMetrics = {};
    const vanUtilization = {};
    const routeTypes = {};
    
    weekRecords.forEach(row => {
      const date = formatDate(row[0]);
      const routeCode = row[1];
      const driverName = row[2];
      const vanId = row[4];
      const packagesDelivered = parseInt(row[17]) || 0;
      const packagesReturned = parseInt(row[18]) || 0;
      
      // Daily aggregation
      if (!dailyMetrics[date]) {
        dailyMetrics[date] = {
          date: date,
          routes: 0,
          drivers: new Set(),
          vans: new Set(),
          packagesDelivered: 0,
          packagesReturned: 0
        };
      }
      
      dailyMetrics[date].routes++;
      if (driverName && driverName !== 'N/A') {
        dailyMetrics[date].drivers.add(driverName);
      }
      if (vanId) {
        dailyMetrics[date].vans.add(vanId);
        
        // Track van utilization
        if (!vanUtilization[vanId]) {
          vanUtilization[vanId] = 0;
        }
        vanUtilization[vanId]++;
      }
      dailyMetrics[date].packagesDelivered += packagesDelivered;
      dailyMetrics[date].packagesReturned += packagesReturned;
    });
    
    // Get vehicle types from vehicle status
    const vehicleTypes = {};
    vehicleData.slice(1).forEach(row => {
      const vanId = row[0];
      const type = row[1];
      const operational = row[2];
      vehicleTypes[vanId] = { type, operational };
    });
    
    // Calculate weekly totals
    const weeklyTotals = {
      totalRoutes: 0,
      totalDrivers: new Set(),
      totalVans: new Set(),
      totalPackagesDelivered: 0,
      totalPackagesReturned: 0,
      daysOperational: Object.keys(dailyMetrics).length
    };
    
    Object.values(dailyMetrics).forEach(day => {
      weeklyTotals.totalRoutes += day.routes;
      day.drivers.forEach(d => weeklyTotals.totalDrivers.add(d));
      day.vans.forEach(v => weeklyTotals.totalVans.add(v));
      weeklyTotals.totalPackagesDelivered += day.packagesDelivered;
      weeklyTotals.totalPackagesReturned += day.packagesReturned;
      
      // Convert sets to counts for report
      day.driverCount = day.drivers.size;
      day.vanCount = day.vans.size;
    });
    
    // Calculate utilization by vehicle type
    const utilizationByType = {
      'Extra Large': { used: 0, total: 0 },
      'Large': { used: 0, total: 0 },
      'Step Van': { used: 0, total: 0 }
    };
    
    Object.entries(vehicleTypes).forEach(([vanId, info]) => {
      if (info.type && utilizationByType[info.type]) {
        utilizationByType[info.type].total++;
        if (vanUtilization[vanId] > 0) {
          utilizationByType[info.type].used++;
        }
      }
    });
    
    // Create report sheet
    const reportSheetName = `Weekly Summary Week ${currentWeek} ${currentYear}`;
    const reportSheet = manager.createSheet(reportSheetName, {
      headers: ['Weekly Summary Report'],
      overwrite: true
    });
    
    // Add report sections
    const reportData = [
      [`Week ${currentWeek}, ${currentYear}`],
      [`${formatDate(startOfWeek)} - ${formatDate(endOfWeek)}`],
      [],
      ['WEEKLY TOTALS'],
      ['Total Routes Completed:', weeklyTotals.totalRoutes],
      ['Unique Drivers:', weeklyTotals.totalDrivers.size],
      ['Unique Vans Used:', weeklyTotals.totalVans.size],
      ['Days Operational:', weeklyTotals.daysOperational],
      ['Total Packages Delivered:', weeklyTotals.totalPackagesDelivered],
      ['Total Packages Returned:', weeklyTotals.totalPackagesReturned],
      ['Overall Delivery Rate:', 
        weeklyTotals.totalPackagesDelivered + weeklyTotals.totalPackagesReturned > 0
          ? ((weeklyTotals.totalPackagesDelivered / (weeklyTotals.totalPackagesDelivered + weeklyTotals.totalPackagesReturned)) * 100).toFixed(1) + '%'
          : 'N/A'],
      [],
      ['DAILY BREAKDOWN'],
      ['Date', 'Routes', 'Drivers', 'Vans', 'Delivered', 'Returned', 'Delivery Rate']
    ];
    
    // Add daily data
    const sortedDays = Object.values(dailyMetrics).sort((a, b) => 
      new Date(a.date) - new Date(b.date)
    );
    
    sortedDays.forEach(day => {
      const totalPackages = day.packagesDelivered + day.packagesReturned;
      const deliveryRate = totalPackages > 0
        ? ((day.packagesDelivered / totalPackages) * 100).toFixed(1) + '%'
        : 'N/A';
      
      reportData.push([
        day.date,
        day.routes,
        day.driverCount,
        day.vanCount,
        day.packagesDelivered,
        day.packagesReturned,
        deliveryRate
      ]);
    });
    
    reportData.push([]);
    reportData.push(['VEHICLE UTILIZATION BY TYPE']);
    reportData.push(['Type', 'Used', 'Total', 'Utilization %']);
    
    Object.entries(utilizationByType).forEach(([type, stats]) => {
      const utilization = stats.total > 0
        ? ((stats.used / stats.total) * 100).toFixed(1) + '%'
        : 'N/A';
      reportData.push([type, stats.used, stats.total, utilization]);
    });
    
    // Add data to sheet
    reportSheet.sheet.getRange(1, 1, reportData.length, 7).setValues(reportData);
    
    // Format headers
    reportSheet.sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
    reportSheet.sheet.getRange(4, 1).setFontWeight('bold').setBackground('#E8F0FE');
    reportSheet.sheet.getRange(13, 1).setFontWeight('bold').setBackground('#E8F0FE');
    reportSheet.sheet.getRange(14, 1, 1, 7).setFontWeight('bold');
    
    const vanUtilRow = reportData.findIndex(row => row[0] === 'VEHICLE UTILIZATION BY TYPE') + 1;
    reportSheet.sheet.getRange(vanUtilRow, 1).setFontWeight('bold').setBackground('#E8F0FE');
    reportSheet.sheet.getRange(vanUtilRow + 1, 1, 1, 4).setFontWeight('bold');
    
    reportSheet.autoResizeColumns(1, 7);
    
    timer.end();
    google.script.host.close();
    
    logger.info('Weekly summary report generated', {
      sheet: reportSheetName,
      weekNumber: currentWeek,
      totalRoutes: weeklyTotals.totalRoutes
    });
    
    UIHelpers.showSuccess(`Weekly summary report generated!\nSheet: ${reportSheetName}`);
    
    SmartDefaults.recordAction('report_generated', {
      type: 'weekly_summary',
      weekNumber: currentWeek,
      totalRoutes: weeklyTotals.totalRoutes
    });
    
  } catch (error) {
    logger.error('Failed to generate weekly summary report', { error: error.message });
    google.script.host.close();
    UIHelpers.showError(error);
  }
}

/**
 * Calculate average delivery completion time
 * Analyzes delivery pace data to determine average completion times
 */
function calculateAverageDeliveryTime() {
  const logger = Logger.createLogger('ReportService');
  
  try {
    const manager = new SheetManager();
    const dailyDetails = manager.getSheet(getConfig('SHEETS.DAILY_DETAILS'));
    const data = dailyDetails.getData();
    
    const timeSlots = getConfig('DELIVERY_TIME_SLOTS');
    const completionTimes = {};
    
    // Initialize time slot data
    timeSlots.forEach(slot => {
      completionTimes[slot.label] = {
        totalStops: 0,
        routeCount: 0,
        averageStops: 0
      };
    });
    
    // Analyze delivery pace columns (L-P, indices 11-15)
    data.slice(1).forEach(row => {
      timeSlots.forEach((slot, index) => {
        const stops = parseInt(row[slot.column - 1]) || 0;
        if (stops > 0) {
          completionTimes[slot.label].totalStops += stops;
          completionTimes[slot.label].routeCount++;
        }
      });
    });
    
    // Calculate averages
    Object.values(completionTimes).forEach(slot => {
      if (slot.routeCount > 0) {
        slot.averageStops = (slot.totalStops / slot.routeCount).toFixed(1);
      }
    });
    
    return completionTimes;
    
  } catch (error) {
    logger.error('Failed to calculate average delivery time', { error: error.message });
    return {};
  }
}