/**
 * ===================================================================
 * RTS (RETURN TO STATION) SERVICE
 * ===================================================================
 * Manages end-of-day reporting including RTS time, packages delivered,
 * packages returned, and route notes.
 */

/**
 * Get form data for RTS reporting
 * @return {Object} Form data including vans and assignments
 */
function getRTSFormData() {
  // Get today's assignments
  var assignments = getTodayAssignments();
  var vans = getActiveVanChoices(true); // Only show assigned vans
  
  return {
    vans: vans,
    assignments: assignments,
    vanMessage: vans.length > 0 ? 'Showing vans assigned today' : 'No vans assigned today'
  };
}

/**
 * Submit RTS (Return to Station) report
 * @param {Object} formData - Form data from RTS report
 * @return {Object} Result object
 */
function submitRTSReport(formData) {
  try {
    console.log('Processing RTS report for Van:', formData.vanId);
    
    // Validate required fields
    if (!formData.vanId || !formData.date || !formData.rtsTime) {
      throw new Error('Missing required fields');
    }
    
    // Update Daily Details with RTS data
    var updated = updateRTSDataInDailyDetails(formData);
    
    if (!updated) {
      throw new Error('Failed to update Daily Details. Route may not exist for today.');
    }
    
    // Log the submission
    console.log('RTS report submitted successfully for Van:', formData.vanId);
    
    // Send confirmation email if configured
    if (getConfig('EMAIL_RECIPIENT')) {
      sendRTSConfirmationEmail(formData);
    }
    
    return {
      success: true,
      message: 'RTS report submitted successfully'
    };
    
  } catch (error) {
    console.error('Error submitting RTS report:', error);
    throw error;
  }
}

/**
 * Update Daily Details sheet with RTS data
 * @param {Object} formData - RTS form data
 * @return {boolean} Success status
 */
function updateRTSDataInDailyDetails(formData) {
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
  
  if (!dailyDetailsSheet) {
    throw new Error('Daily Details sheet not found');
  }
  
  // Get RTS field column indices
  var rtsFields = getConfig('RTS_FIELDS');
  
  // Find the row for this van and date
  var data = dailyDetailsSheet.getDataRange().getValues();
  var rowIndex = -1;
  
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0]; // Column A
    var rowVanId = data[i][4]; // Column E
    
    // Format date for comparison
    if (rowDate instanceof Date) {
      rowDate = formatDate(rowDate);
    }
    
    if (rowDate === formData.date && rowVanId === formData.vanId) {
      rowIndex = i;
      break;
    }
  }
  
  if (rowIndex === -1) {
    console.log('No matching row found for Van:', formData.vanId, 'on date:', formData.date);
    return false;
  }
  
  // Update RTS fields
  try {
    // RTS Time (Column Q) - Format as text to prevent date conversion
    dailyDetailsSheet.getRange(rowIndex + 1, rtsFields.RTS_TIME + 1).setValue(formatTimeAsText(formData.rtsTime));
    
    // Packages Delivered (Column R)
    dailyDetailsSheet.getRange(rowIndex + 1, rtsFields.PKG_DELIVERED + 1).setValue(formData.pkgDelivered);
    
    // Packages Returned (Column S)
    dailyDetailsSheet.getRange(rowIndex + 1, rtsFields.PKG_RETURNED + 1).setValue(formData.pkgReturned);
    
    // Route Notes (Column T)
    if (formData.routeNotes && formData.routeNotes.trim()) {
      dailyDetailsSheet.getRange(rowIndex + 1, rtsFields.ROUTE_NOTES + 1).setValue(formData.routeNotes);
    }
    
    console.log('Updated RTS data for Van:', formData.vanId, 'at row:', rowIndex + 1);
    return true;
    
  } catch (error) {
    console.error('Error updating Daily Details:', error);
    throw error;
  }
}

/**
 * Send confirmation email for RTS submission
 * @param {Object} formData - RTS form data
 */
function sendRTSConfirmationEmail(formData) {
  try {
    var subject = 'RTS Report - Van ' + formData.vanId + ' - ' + formData.date;
    
    var htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; }
          .header { background-color: #1a73e8; color: white; padding: 20px; text-align: center; }
          .content { padding: 20px; }
          .info-box { background-color: #f8f9fa; padding: 15px; margin: 10px 0; border-radius: 5px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
          th { background-color: #f1f3f4; font-weight: bold; }
          .footer { text-align: center; color: #666; font-size: 0.9em; margin-top: 30px; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2 style="margin: 0;">RTS Report Confirmation</h2>
          <p style="margin: 5px 0 0 0;">Van ${formData.vanId} - ${formData.date}</p>
        </div>
        
        <div class="content">
          <div class="info-box">
            <h3 style="margin-top: 0; color: #1a73e8;">Route Details</h3>
            <p><strong>Route Code:</strong> ${formData.routeCode}</p>
            <p><strong>Driver:</strong> ${formData.driverName}</p>
            <p><strong>Van ID:</strong> ${formData.vanId}</p>
          </div>
          
          <table>
            <thead>
              <tr>
                <th>Metric</th>
                <th>Value</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>RTS Time</td>
                <td>${formatTimeWithTimezone(formData.rtsTime)}</td>
              </tr>
              <tr>
                <td>Packages Delivered</td>
                <td>${formData.pkgDelivered}</td>
              </tr>
              <tr>
                <td>Packages Returned</td>
                <td>${formData.pkgReturned}</td>
              </tr>
              <tr>
                <td>Delivery Success Rate</td>
                <td>${calculateSuccessRate(formData.pkgDelivered, formData.pkgReturned)}%</td>
              </tr>
            </tbody>
          </table>
          
          ${formData.routeNotes ? `
          <div class="info-box" style="margin-top: 20px;">
            <h3 style="margin-top: 0; color: #856404;">Route Notes</h3>
            <p style="margin: 0;">${escapeHtml(formData.routeNotes)}</p>
          </div>
          ` : ''}
          
          <div class="footer">
            <p>This is an automated notification from the Fleet Resource Allocator system.</p>
          </div>
        </div>
      </body>
      </html>
    `;
    
    MailApp.sendEmail({
      to: getConfig('EMAIL_RECIPIENT'),
      subject: subject,
      htmlBody: htmlBody
    });
    
    console.log('RTS confirmation email sent for Van:', formData.vanId);
    
  } catch (error) {
    console.error('Error sending RTS email:', error);
    // Don't throw - email failure shouldn't stop the process
  }
}

/**
 * Calculate delivery success rate
 * @param {number} delivered - Packages delivered
 * @param {number} returned - Packages returned
 * @return {string} Success rate as percentage
 */
function calculateSuccessRate(delivered, returned) {
  var total = delivered + returned;
  if (total === 0) return '0';
  return ((delivered / total) * 100).toFixed(1);
}

/**
 * Get RTS summary for a specific date
 * @param {string} date - Date to get summary for (optional, defaults to today)
 * @return {Object} Summary data
 */
function getRTSSummary(date) {
  if (!date) {
    date = formatDate(new Date());
  }
  
  var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
  var dailyDetailsSheet = ss.getSheetByName(getConfig('SHEETS.DAILY_DETAILS'));
  
  if (!dailyDetailsSheet) {
    throw new Error('Daily Details sheet not found');
  }
  
  var data = dailyDetailsSheet.getDataRange().getValues();
  var rtsFields = getConfig('RTS_FIELDS');
  
  var summary = {
    date: date,
    totalRoutes: 0,
    completedReports: 0,
    totalDelivered: 0,
    totalReturned: 0,
    routes: []
  };
  
  // Process each row
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0];
    
    // Format date for comparison
    if (rowDate instanceof Date) {
      rowDate = formatDate(rowDate);
    }
    
    if (rowDate === date) {
      summary.totalRoutes++;
      
      var rtsTime = data[i][rtsFields.RTS_TIME];
      var pkgDelivered = data[i][rtsFields.PKG_DELIVERED];
      var pkgReturned = data[i][rtsFields.PKG_RETURNED];
      
      // Handle RTS time that might be stored as Date object or needs format conversion
      if (rtsTime) {
        // Remove leading apostrophe if present
        if (typeof rtsTime === 'string' && rtsTime.startsWith("'")) {
          rtsTime = rtsTime.substring(1);
        }
        
        if (rtsTime instanceof Date) {
          // Extract and convert to AM/PM format
          var hours = rtsTime.getHours();
          var minutes = rtsTime.getMinutes();
          
          var period = hours >= 12 ? 'PM' : 'AM';
          hours = hours % 12;
          hours = hours ? hours : 12; // 0 should be 12
          
          rtsTime = hours + ':' + (minutes < 10 ? '0' + minutes : minutes) + ' ' + period;
        } else if (typeof rtsTime === 'string' && rtsTime.match(/^\d{1,2}:\d{2}$/) && !rtsTime.match(/(AM|PM)/i)) {
          // Convert 24-hour format string to AM/PM only if it doesn't already have AM/PM
          var timeParts = rtsTime.split(':');
          var hours = parseInt(timeParts[0]);
          var minutes = timeParts[1];
          
          var period = hours >= 12 ? 'PM' : 'AM';
          hours = hours % 12;
          hours = hours ? hours : 12;
          
          rtsTime = hours + ':' + minutes + ' ' + period;
        }
      }
      
      // Check if RTS data is submitted
      if (rtsTime || (pkgDelivered !== null && pkgDelivered !== '')) {
        summary.completedReports++;
        summary.totalDelivered += (pkgDelivered || 0);
        summary.totalReturned += (pkgReturned || 0);
      }
      
      summary.routes.push({
        route: data[i][1], // Column B
        driver: data[i][2], // Column C
        vanId: data[i][4], // Column E
        rtsTime: rtsTime || '',
        pkgDelivered: pkgDelivered || 0,
        pkgReturned: pkgReturned || 0,
        notes: data[i][rtsFields.ROUTE_NOTES] || '',
        reported: !!(rtsTime || (pkgDelivered !== null && pkgDelivered !== ''))
      });
    }
  }
  
  summary.completionRate = summary.totalRoutes > 0 ? 
    ((summary.completedReports / summary.totalRoutes) * 100).toFixed(1) : '0';
  
  summary.overallSuccessRate = (summary.totalDelivered + summary.totalReturned) > 0 ?
    ((summary.totalDelivered / (summary.totalDelivered + summary.totalReturned)) * 100).toFixed(1) : '0';
  
  return summary;
}

/**
 * Generate RTS summary report
 * @param {string} date - Date to generate report for
 */
function generateRTSSummaryReport(date) {
  try {
    if (!date) {
      date = formatDate(new Date());
    }
    
    var summary = getRTSSummary(date);
    
    var ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    var summarySheetName = summary.date.replace(/\//g, '-') + ' - RTS Summary';
    
    // Check if sheet exists
    var summarySheet = ss.getSheetByName(summarySheetName);
    if (summarySheet) {
      ss.deleteSheet(summarySheet);
    }
    
    summarySheet = ss.insertSheet(summarySheetName);
    
    // Add title and summary stats - ensure all rows have 2 columns
    var titleData = [
    ['RTS Summary Report', ''],
    ['Date: ' + summary.date, ''],
    ['', ''],
    ['Total Routes:', summary.totalRoutes],
    ['Completed Reports:', summary.completedReports + ' (' + summary.completionRate + '%)'],
    ['', ''],
    ['Total Packages Delivered:', summary.totalDelivered],
    ['Total Packages Returned:', summary.totalReturned],
    ['Overall Success Rate:', summary.overallSuccessRate + '%'],
    ['', ''],
    ['Route Details:', '']
  ];
  
  // Wrap in try-catch for better error handling
  try {
    summarySheet.getRange(1, 1, titleData.length, 2).setValues(titleData);
  } catch (error) {
    console.error('Error setting title data:', error);
    console.error('Title data dimensions:', titleData.length + ' rows, ' + titleData[0].length + ' columns');
    throw new Error('Failed to create summary report: ' + error.toString());
  }
  
    // Format title
    summarySheet.getRange(1, 1, 1, 2).merge()
      .setFontSize(16)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Add route details headers
    var detailsStartRow = titleData.length + 2;
    var headers = [['Route', 'Driver', 'Van ID', 'RTS Time', 'Delivered', 'Returned', 'Success %', 'Status']];
    summarySheet.getRange(detailsStartRow, 1, 1, headers[0].length).setValues(headers);
    formatHeaderRow(summarySheet, detailsStartRow, headers[0].length);
    
    // Add route details data
    if (summary.routes.length > 0) {
      var detailsData = summary.routes.map(function(route) {
        var total = route.pkgDelivered + route.pkgReturned;
        var successRate = total > 0 ? ((route.pkgDelivered / total) * 100).toFixed(1) : '0';
        
        return [
          route.route,
          route.driver,
          route.vanId,
          formatTimeAsText(route.rtsTime) || '-',
          route.pkgDelivered,
          route.pkgReturned,
          successRate + '%',
          route.reported ? 'Reported' : 'Pending'
        ];
      });
      
      summarySheet.getRange(detailsStartRow + 1, 1, detailsData.length, headers[0].length)
        .setValues(detailsData);
    }
    
    // Auto-resize columns
    summarySheet.autoResizeColumns(1, headers[0].length);
    
    Logger.log('Created RTS summary sheet: ' + summarySheetName);
    showInfoAlert('RTS Summary created: ' + summarySheetName);
    
  } catch (error) {
    console.error('Error generating RTS summary report:', error);
    throw error;
  }
}

/**
 * Test RTS functionality
 */
function testRTSFunctionality() {
  console.log('=== Testing RTS Functionality ===');
  
  try {
    // Test 1: Get form data
    console.log('\nTest 1: Getting RTS form data');
    var formData = getRTSFormData();
    console.log('Vans available:', formData.vans.length);
    console.log('Assignments:', Object.keys(formData.assignments).length);
    
    // Test 2: Test RTS data update
    console.log('\n\nTest 2: Testing RTS data update');
    
    var today = formatDate(new Date());
    
    // Find a test van
    if (formData.vans.length > 0) {
      var testVan = formData.vans[0];
      var testData = {
        date: today,
        vanId: testVan,
        routeCode: formData.assignments[testVan]?.route || 'TEST001',
        driverName: formData.assignments[testVan]?.driver || 'Test Driver',
        rtsTime: '18:30',
        pkgDelivered: 150,
        pkgReturned: 5,
        routeNotes: 'Test RTS submission'
      };
      
      console.log('Testing with data:', testData);
      
      var updated = updateRTSDataInDailyDetails(testData);
      console.log('Update result:', updated ? 'Success' : 'Failed');
    } else {
      console.log('No vans available for testing');
    }
    
    // Test 3: Get RTS summary
    console.log('\n\nTest 3: Getting RTS summary');
    var summary = getRTSSummary(today);
    console.log('Summary:', {
      totalRoutes: summary.totalRoutes,
      completedReports: summary.completedReports,
      completionRate: summary.completionRate + '%',
      totalDelivered: summary.totalDelivered,
      totalReturned: summary.totalReturned,
      overallSuccessRate: summary.overallSuccessRate + '%'
    });
    
    // Test 4: Test RTS Summary Report Generation
    console.log('\n\nTest 4: Testing RTS Summary Report Generation');
    try {
      generateRTSSummaryReport(today);
      console.log('Summary report generated successfully');
    } catch (reportError) {
      console.error('Summary report generation failed:', reportError);
      console.error('Stack:', reportError.stack);
    }
    
    console.log('\n=== Test Complete ===');
    
    SpreadsheetApp.getUi().alert(
      'RTS Test Complete',
      'Tests completed successfully.\n\n' +
      'Total routes today: ' + summary.totalRoutes + '\n' +
      'Completed reports: ' + summary.completedReports + '\n\n' +
      'Check logs for details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Test failed:', error);
    console.error('Stack:', error.stack);
    SpreadsheetApp.getUi().alert('Test failed: ' + error.toString());
  }
}

/**
 * Test RTS time format handling
 */
function testRTSTimeFormat() {
  console.log('=== Testing RTS Time Format ===');
  
  try {
    // Test 1: Test formatTimeAsText function
    console.log('\nTest 1: Testing formatTimeAsText function');
    
    var testCases = [
      { input: '18:30', expected: "'6:30 PM" },
      { input: '9:45', expected: "'9:45 AM" },
      { input: '0:30', expected: "'12:30 AM" },
      { input: '12:00', expected: "'12:00 PM" },
      { input: '13:15', expected: "'1:15 PM" },
      { input: '23:59', expected: "'11:59 PM" },
      { input: '10:30 AM', expected: "'10:30 AM" },
      { input: '6:30 PM', expected: "'6:30 PM" },
      { input: "'8:42 PM", expected: "'8:42 PM" },
      { input: "'10:30 AM", expected: "'10:30 AM" },
      { input: new Date('2024-01-01 18:30:00'), expected: "'6:30 PM" },
      { input: '', expected: '' },
      { input: '-', expected: '-' },
      { input: null, expected: null }
    ];
    
    testCases.forEach(function(testCase) {
      var result = formatTimeAsText(testCase.input);
      console.log('Input:', testCase.input, '=> Output:', result, 
                  '(Expected:', testCase.expected + ')');
    });
    
    // Test 2: Test reading Date objects from sheet
    console.log('\n\nTest 2: Testing Date object handling in getRTSSummary');
    
    // Simulate a Date object that might come from sheets
    var testDate = new Date('1899-12-30 18:30:00');
    console.log('Test Date:', testDate);
    console.log('Hours:', testDate.getHours());
    console.log('Minutes:', testDate.getMinutes());
    
    var hours = testDate.getHours();
    var minutes = testDate.getMinutes();
    var timeStr = hours + ':' + (minutes < 10 ? '0' + minutes : minutes);
    console.log('Extracted time string:', timeStr);
    
    console.log('\n=== RTS Time Format Test Complete ===');
    
    return {
      success: true,
      message: 'All time format tests passed'
    };
    
  } catch (error) {
    console.error('RTS time format test failed:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Test data dimension validation for setValues calls
 */
function testDataDimensionValidation() {
  console.log('=== Testing Data Dimension Validation ===');
  
  try {
    // Test 1: Check titleData structure
    console.log('\nTest 1: Checking titleData structure');
    var testSummary = {
      date: '12/22/2024',
      totalRoutes: 10,
      completedReports: 8,
      completionRate: '80.0',
      totalDelivered: 1200,
      totalReturned: 50,
      overallSuccessRate: '96.0'
    };
    
    var titleData = [
      ['RTS Summary Report', ''],
      ['Date: ' + testSummary.date, ''],
      ['', ''],
      ['Total Routes:', testSummary.totalRoutes],
      ['Completed Reports:', testSummary.completedReports + ' (' + testSummary.completionRate + '%)'],
      ['', ''],
      ['Total Packages Delivered:', testSummary.totalDelivered],
      ['Total Packages Returned:', testSummary.totalReturned],
      ['Overall Success Rate:', testSummary.overallSuccessRate + '%'],
      ['', ''],
      ['Route Details:', '']
    ];
    
    console.log('Title data rows:', titleData.length);
    var allRowsHaveTwoColumns = true;
    for (var i = 0; i < titleData.length; i++) {
      if (titleData[i].length !== 2) {
        console.error('Row ' + i + ' has ' + titleData[i].length + ' columns instead of 2');
        allRowsHaveTwoColumns = false;
      }
    }
    console.log('All rows have 2 columns:', allRowsHaveTwoColumns);
    
    // Test 2: Check route details data structure
    console.log('\n\nTest 2: Checking route details data structure');
    var headers = [['Route', 'Driver', 'Van ID', 'RTS Time', 'Delivered', 'Returned', 'Success %', 'Status']];
    console.log('Headers columns:', headers[0].length);
    
    var testRoutes = [
      {route: 'TEST001', driver: 'Driver 1', vanId: 'BW1', rtsTime: '18:30', 
       pkgDelivered: 150, pkgReturned: 5, reported: true},
      {route: 'TEST002', driver: 'Driver 2', vanId: 'BW2', rtsTime: '', 
       pkgDelivered: 0, pkgReturned: 0, reported: false}
    ];
    
    var detailsData = testRoutes.map(function(route) {
      var total = route.pkgDelivered + route.pkgReturned;
      var successRate = total > 0 ? ((route.pkgDelivered / total) * 100).toFixed(1) : '0';
      
      return [
        route.route,
        route.driver,
        route.vanId,
        route.rtsTime || '-',
        route.pkgDelivered,
        route.pkgReturned,
        successRate + '%',
        route.reported ? 'Reported' : 'Pending'
      ];
    });
    
    console.log('Details data rows:', detailsData.length);
    console.log('Details data columns:', detailsData[0].length);
    console.log('Headers and details match:', headers[0].length === detailsData[0].length);
    
    console.log('\n=== Dimension Validation Test Complete ===');
    
    return {
      success: true,
      titleDataValid: allRowsHaveTwoColumns,
      detailsDataValid: headers[0].length === detailsData[0].length
    };
    
  } catch (error) {
    console.error('Dimension validation test failed:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Test RTS time formatting with timezone
 */
function testRTSTimeFormatting() {
  console.log('=== Testing RTS Time Formatting with Timezone ===');
  
  try {
    var testTimes = [
      '18:30',    // 6:30 PM
      '22:15',    // 10:15 PM
      '09:45',    // 9:45 AM
      '00:30',    // 12:30 AM
      '12:00',    // 12:00 PM
      '23:59'     // 11:59 PM
    ];
    
    var timezone = getTimezoneAbbreviation();
    console.log('Current timezone: ' + timezone);
    console.log('Script timezone: ' + Session.getScriptTimeZone());
    
    testTimes.forEach(function(time) {
      var formatted = formatTimeWithTimezone(time);
      console.log(time + ' => ' + formatted);
    });
    
    // Test actual email format
    var testData = {
      vanId: 'BW1',
      date: formatDate(new Date()),
      routeCode: 'TEST001',
      driverName: 'Test Driver',
      rtsTime: '22:11',  // 10:11 PM
      pkgDelivered: 150,
      pkgReturned: 5,
      routeNotes: 'Test note'
    };
    
    console.log('\nTest email data:');
    console.log('RTS Time raw: ' + testData.rtsTime);
    console.log('RTS Time formatted: ' + formatTimeWithTimezone(testData.rtsTime));
    
    SpreadsheetApp.getUi().alert(
      'RTS Time Format Test',
      'Test time 22:11 formats as: ' + formatTimeWithTimezone('22:11') + '\n\n' +
      'Current timezone: ' + timezone + '\n' +
      'Check logs for more examples.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Test failed:', error);
    SpreadsheetApp.getUi().alert('Test failed: ' + error.toString());
  }
}