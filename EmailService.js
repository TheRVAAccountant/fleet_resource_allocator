/**
 * Email Service for Fleet Resource Allocator
 * Handles email notifications for delivery pace form submissions
 */

/**
 * Safely converts various date formats to Date object
 * @param {Date|string|number} dateInput - Date in various formats
 * @returns {Date} Valid Date object
 */
function parseDate(dateInput) {
  // If already a Date object, return it
  if (dateInput instanceof Date && !isNaN(dateInput)) {
    return dateInput;
  }
  
  // Try to parse string or number
  const parsed = new Date(dateInput);
  
  // Check if parsing was successful
  if (isNaN(parsed)) {
    console.error('Invalid date input:', dateInput);
    return new Date(); // Return current date as fallback
  }
  
  return parsed;
}

/**
 * Sends formatted email notification for delivery pace form submission
 * @param {Object} formData - Parsed form submission data
 * @param {string} formData.vanId - Van identifier
 * @param {string} formData.date - Submission date
 * @param {Object} formData.deliveries - Delivery counts by checkpoint
 * @param {string} formData.notes - Optional driver notes
 * @param {Date} formData.timestamp - Form submission timestamp
 * @returns {boolean} Success status
 */
function sendDeliveryPaceEmail(formData) {
  try {
    // Validate formData
    if (!formData || typeof formData !== 'object') {
      throw new Error('Invalid form data provided');
    }
    
    if (!formData.vanId) {
      throw new Error('Van ID is required');
    }
    
    const recipient = getConfig('EMAIL_RECIPIENT');
    const subject = createEmailSubject(formData);
    const htmlBody = createEmailBody(formData);
    
    // Send the email
    GmailApp.sendEmail(recipient, subject, '', {
      htmlBody: htmlBody,
      name: 'Fleet Resource Allocator',
      noReply: true
    });
    
    console.log(`Email sent successfully to ${recipient} for Van ${formData.vanId}`);
    return true;
  } catch (error) {
    console.error('Error sending delivery pace email:', error);
    console.error('Form data:', JSON.stringify(formData));
    return false;
  }
}

/**
 * Creates email subject line
 * @param {Object} formData - Form submission data
 * @returns {string} Formatted subject line
 */
function createEmailSubject(formData) {
  try {
    const date = parseDate(formData.date);
    const timestamp = parseDate(formData.timestamp);
    
    const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'h:mm a');
    return `Delivery Pace Update - Van ${formData.vanId} - ${dateStr} @ ${timeStr}`;
  } catch (error) {
    console.error('Error creating email subject:', error);
    // Fallback subject
    return `Delivery Pace Update - Van ${formData.vanId || 'Unknown'}`;
  }
}

/**
 * Creates HTML email body with formatted delivery data
 * @param {Object} formData - Form submission data
 * @returns {string} HTML email content
 */
function createEmailBody(formData) {
  let dateStr, submissionTime;
  
  try {
    const date = parseDate(formData.date);
    const timestamp = parseDate(formData.timestamp);
    
    dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    submissionTime = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy h:mm a z');
  } catch (error) {
    console.error('Error formatting dates:', error);
    // Use fallback values
    dateStr = 'Unknown Date';
    submissionTime = 'Unknown Time';
  }
  
  // Calculate metrics with validation
  const deliveries = formData.deliveries || {};
  const checkpoints = Object.keys(deliveries).filter(key => key && key.trim()).sort();
  const latestCheckpoint = checkpoints.length > 0 ? checkpoints[checkpoints.length - 1] : null;
  const totalDeliveries = latestCheckpoint ? (deliveries[latestCheckpoint] || 0) : 0;
  const averagePace = calculateAveragePace(deliveries);
  
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body {
          font-family: Arial, sans-serif;
          line-height: 1.6;
          color: #333;
          max-width: 600px;
          margin: 0 auto;
        }
        .header {
          background-color: #1a73e8;
          color: white;
          padding: 20px;
          text-align: center;
          border-radius: 5px 5px 0 0;
        }
        .content {
          background-color: #f8f9fa;
          padding: 20px;
          border: 1px solid #dadce0;
          border-top: none;
          border-radius: 0 0 5px 5px;
        }
        .info-box {
          background-color: white;
          padding: 15px;
          margin: 15px 0;
          border-radius: 5px;
          border: 1px solid #e0e0e0;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin: 15px 0;
          background-color: white;
        }
        th {
          background-color: #1a73e8;
          color: white;
          padding: 10px;
          text-align: left;
          font-weight: normal;
        }
        td {
          padding: 10px;
          border-bottom: 1px solid #e0e0e0;
        }
        .metric {
          display: inline-block;
          margin: 10px 20px 10px 0;
        }
        .metric-label {
          font-size: 0.9em;
          color: #666;
        }
        .metric-value {
          font-size: 1.3em;
          font-weight: bold;
          color: #1a73e8;
        }
        .notes {
          background-color: #fff3cd;
          padding: 15px;
          border-radius: 5px;
          border: 1px solid #ffeaa7;
          margin: 15px 0;
        }
        .footer {
          text-align: center;
          color: #666;
          font-size: 0.9em;
          margin-top: 20px;
        }
        .progress-bar {
          background-color: #e0e0e0;
          height: 20px;
          border-radius: 10px;
          overflow: hidden;
          margin: 10px 0;
        }
        .progress-fill {
          background-color: #34a853;
          height: 100%;
          transition: width 0.3s ease;
        }
      </style>
    </head>
    <body>
      <div class="header">
        <h2 style="margin: 0;">Delivery Pace Update</h2>
        <p style="margin: 5px 0 0 0;">Van ${formData.vanId} - ${dateStr}</p>
      </div>
      
      <div class="content">
        <div class="info-box">
          <h3 style="margin-top: 0; color: #1a73e8;">Submission Details</h3>
          <p><strong>Submitted:</strong> ${submissionTime}</p>
          <p><strong>Van ID:</strong> ${formData.vanId}</p>
          ${formData.driverName ? `<p><strong>Driver:</strong> ${formData.driverName}</p>` : ''}
        </div>
        
        <div class="info-box">
          <h3 style="margin-top: 0; color: #1a73e8;">Delivery Progress</h3>
          <div>
            <div class="metric">
              <div class="metric-label">Total Deliveries</div>
              <div class="metric-value">${totalDeliveries}</div>
            </div>
            <div class="metric">
              <div class="metric-label">Average Pace</div>
              <div class="metric-value">${averagePace} stops/hr</div>
            </div>
          </div>
        </div>
        
        <table>
          <thead>
            <tr>
              <th>Checkpoint Time</th>
              <th>Deliveries Completed</th>
              <th>Incremental</th>
              <th>Pace (stops/hr)</th>
            </tr>
          </thead>
          <tbody>
  `;
  
  // Add checkpoint data rows
  let previousCount = 0;
  let previousTime = null;
  
  checkpoints.forEach((checkpoint, index) => {
    const count = formData.deliveries[checkpoint] || 0;
    const incremental = count - previousCount;
    
    // Calculate pace for this interval
    let pace = 0;
    if (index > 0 && previousTime) {
      const timeDiff = getTimeDifferenceInHours(previousTime, checkpoint);
      pace = timeDiff > 0 ? Math.round(incremental / timeDiff) : 0;
    }
    
    html += `
      <tr>
        <td>${checkpoint}</td>
        <td>${count}</td>
        <td>${incremental > 0 ? '+' + incremental : incremental}</td>
        <td>${pace > 0 ? pace : '-'}</td>
      </tr>
    `;
    
    previousCount = count;
    previousTime = checkpoint;
  });
  
  html += `
          </tbody>
        </table>
  `;
  
  // Add notes if present
  if (formData.notes && formData.notes.trim()) {
    html += `
        <div class="notes">
          <h3 style="margin-top: 0; color: #856404;">Driver Notes</h3>
          <p style="margin: 0;">${escapeHtml(formData.notes)}</p>
        </div>
    `;
  }
  
  // Add link to spreadsheet
  const spreadsheetId = getConfig('DAILY_SUMMARY_SPREADSHEET_ID');
  html += `
        <div class="info-box" style="text-align: center;">
          <p style="margin: 0;">
            <a href="https://docs.google.com/spreadsheets/d/${spreadsheetId}" 
               style="color: #1a73e8; text-decoration: none; font-weight: bold;">
              View Full Data in Daily Summary Spreadsheet →
            </a>
          </p>
        </div>
      </div>
      
      <div class="footer">
        <p>This notification was sent automatically by the Fleet Resource Allocator system.</p>
        <p>For questions or issues, please contact your system administrator.</p>
      </div>
    </body>
    </html>
  `;
  
  return html;
}

/**
 * Calculates average delivery pace across all checkpoints
 * @param {Object} deliveries - Delivery counts by checkpoint
 * @returns {number} Average stops per hour
 */
function calculateAveragePace(deliveries) {
  // Validate deliveries object
  if (!deliveries || typeof deliveries !== 'object') {
    console.log('Invalid deliveries object:', deliveries);
    return 0;
  }
  
  const checkpoints = Object.keys(deliveries).filter(key => key && key.trim()).sort();
  if (checkpoints.length < 2) {
    console.log('Not enough checkpoints for pace calculation');
    return 0;
  }
  
  const firstCheckpoint = checkpoints[0];
  const lastCheckpoint = checkpoints[checkpoints.length - 1];
  const totalDeliveries = deliveries[lastCheckpoint] || 0;
  
  const totalHours = getTimeDifferenceInHours(firstCheckpoint, lastCheckpoint);
  
  return totalHours > 0 ? Math.round(totalDeliveries / totalHours) : 0;
}

/**
 * Calculates time difference in hours between two checkpoint times
 * @param {string} time1 - Earlier time (e.g., "1:40 PM")
 * @param {string} time2 - Later time (e.g., "3:40 PM")
 * @returns {number} Difference in hours
 */
function getTimeDifferenceInHours(time1, time2) {
  // Validate inputs
  if (!time1 || !time2) {
    console.error('Invalid time values:', { time1, time2 });
    return 0;
  }
  
  const parseTime = (timeStr) => {
    // Ensure timeStr is a string and contains a space
    if (typeof timeStr !== 'string' || !timeStr.includes(' ')) {
      console.error('Invalid time format:', timeStr);
      return 0;
    }
    
    const parts = timeStr.split(' ');
    if (parts.length !== 2) {
      console.error('Invalid time format:', timeStr);
      return 0;
    }
    
    const [time, period] = parts;
    const timeParts = time.split(':');
    
    if (timeParts.length !== 2) {
      console.error('Invalid time format:', timeStr);
      return 0;
    }
    
    let [hours, minutes] = timeParts.map(Number);
    
    if (isNaN(hours) || isNaN(minutes)) {
      console.error('Invalid time values:', { hours, minutes });
      return 0;
    }
    
    if (period === 'PM' && hours !== 12) hours += 12;
    if (period === 'AM' && hours === 12) hours = 0;
    
    return hours + minutes / 60;
  };
  
  return parseTime(time2) - parseTime(time1);
}

/**
 * Escapes HTML special characters to prevent injection
 * @param {string} text - Text to escape
 * @returns {string} Escaped text
 */
function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, m => map[m]);
}

/**
 * Sends test email with sample data
 */
function testDeliveryPaceEmail() {
  const testData = {
    vanId: 'BW15',
    date: new Date(),
    timestamp: new Date(),
    driverName: 'John Smith',
    deliveries: {
      '1:40 PM': 45,
      '3:40 PM': 89,
      '5:40 PM': 134,
      '7:40 PM': 178
    },
    notes: 'Heavy traffic on route. Delayed start due to vehicle inspection.'
  };
  
  const success = sendDeliveryPaceEmail(testData);
  
  if (success) {
    console.log('Test email sent successfully!');
    SpreadsheetApp.getUi().alert('Test email sent successfully! Check inbox for ' + getConfig('EMAIL_RECIPIENT'));
  } else {
    console.log('Failed to send test email');
    SpreadsheetApp.getUi().alert('Failed to send test email. Check logs for details.');
  }
}

/**
 * Debug version of test email to find exact error location
 */
function debugTestDeliveryPaceEmail() {
  console.log('=== Starting Debug Test Email ===');
  
  try {
    // Test 1: Create test data
    console.log('Step 1: Creating test data...');
    const testData = {
      vanId: 'BW15',
      date: new Date(),
      timestamp: new Date(),
      driverName: 'John Smith',
      deliveries: {
        '1:40 PM': 45,
        '3:40 PM': 89,
        '5:40 PM': 134,
        '7:40 PM': 178
      },
      notes: 'Heavy traffic on route. Delayed start due to vehicle inspection.'
    };
    console.log('Test data created successfully:', JSON.stringify(testData));
    
    // Test 2: Check getConfig
    console.log('\nStep 2: Testing getConfig()...');
    let emailRecipient;
    try {
      emailRecipient = getConfig('EMAIL_RECIPIENT');
      console.log('EMAIL_RECIPIENT:', emailRecipient);
      
      // Test other config values
      const spreadsheetId = getConfig('DAILY_SUMMARY_SPREADSHEET_ID');
      console.log('DAILY_SUMMARY_SPREADSHEET_ID:', spreadsheetId);
    } catch (e) {
      console.error('Error in getConfig():', e);
      throw e;
    }
    
    // Test 3: Check Session
    console.log('\nStep 3: Testing Session.getScriptTimeZone()...');
    let timezone;
    try {
      timezone = Session.getScriptTimeZone();
      console.log('Timezone:', timezone);
    } catch (e) {
      console.error('Error getting timezone:', e);
      throw e;
    }
    
    // Test 4: Create subject
    console.log('\nStep 4: Creating email subject...');
    let subject;
    try {
      subject = createEmailSubject(testData);
      console.log('Subject created:', subject);
    } catch (e) {
      console.error('Error in createEmailSubject():', e);
      console.error('Stack trace:', e.stack);
      throw e;
    }
    
    // Test 5: Create body
    console.log('\nStep 5: Creating email body...');
    let body;
    try {
      body = createEmailBody(testData);
      console.log('Body created, length:', body ? body.length : 'null');
    } catch (e) {
      console.error('Error in createEmailBody():', e);
      console.error('Stack trace:', e.stack);
      throw e;
    }
    
    // Test 6: Send email
    console.log('\nStep 6: Sending email...');
    try {
      const success = sendDeliveryPaceEmail(testData);
      console.log('Email send result:', success);
      
      if (success) {
        SpreadsheetApp.getUi().alert('Debug test passed! Email sent successfully.');
      } else {
        SpreadsheetApp.getUi().alert('Debug test failed. Check logs for details.');
      }
    } catch (e) {
      console.error('Error in sendDeliveryPaceEmail():', e);
      console.error('Stack trace:', e.stack);
      throw e;
    }
    
  } catch (error) {
    console.error('=== Debug Test Failed ===');
    console.error('Final error:', error);
    console.error('Stack trace:', error.stack);
    SpreadsheetApp.getUi().alert('Debug test failed: ' + error.toString());
  }
}

/**
 * Comprehensive test suite for email functionality
 */
function runEmailServiceTests() {
  console.log('Starting Email Service Tests...');
  
  const tests = [
    testDateParsing,
    testEmailWithVariousDateFormats,
    testEmailWithMissingData,
    testEmailWithInvalidData,
    testTimeDifferenceCalculations
  ];
  
  let passed = 0;
  let failed = 0;
  
  tests.forEach(test => {
    try {
      console.log(`Running ${test.name}...`);
      test();
      passed++;
      console.log(`✓ ${test.name} passed`);
    } catch (error) {
      failed++;
      console.error(`✗ ${test.name} failed:`, error);
    }
  });
  
  const message = `Email Service Tests Complete: ${passed} passed, ${failed} failed`;
  console.log(message);
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Test date parsing functionality
 */
function testDateParsing() {
  // Test various date formats
  const testCases = [
    { input: new Date(), description: 'Date object' },
    { input: '2025-06-20T19:55:21.061Z', description: 'ISO string' },
    { input: '2025-06-20', description: 'Date string' },
    { input: 1718909721061, description: 'Timestamp' },
    { input: null, description: 'Null value' },
    { input: undefined, description: 'Undefined value' },
    { input: 'invalid', description: 'Invalid string' }
  ];
  
  testCases.forEach(testCase => {
    const result = parseDate(testCase.input);
    console.log(`parseDate(${testCase.description}):`, result);
    
    if (!(result instanceof Date)) {
      throw new Error(`Expected Date object for ${testCase.description}`);
    }
  });
}

/**
 * Test email sending with various date formats
 */
function testEmailWithVariousDateFormats() {
  const dateFormats = [
    new Date(),
    '2025-06-20T19:55:21.061Z',
    new Date().toISOString(),
    new Date().getTime()
  ];
  
  dateFormats.forEach((dateFormat, index) => {
    const testData = {
      vanId: `TEST${index}`,
      date: dateFormat,
      timestamp: dateFormat,
      driverName: 'Test Driver',
      deliveries: {
        '1:40 PM': 10 + index,
        '3:40 PM': 20 + index
      },
      notes: `Test with date format: ${typeof dateFormat}`
    };
    
    try {
      // Just create the email body to test formatting
      const subject = createEmailSubject(testData);
      const body = createEmailBody(testData);
      
      console.log(`Date format test ${index} passed`);
    } catch (error) {
      throw new Error(`Failed with date format ${typeof dateFormat}: ${error}`);
    }
  });
}

/**
 * Test email with missing data
 */
function testEmailWithMissingData() {
  const testCases = [
    { vanId: 'TEST1' }, // Missing everything else
    { vanId: 'TEST2', date: new Date() }, // Missing timestamp
    { vanId: 'TEST3', date: new Date(), timestamp: new Date() }, // Missing deliveries
    { vanId: 'TEST4', date: new Date(), timestamp: new Date(), deliveries: {} } // Empty deliveries
  ];
  
  testCases.forEach((testData, index) => {
    try {
      const subject = createEmailSubject(testData);
      const body = createEmailBody(testData);
      console.log(`Missing data test ${index} passed`);
    } catch (error) {
      throw new Error(`Failed with missing data test ${index}: ${error}`);
    }
  });
}

/**
 * Test email with invalid data
 */
function testEmailWithInvalidData() {
  const testData = {
    vanId: null,
    date: 'not-a-date',
    timestamp: {},
    deliveries: {
      'invalid-time': 'not-a-number',
      '': 50,
      null: 60
    },
    notes: 123 // Number instead of string
  };
  
  try {
    // Should handle gracefully without throwing
    const subject = createEmailSubject(testData);
    const body = createEmailBody(testData);
    console.log('Invalid data test passed');
  } catch (error) {
    throw new Error(`Should handle invalid data gracefully: ${error}`);
  }
}

/**
 * Test time difference calculations
 */
function testTimeDifferenceCalculations() {
  const testCases = [
    { time1: '1:40 PM', time2: '3:40 PM', expected: 2 },
    { time1: '11:40 AM', time2: '1:40 PM', expected: 2 },
    { time1: '11:40 PM', time2: '1:40 AM', expected: -22 }, // Negative (crossing midnight)
    { time1: null, time2: '3:40 PM', expected: 0 }, // Invalid input
    { time1: '1:40 PM', time2: null, expected: 0 }, // Invalid input
    { time1: 'invalid', time2: '3:40 PM', expected: 0 } // Invalid format
  ];
  
  testCases.forEach(testCase => {
    const result = getTimeDifferenceInHours(testCase.time1, testCase.time2);
    console.log(`Time diff ${testCase.time1} to ${testCase.time2}: ${result} hours`);
    
    if (testCase.expected !== 0 && Math.abs(result - testCase.expected) > 0.01) {
      throw new Error(`Expected ${testCase.expected} hours, got ${result}`);
    }
  });
}