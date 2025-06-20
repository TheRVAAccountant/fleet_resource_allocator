/**
 * Email Service for Fleet Resource Allocator
 * Handles email notifications for delivery pace form submissions
 */

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
    const config = getConfig();
    const recipient = config.EMAIL_RECIPIENT;
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
    return false;
  }
}

/**
 * Creates email subject line
 * @param {Object} formData - Form submission data
 * @returns {string} Formatted subject line
 */
function createEmailSubject(formData) {
  const dateStr = Utilities.formatDate(new Date(formData.date), Session.getScriptTimeZone(), 'MM/dd/yyyy');
  const timeStr = Utilities.formatDate(formData.timestamp, Session.getScriptTimeZone(), 'h:mm a');
  return `Delivery Pace Update - Van ${formData.vanId} - ${dateStr} @ ${timeStr}`;
}

/**
 * Creates HTML email body with formatted delivery data
 * @param {Object} formData - Form submission data
 * @returns {string} HTML email content
 */
function createEmailBody(formData) {
  const dateStr = Utilities.formatDate(new Date(formData.date), Session.getScriptTimeZone(), 'MM/dd/yyyy');
  const submissionTime = Utilities.formatDate(formData.timestamp, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy h:mm a z');
  
  // Calculate metrics
  const checkpoints = Object.keys(formData.deliveries).sort();
  const latestCheckpoint = checkpoints[checkpoints.length - 1];
  const totalDeliveries = formData.deliveries[latestCheckpoint] || 0;
  const averagePace = calculateAveragePace(formData.deliveries);
  
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
  const config = getConfig();
  html += `
        <div class="info-box" style="text-align: center;">
          <p style="margin: 0;">
            <a href="https://docs.google.com/spreadsheets/d/${config.DAILY_SUMMARY_SPREADSHEET_ID}" 
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
  const checkpoints = Object.keys(deliveries).sort();
  if (checkpoints.length < 2) return 0;
  
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
  const parseTime = (timeStr) => {
    const [time, period] = timeStr.split(' ');
    let [hours, minutes] = time.split(':').map(Number);
    
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
    SpreadsheetApp.getUi().alert('Test email sent successfully! Check inbox for ' + getConfig().EMAIL_RECIPIENT);
  } else {
    console.log('Failed to send test email');
    SpreadsheetApp.getUi().alert('Failed to send test email. Check logs for details.');
  }
}