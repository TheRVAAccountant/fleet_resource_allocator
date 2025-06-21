/**
 * ===================================================================
 * FORM MANAGEMENT SERVICE
 * ===================================================================
 * Centralized management for all Google Forms used in the system
 */

/**
 * Show form management interface
 * Lists all forms, allows enable/disable, and provides configuration
 */
function showFormManagement() {
  const logger = Logger.createLogger('FormManagementService');
  
  try {
    logger.info('Opening form management interface');
    
    // Get all forms associated with the system
    const forms = getSystemForms();
    
    // Create HTML for form management
    const html = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: 'Google Sans', Arial, sans-serif;">
        <h2>Form Management</h2>
        
        <div style="margin-bottom: 20px;">
          <button onclick="createNewForm()" style="
            background: #1a73e8;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 4px;
            cursor: pointer;
            font-family: 'Google Sans', Arial, sans-serif;
          ">Create New Form</button>
        </div>
        
        <div id="formsList">
          ${forms.map((form, index) => `
            <div style="
              border: 1px solid #dadce0;
              border-radius: 8px;
              padding: 16px;
              margin-bottom: 12px;
              background: ${form.enabled ? '#ffffff' : '#f8f9fa'};
            ">
              <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                  <h3 style="margin: 0 0 8px 0; color: #202124;">
                    ${form.name}
                  </h3>
                  <p style="margin: 0 0 8px 0; color: #5f6368; font-size: 14px;">
                    ${form.description}
                  </p>
                  <div style="font-size: 12px; color: #5f6368;">
                    <span>Type: ${form.type}</span> | 
                    <span>Status: ${form.enabled ? '✅ Enabled' : '❌ Disabled'}</span> |
                    <span>Responses: ${form.responseCount || 0}</span>
                  </div>
                </div>
                <div style="display: flex; gap: 8px;">
                  <button onclick="toggleForm(${index})" style="
                    background: ${form.enabled ? '#EA4335' : '#34A853'};
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 14px;
                  ">${form.enabled ? 'Disable' : 'Enable'}</button>
                  <button onclick="configureForm(${index})" style="
                    background: #ffffff;
                    color: #1a73e8;
                    border: 1px solid #dadce0;
                    padding: 8px 16px;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 14px;
                  ">Configure</button>
                  <button onclick="viewForm(${index})" style="
                    background: #ffffff;
                    color: #1a73e8;
                    border: 1px solid #dadce0;
                    padding: 8px 16px;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 14px;
                  ">View</button>
                </div>
              </div>
              
              ${form.formUrl ? `
                <div style="margin-top: 12px; padding-top: 12px; border-top: 1px solid #e0e0e0;">
                  <div style="display: flex; align-items: center; gap: 12px;">
                    <span style="font-size: 12px; color: #5f6368;">Form URL:</span>
                    <input type="text" value="${form.formUrl}" readonly style="
                      flex: 1;
                      padding: 4px 8px;
                      border: 1px solid #dadce0;
                      border-radius: 4px;
                      font-size: 12px;
                      background: #f8f9fa;
                    ">
                    <button onclick="copyUrl('${form.formUrl}')" style="
                      background: #ffffff;
                      color: #5f6368;
                      border: 1px solid #dadce0;
                      padding: 4px 12px;
                      border-radius: 4px;
                      cursor: pointer;
                      font-size: 12px;
                    ">Copy</button>
                    ${form.qrCodeUrl ? `
                      <button onclick="showQRCode('${form.qrCodeUrl}')" style="
                        background: #ffffff;
                        color: #5f6368;
                        border: 1px solid #dadce0;
                        padding: 4px 12px;
                        border-radius: 4px;
                        cursor: pointer;
                        font-size: 12px;
                      ">QR Code</button>
                    ` : ''}
                  </div>
                </div>
              ` : ''}
            </div>
          `).join('')}
        </div>
        
        ${forms.length === 0 ? `
          <div style="
            text-align: center;
            padding: 40px;
            color: #5f6368;
          ">
            <p>No forms configured yet.</p>
            <p>Click "Create New Form" to get started.</p>
          </div>
        ` : ''}
      </div>
      
      <script>
        function toggleForm(index) {
          google.script.run
            .withSuccessHandler(() => {
              google.script.host.close();
              google.script.run.showFormManagement();
            })
            .withFailureHandler(error => alert('Error: ' + error.message))
            .toggleFormStatus(index);
        }
        
        function configureForm(index) {
          google.script.run
            .withFailureHandler(error => alert('Error: ' + error.message))
            .showFormConfiguration(index);
        }
        
        function viewForm(index) {
          google.script.run
            .withFailureHandler(error => alert('Error: ' + error.message))
            .openFormInNewTab(index);
        }
        
        function createNewForm() {
          google.script.run
            .withFailureHandler(error => alert('Error: ' + error.message))
            .showCreateFormDialog();
        }
        
        function copyUrl(url) {
          const temp = document.createElement('textarea');
          temp.value = url;
          document.body.appendChild(temp);
          temp.select();
          document.execCommand('copy');
          document.body.removeChild(temp);
          alert('URL copied to clipboard!');
        }
        
        function showQRCode(qrUrl) {
          window.open(qrUrl, '_blank');
        }
      </script>
    `)
    .setWidth(800)
    .setHeight(600);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Form Management');
    
  } catch (error) {
    logger.error('Failed to show form management', { error: error.message });
    UIHelpers.showError(error);
  }
}

/**
 * Get all system forms with their status
 * @return {Array} Array of form objects
 */
function getSystemForms() {
  const forms = [];
  
  // Check for existing delivery pace form
  const deliveryPaceFormUrl = PropertiesService.getScriptProperties().getProperty('DELIVERY_PACE_FORM_URL');
  if (deliveryPaceFormUrl) {
    forms.push({
      name: 'Delivery Pace Collection',
      description: 'Collect delivery progress at scheduled checkpoints',
      type: 'Data Collection',
      enabled: getConfig('FEATURES.PACE_TRACKING_ENABLED'),
      formUrl: deliveryPaceFormUrl,
      qrCodeUrl: `https://chart.googleapis.com/chart?chs=200x200&cht=qr&chl=${encodeURIComponent(deliveryPaceFormUrl)}`,
      responseCount: getFormResponseCount('Delivery Pace Data')
    });
  }
  
  // Check for RTS form (embedded in sheet)
  forms.push({
    name: 'End of Day RTS Report',
    description: 'Driver end-of-day return to station reporting',
    type: 'Embedded Form',
    enabled: getConfig('FEATURES.RTS_TRACKING_ENABLED'),
    formUrl: null, // Embedded form doesn't have external URL
    responseCount: getRTSResponseCount()
  });
  
  // Add placeholder for future forms
  const futureFormsConfig = [
    {
      name: 'Vehicle Inspection Form',
      description: 'Daily vehicle inspection checklist',
      type: 'Checklist',
      enabled: false,
      planned: true
    },
    {
      name: 'Incident Report Form',
      description: 'Report accidents, damages, or incidents',
      type: 'Incident Reporting',
      enabled: false,
      planned: true
    }
  ];
  
  forms.push(...futureFormsConfig);
  
  return forms;
}

/**
 * Toggle form enabled status
 * @param {number} index - Form index
 */
function toggleFormStatus(index) {
  const logger = Logger.createLogger('FormManagementService');
  
  try {
    const forms = getSystemForms();
    const form = forms[index];
    
    if (!form) {
      throw new Error('Form not found');
    }
    
    // Update configuration based on form type
    switch (form.name) {
      case 'Delivery Pace Collection':
        const currentStatus = getConfig('FEATURES.PACE_TRACKING_ENABLED');
        // This would need to update the CONFIG object - for now we'll store in properties
        PropertiesService.getScriptProperties().setProperty(
          'PACE_TRACKING_ENABLED', 
          (!currentStatus).toString()
        );
        logger.info('Toggled delivery pace tracking', { enabled: !currentStatus });
        break;
        
      case 'End of Day RTS Report':
        const rtsStatus = getConfig('FEATURES.RTS_TRACKING_ENABLED');
        PropertiesService.getScriptProperties().setProperty(
          'RTS_TRACKING_ENABLED', 
          (!rtsStatus).toString()
        );
        logger.info('Toggled RTS tracking', { enabled: !rtsStatus });
        break;
        
      default:
        throw new Error('Cannot toggle this form type');
    }
    
    UIHelpers.toast(
      `${form.name} ${form.enabled ? 'disabled' : 'enabled'} successfully`,
      3000
    );
    
  } catch (error) {
    logger.error('Failed to toggle form status', { error: error.message });
    throw error;
  }
}

/**
 * Show form configuration dialog
 * @param {number} index - Form index
 */
function showFormConfiguration(index) {
  const forms = getSystemForms();
  const form = forms[index];
  
  if (!form || form.planned) {
    UIHelpers.showError(new Error('This form is not yet available for configuration'));
    return;
  }
  
  const html = HtmlService.createHtmlOutput(`
    <div style="padding: 20px; font-family: 'Google Sans', Arial, sans-serif;">
      <h3>${form.name} Configuration</h3>
      
      <div style="margin-top: 20px;">
        <label style="display: block; margin-bottom: 8px; color: #5f6368;">
          Notification Recipients
        </label>
        <input type="email" id="notificationEmail" placeholder="email@example.com" style="
          width: 100%;
          padding: 8px;
          border: 1px solid #dadce0;
          border-radius: 4px;
        " value="${getFormNotificationEmail(form.name) || ''}">
      </div>
      
      <div style="margin-top: 20px;">
        <label style="display: block; margin-bottom: 8px; color: #5f6368;">
          Auto-Submit Time (for scheduled forms)
        </label>
        <select id="autoSubmitTime" style="
          width: 100%;
          padding: 8px;
          border: 1px solid #dadce0;
          border-radius: 4px;
        ">
          <option value="">Disabled</option>
          <option value="13:40">1:40 PM</option>
          <option value="15:40">3:40 PM</option>
          <option value="17:40">5:40 PM</option>
          <option value="19:40">7:40 PM</option>
          <option value="21:40">9:40 PM</option>
        </select>
      </div>
      
      <div style="margin-top: 20px;">
        <label style="display: flex; align-items: center;">
          <input type="checkbox" id="requireAuth" style="margin-right: 8px;">
          Require authentication
        </label>
      </div>
      
      <div style="margin-top: 20px;">
        <label style="display: flex; align-items: center;">
          <input type="checkbox" id="sendConfirmation" style="margin-right: 8px;" checked>
          Send confirmation email to submitter
        </label>
      </div>
      
      <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 30px;">
        <button onclick="google.script.host.close()" style="
          background: #ffffff;
          color: #5f6368;
          border: 1px solid #dadce0;
          padding: 8px 24px;
          border-radius: 4px;
          cursor: pointer;
        ">Cancel</button>
        <button onclick="saveConfiguration()" style="
          background: #1a73e8;
          color: white;
          border: none;
          padding: 8px 24px;
          border-radius: 4px;
          cursor: pointer;
        ">Save</button>
      </div>
    </div>
    
    <script>
      function saveConfiguration() {
        const config = {
          formName: '${form.name}',
          notificationEmail: document.getElementById('notificationEmail').value,
          autoSubmitTime: document.getElementById('autoSubmitTime').value,
          requireAuth: document.getElementById('requireAuth').checked,
          sendConfirmation: document.getElementById('sendConfirmation').checked
        };
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
            alert('Configuration saved successfully!');
          })
          .withFailureHandler(error => alert('Error: ' + error.message))
          .saveFormConfiguration(config);
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Form Configuration');
}

/**
 * Save form configuration
 * @param {Object} config - Form configuration object
 */
function saveFormConfiguration(config) {
  const logger = Logger.createLogger('FormManagementService');
  
  try {
    // Save configuration to properties
    const configKey = `FORM_CONFIG_${config.formName.replace(/\s+/g, '_').toUpperCase()}`;
    PropertiesService.getScriptProperties().setProperty(
      configKey,
      JSON.stringify(config)
    );
    
    logger.info('Form configuration saved', { formName: config.formName });
    
    // If auto-submit time is set, create trigger
    if (config.autoSubmitTime) {
      setupFormAutoSubmitTrigger(config.formName, config.autoSubmitTime);
    }
    
  } catch (error) {
    logger.error('Failed to save form configuration', { error: error.message });
    throw error;
  }
}

/**
 * Get form response count
 * @param {string} sheetName - Response sheet name
 * @return {number} Number of responses
 */
function getFormResponseCount(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(getConfig('DAILY_SUMMARY_SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(sheetName);
    
    if (sheet) {
      return Math.max(0, sheet.getLastRow() - 1); // Subtract header row
    }
    
    return 0;
  } catch (error) {
    return 0;
  }
}

/**
 * Get RTS response count from Daily Details
 * @return {number} Number of RTS submissions
 */
function getRTSResponseCount() {
  try {
    const manager = new SheetManager();
    const dailyDetails = manager.getSheet(getConfig('SHEETS.DAILY_DETAILS'));
    const data = dailyDetails.getData();
    
    let count = 0;
    data.slice(1).forEach(row => {
      if (row[16]) { // RTS Time column
        count++;
      }
    });
    
    return count;
  } catch (error) {
    return 0;
  }
}

/**
 * Get form notification email
 * @param {string} formName - Form name
 * @return {string} Email address
 */
function getFormNotificationEmail(formName) {
  const configKey = `FORM_CONFIG_${formName.replace(/\s+/g, '_').toUpperCase()}`;
  const config = PropertiesService.getScriptProperties().getProperty(configKey);
  
  if (config) {
    const parsed = JSON.parse(config);
    return parsed.notificationEmail;
  }
  
  return getConfig('EMAIL_RECIPIENT');
}

/**
 * Open form in new tab
 * @param {number} index - Form index
 */
function openFormInNewTab(index) {
  const forms = getSystemForms();
  const form = forms[index];
  
  if (!form) {
    throw new Error('Form not found');
  }
  
  if (form.formUrl) {
    // This would open in a new tab - in Apps Script we show the URL
    const html = HtmlService.createHtmlOutput(`
      <script>
        window.open('${form.formUrl}', '_blank');
        google.script.host.close();
      </script>
    `);
    SpreadsheetApp.getUi().showModalDialog(html, 'Opening form...');
  } else if (form.name === 'End of Day RTS Report') {
    // Show RTS form
    showRTSForm();
  } else {
    UIHelpers.showError(new Error('This form is not available yet'));
  }
}

/**
 * Show create form dialog
 */
function showCreateFormDialog() {
  UIHelpers.showError(
    new Error('Form creation wizard coming soon!'),
    { showDetails: false }
  );
}