/**
 * ===================================================================
 * UI HELPER FUNCTIONS
 * ===================================================================
 * Reusable UI components and utilities
 */

const UIHelpers = {
  /**
   * Show loading dialog with spinner
   */
  showLoading(message = 'Processing...') {
    const html = HtmlService.createHtmlOutput(`
      <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; padding: 20px;">
        <div class="spinner"></div>
        <p style="margin-top: 20px; color: #5f6368; font-family: 'Google Sans', Arial, sans-serif;">
          ${message}
        </p>
      </div>
      <style>
        body { margin: 0; overflow: hidden; }
        .spinner {
          width: 48px;
          height: 48px;
          border: 5px solid #f3f3f3;
          border-top: 5px solid #1a73e8;
          border-radius: 50%;
          animation: spin 1s linear infinite;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      </style>
    `).setWidth(300).setHeight(200);
    
    return SpreadsheetApp.getUi().showModalDialog(html, ' ');
  },
  
  /**
   * Show success message with animation
   */
  showSuccess(message, options = {}) {
    const { autoClose = true, duration = 2000 } = options;
    
    const html = HtmlService.createHtmlOutput(`
      <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; padding: 20px;">
        <div class="success-icon">
          <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="#34A853" stroke-width="3">
            <path d="M20 6L9 17l-5-5" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        </div>
        <h3 style="margin-top: 20px; color: #34A853; font-family: 'Google Sans', Arial, sans-serif;">
          ${message}
        </h3>
      </div>
      <style>
        body { margin: 0; overflow: hidden; }
        .success-icon {
          animation: scaleIn 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        }
        .success-icon svg {
          stroke-dasharray: 100;
          stroke-dashoffset: 100;
          animation: draw 0.5s ease-in-out 0.3s forwards;
        }
        @keyframes scaleIn {
          0% { transform: scale(0); opacity: 0; }
          100% { transform: scale(1); opacity: 1; }
        }
        @keyframes draw {
          to { stroke-dashoffset: 0; }
        }
      </style>
      ${autoClose ? `
      <script>
        setTimeout(() => {
          google.script.host.close();
        }, ${duration});
      </script>
      ` : ''}
    `).setWidth(350).setHeight(250);
    
    SpreadsheetApp.getUi().showModalDialog(html, ' ');
  },
  
  /**
   * Show error message with details
   */
  showError(error, context = {}) {
    const userMessage = ErrorHandler.getUserFriendlyMessage(error);
    const showDetails = context.showDetails !== false;
    
    const html = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: 'Google Sans', Arial, sans-serif;">
        <div style="display: flex; align-items: center; margin-bottom: 20px;">
          <svg width="32" height="32" viewBox="0 0 24 24" fill="#EA4335" style="margin-right: 12px;">
            <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
          </svg>
          <h3 style="margin: 0; color: #EA4335;">Error</h3>
        </div>
        
        <p style="color: #5f6368; margin-bottom: 20px;">${userMessage}</p>
        
        ${showDetails && error.stack ? `
        <details style="margin-top: 20px;">
          <summary style="cursor: pointer; color: #1a73e8;">Technical Details</summary>
          <pre style="background: #f8f9fa; padding: 10px; border-radius: 4px; overflow: auto; font-size: 12px; margin-top: 10px;">
${error.stack}
${context ? '\nContext: ' + JSON.stringify(context, null, 2) : ''}
          </pre>
        </details>
        ` : ''}
        
        <div style="display: flex; justify-content: flex-end; margin-top: 20px;">
          <button onclick="google.script.host.close()" style="
            background: #1a73e8;
            color: white;
            border: none;
            padding: 8px 24px;
            border-radius: 4px;
            cursor: pointer;
            font-family: 'Google Sans', Arial, sans-serif;
          ">OK</button>
        </div>
      </div>
    `).setWidth(450).setHeight(showDetails ? 400 : 300);
    
    SpreadsheetApp.getUi().showModalDialog(html, ' ');
  },
  
  /**
   * Show confirmation dialog
   */
  confirm(title, message, options = {}) {
    const {
      confirmText = 'Confirm',
      cancelText = 'Cancel',
      confirmColor = '#1a73e8',
      dangerous = false
    } = options;
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      title,
      message,
      ui.ButtonSet.YES_NO
    );
    
    return response === ui.Button.YES;
  },
  
  /**
   * Show progress dialog
   */
  showProgress(title, message) {
    const html = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: 'Google Sans', Arial, sans-serif;">
        <h3 style="margin: 0 0 16px 0; color: #202124;">${title}</h3>
        <p style="color: #5f6368; margin-bottom: 20px;">${message}</p>
        
        <div style="background: #e0e0e0; height: 8px; border-radius: 4px; overflow: hidden;">
          <div id="progressBar" style="
            background: #1a73e8;
            height: 100%;
            width: 0%;
            transition: width 0.3s ease;
          "></div>
        </div>
        
        <div style="display: flex; justify-content: space-between; margin-top: 12px; font-size: 14px; color: #5f6368;">
          <span id="progressText">0%</span>
          <span id="progressStatus">Starting...</span>
        </div>
      </div>
      
      <script>
        window.updateProgress = function(percent, status) {
          document.getElementById('progressBar').style.width = percent + '%';
          document.getElementById('progressText').textContent = percent + '%';
          if (status) {
            document.getElementById('progressStatus').textContent = status;
          }
        };
      </script>
    `).setWidth(400).setHeight(200);
    
    return SpreadsheetApp.getUi().showModalDialog(html, ' ');
  },
  
  /**
   * Show input dialog
   */
  promptInput(title, message, defaultValue = '') {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      title,
      message,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      return response.getResponseText();
    }
    
    return null;
  },
  
  /**
   * Show multi-select dialog
   */
  showMultiSelect(title, options, selected = []) {
    const html = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: 'Google Sans', Arial, sans-serif;">
        <h3 style="margin: 0 0 16px 0; color: #202124;">${title}</h3>
        
        <div style="max-height: 300px; overflow-y: auto; border: 1px solid #dadce0; border-radius: 4px; padding: 8px;">
          ${options.map((option, index) => `
            <label style="display: block; padding: 8px; cursor: pointer; hover: background: #f8f9fa;">
              <input type="checkbox" id="option_${index}" value="${option}" 
                ${selected.includes(option) ? 'checked' : ''} 
                style="margin-right: 8px;">
              ${option}
            </label>
          `).join('')}
        </div>
        
        <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 20px;">
          <button onclick="google.script.host.close()" style="
            background: transparent;
            color: #5f6368;
            border: 1px solid #dadce0;
            padding: 8px 24px;
            border-radius: 4px;
            cursor: pointer;
            font-family: 'Google Sans', Arial, sans-serif;
          ">Cancel</button>
          <button onclick="submitSelection()" style="
            background: #1a73e8;
            color: white;
            border: none;
            padding: 8px 24px;
            border-radius: 4px;
            cursor: pointer;
            font-family: 'Google Sans', Arial, sans-serif;
          ">OK</button>
        </div>
      </div>
      
      <script>
        function submitSelection() {
          const selected = [];
          ${options.map((_, index) => `
            if (document.getElementById('option_${index}').checked) {
              selected.push(document.getElementById('option_${index}').value);
            }
          `).join('')}
          google.script.run
            .withSuccessHandler(() => google.script.host.close())
            .processMultiSelectResult(selected);
        }
      </script>
    `).setWidth(400).setHeight(450);
    
    SpreadsheetApp.getUi().showModalDialog(html, title);
  },
  
  /**
   * Show toast notification
   */
  toast(message, duration = 3000) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, '', duration / 1000);
  }
};