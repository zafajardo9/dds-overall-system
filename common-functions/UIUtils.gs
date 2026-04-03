/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * UIUtils.gs - User Interface Helpers
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * PURPOSE:
 * Reusable functions for creating menus, dialogs, sidebars, and other UI
 * components in Google Apps Script projects.
 * 
 * PREREQUISITES:
 * - No special OAuth required
 * 
 * USAGE:
 * Copy this file into your Apps Script project alongside other common-utils.
 * 
 * VERSION: 1.0.0
 * LAST UPDATED: April 2026
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * ==============================================================================
 * FUNCTION: createStandardMenu
 * ==============================================================================
 * 
 * PURPOSE:
 * Creates a standardized custom menu structure with common DDS menu patterns.
 * Reduces boilerplate code for menu creation.
 * 
 * PARAMETERS:
 * @param {string} menuName - Top-level menu name (e.g., "DDS Tools")
 * @param {Array} sections - Array of menu section objects:
 *   - name: {string} - Section name (for separator, can be empty)
 *   - items: {Array} - Menu items in this section:
 *     - label: {string} - Menu item text
 *     - function: {string} - Function to call when clicked
 *     - separator: {boolean} - Add separator before this item
 * 
 * @returns {Menu} The created menu (call .addToUi() to display)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * function onOpen() {
 *   createStandardMenu("DDS Tools", [
 *     {
 *       name: "Setup",
 *       items: [
 *         { label: "Configure API Keys", function: "configureApiKeys" },
 *         { label: "Set Drive Folders", function: "setDriveFolders" }
 *       ]
 *     },
 *     {
 *       name: "Automation",
 *       items: [
 *         { label: "Run Now", function: "runAutomation" },
 *         { label: "---", separator: true },
 *         { label: "Schedule Daily", function: "setupSchedule" }
 *       ]
 *     }
 *   ]).addToUi();
 * }
 * ```
 * ==============================================================================
 */
function createStandardMenu(menuName, sections) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu(menuName);
  
  sections.forEach(function(section, sectionIndex) {
    // Add separator between sections (except first)
    if (sectionIndex > 0 && section.items && section.items.length > 0) {
      menu.addSeparator();
    }
    
    // Add section header if provided
    if (section.name) {
      menu.addItem("📁 " + section.name, "null").setEnabled(false);
    }
    
    // Add items
    if (section.items) {
      section.items.forEach(function(item) {
        if (item.separator) {
          menu.addSeparator();
        } else {
          menu.addItem(item.label, item.function);
        }
      });
    }
  });
  
  return menu;
}

/**
 * ==============================================================================
 * FUNCTION: showAlert
 * ==============================================================================
 * 
 * PURPOSE:
 * Shows a simple alert dialog with consistent styling.
 * Wrapper around SpreadsheetApp.getUi().alert()
 * 
 * PARAMETERS:
 * @param {string} title - Dialog title
 * @param {string} message - Message body (can include \n for newlines)
 * @param {string} type - (Optional) Alert type: "info", "warning", "error", "success"
 *                      (default: "info")
 * 
 * @returns {Button} Button clicked by user
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * showAlert("Setup Complete", "Your configuration has been saved.");
 * showAlert("Error", "Failed to connect to API.", "error");
 * ```
 * ==============================================================================
 */
function showAlert(title, message, type) {
  type = type || "info";
  
  var prefixes = {
    info: "ℹ️ ",
    warning: "⚠️ ",
    error: "❌ ",
    success: "✅ "
  };
  
  var prefixedTitle = (prefixes[type] || "") + title;
  
  return SpreadsheetApp.getUi().alert(prefixedTitle, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * ==============================================================================
 * FUNCTION: showConfirm
 * ==============================================================================
 * 
 * PURPOSE:
 * Shows a confirmation dialog and returns true/false based on user response.
 * 
 * PARAMETERS:
 * @param {string} title - Dialog title
 * @param {string} message - Question to confirm
 * @param {string} confirmLabel - (Optional) Confirm button label (default: "Yes")
 * @param {string} cancelLabel - (Optional) Cancel button label (default: "No")
 * 
 * @returns {boolean} True if user confirmed, false otherwise
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * if (showConfirm("Delete Data", "Are you sure you want to clear all data?")) {
 *   clearAllData();
 * }
 * ```
 * ==============================================================================
 */
function showConfirm(title, message, confirmLabel, cancelLabel) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(title, message, ui.ButtonSet.YES_NO);
  return response === ui.Button.YES;
}

/**
 * ==============================================================================
 * FUNCTION: promptForInput
 * ==============================================================================
 * 
 * PURPOSE:
 * Prompts user for text input with validation and default value support.
 * 
 * PARAMETERS:
 * @param {string} title - Dialog title
 * @param {string} prompt - Prompt message
 * @param {string} defaultValue - (Optional) Default value in input field
 * @param {function} validator - (Optional) Validation function(text) returns true if valid
 * @param {string} errorMessage - (Optional) Message to show if validation fails
 * 
 * @returns {Object} Result:
 *   - cancelled: {boolean} - True if user cancelled
 *   - value: {string} - Input value (if not cancelled)
 *   - valid: {boolean} - Whether input passed validation
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = promptForInput(
 *   "API Key",
 *   "Enter your Gemini API key:",
 *   "",
 *   function(text) { return text.length >= 20; },
 *   "API key must be at least 20 characters"
 * );
 * 
 * if (!result.cancelled && result.valid) {
 *   saveApiKey(result.value);
 * }
 * ```
 * ==============================================================================
 */
function promptForInput(title, prompt, defaultValue, validator, errorMessage) {
  var ui = SpreadsheetApp.getUi();
  
  var maxAttempts = 3;
  var attempts = 0;
  
  while (attempts < maxAttempts) {
    var response = ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);
    var button = response.getSelectedButton();
    var text = response.getResponseText();
    
    if (button === ui.Button.CANCEL) {
      return { cancelled: true, value: null, valid: false };
    }
    
    // Use default if empty
    if (!text && defaultValue) {
      text = defaultValue;
    }
    
    // Validate if validator provided
    if (validator && !validator(text)) {
      ui.alert("Invalid Input", errorMessage || "Please check your input and try again.", ui.ButtonSet.OK);
      attempts++;
      continue;
    }
    
    return { cancelled: false, value: text, valid: true };
  }
  
  return { cancelled: true, value: null, valid: false, reason: "max_attempts" };
}

/**
 * ==============================================================================
 * FUNCTION: showProgressDialog
 * ==============================================================================
 * 
 * PURPOSE:
 * Shows a modeless dialog with progress information.
 * Note: In Apps Script, true progress bars are limited; this creates a sidebar
 * or shows toast notifications for progress updates.
 * 
 * PARAMETERS:
 * @param {string} title - Progress title
 * @param {string} message - Current status message
 * @param {number} percent - (Optional) Percent complete 0-100
 * @param {string} type - (Optional) Display type: "toast" (default), "sidebar", "alert"
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * showProgressDialog("Processing Files", "Processed 5 of 20 files...", 25, "toast");
 * ```
 * 
 * NOTES:
 * - "toast" shows temporary notification (disappears after few seconds)
 * - For true progress tracking, use a Sheet cell or custom sidebar
 * ==============================================================================
 */
function showProgressDialog(title, message, percent, type) {
  type = type || "toast";
  percent = percent || 0;
  
  var ui = SpreadsheetApp.getActiveSpreadsheet();
  
  if (type === "toast") {
    var percentText = percent > 0 ? " (" + percent + "%)" : "";
    ui.toast(title, message + percentText, 3);
  } else if (type === "alert") {
    SpreadsheetApp.getUi().alert(title, message + "\n\nProgress: " + percent + "%", SpreadsheetApp.getUi().ButtonSet.OK);
  }
  // "sidebar" type would require HTML service (not implemented in this util)
}

/**
 * ==============================================================================
 * FUNCTION: createSelectionDialog
 * ==============================================================================
 * 
 * PURPOSE:
 * Shows a dialog with multiple options for user to select from.
 * 
 * PARAMETERS:
 * @param {string} title - Dialog title
 * @param {string} prompt - Instructions for user
 * @param {Array} options - Array of option objects:
 *   - value: {*} - Value to return if selected
 *   - label: {string} - Display text
 *   - description: {string} - (Optional) Additional description text
 * @param {*} defaultValue - (Optional) Pre-selected option value
 * 
 * @returns {Object} Selection result:
 *   - cancelled: {boolean} - True if user cancelled
 *   - selected: {*} - Selected value (if not cancelled)
 *   - selectedLabel: {string} - Label of selected option
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = createSelectionDialog(
 *   "Select Model",
 *   "Choose which Gemini model to use:",
 *   [
 *     { value: "gemini-2.0-flash", label: "Flash (Fast)", description: "Quick responses" },
 *     { value: "gemini-2.0-pro", label: "Pro (Quality)", description: "Better analysis" }
 *   ]
 * );
 * 
 * if (!result.cancelled) {
 *   setSelectedModel(result.selected);
 * }
 * ```
 * ==============================================================================
 */
function createSelectionDialog(title, prompt, options, defaultValue) {
  var ui = SpreadsheetApp.getUi();
  
  // Build option list text
  var listText = options.map(function(opt, index) {
    var marker = (opt.value === defaultValue) ? "→ " : "  ";
    var desc = opt.description ? " (" + opt.description + ")" : "";
    return (index + 1) + ". " + marker + opt.label + desc;
  }).join("\n");
  
  var fullPrompt = prompt + "\n\n" + listText + "\n\nEnter number of your choice:";
  
  var response = ui.prompt(title, fullPrompt, ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.CANCEL) {
    return { cancelled: true, selected: null, selectedLabel: null };
  }
  
  var input = response.getResponseText().trim();
  var index = parseInt(input) - 1;
  
  if (isNaN(index) || index < 0 || index >= options.length) {
    ui.alert("Invalid selection");
    return { cancelled: true, selected: null, selectedLabel: null, reason: "invalid" };
  }
  
  var selected = options[index];
  return {
    cancelled: false,
    selected: selected.value,
    selectedLabel: selected.label
  };
}

/**
 * ==============================================================================
 * FUNCTION: showStatusPanel
 * ==============================================================================
 * 
 * PURPOSE:
 * Displays a formatted status panel showing system configuration and health.
 * Useful for diagnostics and setup verification.
 * 
 * PARAMETERS:
 * @param {Array} statusItems - Array of status objects:
 *   - label: {string} - Item name
 *   - value: {string} - Current value
 *   - status: {string} - "ok", "warning", "error", "info"
 *   - action: {string} - (Optional) Suggested action if not OK
 * @param {string} title - (Optional) Panel title (default: "System Status")
 * 
 * @returns {Button} Button clicked
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var status = [
 *   { label: "API Key", value: "Configured", status: "ok" },
 *   { label: "Drive ID", value: "Not set", status: "error", action: "Run setup" },
 *   { label: "Last Run", value: "2026-01-15 10:30", status: "info" }
 * ];
 * showStatusPanel(status, "DDS System Status");
 * ```
 * ==============================================================================
 */
function showStatusPanel(statusItems, title) {
  title = title || "System Status";
  
  var icons = {
    ok: "✅",
    warning: "⚠️",
    error: "❌",
    info: "ℹ️"
  };
  
  var lines = statusItems.map(function(item) {
    var icon = icons[item.status] || "•";
    var line = icon + " " + item.label + ": " + item.value;
    if (item.action) {
      line += "\n   → " + item.action;
    }
    return line;
  });
  
  var message = lines.join("\n\n");
  
  return SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * ==============================================================================
 * FUNCTION: showSetupWizardStep
 * ==============================================================================
 * 
 * PURPOSE:
 * Displays a single step in a multi-step setup wizard with progress indicator.
 * 
 * PARAMETERS:
 * @param {number} step - Current step number (1-based)
 * @param {number} total - Total number of steps
 * @param {string} title - Step title
 * @param {string} instructions - What the user needs to do
 * @param {string} prompt - (Optional) Additional prompt for input
 * @param {string} defaultValue - (Optional) Default input value
 * 
 * @returns {Object} Step result with user response
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = showSetupWizardStep(
 *   1, 4,
 *   "Step 1: Configure API Key",
 *   "Please enter your Gemini API key. You can get one from:\nhttps://aistudio.google.com/app/apikey",
 *   "API Key:"
 * );
 * ```
 * ==============================================================================
 */
function showSetupWizardStep(step, total, title, instructions, prompt, defaultValue) {
  var progress = "Step " + step + " of " + total;
  var progressBar = "█".repeat(step) + "░".repeat(total - step);
  
  var fullTitle = progress + " " + progressBar + "\n" + title;
  var fullMessage = instructions;
  
  if (prompt) {
    return promptForInput(fullTitle, fullMessage + "\n\n" + prompt, defaultValue);
  } else {
    var result = showConfirm(fullTitle, fullMessage + "\n\nContinue?");
    return { cancelled: !result, value: result };
  }
}
