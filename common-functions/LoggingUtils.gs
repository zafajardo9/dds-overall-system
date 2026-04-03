/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * LoggingUtils.gs - Activity Logging & Audit Trail Functions
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * PURPOSE:
 * Reusable functions for logging system activities, user actions, and audit trails
 * to Google Sheets. Includes automatic log rotation and structured logging formats.
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
 * FUNCTION: initializeLogSheet
 * ==============================================================================
 * 
 * PURPOSE:
 * Creates or initializes a log sheet with standard column headers.
 * Safe to call multiple times - will add headers if sheet exists but is empty.
 * 
 * PARAMETERS:
 * @param {Spreadsheet} spreadsheet - Spreadsheet object to create log in
 * @param {string} sheetName - Name of the log sheet (default: "Activity Log")
 * @param {Array} customColumns - (Optional) Additional column names beyond standard set
 * 
 * @returns {Sheet} The log sheet object
 * 
 * Standard Columns (always included):
 *   - Timestamp: Date/time of the action
 *   - User: Email of user who performed the action
 *   - Action: Type of action performed
 *   - Details: Description of what happened
 *   - Status: Success/failure status
 *   - Source: Script/function that generated the log
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var logSheet = initializeLogSheet(SpreadsheetApp.getActiveSpreadsheet(), "System Logs");
 * // Sheet is ready for logging with proper headers
 * ```
 * 
 * NOTES:
 * - Creates new sheet if it doesn't exist
 * - If sheet exists but has no headers, adds standard headers
 * - Freezes first row for header visibility
 * - Sets up date formatting on Timestamp column
 * ==============================================================================
 */
function initializeLogSheet(spreadsheet, sheetName, customColumns) {
  sheetName = sheetName || "Activity Log";
  
  var standardColumns = ["Timestamp", "User", "Action", "Details", "Status", "Source"];
  var allColumns = customColumns 
    ? standardColumns.concat(customColumns) 
    : standardColumns;
  
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // Create sheet if doesn't exist
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  // Check if headers exist
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    // Add headers
    sheet.getRange(1, 1, 1, allColumns.length).setValues([allColumns]);
    
    // Format header row
    var headerRange = sheet.getRange(1, 1, 1, allColumns.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4285f4");
    headerRange.setFontColor("#ffffff");
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths
    sheet.setColumnWidth(1, 150); // Timestamp
    sheet.setColumnWidth(2, 200); // User
    sheet.setColumnWidth(3, 150); // Action
    sheet.setColumnWidth(4, 400); // Details
    sheet.setColumnWidth(5, 100); // Status
    sheet.setColumnWidth(6, 150); // Source
  }
  
  return sheet;
}

/**
 * ==============================================================================
 * FUNCTION: logActivity
 * ==============================================================================
 * 
 * PURPOSE:
 * Logs a single activity/event to the specified log sheet.
 * Primary logging function for audit trails.
 * 
 * PARAMETERS:
 * @param {Sheet} logSheet - Sheet to log to (from initializeLogSheet)
 * @param {string} action - Short action name/category (e.g., "FILE_SCAN", "EMAIL_SENT")
 * @param {string} details - Detailed description of the activity
 * @param {Object} options - (Optional) Additional log data:
 *   - status: {string} - "SUCCESS", "FAILED", "WARNING", "INFO" (default: "INFO")
 *   - user: {string} - Override user email (default: current user)
 *   - source: {string} - Function/script name (default: auto-detected)
 *   - timestamp: {Date} - Override timestamp (default: now)
 *   - customFields: {Array} - Additional values for custom columns
 * 
 * @returns {number} Row number where log was added, or -1 on error
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var logSheet = initializeLogSheet(ss, "System Logs");
 * 
 * logActivity(logSheet, "FILE_PROCESSED", "Renamed document to '2026-01-15 Contract.pdf'", {
 *   status: "SUCCESS",
 *   source: "processFiles"
 * });
 * 
 * logActivity(logSheet, "EMAIL_SENT", "Reminder sent to client@example.com", {
 *   status: "SUCCESS"
 * });
 * ```
 * 
 * NOTES:
 * - Automatically captures current user email
 * - Truncates details if too long (prevents cell overflow)
 * - Commonly used in: file-center, CAR, litigation for audit trails
 * ==============================================================================
 */
function logActivity(logSheet, action, details, options) {
  options = options || {};
  
  try {
    var timestamp = options.timestamp || new Date();
    var user = options.user || Session.getActiveUser().getEmail() || "system";
    var status = options.status || "INFO";
    var source = options.source || detectCallerFunction() || "unknown";
    
    // Truncate details if extremely long
    var truncatedDetails = details;
    if (details && details.length > 45000) {
      truncatedDetails = details.substring(0, 45000) + "... [truncated]";
    }
    
    var row = [
      timestamp,
      user,
      action,
      truncatedDetails,
      status,
      source
    ];
    
    // Add custom fields if provided
    if (options.customFields && Array.isArray(options.customFields)) {
      row = row.concat(options.customFields);
    }
    
    logSheet.appendRow(row);
    
    return logSheet.getLastRow();
    
  } catch (error) {
    // Silent failure - logging shouldn't break main functionality
    return -1;
  }
}

/**
 * ==============================================================================
 * FUNCTION: logError
 * ==============================================================================
 * 
 * PURPOSE:
 * Specialized logging function for errors with full error details.
 * 
 * PARAMETERS:
 * @param {Sheet} logSheet - Sheet to log to
 * @param {string} context - Where the error occurred
 * @param {Error} error - The error object caught
 * @param {Object} additionalInfo - (Optional) Extra context about the error
 * 
 * @returns {number} Row number where log was added
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * try {
 *   riskyOperation();
 * } catch (error) {
 *   logError(logSheet, "Drive Scan", error, {
 *     folderId: folderId,
 *     retryCount: retryCount
 *   });
 * }
 * ```
 * ==============================================================================
 */
function logError(logSheet, context, error, additionalInfo) {
  var details = "Error: " + error.toString();
  
  if (error.stack) {
    details += "\nStack: " + error.stack;
  }
  
  if (additionalInfo) {
    details += "\nContext: " + JSON.stringify(additionalInfo);
  }
  
  return logActivity(logSheet, "ERROR_" + context.toUpperCase().replace(/\s+/g, "_"), details, {
    status: "FAILED",
    source: context
  });
}

/**
 * ==============================================================================
 * FUNCTION: trimLogSheet
 * ==============================================================================
 * 
 * PURPOSE:
 * Trims old log entries to prevent sheet from growing too large.
 * Keeps most recent entries and archives or deletes older ones.
 * 
 * PARAMETERS:
 * @param {Sheet} logSheet - Log sheet to trim
 * @param {number} keepRows - (Optional) Number of recent rows to keep (default: 1000)
 * @param {boolean} archive - (Optional) Copy deleted rows to archive sheet (default: false)
 * @param {string} archiveSheetName - (Optional) Name of archive sheet
 * 
 * @returns {Object} Trim result:
 *   - trimmed: {number} - Rows removed
 *   - remaining: {number} - Rows kept
 *   - archived: {number} - Rows archived (if archive=true)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = trimLogSheet(logSheet, 500, true, "Old Logs");
 * Logger.log("Trimmed " + result.trimmed + " old log entries");
 * ```
 * 
 * NOTES:
 * - Preserves header row
 * - Archive sheet is created if it doesn't exist
 * - Run periodically via trigger to prevent bloat
 * ==============================================================================
 */
function trimLogSheet(logSheet, keepRows, archive, archiveSheetName) {
  keepRows = keepRows || 1000;
  archive = archive || false;
  archiveSheetName = archiveSheetName || "Log Archive";
  
  var result = {
    trimmed: 0,
    remaining: 0,
    archived: 0
  };
  
  try {
    var lastRow = logSheet.getLastRow();
    
    // Nothing to trim if at or below limit
    if (lastRow <= keepRows + 1) { // +1 for header
      result.remaining = lastRow;
      return result;
    }
    
    var rowsToTrim = lastRow - keepRows - 1; // -1 for header
    var startRow = 2; // First data row (after header)
    var endRow = startRow + rowsToTrim - 1;
    
    // Archive if requested
    if (archive) {
      var spreadsheet = logSheet.getParent();
      var archiveSheet = spreadsheet.getSheetByName(archiveSheetName);
      
      if (!archiveSheet) {
        archiveSheet = initializeLogSheet(spreadsheet, archiveSheetName);
      }
      
      var rangeToArchive = logSheet.getRange(startRow, 1, rowsToTrim, logSheet.getLastColumn());
      var dataToArchive = rangeToArchive.getValues();
      
      // Append to archive
      dataToArchive.forEach(function(row) {
        archiveSheet.appendRow(row);
      });
      
      result.archived = rowsToTrim;
    }
    
    // Delete the old rows
    logSheet.deleteRows(startRow, rowsToTrim);
    
    result.trimmed = rowsToTrim;
    result.remaining = logSheet.getLastRow();
    
  } catch (error) {
    // Silent failure
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: getLogSummary
 * ==============================================================================
 * 
 * PURPOSE:
 * Generates a summary of log activity for reporting or status checks.
 * 
 * PARAMETERS:
 * @param {Sheet} logSheet - Log sheet to analyze
 * @param {Object} options - (Optional) Filter options:
 *   - since: {Date} - Only entries after this date
 *   - actions: {Array} - Filter to specific action types
 *   - status: {string} - Filter to specific status
 *   - user: {string} - Filter to specific user
 * 
 * @returns {Object} Summary statistics:
 *   - totalEntries: {number} - Total log entries matching filters
 *   - byAction: {Object} - Count per action type
 *   - byStatus: {Object} - Count per status
 *   - byUser: {Object} - Count per user
 *   - dateRange: {Object} - First and last entry dates
 *   - recentErrors: {Array} - Recent error entries
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var summary = getLogSummary(logSheet, {
 *   since: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000) // Last 7 days
 * });
 * 
 * Logger.log("Errors this week: " + (summary.byStatus.FAILED || 0));
 * ```
 * ==============================================================================
 */
function getLogSummary(logSheet, options) {
  options = options || {};
  
  var summary = {
    totalEntries: 0,
    byAction: {},
    byStatus: {},
    byUser: {},
    dateRange: { first: null, last: null },
    recentErrors: []
  };
  
  try {
    var data = logSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find column indices
    var timestampCol = headers.indexOf("Timestamp");
    var userCol = headers.indexOf("User");
    var actionCol = headers.indexOf("Action");
    var statusCol = headers.indexOf("Status");
    var detailsCol = headers.indexOf("Details");
    
    // Start from row 1 (skip header)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Date filter
      if (options.since && timestampCol >= 0) {
        var entryDate = new Date(row[timestampCol]);
        if (entryDate < options.since) continue;
      }
      
      // Action filter
      if (options.actions && options.actions.length > 0 && actionCol >= 0) {
        if (options.actions.indexOf(row[actionCol]) === -1) continue;
      }
      
      // Status filter
      if (options.status && statusCol >= 0) {
        if (row[statusCol] !== options.status) continue;
      }
      
      // User filter
      if (options.user && userCol >= 0) {
        if (row[userCol] !== options.user) continue;
      }
      
      // Count this entry
      summary.totalEntries++;
      
      // Count by action
      var action = actionCol >= 0 ? row[actionCol] : "unknown";
      summary.byAction[action] = (summary.byAction[action] || 0) + 1;
      
      // Count by status
      var status = statusCol >= 0 ? row[statusCol] : "unknown";
      summary.byStatus[status] = (summary.byStatus[status] || 0) + 1;
      
      // Count by user
      var user = userCol >= 0 ? row[userCol] : "unknown";
      summary.byUser[user] = (summary.byUser[user] || 0) + 1;
      
      // Track date range
      if (timestampCol >= 0) {
        var ts = new Date(row[timestampCol]);
        if (!summary.dateRange.first || ts < summary.dateRange.first) {
          summary.dateRange.first = ts;
        }
        if (!summary.dateRange.last || ts > summary.dateRange.last) {
          summary.dateRange.last = ts;
        }
      }
      
      // Collect recent errors
      if (status === "FAILED" && summary.recentErrors.length < 10) {
        summary.recentErrors.push({
          timestamp: timestampCol >= 0 ? row[timestampCol] : null,
          action: action,
          details: detailsCol >= 0 ? row[detailsCol] : null
        });
      }
    }
    
  } catch (error) {
    // Return empty summary on error
  }
  
  return summary;
}

/**
 * ==============================================================================
 * HELPER FUNCTION: detectCallerFunction
 * ==============================================================================
 * 
 * Attempts to detect the calling function name for automatic source attribution.
 * @private
 */
function detectCallerFunction() {
  try {
    throw new Error();
  } catch (e) {
    var stack = e.stack;
    var match = stack.match(/at (\S+)/g);
    if (match && match.length >= 3) {
      // match[0] = detectCallerFunction, match[1] = logActivity, match[2] = actual caller
      var caller = match[2].replace("at ", "");
      return caller.split(".")[0]; // Remove object prefix if present
    }
  }
  return null;
}
