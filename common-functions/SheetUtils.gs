/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * SheetUtils.gs - Common Spreadsheet Operations
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * PURPOSE:
 * Reusable functions for Google Sheets operations including data validation,
 * row operations, header management, and sheet formatting.
 * 
 * PREREQUISITES:
 * - Standard SpreadsheetApp access (no special OAuth required)
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
 * FUNCTION: detectHeaderRow
 * ==============================================================================
 * 
 * PURPOSE:
 * Automatically detects which row contains column headers in a sheet.
 * Useful when sheets have title rows or merged cells before actual data headers.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Google Sheet object to analyze
 * @param {Array} expectedHeaders - (Optional) Array of expected header names to match
 * @param {number} maxRows - (Optional) Maximum rows to scan (default: 10)
 * 
 * @returns {Object} Detection result:
 *   - found: {boolean} - Whether header row was detected
 *   - row: {number} - Row number (1-indexed) containing headers
 *   - headers: {Array} - Array of detected header names
 *   - matchScore: {number} - Percentage of expected headers found (if expectedHeaders provided)
 *   - confidence: {string} - "high", "medium", "low" based on detection certainty
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 * var result = detectHeaderRow(sheet, ["Name", "Email", "Status"], 5);
 * 
 * if (result.found && result.confidence === "high") {
 *   Logger.log("Headers found at row " + result.row + ": " + result.headers.join(", "));
 * }
 * ```
 * 
 * NOTES:
 * - Searches for row with most non-empty cells that look like headers
 * - Header cells should contain text (not numbers/dates)
 * - Skips rows that are entirely merged or have very few cells
 * - Commonly used in: CAR, automated-expiry-notification for dynamic setup
 * ==============================================================================
 */
function detectHeaderRow(sheet, expectedHeaders, maxRows) {
  maxRows = maxRows || 10;
  expectedHeaders = expectedHeaders || [];
  
  var result = {
    found: false,
    row: 1,
    headers: [],
    matchScore: 0,
    confidence: "low"
  };
  
  try {
    var dataRange = sheet.getRange(1, 1, Math.min(maxRows, sheet.getLastRow()), sheet.getLastColumn());
    var values = dataRange.getValues();
    
    var bestRow = 0;
    var bestScore = 0;
    var bestHeaders = [];
    
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var nonEmptyCount = 0;
      var headerLikeCount = 0;
      var matchedExpected = 0;
      var rowHeaders = [];
      
      for (var j = 0; j < row.length; j++) {
        var cell = row[j];
        var cellText = cell ? cell.toString().trim() : "";
        
        if (cellText) {
          nonEmptyCount++;
          rowHeaders.push(cellText);
          
          // Check if looks like a header (text, reasonable length, contains letters)
          if (/[a-zA-Z]/.test(cellText) && cellText.length > 1 && cellText.length < 100) {
            headerLikeCount++;
          }
          
          // Check against expected headers
          if (expectedHeaders.length > 0) {
            var cellLower = cellText.toLowerCase();
            for (var k = 0; k < expectedHeaders.length; k++) {
              if (cellLower.indexOf(expectedHeaders[k].toLowerCase()) !== -1 ||
                  expectedHeaders[k].toLowerCase().indexOf(cellLower) !== -1) {
                matchedExpected++;
                break;
              }
            }
          }
        }
      }
      
      // Score this row
      var score = headerLikeCount;
      if (expectedHeaders.length > 0 && matchedExpected > 0) {
        score += matchedExpected * 2; // Bonus for matching expected headers
      }
      
      // Prefer rows that aren't the very first row (title row pattern)
      if (i === 0 && nonEmptyCount < 3) {
        score -= 1;
      }
      
      if (score > bestScore) {
        bestScore = score;
        bestRow = i;
        bestHeaders = rowHeaders;
      }
    }
    
    if (bestScore > 0) {
      result.found = true;
      result.row = bestRow + 1; // Convert to 1-indexed
      result.headers = bestHeaders;
      
      // Calculate match score
      if (expectedHeaders.length > 0) {
        var foundExpected = 0;
        var headersLower = bestHeaders.map(function(h) { return h.toLowerCase(); });
        expectedHeaders.forEach(function(expected) {
          if (headersLower.some(function(h) { return h.indexOf(expected.toLowerCase()) !== -1; })) {
            foundExpected++;
          }
        });
        result.matchScore = Math.round((foundExpected / expectedHeaders.length) * 100);
      }
      
      // Determine confidence
      if (expectedHeaders.length > 0 && result.matchScore >= 80) {
        result.confidence = "high";
      } else if (bestScore >= 3 || result.matchScore >= 50) {
        result.confidence = "medium";
      } else {
        result.confidence = "low";
      }
    }
    
  } catch (error) {
    result.found = false;
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: validateRequiredHeaders
 * ==============================================================================
 * 
 * PURPOSE:
 * Checks if a sheet contains all required column headers.
 * Reports missing columns and suggests alternatives.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to validate
 * @param {Array} requiredHeaders - Array of required header names
 * @param {number} headerRow - (Optional) Row containing headers (default: 1)
 * @param {boolean} caseSensitive - (Optional) Case-sensitive matching (default: false)
 * 
 * @returns {Object} Validation result:
 *   - valid: {boolean} - All required headers present
 *   - present: {Array} - Headers that were found
 *   - missing: {Array} - Headers not found
 *   - suggestions: {Object} - Map of missing header -> suggested alternative found
 *   - headerMap: {Object} - Map of header name -> column index
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var required = ["Name", "Email", "Expiry Date", "Status"];
 * var result = validateRequiredHeaders(sheet, required, 1);
 * 
 * if (!result.valid) {
 *   Logger.log("Missing: " + result.missing.join(", "));
 *   Logger.log("Did you mean: " + result.suggestions["Expiry Date"] + "?");
 * }
 * ```
 * 
 * NOTES:
 * - Returns column indices (1-indexed) in headerMap
 * - Suggests alternatives using fuzzy matching for typos
 * - Commonly used in: automated-expiry-notification, CAR setup validation
 * ==============================================================================
 */
function validateRequiredHeaders(sheet, requiredHeaders, headerRow, caseSensitive) {
  headerRow = headerRow || 1;
  caseSensitive = caseSensitive || false;
  
  var result = {
    valid: false,
    present: [],
    missing: [],
    suggestions: {},
    headerMap: {}
  };
  
  try {
    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      result.missing = requiredHeaders;
      return result;
    }
    
    var headerRange = sheet.getRange(headerRow, 1, 1, lastCol);
    var actualHeaders = headerRange.getValues()[0];
    
    // Normalize headers
    var normalizedActual = actualHeaders.map(function(h) {
      return caseSensitive ? h.toString().trim() : h.toString().trim().toLowerCase();
    });
    
    requiredHeaders.forEach(function(required, index) {
      var normalizedRequired = caseSensitive ? required.trim() : required.trim().toLowerCase();
      var colIndex = normalizedActual.indexOf(normalizedRequired);
      
      if (colIndex !== -1) {
        result.present.push(required);
        result.headerMap[required] = colIndex + 1; // 1-indexed
      } else {
        result.missing.push(required);
        
        // Look for similar header (simple substring matching)
        for (var i = 0; i < normalizedActual.length; i++) {
          var actual = normalizedActual[i];
          if (actual.indexOf(normalizedRequired) !== -1 ||
              normalizedRequired.indexOf(actual) !== -1) {
            result.suggestions[required] = actualHeaders[i];
            break;
          }
        }
      }
    });
    
    result.valid = result.missing.length === 0;
    
  } catch (error) {
    result.missing = requiredHeaders;
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: getColumnIndexByHeader
 * ==============================================================================
 * 
 * PURPOSE:
 * Finds the column index (1-indexed) for a specific header name.
 * Returns -1 if not found.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to search
 * @param {string} headerName - Header text to find
 * @param {number} headerRow - (Optional) Header row number (default: 1)
 * @param {boolean} caseSensitive - (Optional) Case-sensitive match (default: false)
 * 
 * @returns {number} Column index (1-indexed) or -1 if not found
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var col = getColumnIndexByHeader(sheet, "Email Address", 1, false);
 * if (col > 0) {
 *   var emailCell = sheet.getRange(row, col);
 * }
 * ```
 * ==============================================================================
 */
function getColumnIndexByHeader(sheet, headerName, headerRow, caseSensitive) {
  headerRow = headerRow || 1;
  caseSensitive = caseSensitive || false;
  
  try {
    var headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    var searchName = caseSensitive ? headerName : headerName.toLowerCase();
    
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i].toString().trim();
      if (!caseSensitive) header = header.toLowerCase();
      
      if (header === searchName) {
        return i + 1; // Convert to 1-indexed
      }
    }
  } catch (e) {
    return -1;
  }
  
  return -1;
}

/**
 * ==============================================================================
 * FUNCTION: findRowByValue
 * ==============================================================================
 * 
 * PURPOSE:
 * Finds the first row containing a specific value in a given column.
 * Useful for looking up records by ID or unique identifier.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to search
 * @param {number} column - Column index (1-indexed) to search
 * @param {*} value - Value to search for
 * @param {number} startRow - (Optional) First row to check (default: 2, assuming row 1 is headers)
 * 
 * @returns {number} Row index (1-indexed) or -1 if not found
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var row = findRowByValue(sheet, 1, "DOC-2026-001", 2);
 * if (row > 0) {
 *   // Found record at this row
 *   var record = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
 * }
 * ```
 * 
 * NOTES:
 * - Performs exact value match (===)
 * - Stops at first match (use findAllRowsByValue for multiple matches)
 * ==============================================================================
 */
function findRowByValue(sheet, column, value, startRow) {
  startRow = startRow || 2;
  
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < startRow) return -1;
    
    var columnData = sheet.getRange(startRow, column, lastRow - startRow + 1, 1).getValues();
    
    for (var i = 0; i < columnData.length; i++) {
      if (columnData[i][0] === value) {
        return startRow + i; // Convert to 1-indexed
      }
    }
  } catch (e) {
    return -1;
  }
  
  return -1;
}

/**
 * ==============================================================================
 * FUNCTION: findAllRowsByValue
 * ==============================================================================
 * 
 * PURPOSE:
 * Finds all rows containing a specific value in a given column.
 * Returns array of row indices.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to search
 * @param {number} column - Column index (1-indexed) to search
 * @param {*} value - Value to search for
 * @param {number} startRow - (Optional) First row to check (default: 2)
 * 
 * @returns {Array} Array of row indices (1-indexed), empty if none found
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var rows = findAllRowsByValue(sheet, 5, "PENDING", 2);
 * rows.forEach(function(row) {
 *   // Process each pending record
 * });
 * ```
 * ==============================================================================
 */
function findAllRowsByValue(sheet, column, value, startRow) {
  startRow = startRow || 2;
  var rows = [];
  
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < startRow) return rows;
    
    var columnData = sheet.getRange(startRow, column, lastRow - startRow + 1, 1).getValues();
    
    for (var i = 0; i < columnData.length; i++) {
      if (columnData[i][0] === value) {
        rows.push(startRow + i);
      }
    }
  } catch (e) {
    // Return empty array on error
  }
  
  return rows;
}

/**
 * ==============================================================================
 * FUNCTION: appendRowWithTimestamp
 * ==============================================================================
 * 
 * PURPOSE:
 * Appends a row to a sheet with automatic timestamp in first column.
 * Commonly used for logging and audit trails.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to append to
 * @param {Array} rowData - Array of values to append (without timestamp)
 * @param {string} timezone - (Optional) Timezone for timestamp (default: script timezone)
 * @param {string} dateFormat - (Optional) Date format (default: "yyyy-MM-dd HH:mm:ss")
 * 
 * @returns {number} Row number of appended row, or -1 on error
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var rowNum = appendRowWithTimestamp(logSheet, ["Action", "Details", "User"], "Asia/Manila");
 * Logger.log("Logged to row " + rowNum);
 * ```
 * ==============================================================================
 */
function appendRowWithTimestamp(sheet, rowData, timezone, dateFormat) {
  timezone = timezone || Session.getScriptTimeZone();
  dateFormat = dateFormat || "yyyy-MM-dd HH:mm:ss";
  
  try {
    var timestamp = Utilities.formatDate(new Date(), timezone, dateFormat);
    var fullRow = [timestamp].concat(rowData || []);
    
    sheet.appendRow(fullRow);
    return sheet.getLastRow();
  } catch (e) {
    return -1;
  }
}

/**
 * ==============================================================================
 * FUNCTION: clearSheetKeepHeaders
 * ==============================================================================
 * 
 * PURPOSE:
 * Clears all data rows from a sheet while preserving the header row.
 * Useful for resetting data while keeping column structure.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to clear
 * @param {number} headerRow - (Optional) Number of header rows to preserve (default: 1)
 * @param {boolean} clearFormats - (Optional) Also clear formatting (default: false)
 * 
 * @returns {boolean} True if successful, false otherwise
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var success = clearSheetKeepHeaders(dataSheet, 1, false);
 * if (success) {
 *   Logger.log("Sheet cleared, headers preserved");
 * }
 * ```
 * 
 * NOTES:
 * - Preserves header row by default
 * - Only clears data, optionally clears formatting
 * - More efficient than deleting rows for large sheets
 * ==============================================================================
 */
function clearSheetKeepHeaders(sheet, headerRow, clearFormats) {
  headerRow = headerRow || 1;
  clearFormats = clearFormats || false;
  
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow <= headerRow) {
      return true; // Nothing to clear
    }
    
    var dataRange = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol);
    dataRange.clearContent();
    
    if (clearFormats) {
      dataRange.clearFormat();
    }
    
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * ==============================================================================
 * FUNCTION: applyDropdownValidation
 * ==============================================================================
 * 
 * PURPOSE:
 * Applies data validation (dropdown list) to a column or range.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to apply validation
 * @param {number} column - Column index (1-indexed)
 * @param {Array} options - Array of dropdown options
 * @param {string} startRow - (Optional) First row to apply (default: 2)
 * @param {boolean} allowInvalid - (Optional) Allow values not in list (default: false)
 * @param {string} helpText - (Optional) Help text to show
 * 
 * @returns {boolean} True if applied successfully
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var statusOptions = ["TODO", "IN PROGRESS", "DONE", "CANCELLED"];
 * applyDropdownValidation(sheet, 5, statusOptions, 2, false, "Select status");
 * ```
 * 
 * NOTES:
 * - Applies to all rows from startRow to end of sheet
 * - Shows warning if invalid value entered (unless allowInvalid=true)
 * - Commonly used in: CAR, automated-expiry-notification for status columns
 * ==============================================================================
 */
function applyDropdownValidation(sheet, column, options, startRow, allowInvalid, helpText) {
  startRow = startRow || 2;
  allowInvalid = allowInvalid !== undefined ? allowInvalid : false;
  
  try {
    var lastRow = Math.max(sheet.getMaxRows(), 1000); // Apply to many rows
    
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(allowInvalid)
      .setHelpText(helpText || "Please select from the list")
      .build();
    
    var range = sheet.getRange(startRow, column, lastRow - startRow + 1, 1);
    range.setDataValidation(rule);
    
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * ==============================================================================
 * FUNCTION: autoResizeColumns
 * ==============================================================================
 * 
 * PURPOSE:
 * Automatically resizes all columns to fit their content.
 * Optionally sets minimum and maximum widths.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to resize
 * @param {Object} limits - (Optional) Width limits:
 *   - minWidth: {number} - Minimum column width in pixels
 *   - maxWidth: {number} - Maximum column width in pixels
 * @param {Array} excludeColumns - (Optional) Column indices to skip
 * 
 * @returns {boolean} True if successful
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * autoResizeColumns(sheet, { minWidth: 100, maxWidth: 400 }, [1]); // Skip first column
 * ```
 * ==============================================================================
 */
function autoResizeColumns(sheet, limits, excludeColumns) {
  limits = limits || {};
  excludeColumns = excludeColumns || [];
  
  try {
    var lastCol = sheet.getLastColumn();
    
    for (var i = 1; i <= lastCol; i++) {
      if (excludeColumns.indexOf(i) !== -1) continue;
      
      sheet.autoResizeColumn(i);
      
      // Apply limits
      var width = sheet.getColumnWidth(i);
      if (limits.minWidth && width < limits.minWidth) {
        sheet.setColumnWidth(i, limits.minWidth);
      }
      if (limits.maxWidth && width > limits.maxWidth) {
        sheet.setColumnWidth(i, limits.maxWidth);
      }
    }
    
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * ==============================================================================
 * FUNCTION: getDataAsObjects
 * ==============================================================================
 * 
 * PURPOSE:
 * Converts sheet data to an array of objects using headers as keys.
 * Makes data easier to work with than raw 2D arrays.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to read
 * @param {number} headerRow - (Optional) Header row number (default: 1)
 * @param {number} dataStartRow - (Optional) First data row (default: headerRow + 1)
 * 
 * @returns {Array} Array of objects where keys = header names, values = cell values
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var data = getDataAsObjects(sheet, 1, 2);
 * data.forEach(function(row) {
 *   Logger.log(row.Name + " - " + row.Status);
 * });
 * ```
 * 
 * NOTES:
 * - Header names become object keys (spaces and special chars preserved)
 * - Returns empty array if no data rows
 * - More memory-intensive than raw arrays (use for moderate data sizes)
 * ==============================================================================
 */
function getDataAsObjects(sheet, headerRow, dataStartRow) {
  headerRow = headerRow || 1;
  dataStartRow = dataStartRow || (headerRow + 1);
  
  var result = [];
  
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow < dataStartRow || lastCol === 0) {
      return result;
    }
    
    // Get headers
    var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    
    // Get data
    var dataRange = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol);
    var data = dataRange.getValues();
    
    // Convert to objects
    for (var i = 0; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      result.push(row);
    }
    
  } catch (e) {
    // Return empty array on error
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: setRowBackgroundColor
 * ==============================================================================
 * 
 * PURPOSE:
 * Sets the background color of a row based on conditions.
 * Useful for visual status indicators.
 * 
 * PARAMETERS:
 * @param {Sheet} sheet - Sheet to format
 * @param {number} row - Row index (1-indexed)
 * @param {string} color - Hex color code (e.g., "#ff0000") or color name
 * @param {number} startCol - (Optional) First column to color (default: 1)
 * @param {number} endCol - (Optional) Last column to color (default: last column with data)
 * 
 * @returns {boolean} True if successful
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * setRowBackgroundColor(sheet, 5, "#ffcccc", 1, 10); // Light red
 * setRowBackgroundColor(sheet, 6, "#ccffcc", 1, 10); // Light green
 * ```
 * 
 * NOTES:
 * - Common color indicators:
 *   - Green (#d9ead3): Complete/Success
 *   - Yellow (#fff2cc): Warning/Pending
 *   - Red (#f4cccc): Error/Overdue
 *   - Blue (#cfe2f3): Info/Processing
 * ==============================================================================
 */
function setRowBackgroundColor(sheet, row, color, startCol, endCol) {
  startCol = startCol || 1;
  endCol = endCol || sheet.getLastColumn();
  
  try {
    var range = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
    range.setBackground(color);
    return true;
  } catch (e) {
    return false;
  }
}
