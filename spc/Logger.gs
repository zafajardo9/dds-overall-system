/**
 * Logging Manager
 * Version: 3.0 (Production Ready) - Fixed for dynamic CONFIG
 */

const Logger = {
  /**
   * Get config dynamically
   */
  getConfig() {
    const scriptProps = PropertiesService.getScriptProperties();
    return {
      LOG_SHEET_NAME: "Processing Log",
    };
  },
  setupLogSheet(spreadsheet) {
    const CONFIG = this.getConfig();
    let logSheet = spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME);

    if (!logSheet) {
      logSheet = spreadsheet.insertSheet(CONFIG.LOG_SHEET_NAME);

      const headers = ["Timestamp", "User", "Action", "Message", "Level"];
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      logSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      logSheet.setFrozenRows(1);

      logSheet.setColumnWidth(1, 150);
      logSheet.setColumnWidth(2, 200);
      logSheet.setColumnWidth(3, 150);
      logSheet.setColumnWidth(4, 400);
      logSheet.setColumnWidth(5, 80);
    }

    return logSheet;
  },
  log(message, level = "INFO") {
    try {
      const CONFIG = this.getConfig();
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = this.setupLogSheet(spreadsheet);

      const timestamp = new Date();
      const user = this.getCurrentUser();
      const action = "File Processing";

      const newRow = [timestamp, user, action, message, level];
      logSheet.appendRow(newRow);

      const lastRow = logSheet.getLastRow();

      if (level === "ERROR") {
        logSheet.getRange(lastRow, 1, 1, 5).setBackground("#f4c7c3");
      } else if (level === "WARNING") {
        logSheet.getRange(lastRow, 1, 1, 5).setBackground("#fff4c3");
      } else if (level === "SUCCESS") {
        logSheet.getRange(lastRow, 1, 1, 5).setBackground("#d9ead3");
      }

      console.log(`[${level}] ${message}`);
    } catch (error) {
      console.error("Failed to write to log sheet:", error);
    }
  },
  getCurrentUser() {
    try {
      return Session.getActiveUser().getEmail() || "Unknown User";
    } catch (error) {
      return "System";
    }
  },
};
