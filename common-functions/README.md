# Common Functions Library

A reusable library of Google Apps Script functions shared across DDS projects.

---

## Overview

This folder contains modular utility files that provide common functionality across all DDS Apps Script projects. Copy the files you need into your project to get battle-tested, documented functions.

---

## Available Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| **GeminiAPI.gs** | Google Gemini AI integration | `fetchGeminiModels()`, `setGeminiApiKey()`, `callGemini()`, `callGeminiWithFile()`, `extractDocumentTitle()`, `testGeminiConnection()` |
| **DriveUtils.gs** | Google Drive operations | `getFolderByIdOrUrl()`, `listFilesInFolder()`, `ensureDateSubfolder()`, `moveFileToFolder()`, `duplicateFileDetection()` |
| **EmailUtils.gs** | Email sending & tracking | `sendHtmlEmail()`, `sendBatchEmails()`, `checkEmailReplies()`, `createEmailTemplate()`, `getEmailQuotaInfo()` |
| **SheetUtils.gs** | Spreadsheet helpers | `detectHeaderRow()`, `validateRequiredHeaders()`, `findRowByValue()`, `applyDropdownValidation()`, `getDataAsObjects()` |
| **ConfigUtils.gs** | Property management | `getConfigValue()`, `setConfigValue()`, `getUserConfig()`, `getAllConfig()`, `clearAllConfig()` |
| **DateUtils.gs** | Date calculations | `addWorkingDays()`, `getDaysRemaining()`, `formatDateForDisplay()`, `isOverdue()`, `getDeadlineColor()` |
| **UIUtils.gs** | User interface helpers | `createStandardMenu()`, `showAlert()`, `promptForInput()`, `createSelectionDialog()`, `showStatusPanel()` |
| **LoggingUtils.gs** | Activity logging | `initializeLogSheet()`, `logActivity()`, `logError()`, `trimLogSheet()`, `getLogSummary()` |

---

## How to Use

### 1. Copy Required Files

Copy the `.gs` files you need into your Apps Script project alongside your main script.

### 2. Call Functions Directly

All functions are designed to work independently:

```javascript
// Example: Using Gemini API
setGeminiApiKey("your-api-key");
var result = callGemini("Summarize this document", { temperature: 0.3 });

// Example: Using Drive utilities
var folder = getFolderByIdOrUrl("https://drive.google.com/drive/folders/1Bx...");
var files = listFilesInFolder(folder, { mimeTypes: ["application/pdf"] });

// Example: Using Sheet utilities
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var headers = detectHeaderRow(sheet, ["Name", "Email", "Status"]);
```

### 3. Function Documentation

Every function includes detailed JSDoc comments explaining:
- **Purpose**: What the function does
- **Parameters**: What inputs it accepts (with types and descriptions)
- **Returns**: What it returns (with structure details)
- **Usage Example**: Copy-paste ready example code
- **Notes**: Important implementation details

---

## Prerequisites by Module

| Module | Required Scopes | Notes |
|--------|-----------------|-------|
| GeminiAPI.gs | `generative-language` | Requires Gemini API key from AI Studio |
| DriveUtils.gs | `drive` or `drive.readonly` | For folder/file operations |
| EmailUtils.gs | `gmail.send` or `gmail.compose` | For sending emails |
| SheetUtils.gs | `spreadsheets` | Standard SpreadsheetApp access |
| ConfigUtils.gs | None | Uses PropertiesService (built-in) |
| DateUtils.gs | None | Pure JavaScript functions |
| UIUtils.gs | None | Uses SpreadsheetApp.getUi() |
| LoggingUtils.gs | None | Uses SpreadsheetApp |

---

## Quick Reference: Most Common Functions

### Setup & Configuration
```javascript
// Store API key securely
setGeminiApiKey("your-key");

// Store configuration
setConfigValue("DRIVE_FOLDER_ID", "1Bx123...", "string");
var folderId = getConfigValue("DRIVE_FOLDER_ID", null, "string");
```

### AI Integration
```javascript
// Get available models
var models = fetchGeminiModels();

// Extract title from document
var result = extractDocumentTitle(file, { prefixDate: true });
if (result.success) {
  file.setName(result.title);
}
```

### File Operations
```javascript
// Scan folder
var result = listFilesInFolder(folderId, { maxResults: 50 });

// Organize by date
var subfolder = ensureDateSubfolder(parentFolder, new Date(), "YYYY/MMMM/DD MMMM");
file.moveTo(subfolder.folder);
```

### Email
```javascript
// Send notification
sendHtmlEmail({
  to: "client@example.com",
  subject: "Document Ready",
  htmlBody: "<h1>Your document is ready</h1>",
  fromName: "DDS System"
});

// Check for replies
var replies = checkEmailReplies({
  searchQuery: "subject:Reminder",
  keywords: ["acknowledged", "received"]
});
```

### Date Calculations
```javascript
// Calculate deadline
var dueDate = addWorkingDays(new Date(), 10);
var daysLeft = getDaysRemaining(dueDate);

// Check status
if (isOverdue(dueDate)) {
  var color = getDeadlineColor(daysLeft); // Returns hex color
}
```

### Logging
```javascript
// Initialize log sheet
var logSheet = initializeLogSheet(spreadsheet, "System Logs");

// Log activity
logActivity(logSheet, "FILE_PROCESSED", "Processed contract.pdf", {
  status: "SUCCESS"
});

// Log errors
try {
  riskyOperation();
} catch (error) {
  logError(logSheet, "Processing", error, { fileId: id });
}
```

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | April 2026 | Initial release with 8 utility modules |

---

## Contributing

When adding new functions:
1. Follow the existing JSDoc documentation style
2. Include usage examples in comments
3. Add error handling
4. Update this README with new functions
