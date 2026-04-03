# Apps Script Function Similarities Reference

A quick-reference guide showing common function patterns across DDS Google Apps Script projects.

---

## 1. Entry Point Functions

| Pattern | Purpose | Found In |
|---------|---------|----------|
| `onOpen()` | Creates custom menu on spreadsheet open | **All projects**: CAR, file-center (both), litigation, notary (all), spc, automated-expiry-notification, automatic-listing-document, amlc-automation, marketing-seo-website, gsheet-jira |
| `onEdit(e)` | Handles cell edit events | file-center, litigation, marketing-seo-website |

### Menu Creation Pattern
```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Menu Name")
    .addItem("Label", "functionName")
    .addToUi();
}
```

**Projects using this pattern:**
- `CAR/code.gs` — CAR Monitoring menu
- `file-center/code-dds.gs` & `code-filepino.gs` — DDS/FilePino File Center menu
- `litigation/Code-Client-Document-Management.gs` — Client Document Monitor menu
- `notary/Code.gs` (all variants) — Notary menu
- `spc/Code.gs` — Liquidation System menu
- `automated-expiry-notification/code.gs` — Expiry Notifications menu
- `automatic-listing-document/code.gs` — Checklist Sync menu
- `amlc-automation/ExpiryAutomation.gs` — AMLC menu
- `marketing-seo-website/code.gs` — SEO Tools menu

---

## 2. Setup & Initialization Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `initializeBoard()` | Creates headers/tabs for Jira-style board | gsheet-jira |
| `initialSetup()` | First-time system configuration | litigation, CAR |
| `setupWizard()` / `runSetupWizard()` | Guided multi-step setup | automated-expiry-notification, CAR |
| `setupSheet()` / `applySheetSetup()` | Configure sheet headers/formatting | CAR, automated-expiry-notification |
| `configureDrives()` | Set Drive folder IDs | litigation, file-center |
| `linkDriveFolder()` | Link sheet to Drive folder | automatic-listing-document |

### Common Setup Steps Pattern
```javascript
// Found in: CAR, automated-expiry-notification, litigation
function runSetupWizard() {
  // Step 1: Select working sheet
  // Step 2: Configure header/data rows
  // Step 3: Set notification emails
  // Step 4: Apply formatting & validation
}
```

---

## 3. Drive Scanning & File Processing

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `scanDrive()` / `scanFolder()` | Scan Drive for files | litigation, file-center (both), notary |
| `processFiles()` | Main file processing loop | file-center, spc, notary |
| `processFile()` | Process individual file | file-center (both), spc, notary |
| `listFilesInFolder()` | Get files from Drive folder | file-center, litigation, notary |
| `moveFileToDateFolder()` | Organize files by date | file-center (both), notary |

### File Processing Pipeline Pattern
```javascript
// Found in: file-center, notary, spc
function processFiles() {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var newTitle = extractTitle(file);
    renameAndMove(file, newTitle);
    logProcessing(file, newTitle);
  }
}
```

---

## 4. AI/Gemini Integration Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `extractTitle()` / `extractDocumentTitle()` | AI title extraction from file | file-center (both), litigation, notary |
| `callGemini()` / `callGeminiWithFile()` | API call to Gemini AI | file-center, litigation, marketing-seo-website |
| `setGeminiApiKey()` | Store API key | litigation, file-center, marketing-seo-website, spc |
| `selectModel()` / `selectGeminiModel()` | Choose AI model | litigation, file-center, marketing-seo-website |
| `testGeminiIntegration()` | Verify AI connection | litigation, file-center |

### Gemini API Pattern
```javascript
// Found in: file-center, litigation, spc
function extractTitle(file) {
  var blob = file.getBlob();
  var base64 = Utilities.base64Encode(blob.getBytes());
  var payload = {
    contents: [{
      parts: [
        { text: "Extract clean document title" },
        { inline_data: { mime_type: file.getMimeType(), data: base64 } }
      ]
    }]
  };
  // Call Gemini API and parse response
}
```

---

## 5. Email & Notification Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `sendEmail()` / `sendReminderEmail()` | Send notification emails | **Most projects**: automated-expiry-notification, CAR, file-center, amlc-automation |
| `sendBatchEmail()` | Send bulk notifications | file-center, automated-expiry-notification |
| `sendAlert()` | Send error/alert emails | file-center, automated-expiry-notification |
| `checkEmailReplies()` | Scan Gmail for responses | automated-expiry-notification |
| `parseReply()` | Extract acknowledgement from reply | automated-expiry-notification |

### Email Pattern (MailApp)
```javascript
// Found in: CAR, file-center, automated-expiry-notification, amlc-automation
MailApp.sendEmail({
  to: recipient,
  cc: ccList,
  subject: "Reminder: " + taskName,
  htmlBody: htmlContent,
  name: senderName
});
```

---

## 6. Trigger & Scheduling Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `setupTrigger()` / `createDailyTrigger()` | Schedule automation | CAR, automated-expiry-notification, file-center, litigation, automatic-listing-document |
| `deleteTrigger()` / `removeTrigger()` | Remove scheduled runs | CAR, automated-expiry-notification, file-center |
| `setupTriggerForLinking()` | Link-specific trigger | automatic-listing-document |

### Trigger Setup Pattern
```javascript
// Found in: CAR, automated-expiry-notification, file-center, litigation
function setupDailyTrigger() {
  deleteExistingTriggers();
  ScriptApp.newTrigger("dailyFunctionName")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

function deleteExistingTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "targetFunction") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
```

---

## 7. Logging & Audit Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `logActivity()` | Log user/system actions | CAR, litigation, file-center |
| `writeLog()` | Write to Logs sheet | automated-expiry-notification, CAR |
| `addLogEntry()` | Add single log row | file-center, spc |
| `getLogs()` / `viewLogs()` | Retrieve log data | CAR, automated-expiry-notification |

### Logging Pattern
```javascript
// Found in: CAR, litigation, file-center
function logActivity(action, details, user) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Activity Log");
  sheet.appendRow([
    new Date(),
    user || Session.getActiveUser().getEmail(),
    action,
    details
  ]);
}
```

---

## 8. Properties/Configuration Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `saveConfig()` / `setConfig()` | Store settings | file-center, spc, automated-expiry-notification |
| `loadConfig()` / `getConfig()` | Retrieve settings | file-center, spc, litigation |
| `setProperty()` | Save to Script Properties | **All projects** |
| `getProperty()` | Read from Script Properties | **All projects** |
| `setDriveId()` | Store Drive folder ID | litigation, file-center, spc |

### Configuration Pattern
```javascript
// Found in: file-center, spc, litigation
function getConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    DRIVE_FOLDER_ID: props.getProperty("DRIVE_FOLDER_ID"),
    GEMINI_API_KEY: props.getProperty("GEMINI_API_KEY"),
    SHEET_TAB_NAME: props.getProperty("SHEET_TAB_NAME") || "Default"
  };
}
```

---

## 9. UI/Dialog Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `showSidebar()` | Display side panel | automated-expiry-notification, litigation |
| `showModal()` / `showDialog()` | Pop-up dialogs | automated-expiry-notification, gsheet-jira |
| `showAlert()` / `showMessage()` | Simple notifications | **All projects** |
| `promptForInput()` | Get user input | CAR, automated-expiry-notification |

### UI Dialog Pattern
```javascript
// Found in: CAR, automated-expiry-notification, gsheet-jira
function showStatusDialog(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert("Status", message, ui.ButtonSet.OK);
}
```

---

## 10. Sheet Data Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `getDataRange()` | Get all sheet data | **All projects** |
| `appendRow()` | Add new data row | file-center, litigation, spc, automated-expiry-notification |
| `clearSheet()` / `resetSheet()` | Clear for fresh data | litigation, gsheet-jira |
| `findRowById()` | Locate specific record | gsheet-jira, automated-expiry-notification |
| `updateRow()` / `updateStatus()` | Modify existing row | gsheet-jira, CAR, automated-expiry-notification |

### Data Update Pattern
```javascript
// Found in: gsheet-jira, CAR
function updateRowStatus(sheet, rowIndex, newStatus) {
  var statusColumn = getStatusColumn();
  sheet.getRange(rowIndex, statusColumn).setValue(newStatus);
}
```

---

## 11. Validation & Utility Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `validateHeaders()` / `verifyHeaders()` | Check column structure | automated-expiry-notification, CAR |
| `validateConfig()` | Verify settings complete | file-center, litigation |
| `formatDate()` | Consistent date formatting | file-center, spc, CAR |
| `parseDate()` | Date string to object | automated-expiry-notification, CAR |
| `isValidEmail()` | Email format check | CAR, automated-expiry-notification |

### Header Validation Pattern
```javascript
// Found in: automated-expiry-notification, CAR
function validateHeaders(sheet, requiredColumns) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var missing = requiredColumns.filter(function(col) {
    return headers.indexOf(col) === -1;
  });
  return missing.length === 0 ? true : missing;
}
```

---

## 12. Deadline/Date Calculation Functions

| Function Pattern | Purpose | Projects |
|------------------|---------|----------|
| `calculateDueDate()` | Compute deadlines | CAR, automated-expiry-notification |
| `getRemainingDays()` | Days until/overdue | CAR, automated-expiry-notification |
| `isOverdue()` / `checkOverdue()` | Overdue detection | CAR, automated-expiry-notification, amlc-automation |
| `getWorkingDays()` | Business days calc | CAR |

### Deadline Pattern (CAR Example)
```javascript
// Found in: CAR, automated-expiry-notification
function calculateDueDate(startDate, daysToAdd) {
  var date = new Date(startDate);
  date.setDate(date.getDate() + daysToAdd);
  return date;
}

function getRemainingDays(dueDate) {
  var today = new Date();
  var diff = dueDate - today;
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}
```

---

## Function Usage Matrix

| Function Category | CAR | File Center | Litigation | SPC | Auto-Expiry | Notary | AMLC | Marketing | Jira |
|-------------------|:---:|:-----------:|:----------:|:---:|:-----------:|:------:|:----:|:---------:|:----:|
| `onOpen()` / Menu | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| `setupWizard()` | ✅ | ❌ | ⚠️ | ❌ | ✅ | ❌ | ❌ | ❌ | ✅ |
| `scanDrive()` | ❌ | ✅ | ✅ | ✅ | ❌ | ✅ | ❌ | ❌ | ❌ |
| `processFiles()` | ❌ | ✅ | ⚠️ | ✅ | ❌ | ✅ | ❌ | ❌ | ❌ |
| `extractTitle()` (AI) | ❌ | ✅ | ✅ | ❌ | ❌ | ✅ | ❌ | ⚠️ | ❌ |
| `callGemini()` | ❌ | ✅ | ✅ | ✅ | ❌ | ✅ | ❌ | ✅ | ❌ |
| `sendEmail()` | ✅ | ✅ | ❌ | ❌ | ✅ | ❌ | ✅ | ❌ | ❌ |
| `setupTrigger()` | ✅ | ✅ | ✅ | ❌ | ✅ | ❌ | ⚠️ | ❌ | ❌ |
| `logActivity()` | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ | ❌ | ❌ | ❌ |
| `getConfig()` | ✅ | ✅ | ✅ | ✅ | ✅ | ⚠️ | ⚠️ | ⚠️ | ❌ |

**Legend:** ✅ = Full implementation | ⚠️ = Partial/Similar | ❌ = Not present

---

## Common Variable/Config Patterns

### CONFIG Objects
```javascript
// Found in: file-center, spc, automated-expiry-notification
var CONFIG = {
  SHEET_NAME: "Sheet Name",
  LOG_SHEET: "Logs",
  TIMEZONE: "Asia/Manila",
  BATCH_SIZE: 25,
  MAX_RETRIES: 3
};
```

### GEMINI_CONFIG Objects
```javascript
// Found in: file-center, litigation, spc
var GEMINI_CONFIG = {
  FALLBACK_MODEL: "gemini-2.0-flash",
  TEMPERATURE: 0.2,
  API_KEY_PROPERTY: "GEMINI_API_KEY"
};
```

---

## Notes

- **Most Common Pattern:** `onOpen()` with menu creation (appears in ALL 10 Apps Script projects)
- **AI Integration:** 5 projects use Gemini AI (file-center, litigation, notary, marketing-seo-website, spc)
- **Email Automation:** 4 projects send emails (CAR, file-center, automated-expiry-notification, amlc-automation)
- **Drive Scanning:** 4 projects scan Drive folders (file-center, litigation, notary, spc)
- **Setup Wizards:** 3 projects have guided setup (CAR, automated-expiry-notification, gsheet-jira)
