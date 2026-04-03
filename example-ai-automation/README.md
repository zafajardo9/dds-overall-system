# AI Document Analyzer - Example Apps Script Project

A complete working example demonstrating how to use the `common-functions` library to build a Google Sheets automation that uses Gemini AI to analyze documents.

---

## What This Project Does

This automation:
1. **Scans** a Google Drive folder for documents (PDFs, images)
2. **Analyzes** each document using Google's Gemini AI
3. **Extracts** structured information:
   - Document type (Contract, Invoice, Letter, etc.)
   - Main subject/topic
   - Key dates mentioned
   - Important parties/organizations
   - Overall sentiment
4. **Logs** all results to a Google Sheet with color-coded status
5. **Tracks** processing status and errors in an activity log
6. **Notifies** via email when complete (optional)

---

## Files Included

| File | Purpose |
|------|---------|
| `code.gs` | Main automation script (copy this into Apps Script editor) |
| `appsscript.json` | Project manifest with required OAuth scopes |
| `README.md` | This documentation file |

---

## Required Common Functions

Copy these files from `../common-functions/` into your Apps Script project alongside `code.gs`:

| Required File | Why It's Needed |
|---------------|---------------|
| `GeminiAPI.gs` | AI analysis with `callGeminiWithFile()`, API key management |
| `DriveUtils.gs` | Folder scanning with `getFolderByIdOrUrl()`, `listFilesInFolder()` |
| `SheetUtils.gs` | Sheet operations (used internally by common functions) |
| `EmailUtils.gs` | Email notifications with `sendHtmlEmail()` |
| `ConfigUtils.gs` | Property storage for configuration |
| `LoggingUtils.gs` | Activity logging with `logActivity()`, `initializeLogSheet()` |
| `UIUtils.gs` | Dialogs and menus with `showStatusPanel()`, `showProgressDialog()` |

---

## Setup Instructions

### Step 1: Create New Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Click **New Project**
3. Delete the default `myFunction()` code

### Step 2: Copy Files

1. Copy contents of `code.gs` into `Code.gs`
2. Copy contents of `appsscript.json` into `appsscript.json` (see manifest section below)
3. Copy all required common-function files (listed above) into separate `.gs` files

### Step 3: Set OAuth Scopes

In the Apps Script editor:
1. Click **Project Settings** (gear icon)
2. Check **"Show 'appsscript.json' manifest file"**
3. Paste the contents from `appsscript.json` in this folder

### Step 4: Run Setup

1. Save all files (Ctrl+S / Cmd+S)
2. Click **Run** → Select `onOpen` function
3. Authorize the script when prompted
4. Refresh the Google Sheet (if bound) or reload the Apps Script page

You should now see the menu: **🤖 AI Document Analyzer**

### Step 5: Configure

1. Click **🤖 AI Document Analyzer** → **📋 Setup Wizard**
2. Enter your Gemini API key (get free key at [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey))
3. Provide your Drive folder URL/ID containing documents to analyze
4. Complete setup

### Step 6: Run Automation

1. Click **🤖 AI Document Analyzer** → **▶️ Process Documents**
2. Watch the progress toasts
3. View results in the "Document Analysis" sheet
4. Check "Processing Log" sheet for activity tracking

---

## Project Structure

```
Your Apps Script Project
├── code.gs                 ← Your main automation logic
├── GeminiAPI.gs           ← Common: AI functions
├── DriveUtils.gs          ← Common: Drive operations
├── SheetUtils.gs          ← Common: Spreadsheet helpers
├── EmailUtils.gs          ← Common: Email functions
├── ConfigUtils.gs         ← Common: Settings storage
├── LoggingUtils.gs        ← Common: Activity logging
├── UIUtils.gs             ← Common: Dialogs & menus
└── appsscript.json        ← OAuth scopes configuration
```

---

## Configuration Options

Edit the `CONFIG` object at the top of `code.gs`:

```javascript
var CONFIG = {
  SHEET_NAME: "Document Analysis",      // Main results sheet name
  LOG_SHEET_NAME: "Processing Log",      // Activity log sheet name
  
  DEFAULT_PROMPT: "...",                // Change what Gemini analyzes
  
  SUPPORTED_MIME_TYPES: [                // File types to process
    "application/pdf",
    "image/jpeg",
    "image/png"
  ],
  
  MAX_FILE_SIZE_MB: 20,                  // Skip files larger than this
  MAX_FILES_PER_RUN: 10,                 // Limit per execution
  
  NOTIFICATION_EMAIL: "your@email.com", // Set to enable email alerts
  RATE_LIMIT_DELAY_MS: 2000             // Delay between API calls
};
```

---

## How It Works

### 1. Menu System (`onOpen`)
```javascript
function onOpen() {
  // Creates the "🤖 AI Document Analyzer" menu with all options
}
```

### 2. Setup Flow (`setupAutomation`)
Uses common functions:
- `setGeminiApiKey()` - From `GeminiAPI.gs`
- `getFolderByIdOrUrl()` - From `DriveUtils.gs`
- `showStatusPanel()` - From `UIUtils.gs`

### 3. Main Processing (`processDocuments`)
Uses common functions:
- `initializeLogSheet()` - Creates audit trail
- `listFilesInFolder()` - Scans Drive folder
- `callGeminiWithFile()` - AI analysis
- `logActivity()` / `logError()` - Tracking
- `sendHtmlEmail()` - Notifications (optional)
- `showProgressDialog()` - User feedback

### 4. Data Storage
- Results saved to Google Sheet with headers:
  - File ID, Name, Type, Size, Modified Date
  - Status, Raw Analysis, Document Type, Subject
  - Key Dates, Parties, Sentiment, Processed At

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Menu doesn't appear | Run `onOpen()` manually once to authorize |
| "No API key found" | Run **🔑 Set Gemini API Key** from menu |
| "Cannot access folder" | Check folder permissions, ensure it's shared with your account |
| "File too large" | Increase `MAX_FILE_SIZE_MB` or process smaller files |
| API errors | Check quota at [console.cloud.google.com](https://console.cloud.google.com) |
| Email not sending | Verify `NOTIFICATION_EMAIL` is set in CONFIG |

---

## Extending This Example

### Change the Analysis Prompt
Modify `CONFIG.DEFAULT_PROMPT` to ask Gemini different questions:

```javascript
DEFAULT_PROMPT: "Summarize this document in 3 bullet points."
```

### Add Custom Metadata
Extend `appendResultToSheet()` to save additional Gemini outputs.

### Schedule Automatic Runs
1. Click **Triggers** (clock icon in Apps Script)
2. Add trigger for `processDocuments`
3. Choose time-driven (e.g., daily at 9 AM)

### Process Multiple Folders
Modify `processDocuments()` to loop through multiple folder IDs stored in properties.

---

## Security Notes

- API keys are stored in **UserProperties** (not in code)
- Each user has their own separate API key
- Drive access requires explicit OAuth consent
- No sensitive data is logged (only file names and processing status)

---

## Next Steps

Once comfortable with this example:
1. Modify the `DEFAULT_PROMPT` for your specific use case
2. Add your own document processing logic
3. Integrate with other systems (Calendar, Docs, etc.)
4. Deploy as a web app for broader access

For more common functions, see the `../common-functions/` folder.
