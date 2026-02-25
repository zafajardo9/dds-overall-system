# Automated Expiry Notification — Full Documentation

> **Last Updated**: February 2026  
> **Script File**: `code.gs` (1,422 lines)

This document explains how the current Google Apps Script works, how to configure it, and what operations users must follow so reminders are sent correctly.

---

## Table of Contents

1. [Purpose](#1-purpose)
2. [High-Level Behavior](#2-high-level-behavior)
3. [Code Structure Overview](#3-code-structure-overview)
4. [Prerequisites](#4-prerequisites)
5. [Initial Setup (Required)](#5-initial-setup-required)
6. [Menu Reference](#6-menu-reference)
7. [Sheet and Column Requirements](#7-sheet-and-column-requirements)
8. [Status Values](#8-status-values)
9. [Notice Date Parsing Rules](#9-notice-date-parsing-rules)
10. [Email Composition](#10-email-composition)
11. [Logging (LOGS Sheet)](#11-logs-sheet)
12. [Detailed Runtime Flow (runDailyCheck)](#12-detailed-runtime-flow-rundailycheck)
13. [Testing and Diagnostics](#13-testing-and-diagnostics)
14. [Troubleshooting Guide](#14-troubleshooting-guide)
15. [Operational Notes](#15-operational-notes)
16. [Maintenance Checklist](#16-maintenance-checklist)

---

## 1) Purpose

This automation monitors visa/document expiry rows in a Google Sheet and sends reminder emails on the exact notice date.

Core outcomes:

- Sends reminder emails automatically (daily trigger) or manually (menu).
- Uses `Remarks` as the email body template with placeholder substitution.
- Supports Drive file attachments.
- Tracks progress and failures in `Status` and `LOGS`.
- Supports selecting the working sheet tab through Initialization.
- Auto-activates blank status rows for processing.

---

## 2) High-Level Behavior

The automation processes rows where `Status` is **Active** (case-insensitive) or blank.

For each eligible row:

1. Validate required fields.
2. Parse `Notice Date` (dynamic regex parser).
3. Compute `Target Date = Expiry Date - Notice Offset`.
4. If target date is due (today or earlier), send email.
5. On success:
   - Set `Status = Sent`
   - Write sender account into `Staff Email`
   - Add `SENT` log entry
6. On failure, set `Status = Error` and log details.

---

## 3) Code Structure Overview

The script is organized into the following function categories:

### Configuration & Constants
| Item | Description |
|------|-------------|
| `CONFIG` | Sheet names, header row, trigger hour, sender name |
| `HEADERS` | Expected column header names |
| `HEADER_ALIASES` | Alternate header mappings (e.g., "Assigned Staff Email" → "Staff Email") |
| `STATUS` | Valid status values: Active, Sent, Error, Skipped |
| `LOG_COL` | LOGS sheet column indices |

### Menu & UI Functions
| Function | Purpose |
|----------|---------|
| `onOpen()` | Creates "Expiry Notifications" custom menu |
| `showAutomationStatus()` | Alias for checkScheduleStatus() |
| `initializeAutomationSheet()` | Select which sheet tab to use |
| `checkScheduleStatus()` | Shows trigger status and last run summary |
| `manualRunNow()` | Runs daily check manually with confirmation |

### Core Automation
| Function | Purpose |
|----------|---------|
| `runDailyCheck()` | **Main entry point** - processes all eligible rows |
| `buildColumnMap()` | Maps header names to column indices |
| `validateColumnMap()` | Validates required columns exist |

### Date Helpers
| Function | Purpose |
|----------|---------|
| `getMidnight()` | Returns date at 00:00:00 |
| `isSameDay()` | Compares if two dates are the same day |
| `isTargetDateDue()` | Checks if target date is today or earlier |
| `parseNoticeOffset()` | Parses "N days/weeks/months before" |
| `computeTargetDate()` | Calculates target date from expiry + offset |
| `formatDate()` | Formats date as "DD MMM YYYY" |

### Attachment Helpers
| Function | Purpose |
|----------|---------|
| `extractDriveFileId()` | Extracts file ID from various Drive URL formats |
| `resolveAttachments()` | Converts Drive URLs to email attachments |

### Email Helpers
| Function | Purpose |
|----------|---------|
| `buildEmailBody()` | Creates HTML email from Remarks or default template |
| `applyTemplatePlaceholders()` | Replaces [Client Name], [Expiry Date], etc. |
| `buildEmailSubject()` | Creates subject line |
| `sendReminderEmail()` | Sends email via GmailApp |

### Logging
| Function | Purpose |
|----------|---------|
| `ensureLogsSheet()` | Creates LOGS sheet if missing |
| `appendLog()` | Adds log entry with color-coded action |

### Trigger Management
| Function | Purpose |
|----------|---------|
| `installTrigger()` | Creates daily 8 AM trigger |
| `removeTrigger()` | Removes all daily triggers |

### Diagnostic Functions
| Function | Purpose |
|----------|---------|
| `previewTargetDates()` | Shows computed target dates without sending |
| `diagnosticInspectRow()` | Inspects a specific row's parsed values |
| `diagnosticSendTestRow()` | Sends test email by "No." value |
| `testRunNow()` | Direct test runner (alias for runDailyCheck) |

---

## 4) Prerequisites

- Script is bound to the target spreadsheet.
- Script account has access to:
  - Gmail (send email)
  - Drive (open attachment files)
  - Spreadsheet (read/write rows and logs)
- Spreadsheet contains row 2 headers and data starts at row 3.
- Timezone expectation for scheduling: `Asia/Manila`.
- Run once: `installTrigger()` to enable daily scheduling.

---

## 5) Initial Setup (Required)

### 5.1 Reload Spreadsheet to Register Menu

Open/reload the spreadsheet so `onOpen()` adds **Expiry Notifications** menu.

### 5.2 Select Working Sheet Tab (Initialization)

Use:

`Expiry Notifications -> Initialize / Select Working Sheet`

This saves selected sheet tab name in document properties (`AUTOMATION_SHEET_NAME`).

If not initialized, script falls back to default `CONFIG.SHEET_NAME` (`VISA automation`).

### 5.3 Activate Daily Trigger

Use:

`Expiry Notifications -> Activate Daily Schedule (8 AM)`

This installs a daily trigger for `runDailyCheck()` at 8:00 AM Asia/Manila.

---

## 6) Menu Reference

### Expiry Notifications

| Menu Item | Description |
|-----------|-------------|
| **Show Automation Status** | Displays trigger status, configured sheet tab, and latest summary run from `LOGS`. |
| **Initialize / Select Working Sheet** | Lets user select which tab the automation reads. |
| **Run Manual Check Now** | Runs same production logic immediately (`runDailyCheck`). |
| **Check Schedule Status** | Same status dialog as above. |
| **Activate Daily Schedule (8 AM)** | Creates daily trigger. |
| **Deactivate Daily Schedule** | Deletes all existing `runDailyCheck` triggers. |

### Diagnostics Submenu

| Menu Item | Description |
|-----------|-------------|
| **Preview Target Dates (no emails sent)** | Logs computed target dates for active rows only. |
| **Inspect Row...** | Debug by sheet row number. |
| **Send Test Email by No....** | Finds row by `No.` column value and sends a real test email. Ignores `Status` and target-date matching. |

---

## 7) Sheet and Column Requirements

### 7.1 Row Structure

- Row 1: free/visual header (ignored)
- Row 2: actual column headers (used for mapping)
- Row 3+: data rows

### 7.2 Required Header Names (Row 2)

Required for production run:

| Column | Header Name |
|--------|-------------|
| Client Name | `Client Name` |
| Client Email | `Client Email` |
| Expiry Date | `Expiry Date` |
| Notice Date | `Notice Date` |
| Status | `Status` |

Optional but used when present:

| Column | Header Name |
|--------|-------------|
| No. | `No.` |
| Document Type | `Type of ID/Document` |
| Remarks | `Remarks` |
| Attachments | `Attached Files` |
| Staff Email | `Staff Email` |

### 7.3 Header Aliases

The script supports alternate header names (case-insensitive):

| Standard Name | Alias Supported |
|---------------|------------------|
| `Staff Email` | `Assigned Staff Email` |

> Column order can change; script maps by header text, not by fixed column position.

---

## 8) Status Values

Recognized values:

| Status | Behavior |
|--------|----------|
| `Active` (case-insensitive) | Processed in normal runs |
| blank | Auto-set to `Active`, then processed |
| `Sent` | Ignored in normal runs |
| `Error` | Ignored unless manually reset to Active |

During production/manual run:

- Success → `Sent`
- Validation or send failure → `Error`

### Auto-Activation

When a row has **blank Status**:
1. Script automatically sets it to `Active`
2. Processes the row in the same run
3. Logs an INFO message: "Blank Status auto-set to Active for processing."

---

## 9) Notice Date Parsing Rules

`Notice Date` is parsed dynamically using pattern:

```
N day(s)/week(s)/month(s) before
```

Also supports:

- `On expiry date`
- `On the expiry date`

### Supported Formats

| Input | Offset Applied |
|-------|----------------|
| `7 days before` | Minus 7 days |
| `1 day before` | Minus 1 day |
| `2 weeks before` | Minus 14 days |
| `1 week before` | Minus 7 days |
| `2 months before` | Minus 2 calendar months |
| `1 month before` | Minus 1 calendar month |
| `On expiry date` | Minus 0 days (send on expiry) |

If parsing fails, row is set to `Error` and logged.

---

## 10) Email Composition

### 10.1 Subject

```
Reminder: [Type of ID/Document] Expiry on [DD MMM YYYY] - [Client Name]
```

If `Type of ID/Document` is empty, uses `Visa/Permit`.

### 10.2 Body Source

Primary source: `Remarks` column.

If empty, script uses a default fallback message:

```
Good day, [Client Name],

This is a reminder that your visa/permit is expiring on [Expiry Date].

Please take the necessary steps before the expiry date.

Thank you.
```

### 10.3 Supported Placeholders (Case-Insensitive)

| Placeholder | Replaced With |
|--------------|----------------|
| `[Client Name]` | Client Name |
| `[client name]` | Client Name |
| `[Date of Expiration]` | Expiry Date (DD MMM YYYY) |
| `[date of expiration]` | Expiry Date |
| `[Date of Expiry]` | Expiry Date |
| `[date of expiry]` | Expiry Date |
| `[Expiry Date]` | Expiry Date |
| `[expiry date]` | Expiry Date |
| `[Document Type]` | Type of ID/Document (or "Visa/Permit") |

### 10.4 Attachments

`Attached Files` supports comma-separated list of:

- `https://drive.google.com/file/d/FILE_ID/view`
- `https://drive.google.com/open?id=FILE_ID`
- `https://drive.google.com/uc?id=FILE_ID`
- Raw file ID (20+ alphanumeric characters)

If any attachment reference is invalid/inaccessible, row is marked `Error` and send is skipped.

### 10.5 Recipient Logic

- **To**: `Client Email`
- **CC**: Current `Staff Email` value (if present)

After successful send, `Staff Email` is **overwritten** with sender account email (from `Session.getEffectiveUser()`).

---

## 11) Logging (`LOGS` Sheet)

`LOGS` is auto-created if missing.

### Columns

| Column | Index | Description |
|--------|-------|-------------|
| Timestamp | A | Date/time of action |
| Client Name | B | Client name for reference |
| Action | C | Action type (color-coded) |
| Detail | D | Detailed message |

### Action Types

| Action | Color | Description |
|--------|-------|-------------|
| `SENT` | Green | Email sent successfully |
| `ERROR` | Red | Validation or send failure |
| `SKIPPED` | Yellow | Row skipped (not due yet) |
| `SUMMARY` | Blue | End-of-run summary |
| `INFO` | None | General information |

### Summary Format

```
Run complete. Total Rows: X | Eligible (Active/Blank): Y | Auto-Activated: A | Sent: Z | Errors: N
```

Where:
- **Total Rows**: All rows in the sheet
- **Eligible**: Rows with Active or blank status
- **Auto-Activated**: Blank statuses converted to Active
- **Sent**: Successfully sent emails
- **Errors**: Failed validations or sends

---

## 12) Detailed Runtime Flow (`runDailyCheck`)

```
1. Resolve configured sheet tab
      ↓
2. Ensure LOGS sheet exists
      ↓
3. Build and validate column map from row 2 headers
      ↓
4. Read all data rows (row 3+)
      ↓
5. Filter eligible rows (Active + blank Status)
      ↓
6. For each eligible row:
   ├─ 6a. Auto-activate blank status → Active
   ├─ 6b. Validate required fields (Client Name, Client Email, Expiry Date, Notice Date)
   ├─ 6c. Parse expiry date
   ├─ 6d. Parse notice offset
   ├─ 6e. Compute target date
   ├─ 6f. Skip if target date is in the future
   ├─ 6g. Resolve attachments (if any)
   ├─ 6h. Build email subject and body
   ├─ 6i. Send email via GmailApp
   ├─ 6j. On success:
   │    ├─ Set Status = Sent
   │    ├─ Set Staff Email = sender account
   │    └─ Log SENT
   └─ 6k. On failure:
        ├─ Set Status = Error
        └─ Log ERROR
      ↓
7. Write summary log
```

---

## 13) Testing and Diagnostics

### 13.1 Preview Without Sending

**Menu**: `Diagnostics → Preview Target Dates (no emails sent)`

Shows:
- All active/blank rows
- Computed target dates
- Which rows will send now (due/overdue)

Use this to verify date calculations before going live.

### 13.2 Inspect a Row

**Menu**: `Diagnostics → Inspect Row...`

Enter a row number to see:
- All parsed field values
- Target date calculation
- Whether it's send-eligible now

### 13.3 Send Real Test by No.

**Menu**: `Diagnostics → Send Test Email by No....`

1. Enter a value from the `No.` column
2. Script finds the matching row
3. Shows preview of what will be sent
4. Confirms before sending

**Behavior**:
- If no row matches → alert error
- If multiple rows match → warns and uses first match
- **Ignores Status and target date checks** - always sends

---

## 14) Troubleshooting Guide

### Issue: `Configured sheet "..." not found`

**Cause**: Tab renamed/deleted after initialization.

**Fix**: Run `Initialize / Select Working Sheet` and reselect the correct tab.

---

### Issue: `Processed/Active is 0`

**Likely causes**:
- `Status` not Active (must be "Active" or blank)
- Required fields missing
- Wrong configured sheet tab

**Fix**: Use status dialog + preview diagnostics to investigate.

---

### Issue: `Cannot parse Notice Date`

**Cause**: Notice text not matching supported format.

**Fix**: Use one of:
- `On expiry date` or `On the expiry date`
- `N days before` (e.g., "7 days before")
- `N weeks before` (e.g., "2 weeks before")
- `N months before` (e.g., "1 month before")

---

### Issue: `No row found for No. "..."`

**Cause**: No exact/compatible value in `No.` column.

**Fix**: Check value format and whitespace; retry.

---

### Issue: Attachment error

**Cause**: Invalid Drive URL/ID or missing file permission.

**Fix**: Validate Drive links/IDs and sharing access.

---

### Issue: Duplicate triggers

**Cause**: Multiple daily triggers installed.

**Fix**: Run `Deactivate Daily Schedule` then `Activate Daily Schedule (8 AM)`.

---

### Issue: Email not sending but no error

**Cause**: Target date is still in the future.

**Fix**: Check `Preview Target Dates` to see when row will be due.

---

## 15) Operational Notes

- Manual run and scheduled run use the same core logic (`runDailyCheck`).
- Trigger duplicates can happen if manually created multiple times; deactivate then activate once.
- `Staff Email` acts as both CC source and sender-audit field after successful send.
- To resend an item, change `Status` back to `Active` (it will send once target date is due/overdue).
- Blank status rows are automatically activated on first run.
- All actions are logged in the `LOGS` sheet with color-coded action types.

---

## 16) Maintenance Checklist

- [ ] Verify headers in row 2 are intact
- [ ] Keep `Notice Date` values consistent with supported patterns
- [ ] Periodically review `LOGS` summary and errors
- [ ] Re-run initialization if sheet tab names change
- [ ] Confirm trigger remains active after major edits or account changes
- [ ] Check for duplicate triggers monthly
- [ ] Archive or clear old LOGS entries periodically

---

## Quick Reference

### Entry Points

| Function | When Called |
|----------|-------------|
| `runDailyCheck()` | Daily trigger (8 AM) or manual run |
| `installTrigger()` | One-time setup |
| `testRunNow()` | Direct test (alias) |

### Key Variables

```javascript
CONFIG.SHEET_NAME        // Default: "VISA automation"
CONFIG.LOGS_SHEET_NAME   // Default: "LOGS"
CONFIG.HEADER_ROW        // Default: 2
CONFIG.DATA_START_ROW    // Default: 3
CONFIG.TRIGGER_HOUR      // Default: 8 (8 AM)
CONFIG.SENDER_NAME       // Default: "DDS Office"
```

### Status Flow

```
(blank) → [Auto-Activate] → Active → [Send Success] → Sent
                                           → [Failure] → Error
                                           → [Future Date] → Active (skipped)
```
