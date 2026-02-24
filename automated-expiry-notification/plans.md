# Automated Expiry Notification — Full Documentation

This document explains how the current Google Apps Script works, how to configure it, and what operations users must follow so reminders are sent correctly.

---

## 1) Purpose

This automation monitors visa/document expiry rows in a Google Sheet and sends reminder emails on the exact notice date.

Core outcomes:

- Sends reminder emails automatically (daily trigger) or manually (menu).
- Uses `Remarks` as the email body template with placeholder substitution.
- Supports Drive file attachments.
- Tracks progress and failures in `Status` and `LOGS`.
- Supports selecting the working sheet tab through Initialization.

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

## 3) Prerequisites

- Script is bound to the target spreadsheet.
- Script account has access to:
  - Gmail (send email)
  - Drive (open attachment files)
  - Spreadsheet (read/write rows and logs)
- Spreadsheet contains row 2 headers and data starts at row 3.
- Timezone expectation for scheduling: `Asia/Manila`.

---

## 4) Initial Setup (Required)

### 4.1 Reload Spreadsheet to Register Menu

Open/reload the spreadsheet so `onOpen()` adds **Expiry Notifications** menu.

### 4.2 Select Working Sheet Tab (Initialization)

Use:

`Expiry Notifications -> Initialize / Select Working Sheet`

This saves selected sheet tab name in document properties (`AUTOMATION_SHEET_NAME`).

If not initialized, script falls back to default `CONFIG.SHEET_NAME` (`VISA automation`).

### 4.3 Activate Daily Trigger

Use:

`Expiry Notifications -> Activate Daily Schedule (8 AM)`

This installs a daily trigger for `runDailyCheck()` at 8:00 AM Asia/Manila.

---

## 5) Menu Reference

### Expiry Notifications

- **Show Automation Status**
  - Displays trigger status, configured sheet tab, and latest summary run from `LOGS`.

- **Initialize / Select Working Sheet**
  - Lets user select which tab the automation reads.

- **Run Manual Check Now**
  - Runs same production logic immediately (`runDailyCheck`).

- **Check Schedule Status**
  - Same status dialog as above.

- **Activate Daily Schedule (8 AM)**
  - Creates daily trigger.

- **Deactivate Daily Schedule**
  - Deletes all existing `runDailyCheck` triggers.

### Diagnostics Submenu

- **Preview Target Dates (no emails sent)**
  - Logs computed target dates for active rows only.

- **Inspect Row...**
  - Debug by sheet row number.

- **Send Test Email by No....**
  - Finds row by `No.` column value and sends a real test email.
  - Ignores `Status` and target-date matching.

---

## 6) Sheet and Column Requirements

### 6.1 Row Structure

- Row 1: free/visual header (ignored)
- Row 2: actual column headers
- Row 3+: data rows

### 6.2 Required Header Names (Row 2)

Required for production run:

- `Client Name`
- `Client Email`
- `Expiry Date`
- `Notice Date`
- `Status`

Optional but used when present:

- `No.` (used by test-send by No.)
- `Type of ID/Document`
- `Remarks`
- `Attached Files`
- `Staff Email`

Backward-compatible alias supported:

- `Assigned Staff Email` (mapped to `Staff Email` logic)

> Column order can change; script maps by header text, not by fixed column position.

---

## 7) Status Values

Recognized values:

- `Active` (case-insensitive: `active`, `ACTIVE`, etc.) -> processed
- blank status -> auto-set to `Active`, then processed
- `Sent` -> ignored in normal runs
- `Error` -> ignored unless manually reset to Active

During production/manual run:

- Success -> `Sent`
- Validation or send failure -> `Error`

---

## 8) Notice Date Parsing Rules

`Notice Date` is parsed dynamically using pattern:

`N day(s)/week(s)/month(s) before`

Also supports:

- `On expiry date`
- `On the expiry date`

Examples:

- `7 days before` -> minus 7 days
- `2 weeks before` -> minus 14 days
- `2 months before` -> minus 2 calendar months

If parsing fails, row is set to `Error` and logged.

---

## 9) Email Composition

### 9.1 Subject

```
Reminder: [Type of ID/Document] Expiry on [DD MMM YYYY] - [Client Name]
```

If `Type of ID/Document` is empty, uses `Visa/Permit`.

### 9.2 Body Source

Primary source: `Remarks` column.

If empty, script uses a default fallback message.

### 9.3 Supported Placeholders (Case-Insensitive)

- `[Client Name]`
- `[Date of Expiration]`
- `[Date of Expiry]`
- `[Expiry Date]`
- `[Document Type]`

### 9.4 Attachments

`Attached Files` supports comma-separated Drive links or raw file IDs.

If any attachment reference is invalid/inaccessible, row is marked `Error` and send is skipped.

### 9.5 Recipient Logic

- `To`: `Client Email`
- `CC`: current `Staff Email` value (if present)

After successful send, `Staff Email` is overwritten with sender account email (`Session.getEffectiveUser()` fallback `Session.getActiveUser()`).

---

## 10) Logging (`LOGS` Sheet)

`LOGS` is auto-created if missing.

Columns:

1. Timestamp
2. Client Name
3. Action
4. Detail

Common actions:

- `SENT`
- `ERROR`
- `SUMMARY`
- `INFO`

Summary format currently includes:

`Run complete. Total Rows: X | Eligible (Active/Blank): Y | Auto-Activated: A | Sent: Z | Errors: N`

---

## 11) Detailed Runtime Flow (`runDailyCheck`)

1. Resolve configured sheet tab.
2. Ensure `LOGS` sheet exists.
3. Build and validate column map from row 2 headers.
4. Read all data rows.
5. Filter eligible rows (`Active` + blank Status).
6. Validate row required fields.
7. Parse expiry + notice date.
8. Skip row if target date is still in the future.
9. Resolve attachments.
10. Build subject/body.
11. Send email.
12. Update `Status`, `Staff Email`, and logs.
13. Write summary log.

---

## 12) Testing and Diagnostics

### 12.1 Preview Without Sending

Use `Preview Target Dates (no emails sent)` to confirm which rows are send-eligible now (due/overdue).

### 12.2 Inspect a Row

Use `Inspect Row...` by row number for quick field-level debugging.

### 12.3 Send Real Test by `No.`

Use `Send Test Email by No....` and enter a value from `No.` column.

Behavior:

- If no row matches -> alert error.
- If multiple rows match -> warns and uses first match.
- Ignores status and target date checks.

---

## 13) Troubleshooting Guide

### Issue: `Configured sheet "..." not found`

Cause: Tab renamed/deleted after initialization.

Fix: Run `Initialize / Select Working Sheet` and reselect the correct tab.

### Issue: `Processed/Active is 0`

Likely causes:

- `Status` not Active
- Required fields missing
- Wrong configured sheet tab

Use status dialog + preview diagnostics.

### Issue: `Cannot parse Notice Date`

Cause: Notice text not matching supported format.

Fix: Use `On expiry date` or `N days/weeks/months before`.

### Issue: `No row found for No. "..."`

Cause: No exact/compatible value in `No.` column.

Fix: Check value format and whitespace; retry.

### Issue: Attachment error

Cause: Invalid Drive URL/ID or missing file permission.

Fix: Validate Drive links/IDs and sharing access.

---

## 14) Operational Notes

- Manual run and scheduled run use the same core logic (`runDailyCheck`).
- Trigger duplicates can happen if manually created multiple times; deactivate then activate once.
- `Staff Email` acts as both CC source and sender-audit field after successful send.
- To resend an item, change `Status` back to `Active` (it will send once target date is due/overdue).

---

## 15) Recommended Maintenance Checklist

- Verify headers in row 2 are intact.
- Keep `Notice Date` dropdown values consistent with supported patterns.
- Periodically review `LOGS` summary and errors.
- Re-run initialization if sheet tab names are changed.
- Confirm trigger remains active after major edits or account changes.
