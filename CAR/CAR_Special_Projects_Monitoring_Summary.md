# CAR Special Projects Monitoring — Summary

## Overview

This Google Sheet functions as a **project monitoring and compliance tracking system** used to manage legal and corporate service projects, primarily focused on **CAR (Certificate Authorizing Registration) — Transfer of Shares** and related filings.

It acts as a centralized operational tracker for handling multiple client engagements while ensuring deadlines and compliance requirements are monitored. The sheet is **user-created**; `code.gs` is an Apps Script automation layer attached to it.

---

## Implementation Status

`CAR/code.gs` contains a **fully implemented Google Apps Script automation system**.

### Feature Table

| Feature                             | Status          | Description                                                                                               |
| ----------------------------------- | --------------- | --------------------------------------------------------------------------------------------------------- |
| **Deadline Calculation Engine**     | ✅ Active       | Auto-calculates DST Due Date (5th of next month) and CGT/DOD Due Date (30 days after notarization)        |
| **Remaining Days Tracker**          | ✅ Active       | Shows days remaining or overdue with color coding (green/yellow/red)                                      |
| **Reminder Generation**             | ✅ Active       | Auto-generates ⚠️ warning messages when deadlines approach (≤7 days) or are overdue                       |
| **Activity Logging**                | ✅ Active       | Logs all actions to "Activity Log" sheet with timestamp, user, and details                                |
| **Email Notifications**             | ✅ Active       | Sends automated reminder emails to Staff/Client when deadlines are near                                   |
| **Global Notification CC List**     | ✅ Active       | Configurable list of emails always CC'd on every reminder email                                           |
| **Status Auto-Update**              | ✅ Active       | Detects completion keywords in remarks and auto-updates project status                                    |
| **Numbered Setup Wizard Menu**      | ✅ Active       | Guided 4-step setup flow so users know exactly where to start                                             |
| **Working Sheet Selection**         | ✅ Active       | User picks which sheet tab automation targets; persisted in Script Properties                             |
| **Dynamic Header Row Detection**    | ✅ Active       | Auto-detects real header row even when rows 1–2 are merged design/title rows                              |
| **Configurable Header & Data Rows** | ✅ Active       | User can manually set header row and data start row via menu                                              |
| **Status Dropdown from Sheet**      | ✅ Active       | Status column dropdown reads existing values from the sheet; falls back to defaults                       |
| **Merged Row Tolerance**            | ✅ Active       | Setup never breaks apart data rows or divider rows; only unmerges the detected header row before freezing |
| **Daily Automation**                | ✅ Configurable | Optional daily trigger at 8 AM for automatic deadline checks                                              |

---

## Quick Start (First-Time Setup)

Open **CAR Monitoring → ⚙️ Setup (Start Here)** and follow the steps in order:

1. **Step 1 — Select Working Sheet Tab**: Pick which sheet tab the automation should target. Leave blank to run on all detected project sheets.
2. **Step 2 — Set Header & Data Rows**: Tell the script which row has your column names and which row your data starts on. Leave blank for auto-detection.
3. **Step 3 — Manage Notification Emails**: Add email addresses that will be CC'd on every reminder email.
4. **Step 4 — Apply Sheet Setup**: Freezes the header row, applies dropdowns and conditional formatting, and shows a confirmation summary.

Then:

- **CAR Monitoring → Automation → Create Daily Trigger** — enables automatic daily checks at 8 AM
- **CAR Monitoring → Diagnostics** — verify everything is configured correctly

---

## Menu Structure

```
CAR Monitoring
├── ⚙️ Setup (Start Here)
│     ├── Step 1 — Select Working Sheet Tab
│     ├── Step 2 — Set Header & Data Rows
│     ├── Step 3 — Manage Notification Emails
│     └── Step 4 — Apply Sheet Setup
├── Update All Deadlines
├── Generate Reminders
├── Send Reminder Emails
├── Automation → (Run Daily Check, Create/Remove Triggers)
├── Activity Log → (View, Clear Old Logs)
├── Settings → (View, Edit, Reset, Set Working Sheet/Rows/Emails)
├── Advanced → (Setup All Sheets, Validate Headers)
└── Diagnostics
```

---

## Script Properties Reference

All configuration is stored in **Script Properties** and survives script re-deploys.

| Property Key                             | Description                             | Default                               |
| ---------------------------------------- | --------------------------------------- | ------------------------------------- |
| `WORKING_SHEET`                          | Sheet tab the automation targets        | (all detected project sheets)         |
| `HEADER_ROW`                             | Row number of column headers            | auto-detect (densest row in rows 1–5) |
| `DATA_START_ROW`                         | Row number where data records begin     | header row + 1                        |
| `NOTIFICATION_EMAILS`                    | JSON array of global CC email addresses | `[]`                                  |
| `SHEET_CAR`, `SHEET_REAL_PROPERTY`, …    | Override sheet tab names                | see CONFIG_DEFAULTS                   |
| `HEADER_CLIENT_NAME`, `HEADER_STATUS`, … | Override column header names            | see CONFIG_DEFAULTS                   |
| `STATUS_ONGOING`, `STATUS_COMPLETED`     | Override status values                  | "On Going", "Completed"               |
| `HEADER_ROW` (Script Property)           | Manually pin header row                 | (auto-detect)                         |
| `DATA_START_ROW` (Script Property)       | Manually pin data start row             | (auto-detect)                         |

---

## Sheet Structure

The sheet is **user-created**. The script identifies columns by header name (with aliases), not by position.

| Column                | Purpose                                                      |
| --------------------- | ------------------------------------------------------------ |
| No.                   | Row number                                                   |
| Date                  | Entry date                                                   |
| Company / Client Name | Client identifier                                            |
| Seller / Donor        | Transaction party                                            |
| Buyer / Donee         | Transaction party                                            |
| Service Type          | Dropdown: Sale, Donation, Estate, Transfer, Apostille, Other |
| Notary Date           | Reference date for all deadline calculations                 |
| DST Due Date          | Auto-calculated: 5th of the month following notarization     |
| CGT / DOD Due Date    | Auto-calculated: 30 days after notarization                  |
| DST Remaining Days    | Auto-calculated: days until/since DST due date               |
| CGT Remaining Days    | Auto-calculated: days until/since CGT due date               |
| Project Status        | Dropdown: pulled from existing sheet values or defaults      |
| Reminder              | Auto-generated warning messages used in email content        |
| Remarks               | Manual workflow notes; triggers status auto-update           |
| Staff Email           | Primary TO recipient for reminder emails                     |
| Client Email          | Secondary TO recipient (if different from staff)             |

> **Merged rows are tolerated.** Rows used as section dividers (full-width merged) and vertically-merged company cells (one company, multiple seller/buyer rows) are never modified by setup.

---

## Email Notification Behavior

- **TO**: Staff Email + Client Email from each data row (deduplicated)
- **CC**: All emails in the global Notification Email list (those not already in TO)
- **Trigger**: Sent when the `Reminder` column is non-empty for a row
- **Content**: Client name, service type, sheet name, row number, deadline summaries
- Managed via **Settings → Notification Emails** or **Step 3** of the setup wizard

---

## Compliance & Deadline Rules

### DST Due Date

Documentary Stamp Tax: **5th day of the month following notarization.**

### CGT / DOD Due Date

Capital Gains Tax / Donor's Tax: **30 days after notarization.**

### Remaining Days

Positive = days until deadline. Negative = overdue. Color coding:

- 🟢 Green: > 7 days remaining
- 🟡 Yellow: 0–7 days remaining
- 🔴 Red: overdue (negative)

### Reminder Generation

Auto-generated text written to the Reminder column when remaining days ≤ 7 or overdue. This content is used as the email body alert.

---

## Services Covered

- CAR — Transfer of Shares
- Real Property Services
- Title Transfer
- Apostille & Special Projects
- Completed Projects (Archive)
- Dashboard / Summary

---

## Workflow Lifecycle

1. Draft preparation
2. Client review and approval
3. Document signing
4. Notarization ← **Notary Date entered here; triggers all calculations**
5. Tax filing
6. Tax payment
7. BIR submission
8. eCAR release
9. Project completion → Status set to Completed

---

## Systemization Potential

The structure is suitable for conversion into:

- Laravel-based internal system
- Database-driven dashboard
- Automated email reminder system
- Workflow automation platform (AppSheet / Make / n8n)

---

## Conclusion

`code.gs` is a **fully configured Google Apps Script automation layer** for the CAR Special Projects Monitoring Sheet. It handles deadline calculation, reminder generation, email notifications with a global CC list, and provides a guided setup wizard so any user can configure it without editing code.
