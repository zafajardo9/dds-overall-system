# Automated Expiry Notification Plan

## Goal

Send the correct reminder email to each client on the exact notice date derived from their expiry date and notice preference, using the Remarks column as the full email body with dynamic placeholder substitution.

## Scope

### In Scope

- Daily scheduled check of all active records in the `VISA automation` sheet
- Notice date calculation from `Expiry Date` + `Notice Date` (dynamic, regex-based — no hardcoded list)
- Email delivery to client (`To`) and assigned staff (`CC`)
- Remarks column used as the full email body with `[Client Name]` and `[Date of Expiration]` substituted
- Drive file attachments from `Attached Files` column
- Row-level `Status` update after send
- Detailed logging to a separate `LOGS` sheet

### Out of Scope

- Renewal payment processing
- Invoice generation
- Non-expiry campaigns
- Modifying any column other than `Status`

---

## Sheet Structure

- **Tab name**: `VISA automation`
- **Row 1**: Decorative sheet title — ignored by script
- **Row 2**: Column headers
- **Row 3+**: Data rows

### Column Layout (A–M)

| Col | Letter | Header                | Required | Script Use                              |
| --- | ------ | --------------------- | -------- | --------------------------------------- |
| 1   | A      | No.                   | No       | Ignored                                 |
| 2   | B      | Client Name           | Yes      | Email body substitution, subject        |
| 3   | C      | Company               | No       | Ignored                                 |
| 4   | D      | Client Email          | Yes      | `To` address                            |
| 5   | E      | Type of ID/Document   | No       | Email subject                           |
| 6   | F      | ID No. / Document No. | No       | Ignored                                 |
| 7   | G      | Issued Date           | No       | Ignored                                 |
| 8   | H      | Expiry Date           | Yes      | Target date calculation                 |
| 9   | I      | Notice Date           | Yes      | Offset calculation (dynamic)            |
| 10  | J      | Remarks               | No       | Full email body (with placeholders)     |
| 11  | K      | Attached Files        | No       | Drive file attachments                  |
| 12  | L      | Status                | Yes      | Read to filter; written back after send |
| 13  | M      | Assigned Staff Email  | No       | `CC` address                            |

> The script detects column positions by reading the **header names in row 2**, so reordering columns will not break it.

### LOGS Sheet (auto-created if missing)

| Col | Header      |
| --- | ----------- |
| A   | Timestamp   |
| B   | Client Name |
| C   | Action      |
| D   | Detail      |

---

## Notice Date — Dynamic Parsing

The script does **not** use a hardcoded list of notice options. It parses the dropdown value at runtime using a regex, so any new option added to the dropdown works automatically.

### Parsing Rules

| Dropdown value                          | Parsed as                        |
| --------------------------------------- | -------------------------------- |
| `On expiry date` / `On the expiry date` | 0 days offset                    |
| `7 days before`                         | subtract 7 days                  |
| `14 days before`                        | subtract 14 days                 |
| `1 month before`                        | subtract 1 month                 |
| `2 months before`                       | subtract 2 months                |
| `3 months before`                       | subtract 3 months _(auto-works)_ |
| `2 weeks before`                        | subtract 14 days _(auto-works)_  |
| Any `N days/weeks/months before`        | parsed automatically             |

Unrecognized value → row marked `Error`, reason logged.

### Calculation Logic

1. Parse `Expiry Date` from column H.
2. Parse `Notice Date` from column I using regex.
3. Compute `Target Date = Expiry Date − Offset`.
4. Compare `Target Date` against today (midnight, local timezone).
5. Send only when `Target Date == Today`.

---

## Email Composition

### Subject

```
Reminder: [Type of ID/Document] Expiry on [DD MMM YYYY] – [Client Name]
```

### Body

The full text of the `Remarks` column (column J), with these substitutions applied:

| Placeholder            | Replaced with                                |
| ---------------------- | -------------------------------------------- |
| `[Client Name]`        | Value from `Client Name` column (B)          |
| `[Date of Expiration]` | Formatted `Expiry Date` (e.g. `24 Mar 2026`) |

If `Remarks` is empty, a minimal fallback body is used.

### Example Remarks input

```
Good day, [Client Name],

We have received a calendar notification indicating that your/[Client Name] SIRV Probationary
is set to expire on [Date of Expiration].

For your reference, a copy of the ACR and SIRV stamp is attached.

To avoid any penalties, please be reminded that the Conversion of Probationary to Indefinite
SIRV application must be filed within six (6) months prior to the expiration date.

Should you require our assistance with this matter, please do not hesitate to contact us.

Thank you.
```

### Attachments

All valid Drive files listed in `Attached Files` (column K) are attached to the email. The `[Attached Files]` text in the Remarks block is not rendered — actual files are attached.

---

## Daily Automation Workflow

1. Run once daily at 8:00 AM via time-based trigger.
2. Read headers from row 2 to detect column positions.
3. Load all rows from row 3 onward where `Status = Active`.
4. For each row:
   - Validate required fields (`Client Name`, `Client Email`, `Expiry Date`, `Notice Date`).
   - Parse `Notice Date` to compute `Target Date`.
   - Skip silently if `Target Date ≠ Today`.
   - Resolve Drive attachments from `Attached Files`.
   - Substitute placeholders in `Remarks` to build email body.
   - Send email (`To`: Client Email, `CC`: Assigned Staff Email if present).
   - On success → set `Status = Sent`, append LOGS entry.
   - On failure → set `Status = Error`, append LOGS entry with reason.

---

## Error Handling and Safeguards

- **Duplicate prevention**: skip if `Status` is already `Sent` (not `Active`).
- **Missing required field**: set `Status = Error`, log reason.
- **Invalid Notice Date**: set `Status = Error`, log unrecognized value.
- **Attachment error**: set `Status = Error`, log file ID and reason.
- **Send failure**: set `Status = Error`, log GmailApp error message.
- **Auditability**: every send attempt produces a LOGS row.

---

## Testing Checklist

- [ ] Row with `Target Date = Today` sends exactly once.
- [ ] Row with non-matching target date does not send.
- [ ] `[Client Name]` and `[Date of Expiration]` are substituted correctly in email body.
- [ ] New dropdown option (e.g. `3 months before`) is parsed without code changes.
- [ ] Missing `Client Email` is handled and logged.
- [ ] Missing/invalid Drive attachment is handled and logged.
- [ ] `CC` works when present and is omitted when empty.
- [ ] Already-`Sent` rows are not re-sent.
- [ ] `Status` column updates correctly after every run.
- [ ] `LOGS` sheet is auto-created and entries are color-coded.

---

## Success Criteria

- 100% of eligible rows are sent on the correct date.
- 0 duplicate sends (Status-based deduplication).
- All failures are visible with actionable error text in LOGS.
- Adding a new notice option to the dropdown requires zero code changes.
- Staff can trust the sheet as the single source of truth for reminder state.
