/**
 * CAR Special Projects Monitoring System
 * Google Apps Script Implementation
 *
 * Features:
 * - Deadline tracking (DST and CGT/DOD calculations)
 * - Automated reminder generation
 * - Workflow activity logging
 * - Email notifications for approaching deadlines
 * - Multi-sheet support across all service types
 * - Flexible sheet/header names via Script Properties (survives renames)
 */

// ============================================================================
// CONFIGURATION DEFAULTS
// These are fallback values only. All runtime code uses getConfig().
// To change a name without editing code: CAR Monitoring → Settings
// ============================================================================

const CONFIG_DEFAULTS = {
  SHEETS: {
    CAR: "CAR -- Transfer of Shares",
    REAL_PROPERTY: "Real Property Services",
    TITLE_TRANSFER: "Title Transfer",
    APOSTILLE: "Apostille & Special Projects",
    COMPLETED: "Completed Projects",
    DASHBOARD: "Dashboard",
    ACTIVITY_LOG: "Activity Log",
  },
  HEADERS: {
    NO: "No.",
    DATE: "Date",
    CLIENT_NAME: "Company / Client Name",
    SELLER_DONOR: "Seller / Donor",
    BUYER_DONEE: "Buyer / Donee",
    SERVICE_TYPE: "Service Type",
    NOTARY_DATE: "Notary Date",
    DST_DUE_DATE: "DST Due Date",
    CGT_DOD_DUE_DATE: "CGT / DOD Due Date",
    DST_REMAINING: "DST Remaining Days",
    CGT_REMAINING: "CGT Remaining Days",
    STATUS: "Project Status",
    REMINDER: "Reminder",
    REMARKS: "Remarks",
    STAFF_EMAIL: "Staff Email",
    CLIENT_EMAIL: "Client Email",
  },
  STATUS: {
    ONGOING: "On Going",
    COMPLETED: "Completed",
  },
  SERVICE_TYPES: [
    "Sale",
    "Donation",
    "Estate",
    "Transfer",
    "Apostille",
    "Other",
  ],
  REMINDER_DAYS: [7, 3, 1, 0, -1],
  DST_RULE: { day: 5, monthOffset: 1 },
  CGT_RULE: { days: 30 },
};

// Keep CONFIG as alias to defaults so any legacy direct references don't hard-break.
const CONFIG = CONFIG_DEFAULTS;

// ============================================================================
// RUNTIME CONFIG — reads Script Properties, falls back to defaults
// ============================================================================

let _configCache = null;

function getConfig() {
  if (_configCache) return _configCache;

  const props = PropertiesService.getScriptProperties().getProperties();

  const sheets = {};
  Object.keys(CONFIG_DEFAULTS.SHEETS).forEach((key) => {
    const propKey = "SHEET_" + key;
    sheets[key] = props[propKey] || CONFIG_DEFAULTS.SHEETS[key];
  });

  const headers = {};
  Object.keys(CONFIG_DEFAULTS.HEADERS).forEach((key) => {
    const propKey = "HEADER_" + key;
    headers[key] = props[propKey] || CONFIG_DEFAULTS.HEADERS[key];
  });

  const statusOngoing =
    props["STATUS_ONGOING"] || CONFIG_DEFAULTS.STATUS.ONGOING;
  const statusCompleted =
    props["STATUS_COMPLETED"] || CONFIG_DEFAULTS.STATUS.COMPLETED;

  const headerRowProp = parseInt(props["HEADER_ROW"], 10);
  const dataStartRowProp = parseInt(props["DATA_START_ROW"], 10);

  _configCache = {
    SHEETS: sheets,
    HEADERS: headers,
    STATUS: { ONGOING: statusOngoing, COMPLETED: statusCompleted },
    SERVICE_TYPES: CONFIG_DEFAULTS.SERVICE_TYPES,
    REMINDER_DAYS: CONFIG_DEFAULTS.REMINDER_DAYS,
    DST_RULE: CONFIG_DEFAULTS.DST_RULE,
    CGT_RULE: CONFIG_DEFAULTS.CGT_RULE,
    HEADER_ROW: isNaN(headerRowProp) ? null : headerRowProp,
    DATA_START_ROW: isNaN(dataStartRowProp) ? null : dataStartRowProp,
  };

  return _configCache;
}

function invalidateConfigCache() {
  _configCache = null;
}

// ============================================================================
// NOTIFICATION EMAIL LIST — global CC recipients for all reminder emails
// ============================================================================

function getNotificationEmails() {
  const raw = PropertiesService.getScriptProperties().getProperty(
    "NOTIFICATION_EMAILS",
  );
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function setNotificationEmails(arr) {
  PropertiesService.getScriptProperties().setProperty(
    "NOTIFICATION_EMAILS",
    JSON.stringify(arr),
  );
}

function viewNotificationEmails() {
  const ui = SpreadsheetApp.getUi();
  const emails = getNotificationEmails();
  if (emails.length === 0) {
    ui.alert(
      "Notification Emails",
      "No notification emails configured.",
      ui.ButtonSet.OK,
    );
    return;
  }
  const list = emails.map((e, i) => `  ${i + 1}. ${e}`).join("\n");
  ui.alert(
    "Notification Emails",
    `Current global CC recipients:\n\n${list}`,
    ui.ButtonSet.OK,
  );
}

function addNotificationEmail() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Add Notification Email",
    "Enter the email address to add to the global CC list:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const email = response.getResponseText().trim().toLowerCase();
  if (!email) return;
  if (!email.includes("@")) {
    ui.alert("Invalid email address. Please include an '@' sign.");
    return;
  }

  const emails = getNotificationEmails();
  if (emails.includes(email)) {
    ui.alert(`"${email}" is already in the notification list.`);
    return;
  }

  emails.push(email);
  setNotificationEmails(emails);
  logActivity("SYSTEM", "Notification Email Added", email);
  ui.alert(
    `"${email}" has been added to the notification list.\n\nTotal: ${emails.length} email(s).`,
  );
}

function removeNotificationEmail() {
  const ui = SpreadsheetApp.getUi();
  const emails = getNotificationEmails();

  if (emails.length === 0) {
    ui.alert(
      "Notification Emails",
      "No notification emails to remove.",
      ui.ButtonSet.OK,
    );
    return;
  }

  const list = emails.map((e, i) => `  ${i + 1}. ${e}`).join("\n");
  const response = ui.prompt(
    "Remove Notification Email",
    `Current list:\n${list}\n\nEnter the number of the email to remove:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const index = parseInt(input, 10) - 1;

  if (isNaN(index) || index < 0 || index >= emails.length) {
    ui.alert(
      `Invalid selection. Please enter a number between 1 and ${emails.length}.`,
    );
    return;
  }

  const removed = emails.splice(index, 1)[0];
  setNotificationEmails(emails);
  logActivity("SYSTEM", "Notification Email Removed", removed);
  ui.alert(
    `"${removed}" has been removed.\n\nRemaining: ${emails.length} email(s).`,
  );
}

function clearNotificationEmails() {
  const ui = SpreadsheetApp.getUi();
  const emails = getNotificationEmails();

  if (emails.length === 0) {
    ui.alert(
      "Notification Emails",
      "The notification list is already empty.",
      ui.ButtonSet.OK,
    );
    return;
  }

  const response = ui.alert(
    "Clear Notification Emails",
    `This will remove all ${emails.length} email(s) from the global CC list. Continue?`,
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;

  setNotificationEmails([]);
  logActivity(
    "SYSTEM",
    "Notification Emails Cleared",
    `Removed ${emails.length} email(s)`,
  );
  ui.alert("All notification emails have been cleared.");
}

function getWorkingSheet() {
  return (
    PropertiesService.getScriptProperties().getProperty("WORKING_SHEET") || null
  );
}

function setWorkingSheet(sheetName) {
  PropertiesService.getScriptProperties().setProperty(
    "WORKING_SHEET",
    sheetName,
  );
}

function selectWorkingSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();

  const systemSheets = [
    cfg.SHEETS.COMPLETED,
    cfg.SHEETS.DASHBOARD,
    cfg.SHEETS.ACTIVITY_LOG,
  ];

  const allSheets = ss
    .getSheets()
    .map((s) => s.getName())
    .filter((name) => !systemSheets.includes(name));

  if (allSheets.length === 0) {
    ui.alert("No sheets found. Please add a data sheet first.");
    return;
  }

  const current = getWorkingSheet();
  const currentLabel = current
    ? `Current: "${current}"`
    : "Current: (all detected sheets)";
  const list = allSheets.map((name, i) => `  ${i + 1}. ${name}`).join("\n");

  const response = ui.prompt(
    "Step 1 — Select Working Sheet Tab",
    `${currentLabel}\n\nAvailable sheets:\n${list}\n\nEnter the number of the sheet the automation should work on.\nLeave blank to use all auto-detected project sheets:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();

  if (input === "") {
    PropertiesService.getScriptProperties().deleteProperty("WORKING_SHEET");
    invalidateConfigCache();
    logActivity("SYSTEM", "Working Sheet", "Reset to auto-detect all sheets");
    ui.alert(
      "Working sheet cleared. Automation will run on all detected project sheets.",
    );
    return;
  }

  const index = parseInt(input, 10) - 1;
  if (isNaN(index) || index < 0 || index >= allSheets.length) {
    ui.alert(
      `Invalid selection. Enter a number between 1 and ${allSheets.length}.`,
    );
    return;
  }

  const selected = allSheets[index];
  setWorkingSheet(selected);
  invalidateConfigCache();
  logActivity("SYSTEM", "Working Sheet", `Set to "${selected}"`);
  ui.alert(
    `Working sheet set to: "${selected}"\n\nProceed to Step 2 to configure the header and data rows.`,
  );
}

function getProjectSheets() {
  const working = getWorkingSheet();
  if (working) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(working);
    return sheet ? [working] : [];
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const configured = [
    cfg.SHEETS.CAR,
    cfg.SHEETS.REAL_PROPERTY,
    cfg.SHEETS.TITLE_TRANSFER,
    cfg.SHEETS.APOSTILLE,
  ];
  const detected = ss
    .getSheets()
    .map((sheet) => sheet.getName())
    .filter((name) => isProjectSheetName(name));

  return Array.from(new Set(configured.concat(detected))).filter((name) =>
    isProjectSheetName(name),
  );
}

function getHeaderAliases() {
  const cfg = getConfig();
  return {
    [cfg.HEADERS.CLIENT_NAME]: [
      "Company",
      "Company / Client Name",
      "Company Name",
      "Client",
      "Client Name",
    ],
    [cfg.HEADERS.SERVICE_TYPE]: ["Services", "Service", "Service Type"],
    [cfg.HEADERS.NOTARY_DATE]: [
      "NOTARY DATE OF DOCUMENT",
      "Notary Date of Document",
      "Notary Date",
    ],
    [cfg.HEADERS.STATUS]: ["Current Status", "Status", "Project Status"],
    [cfg.HEADERS.REMARKS]: ["Remarks", "Notes"],
    [cfg.HEADERS.STAFF_EMAIL]: [
      "Email Address",
      "Staff Email",
      "Assignee Email",
      "Email",
    ],
    [cfg.HEADERS.CLIENT_EMAIL]: [
      "Email Address",
      "Client Email",
      "Contact Email",
      "Email",
    ],
    [cfg.HEADERS.DST_DUE_DATE]: ["DST Due Date"],
    [cfg.HEADERS.CGT_DOD_DUE_DATE]: ["CGT / DOD Due Date", "CGT/ DOD Due Date"],
    [cfg.HEADERS.DST_REMAINING]: ["Remaining Days", "DST Remaining Days"],
    [cfg.HEADERS.CGT_REMAINING]: ["Remaining Days", "CGT Remaining Days"],
    [cfg.HEADERS.REMINDER]: ["Reminder", "Alert"],
  };
}

function normalizeHeaderName(header) {
  return (header || "")
    .toString()
    .toLowerCase()
    .replace(/["']/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getHeaderRow(sheet) {
  return detectHeaderRow(sheet);
}

function getHeaders(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return [];
  const row = getHeaderRow(sheet);
  return sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
}

function findHeaderIndexByName(headers, headerName) {
  if (!headers || !headers.length) return null;

  const aliases = getHeaderAliases();
  const targetNames = [headerName].concat(aliases[headerName] || []);
  const normalizedTargets = targetNames.map((name) =>
    normalizeHeaderName(name),
  );

  for (let i = 0; i < headers.length; i++) {
    const normalizedHeader = normalizeHeaderName(headers[i]);
    if (!normalizedHeader) continue;

    if (normalizedTargets.includes(normalizedHeader)) return i + 1;

    if (
      headerName === getConfig().HEADERS.DST_DUE_DATE &&
      normalizedHeader.indexOf("dst due date") === 0
    ) {
      return i + 1;
    }

    if (
      headerName === getConfig().HEADERS.CGT_DOD_DUE_DATE &&
      (normalizedHeader.indexOf("cgt / dod due date") === 0 ||
        normalizedHeader.indexOf("cgt/ dod due date") === 0)
    ) {
      return i + 1;
    }
  }

  return null;
}

function isDueDateHeader(header) {
  return normalizeHeaderName(header).indexOf("due date") >= 0;
}

function isProjectSheetName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;

  const cfg = getConfig();
  const systemSheets = [
    cfg.SHEETS.COMPLETED,
    cfg.SHEETS.DASHBOARD,
    cfg.SHEETS.ACTIVITY_LOG,
  ];
  if (systemSheets.includes(sheetName)) return false;

  const headers = getHeaders(sheet);
  if (!headers.length) return false;

  const hasNotaryDate = !!findHeaderIndexByName(
    headers,
    cfg.HEADERS.NOTARY_DATE,
  );
  const hasDueDate = headers.some((header) => isDueDateHeader(header));
  const hasReminder = headers.some(
    (header) => normalizeHeaderName(header) === "reminder",
  );

  return hasNotaryDate || hasDueDate || hasReminder;
}

function getDeadlineGroups(sheet) {
  const headers = getHeaders(sheet);
  const groups = [];

  for (let i = 0; i < headers.length; i++) {
    if (!isDueDateHeader(headers[i])) continue;

    const group = {
      label: headers[i],
      dueDateCol: i + 1,
      remainingCol: null,
      statusCol: null,
      reminderCol: null,
    };

    for (let j = i + 1; j < headers.length; j++) {
      if (isDueDateHeader(headers[j])) break;

      const normalized = normalizeHeaderName(headers[j]);
      if (!group.remainingCol && normalized === "remaining days") {
        group.remainingCol = j + 1;
        continue;
      }

      if (!group.statusCol && normalized === "status") {
        group.statusCol = j + 1;
        continue;
      }

      if (!group.reminderCol && normalized === "reminder") {
        group.reminderCol = j + 1;
      }
    }

    groups.push(group);
  }

  return groups;
}

function parseDueDateRule(headerText) {
  const normalized = normalizeHeaderName(headerText);
  if (!normalized) return null;

  if (normalized.indexOf("every 5th of the following month") >= 0) {
    return { type: "nextMonthFixedDay", day: 5, base: "notary" };
  }

  const dayMatch = normalized.match(
    /(\d+)\s+days?\s+after(?:\s+the)?\s+notary date/,
  );
  if (dayMatch) {
    return { type: "daysAfter", days: Number(dayMatch[1]), base: "notary" };
  }

  const yearMatch = normalized.match(
    /(\d+)\s+years?\s+from(?:\s+the)?\s+decedent(?:'s)?\s+date of death/,
  );
  if (yearMatch) {
    return { type: "yearsAfter", years: Number(yearMatch[1]), base: "death" };
  }

  return null;
}

function calculateDueDateFromRule(baseDate, rule) {
  if (!baseDate || !(baseDate instanceof Date) || !rule) return null;

  const date = new Date(baseDate);
  if (rule.type === "nextMonthFixedDay") {
    date.setMonth(date.getMonth() + 1);
    date.setDate(rule.day);
    return date;
  }

  if (rule.type === "daysAfter") {
    date.setDate(date.getDate() + rule.days);
    return date;
  }

  if (rule.type === "yearsAfter") {
    date.setFullYear(date.getFullYear() + rule.years);
    return date;
  }

  return null;
}

function getBaseDateForRule(sheet, row, rule) {
  if (!rule) return null;

  const cfg = getConfig();
  if (rule.base === "notary") {
    const notaryCol = getColumnIndex(sheet, cfg.HEADERS.NOTARY_DATE);
    return notaryCol ? sheet.getRange(row, notaryCol).getValue() : null;
  }

  if (rule.base === "death") {
    const headers = getHeaders(sheet);
    const deathIndex = headers.findIndex(
      (header) => normalizeHeaderName(header).indexOf("date of death") >= 0,
    );
    return deathIndex >= 0
      ? sheet.getRange(row, deathIndex + 1).getValue()
      : null;
  }

  return null;
}

function buildReminderText(label, remainingDays) {
  if (remainingDays === null || remainingDays === "" || isNaN(remainingDays)) {
    return "";
  }

  const cleanLabel = (label || "Deadline")
    .toString()
    .replace(/\s+/g, " ")
    .trim();

  if (remainingDays < 0) {
    return `${cleanLabel} OVERDUE by ${Math.abs(remainingDays)} days`;
  }

  if (remainingDays <= 7) {
    return `${cleanLabel} due in ${remainingDays} days`;
  }

  return "";
}

// ============================================================================
// SETTINGS — view and edit sheet/header names via menu dialog
// ============================================================================

function showSettings() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties().getProperties();
  const cfg = getConfig();

  let msg = "CAR Monitoring — Current Settings\n";
  msg += "=".repeat(44) + "\n\n";
  msg += "SHEET NAMES (Script Property → Current Value)\n";
  Object.keys(CONFIG_DEFAULTS.SHEETS).forEach((key) => {
    const propKey = "SHEET_" + key;
    const isCustom = !!props[propKey];
    msg += `  ${propKey}: "${cfg.SHEETS[key]}"${isCustom ? " (custom)" : " (default)"}\n`;
  });
  msg += "\nCOLUMN HEADERS (Script Property → Current Value)\n";
  Object.keys(CONFIG_DEFAULTS.HEADERS).forEach((key) => {
    const propKey = "HEADER_" + key;
    const isCustom = !!props[propKey];
    msg += `  ${propKey}: "${cfg.HEADERS[key]}"${isCustom ? " (custom)" : " (default)"}\n`;
  });
  msg += "\nSTATUS VALUES\n";
  msg += `  STATUS_ONGOING: "${cfg.STATUS.ONGOING}"${props["STATUS_ONGOING"] ? " (custom)" : " (default)"}\n`;
  msg += `  STATUS_COMPLETED: "${cfg.STATUS.COMPLETED}"${props["STATUS_COMPLETED"] ? " (custom)" : " (default)"}\n`;
  const workingSheet = getWorkingSheet();
  msg += "\nAUTOMATION TARGET\n";
  msg += `  WORKING_SHEET: ${workingSheet || "(all detected project sheets)"}${workingSheet ? " (custom)" : ""}\n`;
  msg += "\nROW LAYOUT\n";
  msg += `  HEADER_ROW: ${cfg.HEADER_ROW || "(auto-detect)"}${props["HEADER_ROW"] ? " (custom)" : ""}\n`;
  msg += `  DATA_START_ROW: ${cfg.DATA_START_ROW || "(header row + 1)"}${props["DATA_START_ROW"] ? " (custom)" : ""}\n`;
  msg += "\n— To change any value, use: Settings → Edit a Setting\n";
  msg +=
    "— To set row layout, use: Settings → Set Header Row / Set Data Start Row\n";
  msg += "— To reset all to defaults, use: Settings → Reset to Defaults";

  ui.alert("Settings", msg, ui.ButtonSet.OK);
}

function editSetting() {
  const ui = SpreadsheetApp.getUi();

  const keyResponse = ui.prompt(
    "Edit Setting",
    "Enter the Script Property key to change\n(e.g. SHEET_CAR  or  HEADER_CLIENT_NAME  or  STATUS_ONGOING):",
    ui.ButtonSet.OK_CANCEL,
  );
  if (keyResponse.getSelectedButton() !== ui.Button.OK) return;
  const key = keyResponse.getResponseText().trim();
  if (!key) return;

  const cfg = getConfig();
  const props = PropertiesService.getScriptProperties().getProperties();
  const currentValue = props[key] || "(using default)";

  const valResponse = ui.prompt(
    "Edit Setting",
    `Current value for "${key}": ${currentValue}\n\nEnter new value (leave blank to restore default):`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (valResponse.getSelectedButton() !== ui.Button.OK) return;
  const newValue = valResponse.getResponseText().trim();

  const scriptProps = PropertiesService.getScriptProperties();
  if (newValue === "") {
    scriptProps.deleteProperty(key);
    invalidateConfigCache();
    ui.alert(`"${key}" has been reset to its default value.`);
  } else {
    scriptProps.setProperty(key, newValue);
    invalidateConfigCache();
    ui.alert(`"${key}" updated to: "${newValue}"`);
  }

  logActivity(
    "SYSTEM",
    "Settings Changed",
    `${key} → "${newValue || "(reset to default)"}"`,
  );
}

function resetSettingsToDefaults() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Reset Settings",
    "Reset ALL sheet names and column headers to their default values?",
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;

  const scriptProps = PropertiesService.getScriptProperties();
  Object.keys(CONFIG_DEFAULTS.SHEETS).forEach((key) =>
    scriptProps.deleteProperty("SHEET_" + key),
  );
  Object.keys(CONFIG_DEFAULTS.HEADERS).forEach((key) =>
    scriptProps.deleteProperty("HEADER_" + key),
  );
  scriptProps.deleteProperty("STATUS_ONGOING");
  scriptProps.deleteProperty("STATUS_COMPLETED");
  scriptProps.deleteProperty("HEADER_ROW");
  scriptProps.deleteProperty("DATA_START_ROW");
  scriptProps.deleteProperty("WORKING_SHEET");
  invalidateConfigCache();

  ui.alert("All settings have been reset to defaults.");
  logActivity(
    "SYSTEM",
    "Settings Reset",
    "All sheet/header names restored to defaults",
  );
}

function setHeaderRow() {
  const ui = SpreadsheetApp.getUi();
  const current = getConfig().HEADER_ROW;
  const response = ui.prompt(
    "Set Header Row",
    `The header row is the row that contains your column names (Company, Seller/Donor, etc.).

Current value: ${current || "auto-detect"}

Enter the row number (e.g. 3), or leave blank to use auto-detection:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const scriptProps = PropertiesService.getScriptProperties();

  if (input === "") {
    scriptProps.deleteProperty("HEADER_ROW");
    invalidateConfigCache();
    ui.alert("Header row reset to auto-detection.");
    logActivity("SYSTEM", "Header Row", "Reset to auto-detect");
    return;
  }

  const num = parseInt(input, 10);
  if (isNaN(num) || num < 1) {
    ui.alert("Invalid input. Please enter a positive row number.");
    return;
  }

  scriptProps.setProperty("HEADER_ROW", String(num));
  invalidateConfigCache();
  logActivity("SYSTEM", "Header Row", `Set to row ${num}`);
  ui.alert(
    `Header row set to row ${num}.\n\nRun Setup CAR Sheet again to apply the new layout.`,
  );
}

function setDataStartRow() {
  const ui = SpreadsheetApp.getUi();
  const current = getConfig().DATA_START_ROW;
  const response = ui.prompt(
    "Set Data Start Row",
    `The data start row is where your first data record begins (below the header row).

Current value: ${current || "header row + 1"}

Enter the row number (e.g. 4), or leave blank to use header row + 1:`,
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const scriptProps = PropertiesService.getScriptProperties();

  if (input === "") {
    scriptProps.deleteProperty("DATA_START_ROW");
    invalidateConfigCache();
    ui.alert("Data start row reset to header row + 1.");
    logActivity("SYSTEM", "Data Start Row", "Reset to header row + 1");
    return;
  }

  const num = parseInt(input, 10);
  if (isNaN(num) || num < 1) {
    ui.alert("Invalid input. Please enter a positive row number.");
    return;
  }

  scriptProps.setProperty("DATA_START_ROW", String(num));
  invalidateConfigCache();
  logActivity("SYSTEM", "Data Start Row", `Set to row ${num}`);
  ui.alert(
    `Data start row set to row ${num}.\n\nRun Setup CAR Sheet again to apply the new layout.`,
  );
}

// ============================================================================
// SETUP WIZARD WRAPPERS
// ============================================================================

function runRowSetup() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "Step 2 — Header & Data Rows",
    "You will be prompted twice:\n\n  1. Header Row — the row containing your column names (e.g. Company, Seller/Donor)\n  2. Data Start Row — the first row where actual project records begin\n\nLeave either blank to use auto-detection.",
    ui.ButtonSet.OK,
  );
  setHeaderRow();
  setDataStartRow();
}

function manageNotificationEmails() {
  const ui = SpreadsheetApp.getUi();
  const emails = getNotificationEmails();
  const list = emails.length
    ? emails.map((e, i) => `  ${i + 1}. ${e}`).join("\n")
    : "  (none configured)";

  const response = ui.alert(
    "Step 3 — Notification Emails",
    `Reminder emails will be CC'd to these addresses:\n\n${list}\n\nUse Settings → Notification Emails to add or remove addresses.\n\nOpen Settings now?`,
    ui.ButtonSet.YES_NO,
  );

  if (response === ui.Button.YES) {
    addNotificationEmail();
  }
}

function setupSheetWithSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const working = getWorkingSheet();
  const sheetName = working || ss.getActiveSheet().getName();
  const sheet = ss.getSheetByName(sheetName) || ss.getActiveSheet();

  setupSheet(sheetName);

  const cfg = getConfig();
  const hRow = getHeaderRow(sheet);
  const dRow = getDataStartRow(sheet);
  const emails = getNotificationEmails();

  SpreadsheetApp.getUi().alert(
    "Step 4 Complete — Sheet Setup Applied",
    `Sheet: "${sheet.getName()}"\nHeader row: ${hRow}\nData start row: ${dRow}\nNotification emails: ${emails.length > 0 ? emails.join(", ") : "(none)"}\n\nAutomation is ready. Use Automation → Create Daily Trigger to enable automatic daily checks.`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

// ============================================================================
// ON OPEN - CUSTOM MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("CAR Monitoring")
    .addSubMenu(
      ui
        .createMenu("⚙️ Setup (Start Here)")
        .addItem("Step 1 — Select Working Sheet Tab", "selectWorkingSheet")
        .addItem("Step 2 — Set Header & Data Rows", "runRowSetup")
        .addItem(
          "Step 3 — Manage Notification Emails",
          "manageNotificationEmails",
        )
        .addItem("Step 4 — Apply Sheet Setup", "setupSheetWithSummary"),
    )
    .addSeparator()
    .addItem("Update All Deadlines", "updateAllDeadlines")
    .addItem("Generate Reminders", "generateAllReminders")
    .addItem("Send Reminder Emails", "sendReminderEmails")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Automation")
        .addItem("Run Daily Check", "runDailyCheck")
        .addItem("Create Daily Trigger", "createDailyTrigger")
        .addItem("Remove Triggers", "removeTriggers"),
    )
    .addSubMenu(
      ui
        .createMenu("Activity Log")
        .addItem("View Activity Log", "viewActivityLog")
        .addItem("Clear Old Logs", "clearOldLogs"),
    )
    .addSubMenu(
      ui
        .createMenu("Settings")
        .addItem("View Current Settings", "showSettings")
        .addItem("Edit a Setting", "editSetting")
        .addItem("Reset to Defaults", "resetSettingsToDefaults")
        .addSeparator()
        .addItem("Set Working Sheet", "selectWorkingSheet")
        .addItem("Set Header Row", "setHeaderRow")
        .addItem("Set Data Start Row", "setDataStartRow")
        .addSeparator()
        .addItem("View Notification Emails", "viewNotificationEmails")
        .addItem("Add Notification Email", "addNotificationEmail")
        .addItem("Remove Notification Email", "removeNotificationEmail")
        .addItem("Clear All Notification Emails", "clearNotificationEmails"),
    )
    .addSubMenu(
      ui
        .createMenu("Advanced")
        .addItem("Setup All Service Sheets", "setupAllSheets")
        .addItem("Validate Headers", "validateAllHeaders"),
    )
    .addSeparator()
    .addItem("Diagnostics", "showDiagnostics")
    .addToUi();
}

// ============================================================================
// SHEET SETUP & VALIDATION
// ============================================================================

function detectHeaderRow(sheet) {
  const configured = getConfig().HEADER_ROW;
  if (configured && configured >= 1) return configured;

  const maxScan = Math.min(5, sheet.getLastRow());
  if (maxScan < 1) return 1;

  const scanData = sheet
    .getRange(1, 1, maxScan, sheet.getLastColumn())
    .getValues();
  let bestRow = 1;
  let bestCount = 0;

  for (let i = 0; i < scanData.length; i++) {
    const nonEmpty = scanData[i].filter((v) => v !== "" && v !== null).length;
    if (nonEmpty > bestCount) {
      bestCount = nonEmpty;
      bestRow = i + 1;
    }
  }

  return bestRow;
}

function getDataStartRow(sheet) {
  const configured = getConfig().DATA_START_ROW;
  if (configured && configured >= 1) return configured;
  return getHeaderRow(sheet) + 1;
}

function setupSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const name = sheetName || ss.getActiveSheet().getName();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.getActiveSheet();
  }

  const lastCol = sheet.getLastColumn();
  const hasExistingHeaders = lastCol > 0 && sheet.getLastRow() > 0;

  if (!hasExistingHeaders) {
    const defaultHeaders = [
      cfg.HEADERS.NO,
      cfg.HEADERS.DATE,
      cfg.HEADERS.CLIENT_NAME,
      cfg.HEADERS.SELLER_DONOR,
      cfg.HEADERS.BUYER_DONEE,
      cfg.HEADERS.SERVICE_TYPE,
      cfg.HEADERS.NOTARY_DATE,
      cfg.HEADERS.DST_DUE_DATE,
      cfg.HEADERS.CGT_DOD_DUE_DATE,
      cfg.HEADERS.DST_REMAINING,
      cfg.HEADERS.CGT_REMAINING,
      cfg.HEADERS.STATUS,
      cfg.HEADERS.REMINDER,
      cfg.HEADERS.REMARKS,
      cfg.HEADERS.STAFF_EMAIL,
      cfg.HEADERS.CLIENT_EMAIL,
    ];
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    sheet
      .getRange(1, 1, 1, defaultHeaders.length)
      .setFontWeight("bold")
      .setBackground("#4285f4")
      .setFontColor("white");
    sheet.getRange(1, 1, 1, sheet.getMaxColumns()).breakApart();
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, defaultHeaders.length);

    applyDataValidations(sheet, 1);
    applyConditionalFormatting(sheet, 1);

    logActivity(
      "SYSTEM",
      "Sheet Setup",
      `Initialized ${sheet.getName()} with default headers`,
    );
    SpreadsheetApp.getActive().toast(
      `Sheet "${sheet.getName()}" setup complete`,
      "Success",
    );
    return sheet;
  }

  const headerRow = detectHeaderRow(sheet);
  const numCols = sheet.getLastColumn();

  sheet.getRange(headerRow, 1, 1, sheet.getMaxColumns()).breakApart();
  sheet.setFrozenRows(headerRow);
  sheet.autoResizeColumns(1, numCols);

  applyDataValidations(sheet, headerRow);
  applyConditionalFormatting(sheet, headerRow);

  logActivity(
    "SYSTEM",
    "Sheet Setup",
    `Prepared existing sheet ${sheet.getName()} (header row ${headerRow})`,
  );
  SpreadsheetApp.getActive().toast(
    `Sheet "${sheet.getName()}" setup complete`,
    "Success",
  );
  return sheet;
}

function setupAllSheets() {
  getProjectSheets().forEach((name) => setupSheet(name));
  ensureActivityLogSheet();
  ensureCompletedSheet();
  ensureDashboardSheet();

  SpreadsheetApp.getActive().toast(
    "All service sheets initialized",
    "Complete",
  );
}

function getStatusValuesFromSheet(sheet) {
  const cfg = getConfig();
  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);
  if (!statusCol) return [];

  const lastRow = sheet.getLastRow();
  const dataStart = getDataStartRow(sheet);
  if (lastRow < dataStart) return [];

  const values = sheet
    .getRange(dataStart, statusCol, lastRow - dataStart + 1)
    .getValues()
    .flat()
    .map((v) => v.toString().trim())
    .filter((v) => v !== "");

  const unique = Array.from(new Set(values));
  return unique;
}

function applyDataValidations(sheet, headerRow) {
  const hRow = headerRow || getHeaderRow(sheet);
  const dataStart = headerRow ? hRow + 1 : getDataStartRow(sheet);
  const lastRow = Math.max(sheet.getLastRow(), hRow + 100);
  const cfg = getConfig();

  const serviceTypeCol = getColumnIndex(sheet, cfg.HEADERS.SERVICE_TYPE);
  if (serviceTypeCol) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(cfg.SERVICE_TYPES, true)
      .setAllowInvalid(false)
      .build();
    sheet
      .getRange(dataStart, serviceTypeCol, lastRow - hRow)
      .setDataValidation(rule);
  }

  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);
  if (statusCol) {
    const sheetStatusValues = getStatusValuesFromSheet(sheet);
    const statusList =
      sheetStatusValues.length >= 1
        ? sheetStatusValues
        : [cfg.STATUS.ONGOING, cfg.STATUS.COMPLETED];
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusList, true)
      .setAllowInvalid(true)
      .build();
    sheet
      .getRange(dataStart, statusCol, lastRow - hRow)
      .setDataValidation(rule);
  }
}

function applyConditionalFormatting(sheet, headerRow) {
  const hRow = headerRow || getHeaderRow(sheet);
  const dataStart = headerRow ? hRow + 1 : getDataStartRow(sheet);
  const cfg = getConfig();
  const groupedRemainingCols = getDeadlineGroups(sheet)
    .map((group) => group.remainingCol)
    .filter(Boolean);
  const remainingCols = Array.from(
    new Set(
      groupedRemainingCols.concat([
        getColumnIndex(sheet, cfg.HEADERS.DST_REMAINING),
        getColumnIndex(sheet, cfg.HEADERS.CGT_REMAINING),
      ]),
    ),
  ).filter(Boolean);

  const lastRow = Math.max(sheet.getLastRow(), hRow + 100);

  remainingCols.forEach((col) => {
    if (!col) return;
    const range = sheet.getRange(dataStart, col, lastRow - hRow);

    const overdueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground("#ffcccc")
      .setFontColor("#cc0000")
      .setRanges([range])
      .build();

    const warningRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0, 7)
      .setBackground("#ffeb9c")
      .setFontColor("#9c6500")
      .setRanges([range])
      .build();

    const goodRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(7)
      .setBackground("#c6efce")
      .setFontColor("#006100")
      .setRanges([range])
      .build();

    const rules = sheet.getConditionalFormatRules();
    rules.push(overdueRule, warningRule, goodRule);
    sheet.setConditionalFormatRules(rules);
  });
}

function getColumnIndex(sheet, headerName) {
  return findHeaderIndexByName(getHeaders(sheet), headerName);
}

function validateAllHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const issues = [];

  getProjectSheets().forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      issues.push(`Missing sheet: ${sheetName}`);
      return;
    }

    const headers = getHeaders(sheet);
    const required = [
      cfg.HEADERS.CLIENT_NAME,
      cfg.HEADERS.NOTARY_DATE,
      cfg.HEADERS.STATUS,
    ];

    required.forEach((req) => {
      if (!findHeaderIndexByName(headers, req)) {
        issues.push(`${sheetName}: Missing required header "${req}"`);
      }
    });
  });

  if (issues.length > 0) {
    SpreadsheetApp.getUi().alert(
      "Header Validation Issues:\n\n" + issues.join("\n"),
    );
  } else {
    SpreadsheetApp.getActive().toast(
      "All headers validated successfully",
      "Success",
    );
  }

  return issues;
}

// ============================================================================
// DEADLINE CALCULATION ENGINE
// ============================================================================

function calculateDSTDueDate(notaryDate) {
  if (!notaryDate || !(notaryDate instanceof Date)) return null;
  const rule = getConfig().DST_RULE;
  const date = new Date(notaryDate);
  date.setMonth(date.getMonth() + rule.monthOffset);
  date.setDate(rule.day);
  return date;
}

function calculateCGTDueDate(notaryDate) {
  if (!notaryDate || !(notaryDate instanceof Date)) return null;
  const date = new Date(notaryDate);
  date.setDate(date.getDate() + getConfig().CGT_RULE.days);
  return date;
}

function calculateRemainingDays(dueDate) {
  if (!dueDate || !(dueDate instanceof Date)) return null;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const due = new Date(dueDate);
  due.setHours(0, 0, 0, 0);

  const diffTime = due - today;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  return diffDays;
}

function updateRowDeadlines(sheet, row) {
  const cfg = getConfig();
  const groups = getDeadlineGroups(sheet);

  if (groups.length > 0) {
    let updated = false;

    groups.forEach((group) => {
      const rule = parseDueDateRule(group.label);
      const baseDate = getBaseDateForRule(sheet, row, rule);
      let dueDate = sheet.getRange(row, group.dueDateCol).getValue();

      if (rule && baseDate instanceof Date) {
        dueDate = calculateDueDateFromRule(baseDate, rule);
        sheet.getRange(row, group.dueDateCol).setValue(dueDate);
        updated = true;
      }

      if (group.remainingCol && dueDate instanceof Date) {
        sheet
          .getRange(row, group.remainingCol)
          .setValue(calculateRemainingDays(dueDate));
        updated = true;
      }
    });

    if (updated) return true;
  }

  const notaryCol = getColumnIndex(sheet, cfg.HEADERS.NOTARY_DATE);
  const dstDueCol = getColumnIndex(sheet, cfg.HEADERS.DST_DUE_DATE);
  const cgtDueCol = getColumnIndex(sheet, cfg.HEADERS.CGT_DOD_DUE_DATE);
  const dstRemainCol = getColumnIndex(sheet, cfg.HEADERS.DST_REMAINING);
  const cgtRemainCol = getColumnIndex(sheet, cfg.HEADERS.CGT_REMAINING);

  if (!notaryCol) return false;

  const notaryDate = sheet.getRange(row, notaryCol).getValue();
  if (!notaryDate || !(notaryDate instanceof Date)) return false;

  const dstDue = calculateDSTDueDate(notaryDate);
  const cgtDue = calculateCGTDueDate(notaryDate);

  if (dstDueCol) sheet.getRange(row, dstDueCol).setValue(dstDue);
  if (cgtDueCol) sheet.getRange(row, cgtDueCol).setValue(cgtDue);

  if (dstRemainCol) {
    sheet.getRange(row, dstRemainCol).setValue(calculateRemainingDays(dstDue));
  }

  if (cgtRemainCol) {
    sheet.getRange(row, cgtRemainCol).setValue(calculateRemainingDays(cgtDue));
  }

  return true;
}

function updateAllDeadlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getProjectSheets();

  let updatedCount = 0;

  sheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    for (let row = 2; row <= lastRow; row++) {
      if (updateRowDeadlines(sheet, row)) {
        updatedCount++;
      }
    }
  });

  logActivity(
    "SYSTEM",
    "Update Deadlines",
    `Updated ${updatedCount} rows across all sheets`,
  );
  SpreadsheetApp.getActive().toast(
    `Updated ${updatedCount} records`,
    "Deadlines Updated",
  );
}

// ============================================================================
// STATUS & REMINDER SYSTEM
// ============================================================================

function generateReminderMessage(dstRemaining, cgtRemaining, clientName) {
  const reminders = [];

  if (dstRemaining !== null) {
    if (dstRemaining < 0) {
      reminders.push(`DST payment OVERDUE by ${Math.abs(dstRemaining)} days`);
    } else if (dstRemaining <= 7) {
      reminders.push(`DST payment due in ${dstRemaining} days`);
    }
  }

  if (cgtRemaining !== null) {
    if (cgtRemaining < 0) {
      reminders.push(
        `CGT/DOD filing OVERDUE by ${Math.abs(cgtRemaining)} days`,
      );
    } else if (cgtRemaining <= 7) {
      reminders.push(`CGT/DOD filing due in ${cgtRemaining} days`);
    }
  }

  if (reminders.length === 0) return "";

  return "⚠️ " + reminders.join("; ");
}

function updateRowReminder(sheet, row) {
  const cfg = getConfig();
  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);
  const groups = getDeadlineGroups(sheet);

  const currentStatus = statusCol
    ? sheet.getRange(row, statusCol).getValue()
    : "";

  if (groups.length > 0) {
    let hasReminder = false;

    groups.forEach((group) => {
      if (!group.reminderCol) return;

      if (currentStatus === cfg.STATUS.COMPLETED) {
        sheet.getRange(row, group.reminderCol).setValue("");
        return;
      }

      const remainingValue = group.remainingCol
        ? sheet.getRange(row, group.remainingCol).getValue()
        : null;
      const reminder = buildReminderText(group.label, Number(remainingValue));
      sheet.getRange(row, group.reminderCol).setValue(reminder);

      if (reminder) hasReminder = true;
    });

    return hasReminder;
  }

  const clientCol = getColumnIndex(sheet, cfg.HEADERS.CLIENT_NAME);
  const dstRemainCol = getColumnIndex(sheet, cfg.HEADERS.DST_REMAINING);
  const cgtRemainCol = getColumnIndex(sheet, cfg.HEADERS.CGT_REMAINING);
  const reminderCol = getColumnIndex(sheet, cfg.HEADERS.REMINDER);

  if (!reminderCol) return false;

  const clientName = clientCol ? sheet.getRange(row, clientCol).getValue() : "";
  const dstRemaining = dstRemainCol
    ? sheet.getRange(row, dstRemainCol).getValue()
    : null;
  const cgtRemaining = cgtRemainCol
    ? sheet.getRange(row, cgtRemainCol).getValue()
    : null;

  if (currentStatus === cfg.STATUS.COMPLETED) {
    sheet.getRange(row, reminderCol).setValue("");
    return true;
  }

  const reminder = generateReminderMessage(
    dstRemaining,
    cgtRemaining,
    clientName,
  );
  sheet.getRange(row, reminderCol).setValue(reminder);

  return reminder !== "";
}

function generateAllReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getProjectSheets();

  let reminderCount = 0;

  sheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    for (let row = 2; row <= lastRow; row++) {
      if (updateRowReminder(sheet, row)) {
        reminderCount++;
      }
    }
  });

  logActivity(
    "SYSTEM",
    "Generate Reminders",
    `Generated ${reminderCount} reminders`,
  );
  SpreadsheetApp.getActive().toast(
    `${reminderCount} reminders generated`,
    "Reminders",
  );
}

function autoUpdateStatus(sheet, row) {
  const cfg = getConfig();
  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);
  const remarksCol = getColumnIndex(sheet, cfg.HEADERS.REMARKS);

  if (!statusCol || !remarksCol) return;

  const remarks = sheet.getRange(row, remarksCol).getValue();
  const currentStatus = sheet.getRange(row, statusCol).getValue();

  if (!remarks) return;

  const completionKeywords = [
    "completed",
    "done",
    "finished",
    "released",
    "ecar issued",
  ];
  const lowerRemarks = remarks.toString().toLowerCase();

  const isComplete = completionKeywords.some((kw) => lowerRemarks.includes(kw));

  if (isComplete && currentStatus !== cfg.STATUS.COMPLETED) {
    sheet.getRange(row, statusCol).setValue(cfg.STATUS.COMPLETED);

    const clientCol = getColumnIndex(sheet, cfg.HEADERS.CLIENT_NAME);
    const clientName = clientCol
      ? sheet.getRange(row, clientCol).getValue()
      : "Unknown";
    logActivity(
      clientName,
      "Status Auto-Update",
      "Marked as Completed based on remarks",
    );
  }
}

// ============================================================================
// ACTIVITY LOGGING
// ============================================================================

function ensureCompletedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEETS.COMPLETED);

  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEETS.COMPLETED);
    const headers = [
      cfg.HEADERS.NO,
      cfg.HEADERS.DATE,
      cfg.HEADERS.CLIENT_NAME,
      cfg.HEADERS.SELLER_DONOR,
      cfg.HEADERS.BUYER_DONEE,
      cfg.HEADERS.SERVICE_TYPE,
      cfg.HEADERS.NOTARY_DATE,
      cfg.HEADERS.DST_DUE_DATE,
      cfg.HEADERS.CGT_DOD_DUE_DATE,
      cfg.HEADERS.STATUS,
      cfg.HEADERS.REMARKS,
      "Archived Date",
    ];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#34a853");
    headerRange.setFontColor("white");
    sheet.setFrozenRows(1);
    Logger.log("Created Completed Projects sheet");
  }

  return sheet;
}

function ensureDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEETS.DASHBOARD);

  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEETS.DASHBOARD);
    Logger.log("Created Dashboard sheet");
  }

  updateDashboard(ss, sheet);
  return sheet;
}

function updateDashboard(ss, dashSheet) {
  const cfg = getConfig();
  const sheet =
    dashSheet ||
    (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName(
      cfg.SHEETS.DASHBOARD,
    );
  if (!sheet) return;

  sheet.clearContents();

  const title = [["CAR Special Projects — Dashboard"]];
  sheet.getRange(1, 1).setValues(title).setFontWeight("bold").setFontSize(14);

  const generated = [["Generated:", new Date().toLocaleString()]];
  sheet.getRange(2, 1, 1, 2).setValues(generated);

  const summaryHeaders = [
    ["Sheet", "Total", "On Going", "Completed", "Overdue DST", "Overdue CGT"],
  ];
  sheet
    .getRange(4, 1, 1, 6)
    .setValues(summaryHeaders)
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("white");

  const sourceSheets = ss ? ss : SpreadsheetApp.getActiveSpreadsheet();
  let summaryRow = 5;

  getProjectSheets().forEach((sheetName) => {
    const src = sourceSheets.getSheetByName(sheetName);
    if (!src || src.getLastRow() < 2) {
      sheet
        .getRange(summaryRow, 1, 1, 6)
        .setValues([[sheetName, 0, 0, 0, 0, 0]]);
      summaryRow++;
      return;
    }

    const headers = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0];
    const data = src
      .getRange(2, 1, src.getLastRow() - 1, headers.length)
      .getValues();
    const colMap = {};
    headers.forEach((h, i) => (colMap[h] = i));

    let total = 0,
      ongoing = 0,
      completed = 0,
      overdueDst = 0,
      overdueCgt = 0;

    data.forEach((row) => {
      const client =
        colMap[cfg.HEADERS.CLIENT_NAME] !== undefined
          ? row[colMap[cfg.HEADERS.CLIENT_NAME]]
          : "";
      if (!client) return;
      total++;
      const status =
        colMap[cfg.HEADERS.STATUS] !== undefined
          ? row[colMap[cfg.HEADERS.STATUS]]
          : "";
      if (status === cfg.STATUS.COMPLETED) completed++;
      else ongoing++;
      const dst =
        colMap[cfg.HEADERS.DST_REMAINING] !== undefined
          ? row[colMap[cfg.HEADERS.DST_REMAINING]]
          : null;
      const cgt =
        colMap[cfg.HEADERS.CGT_REMAINING] !== undefined
          ? row[colMap[cfg.HEADERS.CGT_REMAINING]]
          : null;
      if (typeof dst === "number" && dst < 0) overdueDst++;
      if (typeof cgt === "number" && cgt < 0) overdueCgt++;
    });

    sheet
      .getRange(summaryRow, 1, 1, 6)
      .setValues([
        [sheetName, total, ongoing, completed, overdueDst, overdueCgt],
      ]);
    summaryRow++;
  });

  sheet.autoResizeColumns(1, 6);
}

function ensureActivityLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  let sheet = ss.getSheetByName(cfg.SHEETS.ACTIVITY_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(cfg.SHEETS.ACTIVITY_LOG);
    const headers = ["Timestamp", "Client Name", "Action", "Detail", "User"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function logActivity(clientName, action, detail) {
  try {
    const sheet = ensureActivityLogSheet();
    const user = Session.getEffectiveUser().getEmail();
    const timestamp = new Date();

    sheet.appendRow([timestamp, clientName, action, detail, user]);
  } catch (e) {
    Logger.log("Failed to log activity: " + e.message);
  }
}

function viewActivityLog() {
  const sheet = ensureActivityLogSheet();
  SpreadsheetApp.setActiveSheet(sheet);
}

function clearOldLogs(daysToKeep = 90) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Clear Old Logs",
    `Delete activity logs older than ${daysToKeep} days?`,
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) return;

  const sheet = ensureActivityLogSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

  const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let deletedCount = 0;

  for (let i = timestamps.length - 1; i >= 0; i--) {
    const rowDate = timestamps[i][0];
    if (rowDate instanceof Date && rowDate < cutoffDate) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    }
  }

  SpreadsheetApp.getActive().toast(
    `Deleted ${deletedCount} old log entries`,
    "Logs Cleared",
  );
}

function getProjectHistory(clientName) {
  const sheet = ensureActivityLogSheet();
  const data = sheet.getDataRange().getValues();
  const history = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === clientName) {
      history.push({
        timestamp: data[i][0],
        action: data[i][2],
        detail: data[i][3],
        user: data[i][4],
      });
    }
  }

  return history.sort((a, b) => b.timestamp - a.timestamp);
}

// ============================================================================
// EMAIL NOTIFICATIONS
// ============================================================================

function sendReminderEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const sheets = getProjectSheets();

  let emailCount = 0;
  const quota = MailApp.getRemainingDailyQuota();

  if (quota < 10) {
    SpreadsheetApp.getUi().alert(
      "Email quota too low (" + quota + "). Try again tomorrow.",
    );
    return;
  }

  sheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const headers = getHeaders(sheet);
    const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    const deadlineGroups = getDeadlineGroups(sheet);

    const genericReminderIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.REMINDER,
    );
    const serviceTypeIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.SERVICE_TYPE,
    );
    const clientNameIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.CLIENT_NAME,
    );
    const staffEmailIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.STAFF_EMAIL,
    );
    const clientEmailIndex = findHeaderIndexByName(
      headers,
      cfg.HEADERS.CLIENT_EMAIL,
    );

    data.forEach((row, idx) => {
      const actualRow = idx + 2;
      const reminders = [];
      const deadlineSummaries = [];

      deadlineGroups.forEach((group) => {
        if (group.reminderCol) {
          const reminderValue = row[group.reminderCol - 1];
          if (reminderValue) reminders.push(reminderValue.toString());
        }

        if (group.remainingCol) {
          const remainingValue = row[group.remainingCol - 1];
          const remainingDays =
            remainingValue === "" || remainingValue === null
              ? null
              : Number(remainingValue);

          if (remainingDays !== null && !isNaN(remainingDays)) {
            deadlineSummaries.push({
              label: group.label,
              remainingDays: remainingDays,
            });
          }
        }
      });

      if (
        genericReminderIndex &&
        (!deadlineGroups.length || !reminders.length) &&
        row[genericReminderIndex - 1]
      ) {
        reminders.push(row[genericReminderIndex - 1].toString());
      }

      if (!reminders.length) return;

      const clientName = clientNameIndex ? row[clientNameIndex - 1] : "";
      const clientEmail = clientEmailIndex ? row[clientEmailIndex - 1] : "";
      const staffEmail = staffEmailIndex ? row[staffEmailIndex - 1] : "";
      const serviceType = serviceTypeIndex ? row[serviceTypeIndex - 1] : "";

      const recipients = [];
      if (staffEmail && staffEmail.toString().includes("@"))
        recipients.push(staffEmail);
      if (
        clientEmail &&
        clientEmail.toString().includes("@") &&
        clientEmail.toString() !== staffEmail.toString()
      ) {
        recipients.push(clientEmail);
      }

      if (recipients.length === 0) return;

      try {
        sendReminderEmail({
          recipients: recipients,
          clientName: clientName,
          serviceType: serviceType,
          reminders: reminders,
          deadlines: deadlineSummaries,
          sheetName: sheetName,
          rowNumber: actualRow,
        });

        emailCount++;
        logActivity(
          clientName,
          "Reminder Email Sent",
          `To: ${recipients.join(", ")}`,
        );
      } catch (e) {
        logActivity(clientName, "Email Failed", e.message);
      }
    });
  });

  SpreadsheetApp.getActive().toast(
    `Sent ${emailCount} reminder emails`,
    "Emails Sent",
  );
}

function sendReminderEmail(data) {
  const subject = `Reminder: ${data.clientName} - ${data.serviceType || "Project"}`;

  let body = `Dear Team,\n\n`;
  body += `This is an automated reminder regarding the following project:\n\n`;
  body += `Client: ${data.clientName}\n`;
  body += `Service: ${data.serviceType || "N/A"}\n`;
  body += `Sheet: ${data.sheetName}\n\n`;
  if (data.rowNumber) {
    body += `Row: ${data.rowNumber}\n\n`;
  }
  body += `ALERT:\n- ${data.reminders.join("\n- ")}\n\n`;

  if (data.deadlines && data.deadlines.length) {
    body += `Deadlines:\n`;
    data.deadlines.forEach((deadline) => {
      const message =
        deadline.remainingDays < 0
          ? `OVERDUE by ${Math.abs(deadline.remainingDays)} days`
          : `${deadline.remainingDays} days remaining`;
      body += `- ${deadline.label}: ${message}\n`;
    });
  }

  body += `\nPlease take appropriate action to ensure compliance.\n\n`;
  body += `---\n`;
  body += `CAR Special Projects Monitoring System\n`;
  body += `This is an automated message. Please do not reply directly to this email.`;

  const notificationEmails = getNotificationEmails();
  const toSet = new Set(data.recipients.map((e) => e.toString().toLowerCase()));
  const ccList = notificationEmails.filter((e) => !toSet.has(e.toLowerCase()));

  const mailOptions = {
    to: data.recipients.join(","),
    subject: subject,
    body: body,
    name: "CAR Monitoring System",
  };
  if (ccList.length > 0) mailOptions.cc = ccList.join(",");

  MailApp.sendEmail(mailOptions);
}

// ============================================================================
// AUTOMATION TRIGGERS
// ============================================================================

function onEdit(e) {
  if (!e) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const cfg = getConfig();

  if (!getProjectSheets().includes(sheetName)) return;

  const row = e.range.getRow();
  if (row < 2) return;

  const notaryCol = getColumnIndex(sheet, cfg.HEADERS.NOTARY_DATE);
  const remarksCol = getColumnIndex(sheet, cfg.HEADERS.REMARKS);
  const statusCol = getColumnIndex(sheet, cfg.HEADERS.STATUS);

  if (notaryCol && e.range.getColumn() === notaryCol) {
    updateRowDeadlines(sheet, row);
    updateRowReminder(sheet, row);

    const clientCol = getColumnIndex(sheet, cfg.HEADERS.CLIENT_NAME);
    const clientName = clientCol
      ? sheet.getRange(row, clientCol).getValue()
      : "Unknown";
    logActivity(
      clientName,
      "Notary Date Updated",
      `Row ${row} in ${sheetName}`,
    );
  }

  if (remarksCol && e.range.getColumn() === remarksCol) {
    autoUpdateStatus(sheet, row);
  }

  if (statusCol && e.range.getColumn() === statusCol) {
    updateRowReminder(sheet, row);
  }
}

function runDailyCheck() {
  const startTime = new Date();
  Logger.log("Starting daily check at " + startTime);

  updateAllDeadlines();
  generateAllReminders();
  sendReminderEmails();
  updateDashboard();

  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;

  logActivity("SYSTEM", "Daily Check Complete", `Duration: ${duration}s`);
  Logger.log("Daily check completed in " + duration + " seconds");
}

function createDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.find(
    (t) => t.getHandlerFunction() === "runDailyCheck",
  );

  if (existing) {
    SpreadsheetApp.getUi().alert("Daily trigger already exists");
    return;
  }

  ScriptApp.newTrigger("runDailyCheck")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getActive().toast(
    "Daily trigger created (runs at 8 AM)",
    "Trigger Created",
  );
  logActivity("SYSTEM", "Trigger Created", "Daily check at 8:00 AM");
}

function removeTriggers() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Remove Triggers",
    "Remove all automation triggers?",
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) return;

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));

  SpreadsheetApp.getActive().toast(
    `Removed ${triggers.length} triggers`,
    "Triggers Removed",
  );
  logActivity("SYSTEM", "Triggers Removed", `Count: ${triggers.length}`);
}

// ============================================================================
// DIAGNOSTICS
// ============================================================================

function showDiagnostics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cfg = getConfig();

  let report = "CAR Monitoring System Diagnostics\n";
  report += "=".repeat(40) + "\n\n";

  const working = getWorkingSheet();
  report += "Setup Configuration:\n";
  report += `  Working Sheet: ${working || "(all detected project sheets)"}\n`;
  report += `  Header Row: ${cfg.HEADER_ROW || "(auto-detect)"}\n`;
  report += `  Data Start Row: ${cfg.DATA_START_ROW || "(header row + 1)"}\n`;
  const notifEmails = getNotificationEmails();
  report += `  Notification Emails: ${notifEmails.length > 0 ? notifEmails.join(", ") : "(none)"}\n`;

  report += "\nSheets Status:\n";
  Object.values(cfg.SHEETS).forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    report += `  ${sheetName}: ${sheet ? "✓ EXISTS" : "✗ MISSING"}\n`;
  });

  report += "\nProject Sheets (automation targets):\n";
  const projectSheets = getProjectSheets();
  if (projectSheets.length === 0) {
    report += "  (none detected)\n";
  } else {
    projectSheets.forEach((name) => {
      report += `  ✓ ${name}\n`;
    });
  }

  report += "\nTriggers:\n";
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    report += "  No triggers configured\n";
  } else {
    triggers.forEach((t) => {
      report += `  - ${t.getHandlerFunction()} (${t.getTriggerSource()})\n`;
    });
  }

  report += "\nEmail Quota:\n";
  report += `  Remaining: ${MailApp.getRemainingDailyQuota()}\n`;

  report += "\nUser:\n";
  report += `  ${Session.getEffectiveUser().getEmail()}\n`;

  ui.alert(report);

  logActivity("SYSTEM", "Diagnostics Run", "Full system check performed");
}

// ============================================================================
// UTILITIES
// ============================================================================

function formatDate(date) {
  if (!date || !(date instanceof Date)) return "";
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "MMM dd, yyyy",
  );
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).map((row) => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = row[i]));
    return obj;
  });

  return { headers, rows };
}

// Quick setup function for initial run
function initializeSystem() {
  setupAllSheets();
  updateAllDeadlines();
  generateAllReminders();
  SpreadsheetApp.getUi().alert(
    "System initialized successfully!\n\nNext steps:\n1. Review sheets and add your data\n2. Set up email triggers if needed\n3. Run diagnostics to verify",
  );
}
