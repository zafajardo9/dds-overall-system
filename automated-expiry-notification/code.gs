/**
 * Alert System - Google Apps Script Automation
 * Multi-tab deadline monitoring with email alerts and comprehensive logging
 */

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

const CONFIG = {
  AUTOMATION_SHEET_NAMES: 'AUTOMATION_SHEET_NAMES',
  DEFAULT_CC_EMAILS: 'DEFAULT_CC_EMAILS',
  DATA_START_ROW: 2,
  TRIGGER_HOUR: 9,
  LOG_SHEET_NAME: 'Automation Logs',
  MAX_LOG_AGE_DAYS: 90,
  BATCH_SIZE: 50
};

const HEADERS = {
  NO: ['No.', 'No', 'ID', 'Number'],
  CLIENT_NAME: ['Client', 'Client Name', 'Company', 'Company Name'],
  TARGET_DATE: ['Target Date', 'Due Date', 'Expiry Date', 'Deadline', 'DST Due', 'CGT Due', 'Notary Date'],
  NOTICE_DATE: ['Notice Date', 'Alert Date', 'Reminder Date', 'Notice'],
  CLIENT_EMAIL: ['Client Email', 'Email', 'E-mail'],
  STAFF_EMAIL: ['Staff Email', 'Assigned Staff Email'],
  CC_EMAIL: ['CC Email', 'CC', 'Carbon Copy'],
  STATUS: ['Status', 'State', 'Progress'],
  REMARKS: ['Remarks', 'Notes', 'Comments', 'Description'],
  ATTACHMENTS: ['Attachments', 'Attachment', 'Files', 'File IDs'],
  DOCUMENT_TYPE: ['Document Type', 'Type', 'Category', 'Service Type']
};

const STATUS = {
  ACTIVE: 'Active',
  SENT: 'Sent',
  SEEN: 'Seen by client',
  COMPLETED: 'Completed',
  BLANK: ''
};

const LOG_COL = {
  TIMESTAMP: 1,
  TAB: 2,
  CLIENT: 3,
  ACTION: 4,
  DETAIL: 5
};

// ============================================================================
// MENU SYSTEM
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Alert System')
    .addSubMenu(ui.createMenu('Initialize / Configure')
      .addItem('Setup Wizard', 'runSetupWizard')
      .addItem('Configure Automation Sheets', 'configureAutomationSheets')
      .addItem('Configure Default CC Emails', 'configureDefaultCcEmails'))
    .addSubMenu(ui.createMenu('Daily Operations')
      .addItem('Run Daily Check Now', 'runDailyCheck')
      .addItem("Preview Today's Targets", 'previewTargetDates')
      .addItem('Check Schedule Status', 'checkScheduleStatus'))
    .addSubMenu(ui.createMenu('Diagnostics')
      .addItem('Inspect Row...', 'diagnosticInspectRow')
      .addItem('Send Test Email...', 'diagnosticSendTestRow')
      .addItem('Validate Column Headers', 'validateHeadersDialog')
      .addItem('View Full Logs', 'viewFullLogs'))
    .addSubMenu(ui.createMenu('Triggers')
      .addItem('Install Daily Trigger', 'installDailyTrigger')
      .addItem('Remove Trigger', 'removeTrigger')
      .addItem('Check Trigger Status', 'checkTriggerStatus'))
    .addSeparator()
    .addItem('About Alert System', 'showAboutDialog')
    .addToUi();
}

// ============================================================================
// SETUP & CONFIGURATION
// ============================================================================

function runSetupWizard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const introHtml = `
    <p>The Setup Wizard will guide you through configuring the Alert System for your sheets.</p>
    <p><strong>Steps:</strong></p>
    <ol>
      <li>Select which tabs to monitor</li>
      <li>Confirm your data starting row</li>
      <li>Verify column headers are detected</li>
      <li>Send a test email to verify</li>
    </ol>
    <p>Ready to begin?</p>
  `;
  
  const response = ui.alert('Alert System - Setup Wizard', introHtml, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  
  configureAutomationSheets();
  
  const confirmHtml = `
    <p>Setup complete! The system is now configured to monitor your selected sheets.</p>
    <p><strong>Next steps:</strong></p>
    <ul>
      <li>Run "Daily Operations > Run Daily Check Now" to test</li>
      <li>Visit "Diagnostics" for troubleshooting tools</li>
      <li>Check "Automation Logs" tab for activity history</li>
    </ul>
  `;
  
  ui.alert('Setup Complete', confirmHtml, ui.ButtonSet.OK);
}

function configureAutomationSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const sheetNames = sheets
    .filter(s => s.getName() !== CONFIG.LOG_SHEET_NAME)
    .map(s => s.getName());
  
  if (sheetNames.length === 0) {
    ui.alert('No sheets available for automation');
    return;
  }
  
  const currentNames = getConfiguredSheetNames();
  const currentDisplay = currentNames.length > 0 
    ? `Currently configured: ${currentNames.join(', ')}` 
    : 'No sheets currently configured';
  
  const selected = selectMultipleFromList(
    `Select sheets to monitor for deadlines:\n\n${currentDisplay}`,
    sheetNames,
    currentNames
  );
  
  if (selected === null) return;
  
  setConfiguredSheetNames(selected);
  ensureLogsSheet();
  
  ui.alert(`Configured ${selected.length} sheet(s) for automation monitoring.`);
}

function getConfiguredSheetNames() {
  const props = PropertiesService.getDocumentProperties();
  const stored = props.getProperty(CONFIG.AUTOMATION_SHEET_NAMES);
  if (!stored) return [];
  try {
    const parsed = JSON.parse(stored);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch (e) {
    return [stored];
  }
}

function setConfiguredSheetNames(names) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(CONFIG.AUTOMATION_SHEET_NAMES, JSON.stringify(names));
}

function getDefaultCcEmails() {
  const props = PropertiesService.getDocumentProperties();
  const stored = props.getProperty(CONFIG.DEFAULT_CC_EMAILS);
  if (!stored) return [];
  try {
    const parsed = JSON.parse(stored);
    return normalizeEmailList(Array.isArray(parsed) ? parsed : [parsed]);
  } catch (e) {
    return normalizeEmailList(String(stored));
  }
}

function setDefaultCcEmails(emails) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(CONFIG.DEFAULT_CC_EMAILS, JSON.stringify(normalizeEmailList(emails)));
}

function configureDefaultCcEmails() {
  const ui = SpreadsheetApp.getUi();
  const current = getDefaultCcEmails();
  const currentDisplay = current.length > 0 ? current.join(', ') : '(none)';
  const response = ui.prompt(
    'Configure Default CC Emails',
    `Enter email addresses to CC on every outgoing email.\nUse commas, semicolons, or new lines.\nLeave blank to clear the list.\n\nCurrent: ${currentDisplay}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText();
  const emails = normalizeEmailList(input);
  const invalid = emails.filter(email => !validateEmail(email));

  if (invalid.length > 0) {
    ui.alert(`Invalid email address(es): ${invalid.join(', ')}`);
    return;
  }

  setDefaultCcEmails(emails);
  ui.alert(
    'Default CC Emails Saved',
    emails.length > 0 ? `These addresses will be CC'd on future emails:\n${emails.join(', ')}` : 'Default CC list cleared.',
    ui.ButtonSet.OK
  );
}

function resolveAutomationSheets(ss) {
  const names = getConfiguredSheetNames();
  const result = [];
  
  for (const name of names) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      result.push({ sheetName: name, sheet: sheet });
    }
  }
  
  return result;
}

function promptSelectConfiguredSheet(ss, title) {
  const configs = resolveAutomationSheets(ss);
  if (configs.length === 0) return null;
  if (configs.length === 1) return configs[0];
  
  const names = configs.map(c => c.sheetName);
  const selected = selectFromList(title || 'Select a sheet:', names);
  if (!selected) return null;
  
  return configs.find(c => c.sheetName === selected);
}

function selectFromList(prompt, items) {
  const ui = SpreadsheetApp.getUi();
  const itemList = items.map((item, i) => `${i + 1}. ${item}`).join('\n');
  
  const response = ui.prompt(
    prompt,
    `${itemList}\n\nEnter number:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  
  const index = parseInt(response.getResponseText()) - 1;
  if (index < 0 || index >= items.length) return null;
  
  return items[index];
}

function selectMultipleFromList(prompt, items, preSelected) {
  const ui = SpreadsheetApp.getUi();
  const itemList = items.map((item, i) => {
    const checked = preSelected.includes(item) ? '[x]' : '[ ]';
    return `${i + 1}. ${checked} ${item}`;
  }).join('\n');
  
  const response = ui.prompt(
    prompt,
    `${itemList}\n\nEnter numbers (comma-separated):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  
  const input = response.getResponseText();
  if (!input.trim()) return [];
  
  const indices = input.split(',').map(s => parseInt(s.trim()) - 1);
  const selected = [];
  
  for (const idx of indices) {
    if (idx >= 0 && idx < items.length) {
      selected.push(items[idx]);
    }
  }
  
  return selected;
}

// ============================================================================
// DATA ACCESS LAYER
// ============================================================================

function getHeaderMap(sheet) {
  const headerRow = CONFIG.DATA_START_ROW - 1;
  if (headerRow < 1) return {};
  
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  
  for (const [key, aliases] of Object.entries(HEADERS)) {
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i] || '').trim();
      if (aliases.some(alias => header.toLowerCase() === alias.toLowerCase())) {
        map[key] = i + 1;
        break;
      }
    }
  }
  
  return map;
}

function getDataRows(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < CONFIG.DATA_START_ROW) return [];
  
  const data = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, lastCol).getValues();
  const rows = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const isEmpty = row.every(cell => cell === '' || cell === null || cell === undefined);
    
    if (!isEmpty) {
      rows.push({
        rowIndex: CONFIG.DATA_START_ROW + i,
        data: row
      });
    }
  }
  
  return rows;
}

function getRowByNumber(sheet, rowNumber) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return null;
  
  try {
    const data = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
    return {
      rowIndex: rowNumber,
      data: data
    };
  } catch (e) {
    return null;
  }
}

function findRowNumberByNo(sheet, noValue) {
  const headers = getHeaderMap(sheet);
  if (!headers.NO) return null;
  
  const rows = getDataRows(sheet);
  const matches = [];
  
  for (const row of rows) {
    const rowNo = normalizeNoValue(row.data[headers.NO - 1]);
    const searchNo = normalizeNoValue(noValue);
    
    if (rowNo === searchNo) {
      matches.push(row.rowIndex);
    }
  }
  
  return matches.length > 0 ? matches[0] : null;
}

function normalizeNoValue(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().replace(/^0+/, '') || '0';
}

function isSameNoValue(val1, val2) {
  return normalizeNoValue(val1) === normalizeNoValue(val2);
}

// ============================================================================
// DATE ENGINE
// ============================================================================

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  
  const str = String(value).trim();
  if (!str) return null;
  
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) return parsed;
  
  const excelDate = parseExcelDate(value);
  if (excelDate) return excelDate;
  
  return null;
}

function parseExcelDate(value) {
  if (typeof value === 'number' && value > 30000 && value < 50000) {
    const epoch = new Date(1899, 11, 30);
    const date = new Date(epoch.getTime() + value * 24 * 60 * 60 * 1000);
    return isNaN(date.getTime()) ? null : date;
  }
  return null;
}

function parseNoticeOffset(value, targetDate) {
  if (!value) return null;
  if (!targetDate) return null;
  
  const str = String(value).toLowerCase().trim();
  
  const relativeMatch = str.match(/^(\d+)\s*(?:days?\s*)?(?:before|prior)?$/);
  if (relativeMatch) {
    const days = parseInt(relativeMatch[1]);
    const noticeDate = new Date(targetDate);
    noticeDate.setDate(noticeDate.getDate() - days);
    return noticeDate;
  }
  
  if (str.includes('on expiry') || str === '0' || str === 'on day') {
    return new Date(targetDate);
  }
  
  const absoluteDate = parseDate(value);
  if (absoluteDate) return absoluteDate;
  
  return null;
}

function isTargetDateDue(targetDate, noticeDate) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const target = new Date(targetDate);
  target.setHours(0, 0, 0, 0);
  
  if (noticeDate) {
    const notice = new Date(noticeDate);
    notice.setHours(0, 0, 0, 0);
    if (today < notice) return false;
  }
  
  return today >= target;
}

function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getDaysRemaining(targetDate) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const target = new Date(targetDate);
  target.setHours(0, 0, 0, 0);
  
  const diff = target.getTime() - today.getTime();
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}

// ============================================================================
// STATUS MANAGER
// ============================================================================

function getStatus(sheet, rowIndex, headerMap) {
  if (!headerMap.STATUS) return STATUS.BLANK;
  
  const value = sheet.getRange(rowIndex, headerMap.STATUS).getValue();
  return String(value || '').trim();
}

function setStatus(sheet, rowIndex, headerMap, status) {
  if (!headerMap.STATUS) return false;
  
  sheet.getRange(rowIndex, headerMap.STATUS).setValue(status);
  return true;
}

function isEligibleForSend(status) {
  const s = String(status || '').trim();
  return s === STATUS.ACTIVE || s === STATUS.BLANK || s === '';
}

function setStaffEmail(sheet, rowIndex, headerMap, email) {
  if (!headerMap.STAFF_EMAIL) return false;
  
  sheet.getRange(rowIndex, headerMap.STAFF_EMAIL).setValue(email);
  return true;
}

// ============================================================================
// EMAIL SERVICE
// ============================================================================

function buildEmailBody(rowData, headerMap, template) {
  const placeholders = {
    '[Document Type]': getCellValue(rowData, headerMap.DOCUMENT_TYPE) || 'Document',
    '[Client]': getCellValue(rowData, headerMap.CLIENT_NAME) || 'Client',
    '[Due Date]': formatDate(getCellValue(rowData, headerMap.TARGET_DATE)),
    '[Days Remaining]': getDaysRemaining(getCellValue(rowData, headerMap.TARGET_DATE)),
    '[Remarks]': getCellValue(rowData, headerMap.REMARKS) || ''
  };
  
  let body = template || buildDefaultTemplate();
  
  for (const [key, value] of Object.entries(placeholders)) {
    body = body.replace(new RegExp(key.replace(/[\[\]]/g, '\\$&'), 'gi'), String(value));
  }
  
  return body;
}

function buildDefaultTemplate() {
  return `Dear Recipient,

This is a reminder regarding the following document:

Document: [Document Type]
Client: [Client]
Due Date: [Due Date]
Days Remaining: [Days Remaining]
Remarks: [Remarks]

Please take appropriate action.

Best regards,
Alert System`;
}

function parseEmailWithCC(emailValue) {
  if (!emailValue) return { primary: null, cc: [] };
  
  const emails = String(emailValue).split(/[,;]/).map(e => e.trim()).filter(e => e);
  if (emails.length === 0) return { primary: null, cc: [] };
  
  return {
    primary: emails[0],
    cc: emails.slice(1)
  };
}

function normalizeEmailList(value) {
  if (Array.isArray(value)) {
    return dedupeEmailList(value.map(entry => String(entry || '').trim()).filter(entry => entry));
  }

  if (!value) return [];

  const emails = String(value)
    .split(/[,;\n]/)
    .map(email => email.trim())
    .filter(email => email);

  return dedupeEmailList(emails);
}

function dedupeEmailList(emails) {
  const seen = {};
  const result = [];

  for (const email of emails) {
    const key = String(email || '').trim().toLowerCase();
    if (!key || seen[key]) continue;
    seen[key] = true;
    result.push(String(email).trim());
  }

  return result;
}

function buildRecipientData(rowData, headerMap) {
  const emailData = parseEmailWithCC(getCellValue(rowData, headerMap.CLIENT_EMAIL));
  const rowCc = headerMap.CC_EMAIL ? normalizeEmailList(getCellValue(rowData, headerMap.CC_EMAIL)) : [];
  const defaultCc = getDefaultCcEmails();
  const cc = dedupeEmailList([].concat(emailData.cc, rowCc, defaultCc))
    .filter(email => String(email).toLowerCase() !== String(emailData.primary || '').toLowerCase());

  return {
    primary: emailData.primary,
    cc: cc
  };
}

function validateEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(String(email));
}

function splitAttachmentEntries(value) {
  if (!value) return [];
  
  const entries = String(value)
    .split(/[,\n]/)
    .map(e => e.trim())
    .filter(e => e);
  
  return [...new Set(entries)];
}

function resolveAttachments(fileEntries) {
  const result = {
    blobs: [],
    warnings: [],
    fatalError: null
  };
  
  if (!fileEntries || fileEntries.length === 0) return result;
  
  for (const entry of fileEntries) {
    try {
      const fileId = extractFileId(entry);
      if (!fileId) {
        result.warnings.push(`Could not extract file ID from: ${entry}`);
        continue;
      }
      
      const file = DriveApp.getFileById(fileId);
      const blob = file.getBlob();
      result.blobs.push({
        blob: blob,
        name: file.getName(),
        fileId: fileId
      });
    } catch (e) {
      result.warnings.push(`Error accessing ${entry}: ${e.message}`);
    }
  }
  
  return result;
}

function extractFileId(entry) {
  const patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /open\?id=([a-zA-Z0-9_-]+)/,
    /file\/d\/([a-zA-Z0-9_-]+)/,
    /^([a-zA-Z0-9_-]{25,})$/
  ];
  
  for (const pattern of patterns) {
    const match = entry.match(pattern);
    if (match) return match[1];
  }
  
  return null;
}

function sendEmailWithAttachments(recipient, cc, subject, body, attachments) {
  const options = {
    htmlBody: body.replace(/\n/g, '<br>'),
    name: 'Alert System'
  };
  
  if (cc && cc.length > 0) {
    options.cc = cc.join(',');
  }
  
  if (attachments && attachments.length > 0) {
    options.attachments = attachments.map(a => a.blob);
  }
  
  MailApp.sendEmail(recipient, subject, body, options);
}

function checkEmailQuota() {
  const remaining = MailApp.getRemainingDailyQuota();
  return {
    remaining: remaining,
    canSend: remaining > 0
  };
}

// ============================================================================
// LOGGING SERVICE
// ============================================================================

function ensureLogsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logsSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  
  if (!logsSheet) {
    logsSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    logsSheet.appendRow(['Timestamp', 'Tab', 'Client', 'Action', 'Detail']);
    logsSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  } else {
    const headers = logsSheet.getRange(1, 1, 1, logsSheet.getLastColumn()).getValues()[0];
    if (headers.length < 5 || !headers[1] || String(headers[1]).toLowerCase() !== 'tab') {
      logsSheet.insertColumnAfter(1);
      logsSheet.getRange(1, 2).setValue('Tab');
    }
  }
  
  return logsSheet;
}

function appendLog(logsSheet, tabName, clientName, action, detail) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  logsSheet.appendRow([timestamp, tabName || '', clientName || '', action, detail || '']);
}

function cleanupOldLogs() {
  const logsSheet = ensureLogsSheet();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - CONFIG.MAX_LOG_AGE_DAYS);
  
  const data = logsSheet.getDataRange().getValues();
  const rowsToDelete = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = new Date(data[i][0]);
    if (rowDate < cutoff) {
      rowsToDelete.push(i + 1);
    }
  }
  
  for (const row of rowsToDelete) {
    logsSheet.deleteRow(row);
  }
}

function getCellValue(rowData, colIndex) {
  if (!colIndex || colIndex < 1 || colIndex > rowData.length) return '';
  return rowData[colIndex - 1] || '';
}

// ============================================================================
// CORE AUTOMATION FLOW
// ============================================================================

function runDailyCheck() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logsSheet = ensureLogsSheet();
  const configs = resolveAutomationSheets(ss);
  
  if (configs.length === 0) {
    ui.alert('No sheets configured. Run Setup Wizard first.');
    return;
  }
  
  const quota = checkEmailQuota();
  if (!quota.canSend) {
    appendLog(logsSheet, 'SYSTEM', 'System', 'ERROR', 'Email quota exhausted');
    ui.alert('Email quota exhausted. Cannot send today.');
    return;
  }
  
  let totalProcessed = 0;
  let totalSent = 0;
  let totalSkipped = 0;
  let totalErrors = 0;
  
  for (const config of configs) {
    const result = processSheet(config.sheet, config.sheetName, logsSheet);
    totalProcessed += result.processed;
    totalSent += result.sent;
    totalSkipped += result.skipped;
    totalErrors += result.errors;
  }
  
  const summary = `Processed ${totalProcessed} rows across ${configs.length} sheet(s).\nSent: ${totalSent}, Skipped: ${totalSkipped}, Errors: ${totalErrors}`;
  appendLog(logsSheet, 'SYSTEM', 'System', 'SUMMARY', summary);
  
  if (configs.length > 1) {
    ui.alert('Multi-Tab Check Complete', summary, ui.ButtonSet.OK);
  }
}

function processSheet(sheet, sheetName, logsSheet) {
  const result = { processed: 0, sent: 0, skipped: 0, errors: 0 };
  const headerMap = getHeaderMap(sheet);
  
  if (!headerMap.TARGET_DATE || !headerMap.CLIENT_EMAIL) {
    appendLog(logsSheet, sheetName, 'System', 'ERROR', 'Missing required columns (Target Date or Client Email)');
    result.errors++;
    return result;
  }
  
  const rows = getDataRows(sheet);
  
  for (const row of rows) {
    result.processed++;
    
    try {
      const targetDate = parseDate(getCellValue(row.data, headerMap.TARGET_DATE));
      const noticeDate = headerMap.NOTICE_DATE 
        ? parseNoticeOffset(getCellValue(row.data, headerMap.NOTICE_DATE), targetDate)
        : null;
      const status = getStatus(sheet, row.rowIndex, headerMap);
      
      if (!isEligibleForSend(status)) {
        result.skipped++;
        continue;
      }
      
      if (!targetDate) {
        result.skipped++;
        continue;
      }
      
      if (!isTargetDateDue(targetDate, noticeDate)) {
        result.skipped++;
        continue;
      }
      
      const emailData = buildRecipientData(row.data, headerMap);
      if (!emailData.primary || !validateEmail(emailData.primary)) {
        appendLog(logsSheet, sheetName, getCellValue(row.data, headerMap.CLIENT_NAME), 'ERROR', 'Invalid email');
        result.errors++;
        continue;
      }
      
      const template = getCellValue(row.data, headerMap.REMARKS) || '';
      const subject = `Reminder: ${getCellValue(row.data, headerMap.DOCUMENT_TYPE) || 'Document'} Due - ${getCellValue(row.data, headerMap.CLIENT_NAME) || 'Client'}`;
      const body = buildEmailBody(row.data, headerMap, template);
      
      const attachments = [];
      if (headerMap.ATTACHMENTS) {
        const attValue = getCellValue(row.data, headerMap.ATTACHMENTS);
        const entries = splitAttachmentEntries(attValue);
        const resolved = resolveAttachments(entries);
        attachments.push(...resolved.blobs);
      }
      
      sendEmailWithAttachments(emailData.primary, emailData.cc, subject, body, attachments);
      
      setStatus(sheet, row.rowIndex, headerMap, STATUS.SENT);
      
      const sender = Session.getEffectiveUser().getEmail();
      setStaffEmail(sheet, row.rowIndex, headerMap, sender);
      
      const detail = `To: ${emailData.primary}${emailData.cc.length > 0 ? ', CC: ' + emailData.cc.join(', ') : ''}${attachments.length > 0 ? ', Attachments: ' + attachments.length : ''}${sender ? ', Sender: ' + sender : ''}`;
      appendLog(logsSheet, sheetName, getCellValue(row.data, headerMap.CLIENT_NAME), 'SENT', detail);
      
      result.sent++;
      
    } catch (e) {
      appendLog(logsSheet, sheetName, getCellValue(row.data, headerMap.CLIENT_NAME), 'ERROR', e.message);
      result.errors++;
    }
  }
  
  return result;
}

// ============================================================================
// DIAGNOSTIC TOOLS
// ============================================================================

function previewTargetDates() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = promptSelectConfiguredSheet(ss, 'Select sheet to preview:');
  if (!config) return;
  
  const sheet = config.sheet;
  const sheetName = config.sheetName;
  const headerMap = getHeaderMap(sheet);
  const rows = getDataRows(sheet);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let preview = `Preview for ${sheetName}:\n\n`;
  let count = 0;
  
  for (const row of rows) {
    const targetDate = parseDate(getCellValue(row.data, headerMap.TARGET_DATE));
    const noticeDate = headerMap.NOTICE_DATE 
      ? parseNoticeOffset(getCellValue(row.data, headerMap.NOTICE_DATE), targetDate)
      : null;
    const status = getStatus(sheet, row.rowIndex, headerMap);
    
    if (!targetDate) continue;
    if (!isEligibleForSend(status)) continue;
    if (!isTargetDateDue(targetDate, noticeDate)) continue;
    
    count++;
    const client = getCellValue(row.data, headerMap.CLIENT_NAME) || 'Unknown';
    const dueStr = formatDate(targetDate);
    const daysLeft = getDaysRemaining(targetDate);
    
    preview += `${count}. Row ${row.rowIndex}: ${client} (Due: ${dueStr}, ${daysLeft} days)\n`;
  }
  
  if (count === 0) {
    preview += 'No items would be sent today (all caught up).\n';
  } else {
    preview += `\nTotal items that would send: ${count}`;
  }
  
  ui.alert("Today's Targets (Due/Overdue)", preview, ui.ButtonSet.OK);
}

function diagnosticInspectRow() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = promptSelectConfiguredSheet(ss, 'Select sheet:');
  if (!config) return;
  
  const response = ui.prompt('Enter row number or No. value:');
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const input = response.getResponseText().trim();
  let rowIndex = parseInt(input);
  
  if (isNaN(rowIndex)) {
    rowIndex = findRowNumberByNo(config.sheet, input);
    if (!rowIndex) {
      ui.alert(`Could not find row with No. = "${input}"`);
      return;
    }
  }
  
  const row = getRowByNumber(config.sheet, rowIndex);
  if (!row) {
    ui.alert(`Row ${rowIndex} not found.`);
    return;
  }
  
  const headerMap = getHeaderMap(config.sheet);
  const targetDate = parseDate(getCellValue(row.data, headerMap.TARGET_DATE));
  const noticeDate = headerMap.NOTICE_DATE 
    ? parseNoticeOffset(getCellValue(row.data, headerMap.NOTICE_DATE), targetDate)
    : null;
  const status = getStatus(config.sheet, rowIndex, headerMap);
  
  let report = `Row ${rowIndex} Inspection:\n\n`;
  report += `Client: ${getCellValue(row.data, headerMap.CLIENT_NAME) || 'N/A'}\n`;
  report += `Status: ${status || '(blank)'}\n`;
  report += `Target Date: ${formatDate(targetDate) || 'N/A'}\n`;
  report += `Notice Date: ${formatDate(noticeDate) || 'N/A'}\n`;
  report += `Send Eligible: ${isEligibleForSend(status) && targetDate && isTargetDateDue(targetDate, noticeDate) ? 'YES' : 'NO'}\n`;
  report += `Client Email: ${getCellValue(row.data, headerMap.CLIENT_EMAIL) || 'N/A'}\n`;
  if (headerMap.CC_EMAIL) {
    report += `Row CC: ${getCellValue(row.data, headerMap.CC_EMAIL) || 'N/A'}\n`;
  }
  const defaultCc = getDefaultCcEmails();
  report += `Default CC: ${defaultCc.length > 0 ? defaultCc.join(', ') : 'N/A'}\n`;
  
  if (headerMap.ATTACHMENTS) {
    const attValue = getCellValue(row.data, headerMap.ATTACHMENTS);
    if (attValue) {
      const entries = splitAttachmentEntries(attValue);
      const resolved = resolveAttachments(entries);
      report += `\nAttachments: ${resolved.blobs.length}/${entries.length} OK`;
      if (resolved.warnings.length > 0) {
        report += `\nWarnings: ${resolved.warnings.join('; ')}`;
      }
    }
  }
  
  ui.alert(report);
}

function diagnosticSendTestRow() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = promptSelectConfiguredSheet(ss, 'Select sheet:');
  if (!config) return;
  
  const response = ui.prompt('Enter row number or No. value:');
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const input = response.getResponseText().trim();
  let rowIndex = parseInt(input);
  
  if (isNaN(rowIndex)) {
    const allRows = getDataRows(config.sheet);
    const matches = [];
    for (const r of allRows) {
      if (isSameNoValue(r.data[0], input)) matches.push(r.rowIndex);
    }
    
    if (matches.length === 0) {
      ui.alert(`No row found with No. = "${input}"`);
      return;
    }
    if (matches.length > 1) {
      ui.alert(`Warning: ${matches.length} rows have No. = "${input}". Using first match.`);
    }
    rowIndex = matches[0];
  }
  
  const row = getRowByNumber(config.sheet, rowIndex);
  if (!row) {
    ui.alert(`Row ${rowIndex} not found.`);
    return;
  }
  
  const headerMap = getHeaderMap(config.sheet);
  const emailData = buildRecipientData(row.data, headerMap);
  
  if (!emailData.primary) {
    ui.alert('No email address found in this row.');
    return;
  }
  
  const confirm = ui.alert(
    'Send Test Email',
    `Send test to: ${emailData.primary}${emailData.cc.length > 0 ? '\nCC: ' + emailData.cc.join(', ') : ''}\n\nProceed?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const logsSheet = ensureLogsSheet();
  
  try {
    const template = getCellValue(row.data, headerMap.REMARKS) || '';
    const subject = `[TEST] Reminder: ${getCellValue(row.data, headerMap.DOCUMENT_TYPE) || 'Document'} Due - ${getCellValue(row.data, headerMap.CLIENT_NAME) || 'Client'}`;
    const body = buildEmailBody(row.data, headerMap, template);
    
    const attachments = [];
    if (headerMap.ATTACHMENTS) {
      const attValue = getCellValue(row.data, headerMap.ATTACHMENTS);
      const entries = splitAttachmentEntries(attValue);
      const resolved = resolveAttachments(entries);
      attachments.push(...resolved.blobs);
      
      if (resolved.warnings.length > 0) {
        ui.alert('Attachment Warnings', resolved.warnings.join('\n'), ui.ButtonSet.OK);
      }
    }
    
    sendEmailWithAttachments(emailData.primary, emailData.cc, subject, body, attachments);
    
    const sender = Session.getEffectiveUser().getEmail();
    setStaffEmail(config.sheet, rowIndex, headerMap, sender);
    
    const detail = `[TEST] To: ${emailData.primary}${emailData.cc.length > 0 ? ', CC: ' + emailData.cc.join(', ') : ''}${attachments.length > 0 ? ', Attachments: ' + attachments.length : ''}${sender ? ', Sender: ' + sender : ''}`;
    appendLog(logsSheet, config.sheetName, getCellValue(row.data, headerMap.CLIENT_NAME), 'TEST_SENT', detail);
    
    ui.alert('Test email sent successfully.');
  } catch (e) {
    appendLog(logsSheet, config.sheetName, getCellValue(row.data, headerMap.CLIENT_NAME), 'TEST_ERROR', e.message);
    ui.alert('Error sending test: ' + e.message);
  }
}

function validateHeadersDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = promptSelectConfiguredSheet(ss, 'Select sheet to validate:');
  if (!config) return;
  
  const headerMap = getHeaderMap(config.sheet);
  const sheetHeaders = config.sheet.getRange(CONFIG.DATA_START_ROW - 1, 1, 1, config.sheet.getLastColumn()).getValues()[0];
  
  let report = `Header Validation for ${config.sheetName}:\n\n`;
  report += `Found headers: ${sheetHeaders.filter(h => h).join(', ')}\n\n`;
  report += 'Mapped columns:\n';
  
  for (const [key, col] of Object.entries(headerMap)) {
    const headerName = col ? sheetHeaders[col - 1] : 'Not found';
    report += `- ${key}: ${headerName} (col ${col || 'N/A'})\n`;
  }
  
  const missing = Object.keys(HEADERS).filter(k => !headerMap[k] && ['TARGET_DATE', 'CLIENT_EMAIL', 'CLIENT_NAME'].includes(k));
  if (missing.length > 0) {
    report += `\nMissing recommended columns: ${missing.join(', ')}`;
  }
  
  ui.alert(report);
}

function viewFullLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logsSheet = ensureLogsSheet();
  ss.setActiveSheet(logsSheet);
}

// ============================================================================
// SCHEDULE & TRIGGERS
// ============================================================================

function checkScheduleStatus() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configs = resolveAutomationSheets(ss);
  
  let status = `Alert System Status:\n\n`;
  status += `Configured Sheets (${configs.length}):\n`;
  
  for (const config of configs) {
    const rowCount = getDataRows(config.sheet).length;
    status += `- ${config.sheetName}: ${rowCount} data rows\n`;
  }
  
  const triggers = ScriptApp.getProjectTriggers();
  const dailyTriggers = triggers.filter(t => t.getHandlerFunction() === 'runDailyCheck');
  
  status += `\nTriggers: ${dailyTriggers.length} installed`;
  if (dailyTriggers.length > 0) {
    const nextRun = dailyTriggers[0].getNextHandlerCallTime;
    status += ` (Next: ${nextRun || 'N/A'})`;
  }
  
  ui.alert(status);
}

function installDailyTrigger() {
  const ui = SpreadsheetApp.getUi();
  
  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.filter(t => t.getHandlerFunction() === 'runDailyCheck');
  
  if (existing.length > 0) {
    ui.alert('Daily trigger already installed.');
    return;
  }
  
  ScriptApp.newTrigger('runDailyCheck')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TRIGGER_HOUR)
    .create();
  
  ui.alert(`Daily trigger installed. Will run at ${CONFIG.TRIGGER_HOUR}:00.`);
}

function removeTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'runDailyCheck') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  ui.alert('All daily triggers removed.');
}

function checkTriggerStatus() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  const dailyTriggers = triggers.filter(t => t.getHandlerFunction() === 'runDailyCheck');
  
  if (dailyTriggers.length === 0) {
    ui.alert('No daily trigger installed. Use "Install Daily Trigger" to enable automation.');
  } else {
    const times = dailyTriggers.map(t => {
      const next = t.getNextHandlerCallTime();
      return next ? Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : 'Unknown';
    });
    ui.alert(`${dailyTriggers.length} trigger(s) installed. Next run(s): ${times.join(', ')}`);
  }
}

// ============================================================================
// ABOUT & HELP
// ============================================================================

function showAboutDialog() {
  const ui = SpreadsheetApp.getUi();
  const html = `
    <p><strong>Alert System v1.0</strong></p>
    <p>Multi-tab deadline monitoring with automated email notifications.</p>
    <p><strong>Key Features:</strong></p>
    <ul>
      <li>Monitor multiple sheet tabs simultaneously</li>
      <li>Flexible column mapping (works with your existing structure)</li>
      <li>Email with CC support and attachment handling</li>
      <li>Comprehensive audit logging</li>
      <li>Built-in diagnostics and test tools</li>
    </ul>
    <p><strong>Quick Start:</strong></p>
    <ol>
      <li>Run "Initialize / Configure > Setup Wizard"</li>
      <li>Select the sheets to monitor</li>
      <li>Run "Run Daily Check Now" to test</li>
      <li>Install trigger for automation</li>
    </ol>
    <p>All activity is logged to the "Automation Logs" tab.</p>
  `;
  ui.alert('About Alert System', html, ui.ButtonSet.OK);
}
