var CONFIG = {
  MENU_NAME: "Checklist Sync",
  SYNC_FUNCTION_NAME: "syncConfiguredSheets",
  TIMEZONE_FALLBACK: "Asia/Manila",
  TRIGGER_INTERVAL_MINUTES: 1,
  PROPERTY_PREFIX: "sheet_config:",
  MIN_CHECKLIST_TEXT_LENGTH: 3,
  PROMPT_TITLE_DRIVE: "Link Gdrive link for this Gsheet",
  HIGHLIGHT_CHECKLIST: "#fff2cc",
  HIGHLIGHT_LINKS: "#cfe2f3",
  HIGHLIGHT_REMARKS: "#d9ead3",
  HIGHLIGHT_COPIES: "#f4cccc",
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.MENU_NAME)
    .addItem(
      "Link Gdrive link for this Gsheet",
      "linkDriveFolderForCurrentSheet",
    )
    .addItem(
      "Setup the checklist column and data",
      "setupChecklistColumnsForSpreadsheet",
    )
    .addItem("Test detected checklist columns", "testDetectedChecklistColumns")
    .addItem("Setup trigger for linking files", "setupTriggerForLinkingFiles")
    .addItem("View setup status", "viewSetupStatus")
    .addItem("Sync linked files now", CONFIG.SYNC_FUNCTION_NAME)
    .addToUi();
}

function linkDriveFolderForCurrentSheet() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var existingConfig = loadSheetConfig_(
    spreadsheet.getId(),
    sheet.getSheetId(),
  );
  var prompt = ui.prompt(
    CONFIG.PROMPT_TITLE_DRIVE,
    'Paste the Google Drive folder link for the current sheet: "' +
      sheet.getName() +
      '"' +
      (existingConfig.folderUrl
        ? "\n\nCurrent linked folder:\n" + existingConfig.folderUrl
        : ""),
    ui.ButtonSet.OK_CANCEL,
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Drive folder linking was cancelled.");
    return;
  }

  var folderUrl = String(prompt.getResponseText() || "").trim();
  var folderId = extractFolderId(folderUrl);
  if (!folderId) {
    ui.alert(
      "The Google Drive folder link is invalid. Please paste a valid folder URL.",
    );
    return;
  }

  try {
    DriveApp.getFolderById(folderId).getName();
  } catch (error) {
    ui.alert(
      "The folder could not be accessed. Please check sharing permissions and the folder link.",
    );
    return;
  }

  var config = existingConfig;
  config.spreadsheetId = spreadsheet.getId();
  config.sheetId = sheet.getSheetId();
  config.sheetName = sheet.getName();
  config.folderUrl = folderUrl;
  config.folderId = folderId;
  saveSheetConfig_(config);

  ui.alert(
    'Google Drive folder linked successfully for sheet "' +
      sheet.getName() +
      '".\n\n' +
      "Saved folder URL:\n" +
      folderUrl,
  );
}

function setupChecklistColumnsForSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var result = detectChecklistLayoutsAcrossSheets_(spreadsheet);

  result.detected.forEach(function (item) {
    var config = loadSheetConfig_(spreadsheet.getId(), item.sheet.getSheetId());
    config.spreadsheetId = spreadsheet.getId();
    config.sheetId = item.sheet.getSheetId();
    config.sheetName = item.sheet.getName();
    config.headerRow = item.detection.headerRow;
    config.checklistColumnIndex = item.detection.checklistColumnIndex;
    config.linksColumnIndex = item.detection.linksColumnIndex;
    config.remarksColumnIndex = item.detection.remarksColumnIndex;
    config.numberOfCopiesColumnIndex =
      item.detection.numberOfCopiesColumnIndex || null;
    saveSheetConfig_(config);
    highlightDetectedColumns_(item.sheet, item.detection);
  });

  if (!result.detected.length) {
    ui.alert("No checklist headers were detected in any sheet tab.");
    return;
  }

  ui.alert(buildDetectionSummary_(result));
}

function testDetectedChecklistColumns() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var result = detectChecklistLayoutsAcrossSheets_(spreadsheet);

  if (!result.detected.length) {
    ui.alert("No checklist headers were detected to test.");
    return;
  }

  result.detected.forEach(function (item) {
    highlightDetectedColumns_(item.sheet, item.detection);
  });

  ui.alert(buildDetectionSummary_(result));
}

function setupTriggerForLinkingFiles() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var hasExistingTrigger = triggers.some(function (trigger) {
    return trigger.getHandlerFunction() === CONFIG.SYNC_FUNCTION_NAME;
  });

  if (!hasExistingTrigger) {
    ScriptApp.newTrigger(CONFIG.SYNC_FUNCTION_NAME)
      .timeBased()
      .everyMinutes(CONFIG.TRIGGER_INTERVAL_MINUTES)
      .create();
    ui.alert(
      "Auto-link trigger created.\n\n" +
        "Handler: " +
        CONFIG.SYNC_FUNCTION_NAME +
        "\n" +
        "Interval: every " +
        CONFIG.TRIGGER_INTERVAL_MINUTES +
        " minute",
    );
    return;
  }

  ui.alert(
    "The auto-link trigger already exists.\n\n" +
      "Handler: " +
      CONFIG.SYNC_FUNCTION_NAME +
      "\n" +
      "Interval: every " +
      CONFIG.TRIGGER_INTERVAL_MINUTES +
      " minute",
  );
}

function viewSetupStatus() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var config = loadSheetConfig_(spreadsheet.getId(), sheet.getSheetId());
  var triggerStatus = getTriggerStatusText_();

  ui.alert(
    'Setup status for sheet "' +
      sheet.getName() +
      '"\n\n' +
      "Detected header config: " +
      (config.headerRow ? "Yes" : "No") +
      "\n" +
      "Header row saved: " +
      (config.headerRow || "Not set") +
      "\n" +
      "Checklist column: " +
      formatColumnStatus_(config.checklistColumnIndex) +
      "\n" +
      "Links column: " +
      formatColumnStatus_(config.linksColumnIndex) +
      "\n" +
      "Remarks column: " +
      formatColumnStatus_(config.remarksColumnIndex) +
      "\n" +
      "Number of copies column: " +
      formatColumnStatus_(config.numberOfCopiesColumnIndex) +
      "\n" +
      "Drive folder linked: " +
      (config.folderUrl ? "Yes" : "No") +
      "\n" +
      (config.folderUrl ? "Folder URL: " + config.folderUrl + "\n" : "") +
      "Every-minute trigger: " +
      triggerStatus,
  );
}

function syncConfiguredSheets() {
  var configs = loadAllSheetConfigs_();
  if (!configs.length) {
    Logger.log("No saved checklist configuration found.");
    return;
  }

  var spreadsheetsById = {};

  configs.forEach(function (config) {
    try {
      if (!config.spreadsheetId) {
        return;
      }

      if (!spreadsheetsById[config.spreadsheetId]) {
        spreadsheetsById[config.spreadsheetId] = SpreadsheetApp.openById(
          config.spreadsheetId,
        );
      }

      var spreadsheet = spreadsheetsById[config.spreadsheetId];
      var sheet = spreadsheet.getSheetByName(config.sheetName);
      if (!sheet || sheet.getSheetId() !== config.sheetId) {
        Logger.log(
          "Skipping config for sheet ID %s: sheet not found or renamed.",
          config.sheetId,
        );
        return;
      }

      syncSheetByConfig_(sheet, config);
    } catch (error) {
      Logger.log(
        "Failed syncing configured sheet %s (%s): %s",
        config.sheetName,
        config.sheetId,
        error && error.message ? error.message : error,
      );
    }
  });
}

function syncSheetByConfig_(sheet, config) {
  if (!config.folderId) {
    Logger.log(
      'Skipping sheet "%s": no linked Drive folder is configured.',
      sheet.getName(),
    );
    return;
  }

  if (
    !config.headerRow ||
    !config.checklistColumnIndex ||
    !config.linksColumnIndex ||
    !config.remarksColumnIndex
  ) {
    Logger.log(
      'Skipping sheet "%s": checklist columns are not configured.',
      sheet.getName(),
    );
    return;
  }

  var folder;
  try {
    folder = DriveApp.getFolderById(config.folderId);
  } catch (error) {
    Logger.log(
      'Skipping sheet "%s": unable to access folder %s. %s',
      sheet.getName(),
      config.folderId,
      error.message,
    );
    return;
  }

  var rows = getChecklistRowsByConfig_(sheet, config);
  if (!rows.length) {
    Logger.log('No checklist rows found for sheet "%s".', sheet.getName());
    return;
  }

  rows.forEach(function (row) {
    var file = findBestMatchingFile(folder, row.label);
    if (!file) {
      Logger.log(
        'No matching file found for row %s ("%s") in sheet "%s".',
        row.rowIndex,
        row.label,
        sheet.getName(),
      );
      return;
    }

    var remark = formatUploadedRemark_(
      file.getDateCreated(),
      sheet.getParent(),
    );
    updateRowWithFile_(
      sheet,
      row.rowIndex,
      config.linksColumnIndex,
      config.remarksColumnIndex,
      file.getUrl(),
      remark,
    );
  });
}

function detectChecklistLayoutsAcrossSheets_(spreadsheet) {
  var detected = [];
  var notDetected = [];

  spreadsheet.getSheets().forEach(function (sheet) {
    var detection = detectChecklistLayoutForSheet_(sheet);
    if (detection) {
      detected.push({
        sheet: sheet,
        detection: detection,
      });
      return;
    }

    notDetected.push(sheet.getName());
  });

  return {
    detected: detected,
    notDetected: notDetected,
  };
}

function detectChecklistLayoutForSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (!lastRow || !lastColumn) {
    return null;
  }

  var values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var bestDetection = null;

  for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var detection = detectChecklistLayoutFromRowValues_(
      values[rowIndex],
      rowIndex + 1,
    );
    if (!detection) {
      continue;
    }

    if (!bestDetection || detection.score > bestDetection.score) {
      bestDetection = detection;
      continue;
    }

    if (
      detection.score === bestDetection.score &&
      detection.headerRow < bestDetection.headerRow
    ) {
      bestDetection = detection;
    }
  }

  return bestDetection;
}

function detectChecklistLayoutFromRowValues_(rowValues, headerRow) {
  var matches = {
    checklist: null,
    links: null,
    remarks: null,
    copies: null,
  };

  for (var columnIndex = 0; columnIndex < rowValues.length; columnIndex++) {
    var meaning = matchHeaderMeaning_(rowValues[columnIndex]);
    if (!meaning) {
      continue;
    }

    if (!matches[meaning]) {
      matches[meaning] = columnIndex + 1;
    }
  }

  if (!matches.checklist || !matches.links || !matches.remarks) {
    return null;
  }

  return {
    headerRow: headerRow,
    checklistColumnIndex: matches.checklist,
    linksColumnIndex: matches.links,
    remarksColumnIndex: matches.remarks,
    numberOfCopiesColumnIndex: matches.copies,
    score: scoreChecklistHeaderRow_(matches),
  };
}

function scoreChecklistHeaderRow_(matches) {
  var columns = [matches.checklist, matches.links, matches.remarks];
  if (matches.copies) {
    columns.push(matches.copies);
  }

  var minColumn = Math.min.apply(null, columns);
  var maxColumn = Math.max.apply(null, columns);
  var span = maxColumn - minColumn;
  var score = 1000 - span;

  if (matches.copies) {
    score += 25;
  }

  return score;
}

function matchHeaderMeaning_(cellText) {
  var normalized = normalizeHeaderMeaningText_(cellText);
  if (!normalized || normalized.length > 60) {
    return "";
  }

  var words = normalized.split(" ");
  var hasWord = function (w) {
    return words.indexOf(w) !== -1;
  };

  if (hasWord("CHECKLIST") || (hasWord("CHECK") && hasWord("LIST"))) {
    return "checklist";
  }

  if (hasWord("LINK") || hasWord("LINKS")) {
    return "links";
  }

  if (hasWord("REMARK") || hasWord("REMARKS")) {
    return "remarks";
  }

  if (
    hasWord("COPIES") ||
    hasWord("COPY") ||
    (hasWord("NO") && hasWord("COPIES")) ||
    (hasWord("NUMBER") && hasWord("COPIES"))
  ) {
    return "copies";
  }

  return "";
}

function getChecklistRowsByConfig_(sheet, config) {
  var startRow = config.headerRow + 1;
  var lastRow = sheet.getLastRow();
  if (startRow > lastRow) {
    return [];
  }

  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastColumn);
  var displayValues = range.getDisplayValues();
  var mergedRanges = range.getMergedRanges();
  var mergedRowMap = buildMergedRowMap_(mergedRanges);
  var rows = [];

  for (var offset = 0; offset < displayValues.length; offset++) {
    var rowIndex = startRow + offset;
    var row = displayValues[offset];
    var label = String(row[config.checklistColumnIndex - 1] || "").trim();

    if (detectChecklistLayoutFromRowValues_(row, rowIndex)) {
      break;
    }

    if (!label) {
      if (rows.length) {
        break;
      }
      continue;
    }

    if (
      isIgnorableConfiguredChecklistRow_(
        row,
        label,
        mergedRowMap[rowIndex],
        config,
      )
    ) {
      continue;
    }

    rows.push({
      rowIndex: rowIndex,
      label: label,
    });
  }

  return rows;
}

function updateRowWithFile_(
  sheet,
  rowIndex,
  linksColumnIndex,
  remarksColumnIndex,
  fileUrl,
  remarkText,
) {
  var linkCell = sheet.getRange(rowIndex, linksColumnIndex);
  var remarksCell = sheet.getRange(rowIndex, remarksColumnIndex);
  var currentLink =
    extractUrlFromCell_(
      linkCell.getDisplayValue(),
      linkCell.getFormula(),
      linkCell.getRichTextValue(),
    ) || String(linkCell.getDisplayValue() || "").trim();
  var currentRemark = String(remarksCell.getDisplayValue() || "").trim();

  if (currentLink === fileUrl && currentRemark === remarkText) {
    Logger.log("Row %s already has the matching link and remark.", rowIndex);
    return;
  }

  linkCell.setValue(fileUrl);
  remarksCell.setValue(remarkText);
  Logger.log(
    'Updated row %s with link %s and remark "%s".',
    rowIndex,
    fileUrl,
    remarkText,
  );
}

function formatUploadedRemark_(dateValue, spreadsheet) {
  var timezone =
    spreadsheet.getSpreadsheetTimeZone() || CONFIG.TIMEZONE_FALLBACK;
  return (
    "Uploaded " +
    Utilities.formatDate(dateValue || new Date(), timezone, "yyyy-MM-dd HH:mm")
  );
}

function findBestMatchingFile(folder, checklistLabel) {
  var normalizedLabel = normalizeText(checklistLabel);
  if (
    !normalizedLabel ||
    normalizedLabel.length < CONFIG.MIN_CHECKLIST_TEXT_LENGTH
  ) {
    return null;
  }

  var files = folder.getFiles();
  var bestFile = null;
  var bestUpdatedTime = 0;

  while (files.hasNext()) {
    var file = files.next();
    var normalizedFileName = normalizeText(file.getName());

    if (
      !normalizedFileName ||
      normalizedFileName.indexOf(normalizedLabel) === -1
    ) {
      continue;
    }

    var updatedTime = file.getLastUpdated().getTime();
    if (!bestFile || updatedTime > bestUpdatedTime) {
      bestFile = file;
      bestUpdatedTime = updatedTime;
    }
  }

  return bestFile;
}

function extractFolderId(url) {
  if (!url) {
    return "";
  }

  var normalizedUrl = String(url).trim();
  var folderMatch = normalizedUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch) {
    return folderMatch[1];
  }

  var openIdMatch = normalizedUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (openIdMatch) {
    return openIdMatch[1];
  }

  return "";
}

function normalizeText(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[\u2018\u2019']/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeHeaderMeaningText_(text) {
  return String(text || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isIgnorableConfiguredChecklistRow_(
  row,
  label,
  hasMergedCells,
  config,
) {
  if (normalizeText(label).length < CONFIG.MIN_CHECKLIST_TEXT_LENGTH) {
    return true;
  }

  if (label === "-" || label === "--") {
    return true;
  }

  var hasOtherContent = row.some(function (cellValue, index) {
    var columnNumber = index + 1;
    if (columnNumber === config.checklistColumnIndex) {
      return false;
    }
    return String(cellValue || "").trim() !== "";
  });

  if (hasMergedCells && !hasOtherContent && label === label.toUpperCase()) {
    return true;
  }

  return false;
}

function extractUrlFromCell_(displayValue, formula, richTextValue) {
  var richTextUrl = getUrlFromRichText_(richTextValue);
  if (richTextUrl) {
    return richTextUrl;
  }

  if (formula) {
    var formulaUrl = getUrlFromFormula_(formula);
    if (formulaUrl) {
      return formulaUrl;
    }
  }

  var textUrlMatch = String(displayValue || "").match(/https?:\/\/[^\s)]+/i);
  return textUrlMatch ? textUrlMatch[0] : "";
}

function getUrlFromRichText_(richTextValue) {
  if (!richTextValue) {
    return "";
  }

  var directUrl = richTextValue.getLinkUrl();
  if (directUrl) {
    return directUrl;
  }

  var runs = richTextValue.getRuns();
  for (var i = 0; i < runs.length; i++) {
    var runUrl = runs[i].getLinkUrl();
    if (runUrl) {
      return runUrl;
    }
  }

  return "";
}

function getUrlFromFormula_(formula) {
  var hyperlinkMatch = String(formula).match(/=HYPERLINK\("([^"]+)"/i);
  return hyperlinkMatch ? hyperlinkMatch[1] : "";
}

function buildMergedRowMap_(mergedRanges) {
  var mergedRowMap = {};

  mergedRanges.forEach(function (mergedRange) {
    var startRow = mergedRange.getRow();
    var endRow = mergedRange.getLastRow();
    for (var row = startRow; row <= endRow; row++) {
      mergedRowMap[row] = true;
    }
  });

  return mergedRowMap;
}

function highlightDetectedColumns_(sheet, detection) {
  var headerRow = detection.headerRow;
  sheet
    .getRange(headerRow, detection.checklistColumnIndex)
    .setBackground(CONFIG.HIGHLIGHT_CHECKLIST);
  sheet
    .getRange(headerRow, detection.linksColumnIndex)
    .setBackground(CONFIG.HIGHLIGHT_LINKS);
  sheet
    .getRange(headerRow, detection.remarksColumnIndex)
    .setBackground(CONFIG.HIGHLIGHT_REMARKS);

  if (detection.numberOfCopiesColumnIndex) {
    sheet
      .getRange(headerRow, detection.numberOfCopiesColumnIndex)
      .setBackground(CONFIG.HIGHLIGHT_COPIES);
  }
}

function buildDetectionSummary_(result) {
  var lines = ["Checklist detection completed.", "", "Detected tabs:"];

  result.detected.forEach(function (item) {
    var detection = item.detection;
    lines.push(
      "- " +
        item.sheet.getName() +
        ": row " +
        detection.headerRow +
        ", Checklist " +
        formatColumnStatus_(detection.checklistColumnIndex) +
        ", Links " +
        formatColumnStatus_(detection.linksColumnIndex) +
        ", Remarks " +
        formatColumnStatus_(detection.remarksColumnIndex) +
        (detection.numberOfCopiesColumnIndex
          ? ", Copies " +
            formatColumnStatus_(detection.numberOfCopiesColumnIndex)
          : ""),
    );
  });

  lines.push("");
  lines.push("Highlight colors:");
  lines.push("Checklist " + CONFIG.HIGHLIGHT_CHECKLIST);
  lines.push("Links " + CONFIG.HIGHLIGHT_LINKS);
  lines.push("Remarks " + CONFIG.HIGHLIGHT_REMARKS);
  lines.push("Copies " + CONFIG.HIGHLIGHT_COPIES);

  if (result.notDetected.length) {
    lines.push("");
    lines.push("Not detected: " + result.notDetected.join(", "));
  }

  return lines.join("\n");
}

function getTriggerStatusText_() {
  var triggers = ScriptApp.getProjectTriggers();
  var hasTrigger = triggers.some(function (trigger) {
    return trigger.getHandlerFunction() === CONFIG.SYNC_FUNCTION_NAME;
  });

  return hasTrigger
    ? "Enabled (every " + CONFIG.TRIGGER_INTERVAL_MINUTES + " minute)"
    : "Not enabled";
}

function formatColumnStatus_(columnIndex) {
  if (!columnIndex) {
    return "Not set";
  }

  return columnNumberToLetter_(columnIndex) + " (" + columnIndex + ")";
}

function columnNumberToLetter_(columnNumber) {
  var letter = "";

  while (columnNumber > 0) {
    var modulo = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    columnNumber = Math.floor((columnNumber - modulo) / 26);
  }

  return letter;
}

function loadSheetConfig_(spreadsheetId, sheetId) {
  var propertyKey = getSheetConfigKey_(spreadsheetId, sheetId);
  var storedValue =
    PropertiesService.getDocumentProperties().getProperty(propertyKey);
  return storedValue ? JSON.parse(storedValue) : {};
}

function saveSheetConfig_(config) {
  var propertyKey = getSheetConfigKey_(config.spreadsheetId, config.sheetId);
  PropertiesService.getDocumentProperties().setProperty(
    propertyKey,
    JSON.stringify(config),
  );
}

function loadAllSheetConfigs_() {
  var prefix = CONFIG.PROPERTY_PREFIX;
  var properties = PropertiesService.getDocumentProperties().getProperties();
  var configs = [];

  Object.keys(properties).forEach(function (key) {
    if (key.indexOf(prefix) !== 0) {
      return;
    }

    try {
      configs.push(JSON.parse(properties[key]));
    } catch (error) {
      Logger.log('Skipping invalid saved config in property "%s".', key);
    }
  });

  return configs;
}

function getSheetConfigKey_(spreadsheetId, sheetId) {
  return CONFIG.PROPERTY_PREFIX + spreadsheetId + ":" + sheetId;
}
