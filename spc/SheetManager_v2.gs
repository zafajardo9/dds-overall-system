/**
 * Google Sheets Manager
 * Version: 3.0 (Production Ready) - Fixed CONFIG access for dynamic configuration
 */

const SheetManager_v2 = {
  /**
   * Get config dynamically
   */
  getConfig() {
    const scriptProps = PropertiesService.getScriptProperties();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    return {
      DRIVE_FOLDER_ID: scriptProps.getProperty("DRIVE_FOLDER_ID") || "",
      SHEET_ID: spreadsheet.getId(),
      SHEET_TAB_NAME:
        scriptProps.getProperty("SHEET_TAB_NAME") ||
        "Liquidation (unreplenished)",
      LOG_SHEET_NAME: "Processing Log",
      SUPPORTED_FORMATS: [
        "pdf",
        "jpg",
        "jpeg",
        "png",
        "doc",
        "docx",
        "xls",
        "xlsx",
      ],
      MAX_VOUCHERS_PER_BATCH: 30,
      MULTI_VOUCHER_NOTE: "Part of multi-voucher file",
      DATE_SUBFOLDER_FORMAT: "MM Month",
      MAX_FILES_PER_RUN: 10,
    };
  },
  createLiquidationSheet(spreadsheet = null) {
    if (!spreadsheet) {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }

    const CONFIG = this.getConfig();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_TAB_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.SHEET_TAB_NAME);
    }

    const headers = [
      "Encored Ref. No.", // A
      "Voucher Reference", // B
      "Client", // C
      "DATE", // D
      "DETAILS", // E
      "Location", // F
      "Amount", // G
      "Expense Classification", // H
      "STAFF", // I
      "Service", // J
      "Interco Classification", // K
      "DRAFT/BILLED", // L
      "SOA LINK", // M
      "Remarks", // N
      "Checked by Ms. Ela", // O
      "Liquidated?", // P
      "AP (Reconciliation)", // Q
      "Multi-Voucher File", // R
      "Original Filename", // S
      "Processed Date", // T
      "Processed By", // U
    ];

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.getRange(1, 1, 1, headers.length).setBackground("#f3f3f3");
      sheet.setFrozenRows(1);

      this.setupDropdownValidation(sheet);
    }

    return sheet;
  },
  addLiquidationRecord(
    validatedData,
    sourceFileName,
    fileUrl,
    folderUrl,
    isMultiVoucher,
    originalFilename,
    processedDate,
  ) {
    try {
      const CONFIG = this.getConfig(); // Get config dynamically

      console.log("SheetManager_v2.addLiquidationRecord v2.5 called");

      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = this.createLiquidationSheet(spreadsheet);

      const nextRow = sheet.getLastRow() + 1;
      const processedBy = this.getProcessedByUser();

      const multiVoucherIndicator = isMultiVoucher ? "YES" : "NO";
      const originalFileForDisplay = originalFilename || sourceFileName || "";

      console.log("Writing to row:", nextRow);
      console.log("Data:", {
        voucher: validatedData.voucherNo,
        company: validatedData.company,
        location: validatedData.mainLocation,
        expenseClass: validatedData.expenseClassification,
        service: validatedData.service,
        originalFile: originalFileForDisplay,
      });

      const rowData = [
        "", // A: Encored Ref. No.
        validatedData.voucherNo || "", // B: Voucher Reference
        validatedData.company || "", // C: Client
        validatedData.errandDate || "", // D: DATE
        validatedData.details || "", // E: DETAILS
        validatedData.mainLocation || "", // F: Location
        parseFloat(validatedData.total) || 0, // G: Amount
        validatedData.expenseClassification || "", // H: Expense Classification
        validatedData.errandBy || "", // I: STAFF
        validatedData.service || "", // J: Service
        "", // K: Interco Classification
        "", // L: DRAFT/BILLED
        "", // M: SOA LINK
        "", // N: Remarks
        "", // O: Checked by Ms. Ela
        "", // P: Liquidated?
        "", // Q: AP (Reconciliation)
        multiVoucherIndicator, // R: Multi-Voucher File
        originalFileForDisplay, // S: Original Filename
        processedDate, // T: Processed Date
        processedBy, // U: Processed By
      ];

      console.log("Row data array length:", rowData.length);

      sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
      console.log("✓ Data written to sheet");

      // Create Voucher Reference hyperlink (column B)
      if (fileUrl && validatedData.voucherNo) {
        const voucherText = isMultiVoucher
          ? `${validatedData.voucherNo} (${CONFIG.MULTI_VOUCHER_NOTE})`
          : validatedData.voucherNo;
        this.createHyperlink(sheet, nextRow, 2, voucherText, fileUrl, false);
        console.log("✓ Created voucher hyperlink");
      }

      // Create View Folder hyperlink (column M)
      if (folderUrl) {
        this.createHyperlink(
          sheet,
          nextRow,
          13,
          "View Folder",
          folderUrl,
          true,
        );
        console.log("✓ Created folder hyperlink");
      }

      this.formatNewRow(sheet, nextRow, isMultiVoucher);
      console.log("✓ Row formatted");

      this.applyRowValidation(sheet, nextRow);
      console.log("✓ Validations applied");

      console.log(`✓ COMPLETE: Record added at row ${nextRow}`);

      return nextRow;
    } catch (error) {
      console.error("❌ SheetManager_v2 ERROR:", error);
      throw new Error(`Failed to add record: ${error.message}`);
    }
  },
  createHyperlink(
    sheet,
    rowNumber,
    columnNumber,
    displayText,
    url,
    isFolderLink = false,
  ) {
    try {
      const cell = sheet.getRange(rowNumber, columnNumber);

      const escapedText = displayText.replace(/"/g, '""');
      const escapedUrl = url.replace(/"/g, "%22");

      const hyperlinkFormula = `=HYPERLINK("${escapedUrl}","${escapedText}")`;
      cell.setFormula(hyperlinkFormula);
      cell.setFontColor("#1155cc");
      cell.setFontLine("underline");

      if (isFolderLink) {
        cell.setFontColor("#0b5394");
        cell.setFontWeight("bold");
        cell.setHorizontalAlignment("center");
      }

      console.log(
        `✓ Created hyperlink at row ${rowNumber}, col ${columnNumber}`,
      );
    } catch (error) {
      console.error("Hyperlink creation failed:", error);
      const cell = sheet.getRange(rowNumber, columnNumber);
      cell.setValue(displayText);
    }
  },
  setupDropdownValidation(sheet) {
    const elaRange = sheet.getRange("O2:O");
    const elaRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Yes", "No"], true)
      .setAllowInvalid(false)
      .build();
    elaRange.setDataValidation(elaRule);

    const draftRange = sheet.getRange("L2:L");
    const draftRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["DRAFT", "BILLED"], true)
      .setAllowInvalid(false)
      .build();
    draftRange.setDataValidation(draftRule);

    const multiRange = sheet.getRange("R2:R");
    const multiRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["YES", "NO"], true)
      .setAllowInvalid(false)
      .build();
    multiRange.setDataValidation(multiRule);
  },
  applyRowValidation(sheet, rowNumber) {
    const elaCell = sheet.getRange(rowNumber, 15);
    const elaRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Yes", "No"], true)
      .setAllowInvalid(false)
      .build();
    elaCell.setDataValidation(elaRule);

    const draftCell = sheet.getRange(rowNumber, 12);
    const draftRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["DRAFT", "BILLED"], true)
      .setAllowInvalid(false)
      .build();
    draftCell.setDataValidation(draftRule);

    const multiCell = sheet.getRange(rowNumber, 18);
    const multiRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["YES", "NO"], true)
      .setAllowInvalid(false)
      .build();
    multiCell.setDataValidation(multiRule);
  },
  formatNewRow(sheet, rowNumber, isMultiVoucher = false) {
    const range = sheet.getRange(rowNumber, 1, 1, 21);

    if (isMultiVoucher) {
      range.setBackground("#e8f4fd");
      range.setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#4285f4",
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
      );
    } else {
      if (rowNumber % 2 === 0) {
        range.setBackground("#f8f9fa");
      }
      range.setBorder(true, true, true, true, true, true);
    }

    sheet.getRange(rowNumber, 4).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(rowNumber, 20).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet.getRange(rowNumber, 7).setNumberFormat("#,##0.00");

    sheet.getRange(rowNumber, 1).setBackground("#fffbf0");
    sheet.getRange(rowNumber, 11).setBackground("#fffbf0");
    sheet.getRange(rowNumber, 14).setBackground("#fffbf0");
    sheet.getRange(rowNumber, 13).setBackground("#e6f4ea");

    if (isMultiVoucher) {
      sheet.getRange(rowNumber, 18).setBackground("#fff3cd");
      sheet.getRange(rowNumber, 19).setBackground("#fff3cd");
    }
  },
  getProcessedByUser() {
    try {
      const email = Session.getActiveUser().getEmail();
      return email || "System User";
    } catch (error) {
      return "Automated System";
    }
  },
};
