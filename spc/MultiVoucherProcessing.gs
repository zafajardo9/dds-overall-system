/**
 * Version: 3.0 (Production Ready) -
 * Multi-Voucher Processor
 * Handles processing multiple vouchers from single files
 */

const MultiVoucherProcessor = {
  /**
   * Create batches of vouchers for processing
   */
  createBatches(vouchersData, batchSize) {
    const batches = [];
    for (let i = 0; i < vouchersData.length; i += batchSize) {
      batches.push(vouchersData.slice(i, i + batchSize));
    }
    return batches;
  },
  /**
   * Handle duplicate voucher numbers within the same file
   */
  handleDuplicateVouchers(validatedData, allVouchersInFile, currentIndex) {
    const currentVoucherNo = validatedData.voucherNo;

    if (!currentVoucherNo) {
      return validatedData;
    }

    // Count occurrences of this voucher number before current index
    let duplicateCount = 0;
    for (let i = 0; i < currentIndex; i++) {
      if (
        allVouchersInFile[i] &&
        allVouchersInFile[i].voucherNo === currentVoucherNo
      ) {
        duplicateCount++;
      }
    }

    // If duplicates found, append version suffix
    if (duplicateCount > 0) {
      const modifiedData = { ...validatedData };
      modifiedData.voucherNo = `${currentVoucherNo}_v${duplicateCount + 1}`;
      return modifiedData;
    }

    return validatedData;
  },
  /**
   * Create partial voucher data for failed extractions
   */
  createPartialVoucherData(rawVoucherData, errorMessage) {
    return {
      name: rawVoucherData?.name || "",
      errandDate: rawVoucherData?.errandDate || "",
      voucherNo: rawVoucherData?.voucherNo || "EXTRACTION_FAILED",
      company: rawVoucherData?.company || "",
      errandBy: rawVoucherData?.errandBy || "",
      service: rawVoucherData?.service || "",
      details: `EXTRACTION ERROR: ${errorMessage}`,
      expenseClassification: rawVoucherData?.expenseClassification || "",
      total: 0,
      mainLocation: rawVoucherData?.mainLocation || "",
    };
  },
};

/**
 * DEBUG FUNCTION - Test multi-voucher parsing with sample data
 * Run this function to test if the parsing works correctly
 */
function debugMultiVoucherParsing() {
  // Simulate Gemini API response with multiple vouchers
  const sampleResponse = `[
   {
     "name": "John Doe",
     "errandDate": "2024-01-15",
     "voucherNo": "2024-001-A",
     "company": "ABC Corp",
     "errandBy": "Staff A",
     "service": "Delivery",
     "details": "Package delivery to client",
     "expenseClassification": "Transportation",
     "total": 150,
     "mainLocation": "Downtown Office"
   },
   {
     "name": "Jane Smith",
     "errandDate": "2024-01-16",
     "voucherNo": "2024-002-B",
     "company": "ABC Corp",
     "errandBy": "Staff B",
     "service": "Document Pickup",
     "details": "Legal document retrieval",
     "expenseClassification": "Transportation",
     "total": 200,
     "mainLocation": "Law Office"
   },
   {
     "name": "Bob Johnson",
     "errandDate": "2024-01-17",
     "voucherNo": "2024-003-C",
     "company": "XYZ Inc",
     "errandBy": "Staff C",
     "service": "Meeting Attendance",
     "details": "Client meeting representation",
     "expenseClassification": "Professional Services",
     "total": 300,
     "mainLocation": "Client Office"
   }
 ]`;
  try {
    console.log("Testing multi-voucher JSON parsing...");

    const parsedVouchers =
      GeminiParser.parseMultiVoucherJsonResponse(sampleResponse);

    console.log(`DEBUG: Successfully parsed ${parsedVouchers.length} vouchers`);
    Logger.log(
      `DEBUG: Multi-voucher test - parsed ${parsedVouchers.length} vouchers`,
    );

    parsedVouchers.forEach((voucher, index) => {
      console.log(
        `DEBUG Voucher ${index + 1}: ${voucher.voucherNo} - ${voucher.name} - $${voucher.total}`,
      );
      Logger.log(
        `DEBUG Voucher ${index + 1}: VoucherNo=${voucher.voucherNo}, Name=${voucher.name}, Company=${voucher.company}, Total=${voucher.total}`,
      );
    });

    SpreadsheetApp.getUi().alert(
      "Debug Results",
      `Multi-voucher parsing test completed!\n\nParsed ${parsedVouchers.length} vouchers successfully.\n\nCheck the execution log (View > Execution Log) for detailed results.`,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );

    return parsedVouchers;
  } catch (error) {
    console.error("DEBUG: Multi-voucher parsing test failed:", error.message);
    Logger.log(
      `DEBUG: Multi-voucher parsing test failed: ${error.message}`,
      "ERROR",
    );
    SpreadsheetApp.getUi().alert(
      "Debug Error",
      `Test failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
    return [];
  }
}
