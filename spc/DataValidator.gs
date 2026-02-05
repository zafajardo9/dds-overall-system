/**
 * Version: 3.0 (Production Ready)
 * Data Validator
 * Validates and cleans extracted data
 */

const DataValidator = {
  /**
   * Validate and clean extracted data
   */
  validateAndClean(rawData) {
    const cleanData = {
      name: this.validateName(rawData.name),
      errandDate: this.validateDate(rawData.errandDate),
      voucherNo: this.validateVoucherNo(rawData.voucherNo),
      company: this.validateText(rawData.company, "Company"),
      errandBy: this.validateText(rawData.errandBy, "Staff"),
      service: this.validateText(rawData.service, "Service"),
      details: this.validateText(rawData.details, "Details"),
      expenseClassification: this.validateText(
        rawData.expenseClassification,
        "Expense Classification",
      ),
      total: this.validateAmount(rawData.total),
      mainLocation: this.validateText(rawData.mainLocation, "Location"),
    };

    return cleanData;
  },
  /**
   * Validate name field
   */
  validateName(name) {
    if (!name || typeof name !== "string") {
      return "";
    }
    return name.trim().substring(0, 100);
  },
  /**
   * Validate date field
   */
  validateDate(dateStr) {
    if (!dateStr) return "";

    try {
      // Try to parse various date formats
      const date = new Date(dateStr);
      if (isNaN(date.getTime())) {
        return "";
      }

      // Return in YYYY-MM-DD format
      return (
        date.getFullYear() +
        "-" +
        String(date.getMonth() + 1).padStart(2, "0") +
        "-" +
        String(date.getDate()).padStart(2, "0")
      );
    } catch (error) {
      return "";
    }
  },
  /**
   * Validate voucher number format
   */
  validateVoucherNo(voucherNo) {
    if (!voucherNo || typeof voucherNo !== "string") {
      return "";
    }

    const cleaned = voucherNo.trim();

    // Check if it matches expected format: YYYY-number-letter
    const voucherPattern = /^\d{4}-\d+-[A-Za-z]$/;
    if (voucherPattern.test(cleaned)) {
      return cleaned;
    }

    // Return as-is if doesn't match expected format
    return cleaned.substring(0, 50);
  },
  /**
   * Validate text fields
   */
  validateText(text, fieldName) {
    if (!text || typeof text !== "string") {
      return "";
    }

    const cleaned = text.trim();

    // Limit length based on field
    const maxLength = this.getMaxLength(fieldName);
    return cleaned.substring(0, maxLength);
  },
  /**
   * Get maximum length for different fields
   */
  getMaxLength(fieldName) {
    const lengths = {
      Company: 200,
      Staff: 100,
      Service: 150,
      Details: 500,
      "Expense Classification": 100,
      Location: 200,
    };

    return lengths[fieldName] || 100;
  },
  /**
   * Validate amount field
   */
  validateAmount(amount) {
    if (!amount) return 0;

    // If string, try to parse
    if (typeof amount === "string") {
      // Remove currency symbols and commas
      const cleaned = amount.replace(/[₱$,\s]/g, "");
      const parsed = parseFloat(cleaned);
      return isNaN(parsed) ? 0 : Math.abs(parsed);
    }

    // If number, validate
    if (typeof amount === "number") {
      return isNaN(amount) ? 0 : Math.abs(amount);
    }

    return 0;
  },
};
