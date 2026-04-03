/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * DateUtils.gs - Date & Deadline Calculation Functions
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * PURPOSE:
 * Reusable functions for date manipulation, deadline calculations, working days,
 * and timezone handling across DDS Apps Script projects.
 * 
 * PREREQUISITES:
 * - No special requirements
 * 
 * USAGE:
 * Copy this file into your Apps Script project alongside other common-utils.
 * 
 * VERSION: 1.0.0
 * LAST UPDATED: April 2026
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * ==============================================================================
 * FUNCTION: addWorkingDays
 * ==============================================================================
 * 
 * PURPOSE:
 * Adds a specified number of working days (business days) to a date.
 * Excludes weekends (Saturday/Sunday) and optionally holidays.
 * 
 * PARAMETERS:
 * @param {Date} startDate - Starting date
 * @param {number} days - Number of working days to add (can be negative)
 * @param {Array} holidays - (Optional) Array of holiday dates to exclude
 * @param {boolean} includeStart - (Optional) Count start date as day 1 (default: false)
 * 
 * @returns {Date} Resulting date after adding working days
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * // Due in 5 working days
 * var dueDate = addWorkingDays(new Date(), 5);
 * 
 * // With holidays
 * var holidays = [new Date("2026-01-01"), new Date("2026-04-09")];
 * var dueDate = addWorkingDays(new Date(), 10, holidays);
 * 
 * // 30 days ago (working days)
 * var pastDate = addWorkingDays(new Date(), -30);
 * ```
 * 
 * NOTES:
 * - Working days = Monday-Friday only
 * - Holidays array should contain Date objects
 * - Handles negative days (goes backwards)
 * - Commonly used in: CAR, automated-expiry-notification for deadline calc
 * ==============================================================================
 */
function addWorkingDays(startDate, days, holidays, includeStart) {
  if (!startDate || isNaN(startDate.getTime())) return null;
  
  holidays = holidays || [];
  includeStart = includeStart || false;
  
  var result = new Date(startDate);
  var direction = days >= 0 ? 1 : -1;
  var remaining = Math.abs(days);
  
  // Normalize holidays to string format for comparison
  var holidayStrings = holidays.map(function(h) {
    return Utilities.formatDate(h, Session.getScriptTimeZone(), "yyyy-MM-dd");
  });
  
  if (!includeStart) {
    result.setDate(result.getDate() + direction);
  }
  
  while (remaining > 0) {
    var dayOfWeek = result.getDay();
    var isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
    var dateStr = Utilities.formatDate(result, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var isHoliday = holidayStrings.indexOf(dateStr) !== -1;
    
    if (!isWeekend && !isHoliday) {
      remaining--;
    }
    
    if (remaining > 0) {
      result.setDate(result.getDate() + direction);
    }
  }
  
  return result;
}

/**
 * ==============================================================================
 * FUNCTION: getWorkingDaysBetween
 * ==============================================================================
 * 
 * PURPOSE:
 * Calculates the number of working days between two dates.
 * 
 * PARAMETERS:
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @param {Array} holidays - (Optional) Array of holiday dates
 * @param {boolean} inclusive - (Optional) Include both start and end dates (default: false)
 * 
 * @returns {number} Count of working days between dates
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var daysRemaining = getWorkingDaysBetween(new Date(), dueDate);
 * if (daysRemaining <= 3) {
 *   // Send urgent reminder
 * }
 * ```
 * ==============================================================================
 */
function getWorkingDaysBetween(startDate, endDate, holidays, inclusive) {
  if (!startDate || !endDate) return 0;
  
  holidays = holidays || [];
  inclusive = inclusive || false;
  
  var start = new Date(startDate);
  var end = new Date(endDate);
  
  if (start > end) {
    var temp = start;
    start = end;
    end = temp;
  }
  
  var holidayStrings = holidays.map(function(h) {
    return Utilities.formatDate(h, Session.getScriptTimeZone(), "yyyy-MM-dd");
  });
  
  var workingDays = 0;
  var current = new Date(start);
  
  while (current <= end) {
    var dayOfWeek = current.getDay();
    var dateStr = Utilities.formatDate(current, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    if (dayOfWeek !== 0 && dayOfWeek !== 6 && holidayStrings.indexOf(dateStr) === -1) {
      if (inclusive || (current > start && current < end)) {
        workingDays++;
      }
    }
    
    current.setDate(current.getDate() + 1);
  }
  
  return workingDays;
}

/**
 * ==============================================================================
 * FUNCTION: getDaysRemaining
 * ==============================================================================
 * 
 * PURPOSE:
 * Calculates days remaining until a deadline. Returns negative if overdue.
 * 
 * PARAMETERS:
 * @param {Date} deadline - Target deadline date
 * @param {Date} fromDate - (Optional) Calculate from this date (default: today)
 * @param {boolean} workingDays - (Optional) Count only working days (default: false)
 * @param {Array} holidays - (Optional) Holidays for working day calc
 * 
 * @returns {number} Days remaining (negative if overdue, 0 if today)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var daysLeft = getDaysRemaining(dueDate);
 * if (daysLeft < 0) {
 *   Logger.log("Overdue by " + Math.abs(daysLeft) + " days");
 * } else if (daysLeft <= 7) {
 *   Logger.log("Due soon: " + daysLeft + " days left");
 * }
 * ```
 * 
 * NOTES:
 * - Returns 0 if deadline is today
 * - Returns negative number if deadline passed
 * - Commonly used in: CAR, automated-expiry-notification for countdowns
 * ==============================================================================
 */
function getDaysRemaining(deadline, fromDate, workingDays, holidays) {
  if (!deadline || isNaN(deadline.getTime())) return null;
  
  var today = fromDate || new Date();
  
  // Strip time components for date-only comparison
  var deadlineDate = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  var todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  
  if (workingDays) {
    return getWorkingDaysBetween(todayDate, deadlineDate, holidays, false) * 
           (deadlineDate >= todayDate ? 1 : -1);
  }
  
  var diffMs = deadlineDate - todayDate;
  var diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));
  
  return diffDays;
}

/**
 * ==============================================================================
 * FUNCTION: formatDateForDisplay
 * ==============================================================================
 * 
 * PURPOSE:
 * Formats a date for user-friendly display with various format options.
 * 
 * PARAMETERS:
 * @param {Date} date - Date to format
 * @param {string} format - (Optional) Format type:
 *   - "short": "Jan 5, 2026" (default)
 *   - "long": "January 5, 2026"
 *   - "iso": "2026-01-05"
 *   - "friendly": "Today", "Tomorrow", "Yesterday", or "Jan 5"
 *   - "datetime": "Jan 5, 2026 2:30 PM"
 * @param {string} timezone - (Optional) Timezone (default: script timezone)
 * 
 * @returns {string} Formatted date string
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var display = formatDateForDisplay(dueDate, "friendly");
 * // Returns: "Today", "Tomorrow", "Yesterday", or "Jan 5"
 * 
 * var iso = formatDateForDisplay(dueDate, "iso");
 * // Returns: "2026-01-05"
 * ```
 * ==============================================================================
 */
function formatDateForDisplay(date, format, timezone) {
  if (!date || isNaN(date.getTime())) return "Invalid date";
  
  format = format || "short";
  timezone = timezone || Session.getScriptTimeZone();
  
  if (format === "friendly") {
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var target = new Date(date);
    target.setHours(0, 0, 0, 0);
    
    var diffDays = Math.round((target - today) / (1000 * 60 * 60 * 24));
    
    if (diffDays === 0) return "Today";
    if (diffDays === 1) return "Tomorrow";
    if (diffDays === -1) return "Yesterday";
    
    // Fall through to short format for other dates
    format = "short";
  }
  
  switch (format) {
    case "iso":
      return Utilities.formatDate(date, timezone, "yyyy-MM-dd");
    case "long":
      return Utilities.formatDate(date, timezone, "MMMM d, yyyy");
    case "datetime":
      return Utilities.formatDate(date, timezone, "MMM d, yyyy h:mm a");
    case "short":
    default:
      return Utilities.formatDate(date, timezone, "MMM d, yyyy");
  }
}

/**
 * ==============================================================================
 * FUNCTION: parseDateString
 * ==============================================================================
 * 
 * PURPOSE:
 * Parses various date string formats into a Date object.
 * Handles common formats found in documents and user input.
 * 
 * PARAMETERS:
 * @param {string} dateString - String to parse
 * @param {string} format - (Optional) Expected format hint:
 *   - "auto": Try common formats (default)
 *   - "us": MM/DD/YYYY
 *   - "intl": DD/MM/YYYY
 *   - "iso": YYYY-MM-DD
 * 
 * @returns {Date|null} Parsed Date object or null if invalid
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var date1 = parseDateString("2026-01-15"); // ISO
 * var date2 = parseDateString("01/15/2026", "us"); // US format
 * var date3 = parseDateString("15 Jan 2026", "auto"); // Auto-detect
 * ```
 * 
 * NOTES:
 * - Returns null for unparseable strings
 * - Be careful with ambiguous dates like "01/02/2026" - use format hint
 * ==============================================================================
 */
function parseDateString(dateString, format) {
  if (!dateString) return null;
  
  format = format || "auto";
  var str = dateString.toString().trim();
  
  // Try ISO format first (YYYY-MM-DD)
  var isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    var date = new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]));
    if (!isNaN(date.getTime())) return date;
  }
  
  // Try US format (MM/DD/YYYY)
  if (format === "us" || format === "auto") {
    var usMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (usMatch) {
      var date = new Date(parseInt(usMatch[3]), parseInt(usMatch[1]) - 1, parseInt(usMatch[2]));
      if (!isNaN(date.getTime())) return date;
    }
  }
  
  // Try international format (DD/MM/YYYY)
  if (format === "intl" || format === "auto") {
    var intlMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (intlMatch) {
      var date = new Date(parseInt(intlMatch[3]), parseInt(intlMatch[2]) - 1, parseInt(intlMatch[1]));
      if (!isNaN(date.getTime())) return date;
    }
  }
  
  // Try natural language parsing
  var natural = new Date(str);
  if (!isNaN(natural.getTime())) {
    return natural;
  }
  
  return null;
}

/**
 * ==============================================================================
 * FUNCTION: getFirstDayOfMonth
 * ==============================================================================
 * 
 * PURPOSE:
 * Returns the first day of the month for a given date.
 * 
 * PARAMETERS:
 * @param {Date} date - Reference date (default: today)
 * @param {number} offset - (Optional) Month offset: 0=current, 1=next, -1=previous
 * 
 * @returns {Date} First day of target month
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var first = getFirstDayOfMonth(new Date()); // This month
 * var nextFirst = getFirstDayOfMonth(new Date(), 1); // Next month
 * ```
 * 
 * NOTES:
 * - Time is set to 00:00:00
 * - Used in CAR for DST due date calculation (5th of next month)
 * ==============================================================================
 */
function getFirstDayOfMonth(date, offset) {
  date = date || new Date();
  offset = offset || 0;
  
  var year = date.getFullYear();
  var month = date.getMonth() + offset;
  
  // Handle year rollover
  while (month < 0) {
    month += 12;
    year--;
  }
  while (month > 11) {
    month -= 12;
    year++;
  }
  
  return new Date(year, month, 1, 0, 0, 0, 0);
}

/**
 * ==============================================================================
 * FUNCTION: getLastDayOfMonth
 * ==============================================================================
 * 
 * PURPOSE:
 * Returns the last day of the month for a given date.
 * 
 * PARAMETERS:
 * @param {Date} date - Reference date (default: today)
 * @param {number} offset - (Optional) Month offset: 0=current, 1=next, -1=previous
 * 
 * @returns {Date} Last day of target month
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var last = getLastDayOfMonth(new Date()); // End of this month
 * var nextLast = getLastDayOfMonth(new Date(), 1); // End of next month
 * ```
 * ==============================================================================
 */
function getLastDayOfMonth(date, offset) {
  date = date || new Date();
  offset = (offset || 0) + 1; // Add 1 to get first day of next month, then subtract 1
  
  var firstOfNext = getFirstDayOfMonth(date, offset);
  firstOfNext.setDate(0); // Go back to last day of previous month
  
  return firstOfNext;
}

/**
 * ==============================================================================
 * FUNCTION: isOverdue
 * ==============================================================================
 * 
 * PURPOSE:
 * Checks if a deadline has passed (is overdue).
 * 
 * PARAMETERS:
 * @param {Date} deadline - Deadline date to check
 * @param {Date} fromDate - (Optional) Compare against this date (default: today)
 * @param {boolean} endOfDay - (Optional) Consider EOD as cutoff (default: true)
 * 
 * @returns {boolean} True if deadline is in the past
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * if (isOverdue(dueDate)) {
 *   sendUrgentReminder();
 * }
 * ```
 * ==============================================================================
 */
function isOverdue(deadline, fromDate, endOfDay) {
  if (!deadline) return false;
  
  endOfDay = endOfDay !== false; // Default true
  
  var compareDate = fromDate || new Date();
  
  if (endOfDay) {
    // Compare dates only (ignore time)
    deadline = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
    compareDate = new Date(compareDate.getFullYear(), compareDate.getMonth(), compareDate.getDate());
  }
  
  return deadline < compareDate;
}

/**
 * ==============================================================================
 * FUNCTION: getDeadlineColor
 * ==============================================================================
 * 
 * PURPOSE:
 * Returns a color code based on days remaining until deadline.
 * Useful for conditional formatting.
 * 
 * PARAMETERS:
 * @param {number} daysRemaining - Days until deadline (negative = overdue)
 * @param {Object} thresholds - (Optional) Custom thresholds:
 *   - danger: {number} - Red if days <= this (default: 0)
 *   - warning: {number} - Yellow if days <= this (default: 7)
 *   - success: {number} - Green if days >= this (default: 14)
 * 
 * @returns {string} Hex color code for status
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var daysLeft = getDaysRemaining(dueDate);
 * var color = getDeadlineColor(daysLeft);
 * setRowBackgroundColor(sheet, row, color);
 * ```
 * 
 * NOTES:
 * - Default colors: #f4cccc (red/danger), #fff2cc (yellow/warning), #d9ead3 (green/safe)
 * - Override thresholds for custom urgency levels
 * ==============================================================================
 */
function getDeadlineColor(daysRemaining, thresholds) {
  thresholds = thresholds || {};
  
  var danger = thresholds.danger !== undefined ? thresholds.danger : 0;
  var warning = thresholds.warning !== undefined ? thresholds.warning : 7;
  var success = thresholds.success !== undefined ? thresholds.success : 14;
  
  if (daysRemaining <= danger) {
    return "#f4cccc"; // Light red - overdue/immediate
  } else if (daysRemaining <= warning) {
    return "#fff2cc"; // Light yellow - approaching
  } else {
    return "#d9ead3"; // Light green - safe
  }
}

/**
 * ==============================================================================
 * FUNCTION: generatePhilippinesHolidays
 * ==============================================================================
 * 
 * PURPOSE:
 * Returns an array of Philippine national holidays for a given year.
 * Includes regular and special non-working holidays.
 * 
 * PARAMETERS:
 * @param {number} year - Year to generate holidays for (default: current year)
 * @param {boolean} includeSpecial - (Optional) Include special non-working holidays
 *                                   (default: true)
 * 
 * @returns {Array} Array of Date objects for holidays
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var holidays = generatePhilippinesHolidays(2026);
 * var dueDate = addWorkingDays(new Date(), 30, holidays);
 * ```
 * 
 * NOTES:
 * - Includes fixed holidays (New Year, Christmas, etc.)
 * - Includes Holy Week (Maundy Thursday, Good Friday)
 * - Includes Ninoy Aquino Day, All Saints Day, Bonifacio Day, etc.
 * - Holy Week dates are approximate (varies by year)
 * ==============================================================================
 */
function generatePhilippinesHolidays(year, includeSpecial) {
  year = year || new Date().getFullYear();
  includeSpecial = includeSpecial !== false;
  
  var holidays = [
    // Fixed regular holidays
    new Date(year, 0, 1),   // New Year's Day
    new Date(year, 3, 9),   // Araw ng Kagitingan (Day of Valor)
    new Date(year, 4, 1),   // Labor Day
    new Date(year, 5, 12),  // Independence Day
    new Date(year, 7, 26),  // National Heroes Day (last Monday of Aug - approximate)
    new Date(year, 10, 30), // Bonifacio Day
    new Date(year, 11, 25), // Christmas Day
    new Date(year, 11, 30), // Rizal Day
  ];
  
  if (includeSpecial) {
    // Additional special non-working holidays
    var special = [
      new Date(year, 0, 2),   // New Year additional
      new Date(year, 2, 28),  // Maundy Thursday (approximate - varies yearly)
      new Date(year, 2, 29),  // Good Friday (approximate - varies yearly)
      new Date(year, 7, 21),  // Ninoy Aquino Day
      new Date(year, 10, 1),  // All Saints Day
      new Date(year, 11, 24), // Christmas Eve
      new Date(year, 11, 31), // New Year's Eve
    ];
    holidays = holidays.concat(special);
  }
  
  return holidays;
}
