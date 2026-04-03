/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * ConfigUtils.gs - Property & Configuration Management
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * PURPOSE:
 * Reusable functions for managing script properties, user settings, and
 * configuration storage across DDS Apps Script projects.
 * 
 * PREREQUISITES:
 * - No special OAuth required for PropertiesService
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
 * FUNCTION: getConfigValue
 * ==============================================================================
 * 
 * PURPOSE:
 * Retrieves a configuration value from ScriptProperties with optional default fallback.
 * Type-safe wrapper around PropertiesService.
 * 
 * PARAMETERS:
 * @param {string} key - Property key name
 * @param {*} defaultValue - (Optional) Default value if property not found
 * @param {string} type - (Optional) Expected type: "string", "number", "boolean", "json", "date"
 *                      (default: "string")
 * 
 * @returns {*} The stored value converted to requested type, or defaultValue if not found
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var driveId = getConfigValue("DRIVE_FOLDER_ID", null, "string");
 * var batchSize = getConfigValue("BATCH_SIZE", 25, "number");
 * var enabled = getConfigValue("FEATURE_ENABLED", false, "boolean");
 * var settings = getConfigValue("USER_SETTINGS", {}, "json");
 * ```
 * 
 * NOTES:
 * - Stores values as strings internally, converts on retrieval
 * - "json" type parses stored JSON string to object
 * - "date" type parses ISO date strings
 * ==============================================================================
 */
function getConfigValue(key, defaultValue, type) {
  type = type || "string";
  
  try {
    var value = PropertiesService.getScriptProperties().getProperty(key);
    
    if (value === null) {
      return defaultValue;
    }
    
    switch (type) {
      case "number":
        var num = parseFloat(value);
        return isNaN(num) ? defaultValue : num;
      
      case "boolean":
        return value === "true" || value === "1";
      
      case "json":
        try {
          return JSON.parse(value);
        } catch (e) {
          return defaultValue;
        }
      
      case "date":
        var date = new Date(value);
        return isNaN(date.getTime()) ? defaultValue : date;
      
      case "string":
      default:
        return value;
    }
  } catch (e) {
    return defaultValue;
  }
}

/**
 * ==============================================================================
 * FUNCTION: setConfigValue
 * ==============================================================================
 * 
 * PURPOSE:
 * Stores a configuration value in ScriptProperties with automatic type conversion.
 * 
 * PARAMETERS:
 * @param {string} key - Property key name
 * @param {*} value - Value to store (will be converted to string)
 * @param {string} type - (Optional) Type hint: "string", "number", "boolean", "json", "date"
 * 
 * @returns {boolean} True if stored successfully, false otherwise
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * setConfigValue("DRIVE_FOLDER_ID", "1Bx123...", "string");
 * setConfigValue("BATCH_SIZE", 50, "number");
 * setConfigValue("USER_SETTINGS", {theme: "dark"}, "json");
 * setConfigValue("LAST_RUN", new Date(), "date");
 * ```
 * 
 * NOTES:
 * - Objects are serialized to JSON
 * - Dates are stored as ISO strings
 * - Maximum property storage: 500KB per script
 * ==============================================================================
 */
function setConfigValue(key, value, type) {
  type = type || "string";
  
  try {
    var storeValue;
    
    switch (type) {
      case "json":
        storeValue = JSON.stringify(value);
        break;
      case "date":
        storeValue = value instanceof Date ? value.toISOString() : value.toString();
        break;
      case "boolean":
        storeValue = value ? "true" : "false";
        break;
      default:
        storeValue = value.toString();
    }
    
    PropertiesService.getScriptProperties().setProperty(key, storeValue);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * ==============================================================================
 * FUNCTION: deleteConfigValue
 * ==============================================================================
 * 
 * PURPOSE:
 * Removes a configuration value from ScriptProperties.
 * 
 * PARAMETERS:
 * @param {string} key - Property key to delete
 * 
 * @returns {boolean} True if deleted or didn't exist, false on error
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * deleteConfigValue("OLD_SETTING");
 * ```
 * ==============================================================================
 */
function deleteConfigValue(key) {
  try {
    PropertiesService.getScriptProperties().deleteProperty(key);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * ==============================================================================
 * FUNCTION: getAllConfig
 * ==============================================================================
 * 
 * PURPOSE:
 * Retrieves all configuration values from ScriptProperties as an object.
 * 
 * PARAMETERS: None
 * 
 * @returns {Object} All properties as key-value pairs
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var all = getAllConfig();
 * Logger.log("Configured drives: " + Object.keys(all).filter(function(k) {
 *   return k.indexOf("DRIVE_") === 0;
 * }).join(", "));
 * ```
 * 
 * NOTES:
 * - Returns empty object if no properties set
 * - Useful for debugging or configuration export
 * ==============================================================================
 */
function getAllConfig() {
  try {
    return PropertiesService.getScriptProperties().getProperties();
  } catch (e) {
    return {};
  }
}

/**
 * ==============================================================================
 * FUNCTION: clearAllConfig
 * ==============================================================================
 * 
 * PURPOSE:
 * Deletes all ScriptProperties. Use with caution - typically for reset/initial setup.
 * 
 * PARAMETERS: None
 * 
 * @returns {boolean} True if cleared successfully
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var ui = SpreadsheetApp.getUi();
 * var response = ui.alert("Clear all settings?", ui.ButtonSet.YES_NO);
 * if (response === ui.Button.YES) {
 *   clearAllConfig();
 * }
 * ```
 * 
 * NOTES:
 * - Irreversible operation
 * - Consider backing up config before clearing
 * ==============================================================================
 */
function clearAllConfig() {
  try {
    PropertiesService.getScriptProperties().deleteAllProperties();
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * ==============================================================================
 * FUNCTION: getUserConfig / setUserConfig
 * ==============================================================================
 * 
 * PURPOSE:
 * User-scoped property storage (each user has their own settings).
 * Same interface as ScriptProperties but stored per-user.
 * 
 * PARAMETERS:
 * @param {string} key - Property key
 * @param {*} value - Value to store (for setUserConfig)
 * @param {*} defaultValue - Default if not found (for getUserConfig)
 * @param {string} type - Type conversion hint
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * // Store user's API key (only they can see it)
 * setUserConfig("PERSONAL_API_KEY", "abc123", "string");
 * 
 * // Retrieve user's preference
 * var theme = getUserConfig("THEME", "light", "string");
 * ```
 * 
 * NOTES:
 * - UserProperties are separate for each Google account
 * - Useful for API keys, personal preferences
 * - Other users cannot see your UserProperties
 * ==============================================================================
 */
function getUserConfig(key, defaultValue, type) {
  type = type || "string";
  
  try {
    var value = PropertiesService.getUserProperties().getProperty(key);
    if (value === null) return defaultValue;
    
    switch (type) {
      case "number":
        var num = parseFloat(value);
        return isNaN(num) ? defaultValue : num;
      case "boolean":
        return value === "true";
      case "json":
        return JSON.parse(value);
      default:
        return value;
    }
  } catch (e) {
    return defaultValue;
  }
}

function setUserConfig(key, value, type) {
  type = type || "string";
  
  try {
    var storeValue = type === "json" ? JSON.stringify(value) : value.toString();
    PropertiesService.getUserProperties().setProperty(key, storeValue);
    return true;
  } catch (e) {
    return false;
  }
}
