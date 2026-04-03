/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * GeminiAPI.gs - Common Gemini AI Integration Functions
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * PURPOSE:
 * This file contains reusable functions for interacting with Google's Gemini AI API.
 * Use these functions across all DDS projects that require AI-powered document 
 * analysis, title extraction, or content generation.
 * 
 * PREREQUISITES:
 * - Gemini API key must be stored via setGeminiApiKey() before using other functions
 * - Required OAuth scope: https://www.googleapis.com/auth/generative-language
 * 
 * USAGE:
 * Copy this file into your Apps Script project and call functions as needed.
 * All functions include error handling and rate limiting considerations.
 * 
 * VERSION: 1.0.0
 * LAST UPDATED: April 2026
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/**
 * ==============================================================================
 * FUNCTION: fetchGeminiModels
 * ==============================================================================
 * 
 * PURPOSE:
 * Fetches the complete list of available Gemini AI models from Google's API.
 * This allows users to select the most appropriate model for their use case
 * (e.g., gemini-2.0-flash for speed, gemini-2.0-pro for complex analysis).
 * 
 * PARAMETERS:
 * @param {string} apiKey - (Optional) Gemini API key. If not provided, reads from 
 *                          UserProperties with key "GEMINI_API_KEY"
 * 
 * @returns {Object} Result object with structure:
 *   - success: {boolean} - Whether the fetch succeeded
 *   - models: {Array} - List of available models (if success=true)
 *   - error: {string} - Error message (if success=false)
 *   - recommended: {string} - The model marked as recommended by the API
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = fetchGeminiModels();
 * if (result.success) {
 *   result.models.forEach(function(model) {
 *     Logger.log(model.name + " - " + model.description);
 *   });
 * }
 * ```
 * 
 * NOTES:
 * - Requires valid Gemini API key to be set first
 * - API endpoint: https://generativelanguage.googleapis.com/v1beta/models
 * - Models are filtered to only show Gemini family models (excludes embedding models)
 * - Each model object contains: name, version, displayName, description, inputTokenLimit, 
 *   outputTokenLimit, supportedGenerationMethods, temperature, topP, topK
 * ==============================================================================
 */
function fetchGeminiModels(apiKey) {
  var key = apiKey || PropertiesService.getUserProperties().getProperty("GEMINI_API_KEY");
  
  if (!key) {
    return {
      success: false,
      error: "No API key found. Please run setGeminiApiKey() first.",
      models: [],
      recommended: null
    };
  }
  
  var url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + key;
  
  var options = {
    method: "GET",
    headers: {
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    
    if (responseCode !== 200) {
      return {
        success: false,
        error: "API Error (" + responseCode + "): " + responseBody,
        models: [],
        recommended: null
      };
    }
    
    var data = JSON.parse(responseBody);
    var models = data.models || [];
    
    // Filter to Gemini models only and format
    var geminiModels = models
      .filter(function(model) {
        return model.name && model.name.indexOf("gemini") !== -1;
      })
      .map(function(model) {
        return {
          name: model.name.replace("models/", ""),
          fullName: model.name,
          version: model.version || "unknown",
          displayName: model.displayName || model.name,
          description: model.description || "No description available",
          inputTokenLimit: model.inputTokenLimit || 0,
          outputTokenLimit: model.outputTokenLimit || 0,
          temperature: model.temperature || 0.9,
          topP: model.topP || 0.95,
          topK: model.topK || 40,
          supportedMethods: model.supportedGenerationMethods || []
        };
      });
    
    // Determine recommended model (prefer flash for speed, pro for quality)
    var recommended = geminiModels.find(function(m) {
      return m.name.indexOf("flash") !== -1 && m.name.indexOf("2.0") !== -1;
    });
    
    if (!recommended && geminiModels.length > 0) {
      recommended = geminiModels[0];
    }
    
    return {
      success: true,
      models: geminiModels,
      recommended: recommended ? recommended.name : null,
      count: geminiModels.length
    };
    
  } catch (error) {
    return {
      success: false,
      error: "Fetch failed: " + error.toString(),
      models: [],
      recommended: null
    };
  }
}

/**
 * ==============================================================================
 * FUNCTION: setGeminiApiKey
 * ==============================================================================
 * 
 * PURPOSE:
 * Securely stores the Gemini API key in UserProperties for use across all
 * Gemini-related functions. This prevents hardcoding API keys in scripts.
 * 
 * PARAMETERS:
 * @param {string} apiKey - The Gemini API key from https://aistudio.google.com/app/apikey
 *                          If not provided, shows a prompt dialog for input.
 * 
 * @returns {boolean} - True if key was stored successfully, false otherwise
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * // Store key programmatically
 * setGeminiApiKey("your-api-key-here");
 * 
 * // Or let user enter via dialog
 * setGeminiApiKey(); // Shows input prompt
 * ```
 * 
 * NOTES:
 * - Stores in UserProperties (not ScriptProperties) so each user has their own key
 * - Key is validated with a test API call before storing
 * - Use deleteGeminiApiKey() to remove stored key
 * ==============================================================================
 */
function setGeminiApiKey(apiKey) {
  var key = apiKey;
  
  // If no key provided, prompt user
  if (!key) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "Set Gemini API Key",
      "Enter your Gemini API key from https://aistudio.google.com/app/apikey:\n\n" +
      "(Leave blank to clear existing key)",
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return false;
    }
    
    key = response.getResponseText().trim();
    
    // Empty input means delete
    if (!key) {
      return deleteGeminiApiKey();
    }
  }
  
  // Validate key with test call
  var testResult = fetchGeminiModels(key);
  
  if (!testResult.success) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Invalid API Key", "The provided key failed validation: " + testResult.error, ui.ButtonSet.OK);
    return false;
  }
  
  // Store valid key
  PropertiesService.getUserProperties().setProperty("GEMINI_API_KEY", key);
  
  // Also store default model if not set
  var currentModel = PropertiesService.getUserProperties().getProperty("GEMINI_SELECTED_MODEL");
  if (!currentModel && testResult.recommended) {
    PropertiesService.getUserProperties().setProperty("GEMINI_SELECTED_MODEL", testResult.recommended);
  }
  
  SpreadsheetApp.getUi().alert(
    "Success",
    "API key validated and stored.\nDetected " + testResult.count + " models.\nRecommended: " + testResult.recommended,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  return true;
}

/**
 * ==============================================================================
 * FUNCTION: deleteGeminiApiKey
 * ==============================================================================
 * 
 * PURPOSE:
 * Removes the stored Gemini API key from UserProperties.
 * Use this when rotating keys or clearing credentials.
 * 
 * PARAMETERS: None
 * 
 * @returns {boolean} - True if key was deleted or didn't exist
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * deleteGeminiApiKey();
 * ```
 * ==============================================================================
 */
function deleteGeminiApiKey() {
  PropertiesService.getUserProperties().deleteProperty("GEMINI_API_KEY");
  PropertiesService.getUserProperties().deleteProperty("GEMINI_SELECTED_MODEL");
  return true;
}

/**
 * ==============================================================================
 * FUNCTION: getStoredGeminiModel
 * ==============================================================================
 * 
 * PURPOSE:
 * Retrieves the currently selected Gemini model from UserProperties.
 * Falls back to a default if none is stored.
 * 
 * PARAMETERS: None
 * 
 * @returns {string} - The model name to use for API calls (e.g., "gemini-2.0-flash")
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var model = getStoredGeminiModel();
 * // Use model in API call...
 * ```
 * 
 * NOTES:
 * - Default fallback: "gemini-2.0-flash"
 * - Model is set automatically when storing API key, or via selectGeminiModel()
 * ==============================================================================
 */
function getStoredGeminiModel() {
  return PropertiesService.getUserProperties().getProperty("GEMINI_SELECTED_MODEL") || "gemini-2.0-flash";
}

/**
 * ==============================================================================
 * FUNCTION: selectGeminiModel
 * ==============================================================================
 * 
 * PURPOSE:
 * Shows a dialog allowing user to select from available Gemini models.
 * Selected model is stored in UserProperties for future API calls.
 * 
 * PARAMETERS: None (interactive function)
 * 
 * @returns {boolean} - True if a model was selected, false if cancelled
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * selectGeminiModel(); // Shows selection dialog
 * ```
 * 
 * NOTES:
 * - Fetches live model list from API
 * - Displays models with ★ marker for recommended option
 * - Requires valid API key to be set first
 * ==============================================================================
 */
function selectGeminiModel() {
  var apiKey = PropertiesService.getUserProperties().getProperty("GEMINI_API_KEY");
  
  if (!apiKey) {
    SpreadsheetApp.getUi().alert("Please set your API key first (setGeminiApiKey)");
    return false;
  }
  
  var result = fetchGeminiModels(apiKey);
  
  if (!result.success) {
    SpreadsheetApp.getUi().alert("Failed to fetch models: " + result.error);
    return false;
  }
  
  var currentModel = getStoredGeminiModel();
  var modelsList = result.models.map(function(m) {
    var marker = (m.name === result.recommended) ? "★ " : "";
    var selected = (m.name === currentModel) ? " (current)" : "";
    return marker + m.name + " - " + m.description.substring(0, 60) + "..." + selected;
  }).join("\n");
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Select Gemini Model",
    "Available models (★ = recommended):\n\n" + modelsList + "\n\nEnter model name:",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return false;
  }
  
  var selected = response.getResponseText().trim();
  
  // Validate selection
  var isValid = result.models.some(function(m) { return m.name === selected; });
  if (!isValid) {
    ui.alert("Invalid model name. Please choose from the list above.");
    return false;
  }
  
  PropertiesService.getUserProperties().setProperty("GEMINI_SELECTED_MODEL", selected);
  ui.alert("Model set to: " + selected);
  return true;
}

/**
 * ==============================================================================
 * FUNCTION: callGemini
 * ==============================================================================
 * 
 * PURPOSE:
 * Makes a Gemini API call with text-only prompt. Use this for general
 * text generation, analysis, or question answering tasks.
 * 
 * PARAMETERS:
 * @param {string} prompt - The text prompt to send to Gemini
 * @param {Object} options - (Optional) Configuration options:
 *   - model: {string} - Model name (default: stored model or gemini-2.0-flash)
 *   - temperature: {number} - Creativity level 0-2 (default: 0.2)
 *   - maxOutputTokens: {number} - Max response length (default: 1024)
 *   - apiKey: {string} - Override stored API key
 * 
 * @returns {Object} Result object:
 *   - success: {boolean} - Whether call succeeded
 *   - text: {string} - Generated text (if success=true)
 *   - error: {string} - Error message (if success=false)
 *   - usage: {Object} - Token usage info (promptTokenCount, candidatesTokenCount)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var result = callGemini("Summarize this: " + documentText, {
 *   temperature: 0.3,
 *   maxOutputTokens: 500
 * });
 * if (result.success) {
 *   Logger.log(result.text);
 * }
 * ```
 * 
 * NOTES:
 * - For file analysis, use callGeminiWithFile() instead
 * - Temperature: Lower (0.1) = more focused, Higher (0.9) = more creative
 * - Automatically handles API key retrieval
 * ==============================================================================
 */
function callGemini(prompt, options) {
  options = options || {};
  
  var apiKey = options.apiKey || PropertiesService.getUserProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    return { success: false, error: "No API key found. Run setGeminiApiKey() first.", text: "" };
  }
  
  var model = options.model || getStoredGeminiModel();
  var temperature = options.temperature !== undefined ? options.temperature : 0.2;
  var maxTokens = options.maxOutputTokens || 1024;
  
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;
  
  var payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: temperature,
      maxOutputTokens: maxTokens,
      topP: 0.95,
      topK: 40
    }
  };
  
  var fetchOptions = {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, fetchOptions);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode !== 200) {
      return { 
        success: false, 
        error: "API Error " + responseCode + ": " + responseText,
        text: ""
      };
    }
    
    var data = JSON.parse(responseText);
    
    if (!data.candidates || data.candidates.length === 0) {
      return { success: false, error: "No response generated", text: "" };
    }
    
    var generatedText = data.candidates[0].content.parts[0].text || "";
    
    return {
      success: true,
      text: generatedText,
      usage: data.usageMetadata || {}
    };
    
  } catch (error) {
    return { success: false, error: error.toString(), text: "" };
  }
}

/**
 * ==============================================================================
 * FUNCTION: callGeminiWithFile
 * ==============================================================================
 * 
 * PURPOSE:
 * Makes a Gemini API call with both text prompt AND file attachment.
 * Use this for document analysis, OCR, title extraction, or any file-based AI task.
 * 
 * PARAMETERS:
 * @param {File} file - Google Drive File object or Blob to analyze
 * @param {string} prompt - The text prompt describing what to extract/analyze
 * @param {Object} options - (Optional) Configuration options:
 *   - model: {string} - Model name (default: gemini-2.0-flash)
 *   - temperature: {number} - Creativity level 0-2 (default: 0.2)
 *   - maxOutputTokens: {number} - Max response length (default: 1024)
 *   - mimeType: {string} - Force specific MIME type (auto-detected if not set)
 * 
 * @returns {Object} Result object:
 *   - success: {boolean} - Whether call succeeded
 *   - text: {string} - Generated text/analysis (if success=true)
 *   - error: {string} - Error message (if success=false)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var file = DriveApp.getFileById(fileId);
 * var result = callGeminiWithFile(file, 
 *   "Extract the document title and format as: YYYY-MM-DD Descriptive Title",
 *   { temperature: 0.1 }
 * );
 * if (result.success) {
 *   var cleanTitle = result.text.trim();
 * }
 * ```
 * 
 * NOTES:
 * - Automatically converts file to base64 for API transmission
 * - Supported formats: PDF, images (jpg/png), DOC, DOCX (via conversion)
 * - Large files may take longer to process
 * - For multiple files, call separately or use callGeminiWithMultipleFiles()
 * ==============================================================================
 */
function callGeminiWithFile(file, prompt, options) {
  options = options || {};
  
  try {
    var blob = file.getBlob ? file.getBlob() : file;
    var mimeType = options.mimeType || blob.getContentType();
    var base64Data = Utilities.base64Encode(blob.getBytes());
    
    var apiKey = PropertiesService.getUserProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
      return { success: false, error: "No API key found", text: "" };
    }
    
    var model = options.model || getStoredGeminiModel();
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;
    
    var payload = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: mimeType,
              data: base64Data
            }
          }
        ]
      }],
      generationConfig: {
        temperature: options.temperature !== undefined ? options.temperature : 0.2,
        maxOutputTokens: options.maxOutputTokens || 1024,
        topP: 0.95,
        topK: 40
      }
    };
    
    var fetchOptions = {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, fetchOptions);
    var responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      return { 
        success: false, 
        error: "API Error " + responseCode + ": " + response.getContentText(),
        text: ""
      };
    }
    
    var data = JSON.parse(response.getContentText());
    
    if (!data.candidates || data.candidates.length === 0) {
      return { success: false, error: "No response generated", text: "" };
    }
    
    return {
      success: true,
      text: data.candidates[0].content.parts[0].text || "",
      usage: data.usageMetadata || {}
    };
    
  } catch (error) {
    return { success: false, error: "File processing error: " + error.toString(), text: "" };
  }
}

/**
 * ==============================================================================
 * FUNCTION: extractDocumentTitle
 * ==============================================================================
 * 
 * PURPOSE:
 * Specialized function for extracting and formatting document titles from files.
 * Uses Gemini AI to analyze document content and produce clean, standardized titles.
 * 
 * PARAMETERS:
 * @param {File} file - Google Drive File to analyze
 * @param {Object} options - (Optional) Formatting options:
 *   - prefixDate: {boolean} - Prepend YYYY-MM-DD date (default: true)
 *   - dateSource: {string} - "file" (creation date), "today", or "none" (default: "file")
 *   - maxLength: {number} - Maximum title length (default: 100)
 *   - customPrompt: {string} - Override default extraction prompt
 * 
 * @returns {Object} Result object:
 *   - success: {boolean} - Whether extraction succeeded
 *   - title: {string} - Formatted title (if success=true)
 *   - rawText: {string} - Raw AI response before formatting
 *   - error: {string} - Error message (if success=false)
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var file = DriveApp.getFileById("abc123");
 * var result = extractDocumentTitle(file, {
 *   prefixDate: true,
 *   dateSource: "file",
 *   maxLength: 80
 * });
 * if (result.success) {
 *   file.setName(result.title); // Rename file
 * }
 * ```
 * 
 * NOTES:
 * - Default prompt focuses on legal/business document types
 * - Date extraction attempts to find dates within document content
 * - Falls back to existing filename if AI extraction fails
 * - Commonly used in: file-center, litigation, notary projects
 * ==============================================================================
 */
function extractDocumentTitle(file, options) {
  options = options || {};
  
  var defaultPrompt = "Analyze this document and extract a clean, descriptive title. " +
    "The title should:\n" +
    "1. Start with the most relevant date found in the document (YYYY-MM-DD format)\n" +
    "2. Include the document type (e.g., Contract, Invoice, Letter, Pleading)\n" +
    "3. Include key parties or subject matter\n" +
    "4. Be concise but informative (under 80 characters)\n" +
    "5. Remove unnecessary words like 'the', 'a', 'an' when possible\n\n" +
    "Return ONLY the title, nothing else.";
  
  var prompt = options.customPrompt || defaultPrompt;
  
  var result = callGeminiWithFile(file, prompt, {
    temperature: 0.1,  // Low temperature for consistency
    maxOutputTokens: 150
  });
  
  if (!result.success) {
    return {
      success: false,
      title: file.getName(),
      rawText: "",
      error: result.error
    };
  }
  
  var rawTitle = result.text.trim();
  
  // Clean up the title
  var cleanTitle = rawTitle
    .replace(/[<>:"/\\|?*]/g, "")  // Remove illegal filename characters
    .replace(/\s+/g, " ")            // Normalize whitespace
    .trim();
  
  // Truncate if too long
  if (options.maxLength && cleanTitle.length > options.maxLength) {
    cleanTitle = cleanTitle.substring(0, options.maxLength).trim() + "...";
  }
  
  // Add date prefix if requested
  if (options.prefixDate !== false) {
    var date;
    if (options.dateSource === "today") {
      date = new Date();
    } else if (options.dateSource === "file" || !options.dateSource) {
      date = file.getDateCreated ? file.getDateCreated() : new Date();
    }
    
    if (date) {
      var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      cleanTitle = dateStr + " " + cleanTitle;
    }
  }
  
  return {
    success: true,
    title: cleanTitle,
    rawText: rawTitle,
    error: null
  };
}

/**
 * ==============================================================================
 * FUNCTION: testGeminiConnection
 * ==============================================================================
 * 
 * PURPOSE:
 * Diagnostic function to verify Gemini API connectivity and configuration.
 * Tests API key, model availability, and basic generation capability.
 * 
 * PARAMETERS: None
 * 
 * @returns {Object} Diagnostic result:
 *   - success: {boolean} - Overall test passed
 *   - apiKeySet: {boolean} - API key is stored
 *   - modelAvailable: {boolean} - Selected model is accessible
 *   - testGeneration: {boolean} - Basic text generation works
 *   - selectedModel: {string} - Currently selected model
 *   - errors: {Array} - List of any errors encountered
 * 
 * USAGE EXAMPLE:
 * ```javascript
 * var test = testGeminiConnection();
 * if (!test.success) {
 *   Logger.log("Issues found: " + test.errors.join(", "));
 * }
 * ```
 * 
 * NOTES:
 * - Safe to run multiple times (idempotent)
 * - Shows UI alert with results when run from menu
 * - Useful for initial setup troubleshooting
 * ==============================================================================
 */
function testGeminiConnection() {
  var results = {
    success: false,
    apiKeySet: false,
    modelAvailable: false,
    testGeneration: false,
    selectedModel: null,
    errors: []
  };
  
  // Check API key
  var apiKey = PropertiesService.getUserProperties().getProperty("GEMINI_API_KEY");
  results.apiKeySet = !!apiKey;
  if (!results.apiKeySet) {
    results.errors.push("No API key stored");
  }
  
  // Check model selection
  var model = getStoredGeminiModel();
  results.selectedModel = model;
  
  // Test model list fetch
  if (results.apiKeySet) {
    var modelsResult = fetchGeminiModels(apiKey);
    results.modelAvailable = modelsResult.success;
    if (!modelsResult.success) {
      results.errors.push("Model fetch failed: " + modelsResult.error);
    }
  }
  
  // Test simple generation
  if (results.apiKeySet && results.modelAvailable) {
    var genResult = callGemini("Say 'test successful' only", { maxOutputTokens: 10 });
    results.testGeneration = genResult.success && genResult.text.toLowerCase().indexOf("test successful") !== -1;
    if (!genResult.success) {
      results.errors.push("Generation test failed: " + genResult.error);
    }
  }
  
  results.success = results.apiKeySet && results.modelAvailable && results.testGeneration;
  
  // Show UI alert
  var ui = SpreadsheetApp.getUi();
  var message = "Gemini Connection Test Results:\n\n" +
    "API Key: " + (results.apiKeySet ? "✓ Set" : "✗ Missing") + "\n" +
    "Model Access: " + (results.modelAvailable ? "✓ OK" : "✗ Failed") + "\n" +
    "Generation: " + (results.testGeneration ? "✓ Working" : "✗ Failed") + "\n" +
    "Selected Model: " + results.selectedModel + "\n\n" +
    (results.success ? "All tests passed!" : "Issues:\n" + results.errors.join("\n"));
  
  ui.alert("Gemini Diagnostics", message, ui.ButtonSet.OK);
  
  return results;
}
