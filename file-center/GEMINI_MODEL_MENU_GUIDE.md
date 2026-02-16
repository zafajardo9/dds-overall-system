# Gemini Model Selection Menu Guide (Google Apps Script)

This guide explains how to add a **user-selectable Gemini model** in Google Sheets Apps Script, while keeping your existing extraction logic unchanged.

---

## What this solves

If your script hardcodes a Gemini model (for example `gemini-2.0-flash-exp`), model deprecations can break your flow.

This pattern lets users:

- keep the same API key
- change the Gemini model from the Sheet menu
- avoid code edits whenever Google updates model availability

---

## Architecture used

1. **Menu item** in `onOpen()` for model selection
2. **Menu item** for API key/token input
3. **Script Properties** to store selected model (`GEMINI_MODEL`) and secret (`GEMINI_API_KEY`)
4. **Simple prompt dialog** (`ui.prompt`) for user input
5. **Single URL builder helper** so API calls are always model-driven
6. Existing extraction/fallback logic remains intact

---

## API key/token input pattern (plain prompt)

Use the same minimal dialog style (`ui.prompt`) for secret input.

### Add menu entry

```javascript
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Notary Automation")
    .addItem("Set API Key", "showApiKeyDialog")
    .addItem("Set Gemini Model", "showGeminiModelDialog")
    .addToUi();
}
```

### Prompt user for API key/token

```javascript
function showApiKeyDialog() {
  var ui = SpreadsheetApp.getUi();
  var promptResult = ui.prompt(
    "Set Gemini API Key",
    "Paste your API key (starts with AIza):",
    ui.ButtonSet.OK_CANCEL,
  );

  if (promptResult.getSelectedButton() !== ui.Button.OK) return;

  var apiKey = (promptResult.getResponseText() || "").trim();
  var result = saveApiKey(apiKey);

  if (result.success) {
    ui.alert("API Key Saved", result.message, ui.ButtonSet.OK);
  } else {
    ui.alert("API Key Error", result.message, ui.ButtonSet.OK);
  }
}
```

### Save and retrieve secret from Script Properties

```javascript
function saveApiKey(apiKey) {
  if (!apiKey || apiKey.trim() === "") {
    return { success: false, message: "API key cannot be empty" };
  }

  apiKey = apiKey.trim();

  // Optional Gemini-specific validation
  if (!apiKey.startsWith("AIza")) {
    return {
      success: false,
      message: 'Invalid API key format. Should start with "AIza"',
    };
  }

  PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", apiKey);
  return { success: true, message: "API key saved successfully" };
}

function getApiKey() {
  var apiKey =
    PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  return apiKey || null;
}
```

### If using non-Gemini token

Use a separate key name, for example:

```javascript
PropertiesService.getScriptProperties().setProperty(
  "SERVICE_TOKEN",
  tokenValue,
);
var tokenValue =
  PropertiesService.getScriptProperties().getProperty("SERVICE_TOKEN");
```

---

## Step 1) Add menu item

Add a model setting entry in your custom menu:

```javascript
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Notary Automation")
    .addItem("Set API Key", "showApiKeyDialog")
    .addItem("Set Gemini Model", "showGeminiModelDialog")
    .addItem("Test API Key", "testApiKey")
    .addToUi();
}
```

---

## Step 2) Add model config helpers

Use Script Properties so the model persists without code changes.

```javascript
function getDefaultGeminiModel() {
  return "gemini-3.0-flash";
}

function normalizeGeminiModelName(modelName) {
  if (!modelName) return "";
  return modelName
    .toString()
    .trim()
    .replace(/^models\//i, "");
}

function saveGeminiModel(modelName) {
  var normalizedModel = normalizeGeminiModelName(modelName);

  if (!normalizedModel) {
    return { success: false, message: "Model name cannot be empty" };
  }

  if (!/^[-a-zA-Z0-9._]+$/.test(normalizedModel)) {
    return {
      success: false,
      message: "Invalid model name format. Use values like gemini-3.0-flash",
    };
  }

  if (!/^gemini-/i.test(normalizedModel)) {
    return {
      success: false,
      message: "Model name should start with 'gemini-'",
    };
  }

  PropertiesService.getScriptProperties().setProperty(
    "GEMINI_MODEL",
    normalizedModel,
  );

  return { success: true, message: "Gemini model saved: " + normalizedModel };
}

function getGeminiModel() {
  var model =
    PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
  var normalizedModel = normalizeGeminiModelName(model);
  return normalizedModel || getDefaultGeminiModel();
}
```

---

## Step 3) Add user input dialog (plain Apps Script UI)

Use a simple prompt to match a minimal and functional UX.

```javascript
function showGeminiModelDialog() {
  var ui = SpreadsheetApp.getUi();
  var currentModel = getGeminiModel();

  var promptResult = ui.prompt(
    "Set Gemini Model",
    "Current model: " +
      currentModel +
      "\n\nEnter model (example: gemini-3.0-flash)",
    ui.ButtonSet.OK_CANCEL,
  );

  if (promptResult.getSelectedButton() !== ui.Button.OK) return;

  var modelInput = (promptResult.getResponseText() || "").trim();
  var result = saveGeminiModel(modelInput);

  if (result.success) {
    ui.alert("Model Saved", result.message, ui.ButtonSet.OK);
  } else {
    ui.alert("Model Error", result.message, ui.ButtonSet.OK);
  }
}
```

---

## Step 4) Centralize Gemini endpoint building

Never hardcode model names in multiple places.

```javascript
function buildGeminiGenerateContentUrl(apiKey) {
  var model = getGeminiModel();
  return (
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model +
    ":generateContent?key=" +
    apiKey
  );
}
```

---

## Step 5) Replace hardcoded model URLs

Wherever you call Gemini, replace hardcoded URLs like:

```javascript
// old
".../models/gemini-2.0-flash-exp:generateContent?key=" + apiKey;
```

with:

```javascript
// new
var url = buildGeminiGenerateContentUrl(apiKey);
```

Apply this to:

- API key test endpoint
- document analysis endpoint
- any other Gemini calls

---

## Step 6) (Optional but recommended) Show active model in status/test dialogs

Include the selected model in your status and test messages so users can verify which model is active.

---

## Drop-in checklist for other projects

- [ ] Add menu item: `Set API Key` (or `Set Token`)
- [ ] Add `showApiKeyDialog()` prompt flow
- [ ] Add `saveApiKey()` (or `saveServiceToken()`) with validation
- [ ] Add `getApiKey()` (or `getServiceToken()`)
- [ ] Add menu item: `Set Gemini Model`
- [ ] Add `getDefaultGeminiModel()`
- [ ] Add `normalizeGeminiModelName()`
- [ ] Add `saveGeminiModel()`
- [ ] Add `getGeminiModel()`
- [ ] Add `buildGeminiGenerateContentUrl(apiKey)`
- [ ] Replace all hardcoded Gemini model URLs
- [ ] Keep existing extraction/fallback logic untouched

---

## Notes

- Script Properties are project-wide for the Apps Script project.
- If user input is blank/invalid, keep previous model unchanged.
- You can change default model later by editing only `getDefaultGeminiModel()`.
- Avoid logging full API keys/tokens in `console.log`; only log masked values if needed.
