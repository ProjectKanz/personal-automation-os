// Shared utilities and globals (Phase 1-2 split only).
// TODO(v2.2): Add centralized property validation and safer error redaction.
function getRequiredProperty(propertyName) {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty(propertyName);

  if (value === null || value === undefined || value === "") {
    throw new Error("Missing required Script Property: " + propertyName);
  }

  return value;
}

const GEMINI_API_KEY = getRequiredProperty("GEMINI_API_KEY");

function writeSystemLog(module, action, status, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("System_Log");


  if (!logSheet) {
    logSheet = ss.insertSheet("System_Log");
    logSheet.appendRow(["Timestamp", "Module", "Action", "Status", "Message"]);
  }


  logSheet.appendRow([
    new Date(),
    module,
    action,
    status,
    message
  ]);
}
