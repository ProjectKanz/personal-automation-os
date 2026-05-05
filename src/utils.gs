// Shared utilities and globals (Phase 1-2 split only).
// TODO(v2.2): Add centralized property validation and safer error redaction.
const GEMINI_API_KEY = PropertiesService
.getScriptProperties()
.getProperty("GEMINI_API_KEY"); 

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
