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


function applyPortfolioSheetTheme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();


  themeSheetIfExists_(ss, "Dashboard", themeDashboardSheet_);
  themeSheetIfExists_(ss, "Applications", themeApplicationsSheet_);
  themeSheetIfExists_(ss, "AI_Audit", themeAIAuditSheet_);
  // Log is intentionally excluded because it has a specialized operational layout.
}


function themeSheetIfExists_(ss, sheetName, themeFn) {
  const sheet = ss.getSheetByName(sheetName);


  if (!sheet) {
    writeSystemLog(
      "Spreadsheet UI",
      "Apply Portfolio Sheet Theme",
      "Skipped",
      "Sheet not found: " + sheetName
    );
    return;
  }


  try {
    themeFn(sheet);
  } catch (error) {
    if (isTypedColumnsError_(error)) {
      console.warn(
        "Skipped theme for " + sheetName +
        " because typed columns do not allow this formatting operation: " +
        getErrorMessage_(error)
      );
      return;
    }

    throw error;
  }
}


function themeDashboardSheet_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const lastRow = sheet.getLastRow();


  if (lastRow >= 1) {
    sheet.setFrozenRows(1);
    styleHeaderRow_(sheet, 1, lastColumn);
  }


  styleUsedRange_(sheet);
  setColumnWidths_(sheet, [
    120,
    150,
    260,
    110,
    140,
    140,
    260
  ]);


  if (lastColumn >= 1 && lastRow > 0) {
    applyNumberFormatToUsedRows_(sheet, 1, "dd mmm yyyy");
  }
}


function applyNumberFormatToUsedRows_(sheet, column, numberFormat) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 0) return;

  try {
    sheet
      .getRange(1, column, lastRow, 1)
      .setNumberFormat(numberFormat);
  } catch (error) {
    if (isTypedColumnsError_(error)) {
      console.warn(
        "Skipped number format for " + sheet.getName() +
        " column " + column +
        " because of typed columns: " + getErrorMessage_(error)
      );
      return;
    }

    throw error;
  }
}


function isTypedColumnsError_(error) {
  return getErrorMessage_(error).toLowerCase().indexOf("typed columns") !== -1;
}


function getErrorMessage_(error) {
  return String(error && error.message ? error.message : error);
}


function themeApplicationsSheet_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const lastRow = sheet.getLastRow();


  if (lastRow >= 1) {
    sheet.setFrozenRows(1);
    styleHeaderRow_(sheet, 1, lastColumn);
  }


  styleUsedRange_(sheet);
  setColumnWidths_(sheet, [
    180,
    220,
    180,
    120,
    160,
    180,
    180,
    130,
    260,
    260
  ]);


  if (lastColumn >= 4) {
    applyNumberFormatToUsedRows_(sheet, 4, "dd mmm yyyy");
  }


  if (lastColumn >= 8) {
    applyStatusConditionalFormatting_(sheet, sheet.getRange(2, 8, Math.max(lastRow - 1, 1), 1));
  }
}


function themeAIAuditSheet_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const lastRow = sheet.getLastRow();


  if (lastRow >= 1) {
    sheet.setFrozenRows(1);
    styleHeaderRow_(sheet, 1, lastColumn);
  }


  styleUsedRange_(sheet);
  sheet.setColumnWidth(1, 160);


  if (lastColumn >= 2) {
    sheet.setColumnWidth(2, 720);
    if (lastRow > 0) {
      sheet
        .getRange(1, 2, lastRow, 1)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setVerticalAlignment("top");
    }
  }


  applyNumberFormatToUsedRows_(sheet, 1, "dd mmm yyyy hh:mm");
}


function styleUsedRange_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();


  if (lastRow < 1 || lastColumn < 1) return;


  sheet
    .getRange(1, 1, lastRow, lastColumn)
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("top")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setBorder(true, true, true, true, true, true, "#d9e2ec", SpreadsheetApp.BorderStyle.SOLID);
}


function styleHeaderRow_(sheet, row, lastColumn) {
  if (lastColumn < 1) return;


  sheet
    .getRange(row, 1, 1, lastColumn)
    .setBackground("#1f2937")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);


  sheet.setRowHeight(row, 34);
}


function setColumnWidths_(sheet, widths) {
  const maxColumns = sheet.getMaxColumns();


  widths.forEach((width, index) => {
    const column = index + 1;


    if (column <= maxColumns) {
      sheet.setColumnWidth(column, width);
    }
  });
}


function applyStatusConditionalFormatting_(sheet, range) {
  const existingRules = sheet.getConditionalFormatRules();
  const statusRules = [
    { text: "Applied", background: "#dbeafe", color: "#1e3a8a" },
    { text: "Saved", background: "#fef3c7", color: "#92400e" },
    { text: "Interview", background: "#ede9fe", color: "#5b21b6" },
    { text: "Rejected", background: "#fee2e2", color: "#991b1b" },
    { text: "Offer", background: "#dcfce7", color: "#166534" },
    { text: "OK", background: "#dcfce7", color: "#166534" },
    { text: "Needs Review", background: "#ffedd5", color: "#9a3412" },
    { text: "Success", background: "#dcfce7", color: "#166534" },
    { text: "Error", background: "#fee2e2", color: "#991b1b" }
  ];


  const newRules = statusRules.map(statusRule =>
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo(statusRule.text)
      .setBackground(statusRule.background)
      .setFontColor(statusRule.color)
      .setRanges([range])
      .build()
  );


  sheet.setConditionalFormatRules(existingRules.concat(newRules));
}
