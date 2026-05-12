// Career module: Application tracking DSS standardization.
const APPLICATIONS_SHEET_NAME = "Applications";
const APPLICATIONS_V3_HEADERS = [
  "Company Name",
  "Job Title",
  "Category",
  "Date Applied",
  "Status",
  "Next Action",
  "Target Date",
  "FU Required?",
  "Follow Up Date",
  "Notes"
];

const APPLICATIONS_STATUS_OPTIONS = [
  "Saved",
  "Applied",
  "Assessment",
  "Interview",
  "Rejected",
  "Offer",
  "Withdrawn"
];

const APPLICATIONS_FU_OPTIONS = ["Yes", "No"];

function standardizeApplicationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(APPLICATIONS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(APPLICATIONS_SHEET_NAME);
  }

  const existingValues = getApplicationsExistingValues_(sheet);
  const migratedRows = migrateApplicationsRowsToV3_(existingValues);

  sheet.clearContents();
  sheet
    .getRange(1, 1, 1, APPLICATIONS_V3_HEADERS.length)
    .setValues([APPLICATIONS_V3_HEADERS]);

  if (migratedRows.length > 0) {
    sheet
      .getRange(2, 1, migratedRows.length, APPLICATIONS_V3_HEADERS.length)
      .setValues(migratedRows);
  }

  ensureApplicationsMinimumRows_(sheet);
  applyApplicationsDataValidation_(sheet);
  applyPortfolioSheetTheme();
  writeSystemLog(
    "Career",
    "Standardize Sheet",
    "Success",
    "Applications sheet migrated to V3.0 standard"
  );
}

function getApplicationsExistingValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 1 || lastColumn < 1) {
    return [];
  }

  return sheet.getRange(1, 1, lastRow, lastColumn).getValues();
}

function migrateApplicationsRowsToV3_(values) {
  if (values.length < 2) {
    return [];
  }

  const headerMap = buildApplicationsHeaderMap_(values[0]);

  return values
    .slice(1)
    .filter(row => row.some(value => value !== "" && value !== null))
    .map(row => [
      getApplicationValue_(row, headerMap, ["Company Name", "Company"]),
      getApplicationValue_(row, headerMap, ["Job Title", "Position"]),
      getApplicationValue_(row, headerMap, ["Category"]),
      getApplicationValue_(row, headerMap, ["Date Applied", "Date"]),
      getApplicationValue_(row, headerMap, ["Status"]),
      getApplicationValue_(row, headerMap, ["Next Action"]),
      getApplicationValue_(row, headerMap, ["Target Date"]),
      getApplicationValue_(row, headerMap, ["FU Required?", "FU"]),
      getApplicationValue_(row, headerMap, ["Follow Up Date", "Date FU"]),
      buildApplicationsNotes_(row, headerMap)
    ]);
}

function buildApplicationsHeaderMap_(headers) {
  return headers.reduce((map, header, index) => {
    const normalizedHeader = normalizeApplicationsHeader_(header);

    if (normalizedHeader) {
      map[normalizedHeader] = index;
    }

    return map;
  }, {});
}

function getApplicationValue_(row, headerMap, candidateHeaders) {
  for (let i = 0; i < candidateHeaders.length; i++) {
    const index = headerMap[normalizeApplicationsHeader_(candidateHeaders[i])];

    if (index !== undefined) {
      return row[index];
    }
  }

  return "";
}

function buildApplicationsNotes_(row, headerMap) {
  const noteParts = [];
  const legacyFields = ["Location", "Link", "Review"];
  const hasLegacyNoteSources = legacyFields.some(
    fieldName => headerMap[normalizeApplicationsHeader_(fieldName)] !== undefined
  );

  if (!hasLegacyNoteSources) {
    return getApplicationValue_(row, headerMap, ["Notes"]);
  }

  const fieldsToCombine = ["Location", "Link", "Notes", "Review"];

  fieldsToCombine.forEach(fieldName => {
    const value = getApplicationValue_(row, headerMap, [fieldName]);

    if (value !== "" && value !== null) {
      noteParts.push(fieldName + ": " + value);
    }
  });

  return noteParts.join("\n");
}

function ensureApplicationsMinimumRows_(sheet) {
  if (sheet.getMaxRows() < 2) {
    sheet.insertRowsAfter(1, 1);
  }
}

function applyApplicationsDataValidation_(sheet) {
  const validationRowCount = sheet.getMaxRows() - 1;
  const statusRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(APPLICATIONS_STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  const fuRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(APPLICATIONS_FU_OPTIONS, true)
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(2, 5, validationRowCount, 1)
    .setDataValidation(statusRule);
  sheet
    .getRange(2, 8, validationRowCount, 1)
    .setDataValidation(fuRule);
}

function normalizeApplicationsHeader_(header) {
  return String(header || "").trim().toLowerCase();
}
