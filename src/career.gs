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
  cleanupApplicationsFormatting_(sheet);
  applyApplicationsDataValidation_(sheet);
  applyPortfolioSheetTheme();
  writeSystemLog(
    "Career",
    "Standardize Sheet",
    "Success",
    "Applications sheet migrated to V3.0 standard"
  );
}

function getFollowUpList() {
  const followUpItems = getCareerFollowUpItems_();
  const today = normalizeApplicationDate_(new Date());

  if (followUpItems.length === 0) {
    return "✅ All clear! No applications require follow-up today.";
  }

  return (
    "📌 *Executive Follow-Up Memo*\n" +
    "Date: " + escapeTelegramMarkdown(formatApplicationDate_(today)) + "\n" +
    "Pending items: " + followUpItems.length + "\n\n" +
    followUpItems.map(formatApplicationFollowUpMemo_).join("\n\n")
  );
}

function getCareerDashboardStats() {
  const rows = getApplicationDataRows_();
  const totalApplications = rows.length;
  const offerCount = rows.filter(item => item.status === "offer").length;
  const rejectedCount = rows.filter(item => item.status === "rejected").length;
  const interviewPipeline = rows.filter(item => item.status === "interview").length;

  return {
    totalApplications: totalApplications,
    offerCount: offerCount,
    rejectedCount: rejectedCount,
    interviewPipeline: interviewPipeline,
    followUpsDue: getCareerFollowUpItems_().length,
    successRate: calculateApplicationRate_(offerCount, totalApplications),
    rejectionRate: calculateApplicationRate_(rejectedCount, totalApplications)
  };
}

function buildCareerDashboardMessage() {
  const stats = getCareerDashboardStats();

  return (
    "📊 *Executive Career Dashboard*\n" +
    "Total Applications: " + stats.totalApplications + "\n" +
    "Success Rate: " + stats.successRate + "% (" + stats.offerCount + " offer)\n" +
    "Rejection Rate: " + stats.rejectionRate + "% (" + stats.rejectedCount + " rejected)\n" +
    "Interview Pipeline: " + stats.interviewPipeline + "\n" +
    "Follow-Ups Due: " + stats.followUpsDue
  );
}

function getCareerFollowUpCount() {
  return getCareerFollowUpItems_().length;
}

function getOldestAppliedApplications(limit) {
  return getApplicationDataRows_()
    .filter(item => item.status === "applied" && item.daysSinceApplied !== null)
    .sort((a, b) => b.daysSinceApplied - a.daysSinceApplied)
    .slice(0, limit || 3);
}

function getApplicationDataRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APPLICATIONS_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const headerMap = buildApplicationsHeaderMap_(values[0]);
  const today = normalizeApplicationDate_(new Date());

  return values
    .slice(1)
    .filter(row => row.some(value => value !== "" && value !== null))
    .map(row => buildApplicationDataItem_(row, headerMap, today));
}

function getCareerFollowUpItems_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APPLICATIONS_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const headerMap = buildApplicationsHeaderMap_(values[0]);
  const today = normalizeApplicationDate_(new Date());
  const followUpItems = values
    .slice(1)
    .filter(row => row.some(value => value !== "" && value !== null))
    .map(row => buildApplicationFollowUpItem_(row, headerMap, today))
    .filter(item => item !== null);

  return followUpItems;
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

function cleanupApplicationsFormatting_(sheet) {
  const rowCount = sheet.getMaxRows();

  sheet
    .getRange(1, 1, rowCount, APPLICATIONS_V3_HEADERS.length)
    .clearDataValidations();
  applyNumberFormatToUsedRows_(sheet, 5, "@");
  applyNumberFormatToUsedRows_(sheet, 7, "dd mmm yyyy");
  applyNumberFormatToUsedRows_(sheet, 9, "dd mmm yyyy");
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

function buildApplicationFollowUpItem_(row, headerMap, today) {
  const status = String(
    getApplicationValue_(row, headerMap, ["Status"]) || ""
  ).trim().toLowerCase();
  const fuRequired = String(
    getApplicationValue_(row, headerMap, ["FU Required?", "FU"]) || ""
  ).trim().toLowerCase();
  const closedStatuses = ["rejected", "withdrawn", "offer"];

  if (closedStatuses.indexOf(status) !== -1) {
    return null;
  }

  if (fuRequired !== "" && fuRequired !== "yes") {
    return null;
  }

  const dateApplied = parseApplicationDate_(
    getApplicationValue_(row, headerMap, ["Date Applied", "Date"])
  );
  const followUpDate = parseApplicationDate_(
    getApplicationValue_(row, headerMap, ["Follow Up Date", "Date FU"])
  );
  const followUpDue = followUpDate && followUpDate.getTime() <= today.getTime();
  const daysSinceApplied = dateApplied
    ? Math.floor((today.getTime() - dateApplied.getTime()) / 86400000)
    : null;
  const staleApplication = dateApplied &&
    daysSinceApplied !== null &&
    daysSinceApplied > 7;

  if (!followUpDue && !staleApplication) {
    return null;
  }

  return {
    companyName: getApplicationValue_(row, headerMap, ["Company Name", "Company"]) || "-",
    jobTitle: getApplicationValue_(row, headerMap, ["Job Title", "Position"]) || "-",
    daysSinceApplied: daysSinceApplied,
    notes: getApplicationValue_(row, headerMap, ["Notes"]) || "No notes provided."
  };
}

function buildApplicationDataItem_(row, headerMap, today) {
  const dateApplied = parseApplicationDate_(
    getApplicationValue_(row, headerMap, ["Date Applied", "Date"])
  );
  const daysSinceApplied = dateApplied
    ? Math.floor((today.getTime() - dateApplied.getTime()) / 86400000)
    : null;

  return {
    companyName: getApplicationValue_(row, headerMap, ["Company Name", "Company"]) || "-",
    jobTitle: getApplicationValue_(row, headerMap, ["Job Title", "Position"]) || "-",
    status: String(getApplicationValue_(row, headerMap, ["Status"]) || "").trim().toLowerCase(),
    dateApplied: dateApplied,
    daysSinceApplied: daysSinceApplied,
    notes: getApplicationValue_(row, headerMap, ["Notes"]) || ""
  };
}

function formatApplicationFollowUpMemo_(item, index) {
  const daysSinceApplied = item.daysSinceApplied === null
    ? "Unknown"
    : item.daysSinceApplied + " days";

  return (
    "*Memo " + (index + 1) + "*\n" +
    "*Company Name:* " + escapeTelegramMarkdown(item.companyName) + "\n" +
    "*Job Title:* " + escapeTelegramMarkdown(item.jobTitle) + "\n" +
    "*Days since Applied:* " + escapeTelegramMarkdown(daysSinceApplied) + "\n" +
    "*Context:* " + escapeTelegramMarkdown(item.notes)
  );
}

function parseApplicationDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return normalizeApplicationDate_(value);
  }

  if (value === "" || value === null || value === undefined) {
    return null;
  }

  const parsedDate = new Date(value);

  if (isNaN(parsedDate.getTime())) {
    return null;
  }

  return normalizeApplicationDate_(parsedDate);
}

function normalizeApplicationDate_(date) {
  const normalizedDate = new Date(date);
  normalizedDate.setHours(0, 0, 0, 0);
  return normalizedDate;
}

function formatApplicationDate_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "dd MMM yyyy"
  );
}

function calculateApplicationRate_(count, total) {
  if (!total) return "0.0";

  return ((count / total) * 100).toFixed(1);
}
