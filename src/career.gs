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
  "Link Job posting",
  "CV VERSION",
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
const APPLICATIONS_COL = {
  COMPANY_NAME: 0,
  JOB_TITLE: 1,
  DATE_APPLIED: 3,
  STATUS: 4,
  FU_REQUIRED: 7,
  FOLLOW_UP_DATE: 8,
  LINK_JOB_POSTING: 9,
  CV_VERSION: 10,
  NOTES: 11
};

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
  const today = getTodayWibDate_();

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

function analyzeCareerStrategy() {
  const rows = getApplicationDataRows_();

  if (rows.length === 0) {
    return (
      "📊 *AI Career Advisor*\n" +
      "No application data found in `Applications` yet."
    );
  }

  const diversification = analyzeCareerDiversification_(rows);
  const cvPerformance = analyzeCareerCvPerformance_(rows);
  const coldLeads = getCareerColdLeads_(rows, 3);
  const nextMove = buildCareerNextMove_(diversification, cvPerformance, coldLeads);

  return (
    "📊 *AI Career Advisor Memo*\n" +
    "Applications reviewed: " + rows.length + "\n\n" +
    "*📊 Diversification Audit:*\n" +
    buildCareerDiversificationMemo_(diversification) +
    "\n\n*💡 CV Insight:*\n" +
    buildCareerCvInsightMemo_(cvPerformance) +
    "\n\n*⏳ Stagnancy Alert:*\n" +
    buildCareerColdLeadMemo_(coldLeads) +
    "\n\n*🎯 Next Move:*\n" +
    nextMove
  );
}

function addApplicationFromTelegram(inputString) {
  const parsedInput = parseCareerAddInput_(inputString);

  if (!parsedInput.isValid) {
    return buildCareerAddUsageMessage_(parsedInput.error);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(APPLICATIONS_SHEET_NAME);

  if (!sheet) {
    standardizeApplicationsSheet();
    sheet = ss.getSheetByName(APPLICATIONS_SHEET_NAME);
  }

  ensureApplicationsMinimumRows_(sheet);
  ensureApplicationsHeaders_(sheet);

  const nextRow = findNextApplicationsRow_(sheet);
  ensureApplicationsRowCapacity_(sheet, nextRow);
  const today = getTodayWibDate_();
  const rowValues = [
    parsedInput.company,
    parsedInput.role,
    "",
    today,
    parsedInput.status,
    "",
    "",
    "",
    "",
    "",
    "",
    parsedInput.notes
  ];

  sheet
    .getRange(nextRow, 1, 1, APPLICATIONS_V3_HEADERS.length)
    .setValues([rowValues]);
  applyApplicationsDataValidation_(sheet);
  applyNumberFormatToUsedRows_(sheet, 4, "dd mmm yyyy");
  applyNumberFormatToUsedRows_(sheet, 5, "@");
  applyNumberFormatToUsedRows_(sheet, 7, "dd mmm yyyy");
  applyNumberFormatToUsedRows_(sheet, 9, "dd mmm yyyy");
  applyNumberFormatToUsedRows_(sheet, 10, "@");
  applyNumberFormatToUsedRows_(sheet, 11, "@");
  applyNumberFormatToUsedRows_(sheet, 12, "@");
  writeSystemLog(
    "Career",
    "Telegram Add Application",
    "Success",
    "Added application from Telegram: " + parsedInput.company + " - " + parsedInput.role
  );

  return "✅ Successfully added " + escapeTelegramMarkdown(parsedInput.company) + " to your Career Pipeline!";
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
      getApplicationValue_(row, headerMap, ["Link Job posting", "Link"]),
      getApplicationValue_(row, headerMap, ["CV VERSION", "CV", "CV Version"]),
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

function getApplicationField_(row, headerMap, candidateHeaders, fallbackIndex) {
  const headerValue = getApplicationValue_(row, headerMap, candidateHeaders);

  if (headerValue !== "" && headerValue !== null && headerValue !== undefined) {
    return headerValue;
  }

  if (fallbackIndex !== undefined && fallbackIndex < row.length) {
    return row[fallbackIndex];
  }

  return "";
}

function buildApplicationsNotes_(row, headerMap) {
  const noteParts = [];
  const legacyFields = ["Location", "Review"];
  const hasLegacyNoteSources = legacyFields.some(
    fieldName => headerMap[normalizeApplicationsHeader_(fieldName)] !== undefined
  );

  if (!hasLegacyNoteSources) {
    return getApplicationField_(row, headerMap, ["Notes"], APPLICATIONS_COL.NOTES);
  }

  const fieldsToCombine = ["Location", "Notes", "Review"];

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
  applyNumberFormatToUsedRows_(sheet, 4, "dd mmm yyyy");
  applyNumberFormatToUsedRows_(sheet, 5, "@");
  applyNumberFormatToUsedRows_(sheet, 7, "dd mmm yyyy");
  applyNumberFormatToUsedRows_(sheet, 9, "dd mmm yyyy");
  applyNumberFormatToUsedRows_(sheet, 10, "@");
  applyNumberFormatToUsedRows_(sheet, 11, "@");
  applyNumberFormatToUsedRows_(sheet, 12, "@");
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

function ensureApplicationsHeaders_(sheet) {
  const existingHeaders = sheet
    .getRange(1, 1, 1, APPLICATIONS_V3_HEADERS.length)
    .getValues()[0];
  const headersAreValid = APPLICATIONS_V3_HEADERS.every((header, index) =>
    existingHeaders[index] === header
  );

  if (!headersAreValid) {
    sheet
      .getRange(1, 1, 1, APPLICATIONS_V3_HEADERS.length)
      .setValues([APPLICATIONS_V3_HEADERS]);
  }
}

function findNextApplicationsRow_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);

  if (lastRow < 2) {
    return 2;
  }

  const rowValues = sheet
    .getRange(2, 1, lastRow - 1, APPLICATIONS_V3_HEADERS.length)
    .getValues();

  for (let i = 0; i < rowValues.length; i++) {
    if (rowValues[i].every(value => value === "" || value === null)) {
      return i + 2;
    }
  }

  return lastRow + 1;
}

function ensureApplicationsRowCapacity_(sheet, targetRow) {
  const maxRows = sheet.getMaxRows();

  if (targetRow > maxRows) {
    sheet.insertRowsAfter(maxRows, targetRow - maxRows);
  }
}

function parseCareerAddInput_(inputString) {
  const parts = String(inputString || "")
    .split("|")
    .map(part => part.trim());

  if (parts.length < 4) {
    return {
      isValid: false,
      error: "Format must be: Company | Role | Status | Notes"
    };
  }

  const company = parts[0];
  const role = parts[1];
  const status = normalizeCareerStatus_(parts[2]);
  const notes = parts.slice(3).join(" | ").trim();

  if (!company || !role || !parts[2].trim() || !notes) {
    return {
      isValid: false,
      error: "Company, Role, Status, and Notes are all required."
    };
  }

  if (!status) {
    return {
      isValid: false,
      error: "Status must be one of: " + APPLICATIONS_STATUS_OPTIONS.join(", ")
    };
  }

  return {
    isValid: true,
    company: company,
    role: role,
    status: status,
    notes: notes
  };
}

function normalizeCareerStatus_(status) {
  const normalizedStatus = String(status || "").trim().toLowerCase();

  for (let i = 0; i < APPLICATIONS_STATUS_OPTIONS.length; i++) {
    if (APPLICATIONS_STATUS_OPTIONS[i].toLowerCase() === normalizedStatus) {
      return APPLICATIONS_STATUS_OPTIONS[i];
    }
  }

  return "";
}

function buildCareerAddUsageMessage_(error) {
  return (
    "⚠️ Unable to add application.\n" +
    escapeTelegramMarkdown(error) + "\n\n" +
    "Usage:\n" +
    "`/careeradd Google | Product Manager | Applied | Referral from John`"
  );
}

function normalizeApplicationsHeader_(header) {
  return String(header || "").trim().toLowerCase();
}

function buildApplicationFollowUpItem_(row, headerMap, today) {
  const status = String(
    getApplicationField_(row, headerMap, ["Status"], APPLICATIONS_COL.STATUS) || ""
  ).trim().toLowerCase();
  const fuRequired = String(
    getApplicationField_(row, headerMap, ["FU Required?", "FU"], APPLICATIONS_COL.FU_REQUIRED) || ""
  ).trim().toLowerCase();
  const closedStatuses = ["rejected", "withdrawn", "offer"];

  if (closedStatuses.indexOf(status) !== -1) {
    return null;
  }

  if (fuRequired !== "" && fuRequired !== "yes") {
    return null;
  }

  const dateApplied = parseApplicationDate_(
    getApplicationField_(row, headerMap, ["Date Applied", "Date"], APPLICATIONS_COL.DATE_APPLIED)
  );
  const followUpDate = parseApplicationDate_(
    getApplicationField_(row, headerMap, ["Follow Up Date", "Date FU"], APPLICATIONS_COL.FOLLOW_UP_DATE)
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
    companyName: getApplicationField_(row, headerMap, ["Company Name", "Company"], APPLICATIONS_COL.COMPANY_NAME) || "-",
    jobTitle: getApplicationField_(row, headerMap, ["Job Title", "Position"], APPLICATIONS_COL.JOB_TITLE) || "-",
    jobLink: getApplicationField_(row, headerMap, ["Link Job posting", "Link"], APPLICATIONS_COL.LINK_JOB_POSTING) || "",
    daysSinceApplied: daysSinceApplied,
    notes: getApplicationField_(row, headerMap, ["Notes"], APPLICATIONS_COL.NOTES) || "No notes provided."
  };
}

function buildApplicationDataItem_(row, headerMap, today) {
  const dateApplied = parseApplicationDate_(
    getApplicationField_(row, headerMap, ["Date Applied", "Date"], APPLICATIONS_COL.DATE_APPLIED)
  );
  const daysSinceApplied = dateApplied
    ? Math.floor((today.getTime() - dateApplied.getTime()) / 86400000)
    : null;

  return {
    companyName: getApplicationField_(row, headerMap, ["Company Name", "Company"], APPLICATIONS_COL.COMPANY_NAME) || "-",
    jobTitle: getApplicationField_(row, headerMap, ["Job Title", "Position"], APPLICATIONS_COL.JOB_TITLE) || "-",
    status: String(getApplicationField_(row, headerMap, ["Status"], APPLICATIONS_COL.STATUS) || "").trim().toLowerCase(),
    dateApplied: dateApplied,
    daysSinceApplied: daysSinceApplied,
    cvVersion: getApplicationField_(row, headerMap, ["CV VERSION", "CV", "CV Version"], APPLICATIONS_COL.CV_VERSION) || "",
    jobLink: getApplicationField_(row, headerMap, ["Link Job posting", "Link"], APPLICATIONS_COL.LINK_JOB_POSTING) || "",
    notes: getApplicationField_(row, headerMap, ["Notes"], APPLICATIONS_COL.NOTES) || ""
  };
}

function analyzeCareerDiversification_(rows) {
  const sectorCounts = {};

  rows.forEach(row => {
    const sector = classifyCareerSector_(row.companyName);
    sectorCounts[sector] = (sectorCounts[sector] || 0) + 1;
  });

  const sectors = Object.keys(sectorCounts)
    .map(sector => ({
      sector: sector,
      count: sectorCounts[sector],
      percentage: Math.round((sectorCounts[sector] / rows.length) * 100)
    }))
    .sort((a, b) => b.count - a.count || a.sector.localeCompare(b.sector));

  return {
    total: rows.length,
    sectors: sectors,
    dominantSector: sectors.length > 0 ? sectors[0] : null,
    isConcentrated: sectors.length > 0 && sectors[0].percentage > 50
  };
}

function classifyCareerSector_(companyName) {
  const company = String(companyName || "").toLowerCase();

  if (/(bank|mandiri|ocbc|bca|bri|bni|btn|cimb|danamon|permata|maybank|uob|hsbc)/i.test(company)) {
    return "Banking";
  }

  if (/(astra|auto2000|toyota|daihatsu|isuzu|honda|mitsubishi|adira)/i.test(company)) {
    return "Automotive/Conglomerate";
  }

  if (/(nielsen|abeam|accenture|deloitte|pwc|kpmg|ey|mckinsey|bcg|bain|data|analytics)/i.test(company)) {
    return "Consulting/Data";
  }

  if (/(unilever|nestle|indofood|wings|mayora|danone|fmcg|p&g|procter)/i.test(company)) {
    return "FMCG";
  }

  if (/(telkom|gojek|tokopedia|shopee|grab|traveloka|bukalapak|tech|digital)/i.test(company)) {
    return "Technology/Digital";
  }

  return "Other";
}

function analyzeCareerCvPerformance_(rows) {
  const tagStats = {};

  rows.forEach(row => {
    const tags = extractCareerCvTags_(row.cvVersion);
    const converted = isCareerConvertedStatus_(row.status);

    tags.forEach(tag => {
      if (!tagStats[tag]) {
        tagStats[tag] = {
          tag: tag,
          total: 0,
          converted: 0
        };
      }

      tagStats[tag].total++;

      if (converted) {
        tagStats[tag].converted++;
      }
    });
  });

  const tags = Object.keys(tagStats)
    .map(tag => ({
      tag: tag,
      total: tagStats[tag].total,
      converted: tagStats[tag].converted,
      rate: tagStats[tag].total === 0
        ? 0
        : Math.round((tagStats[tag].converted / tagStats[tag].total) * 100)
    }))
    .sort((a, b) => b.rate - a.rate || b.total - a.total || a.tag.localeCompare(b.tag));

  return {
    tags: tags,
    bestTag: tags.length > 0 ? tags[0] : null,
    baselineTag: tags.filter(item => item.tag === "General")[0] || null
  };
}

function extractCareerCvTags_(cvVersion) {
  const cvText = String(cvVersion || "");
  const tagPatterns = [
    { tag: "MT", pattern: /\bmt\b|management trainee/i },
    { tag: "ODP", pattern: /\bodp\b|officer development/i },
    { tag: "Data", pattern: /\bdata\b/i },
    { tag: "Automation", pattern: /automation/i },
    { tag: "Analyst", pattern: /analyst/i },
    { tag: "Business", pattern: /business/i }
  ];
  const tags = tagPatterns
    .filter(item => item.pattern.test(cvText))
    .map(item => item.tag);

  return tags.length > 0 ? tags : ["General"];
}

function isCareerConvertedStatus_(status) {
  const normalizedStatus = String(status || "").trim().toLowerCase();

  return normalizedStatus === "assessment" || normalizedStatus === "interview";
}

function getCareerColdLeads_(rows, limit) {
  return rows
    .filter(row =>
      row.status === "applied" &&
      row.daysSinceApplied !== null &&
      row.daysSinceApplied >= 7
    )
    .map(row => ({
      companyName: row.companyName,
      jobTitle: row.jobTitle,
      daysSinceApplied: row.daysSinceApplied,
      zone: row.daysSinceApplied > 14 ? "Red Zone" : "Yellow Zone"
    }))
    .sort((a, b) => b.daysSinceApplied - a.daysSinceApplied)
    .slice(0, limit || 3);
}

function buildCareerDiversificationMemo_(analysis) {
  if (!analysis.dominantSector) {
    return "No sector pattern detected yet.";
  }

  const topSectors = analysis.sectors
    .slice(0, 3)
    .map(item =>
      "- " + escapeTelegramMarkdown(item.sector) + ": " + item.count + " apps (" + item.percentage + "%)"
    )
    .join("\n");
  const concentrationText = analysis.isConcentrated
    ? "\nPortfolio is concentrated in " + escapeTelegramMarkdown(analysis.dominantSector.sector) + " (>50%)."
    : "\nPortfolio spread is balanced enough for now.";

  return topSectors + concentrationText;
}

function buildCareerCvInsightMemo_(analysis) {
  if (!analysis.bestTag) {
    return "No CV version signal detected yet. Start labeling CV Version with tags like Data, Automation, Analyst, MT, or ODP.";
  }

  const bestTagText = (
    escapeTelegramMarkdown(analysis.bestTag.tag) +
    " is currently strongest: " +
    analysis.bestTag.converted + "/" +
    analysis.bestTag.total +
    " converted (" +
    analysis.bestTag.rate +
    "%)."
  );

  if (analysis.baselineTag && analysis.bestTag.tag !== "General") {
    const delta = analysis.bestTag.rate - analysis.baselineTag.rate;

    if (delta > 0) {
      return bestTagText + "\nThis is +" + delta + " pts above General CVs.";
    }
  }

  return bestTagText;
}

function buildCareerColdLeadMemo_(coldLeads) {
  if (coldLeads.length === 0) {
    return "No cold leads older than 7 days in Applied status.";
  }

  return coldLeads.map(item =>
    "- " +
    escapeTelegramMarkdown(item.companyName) +
    " | " +
    escapeTelegramMarkdown(item.jobTitle) +
    " | " +
    item.daysSinceApplied +
    " days | " +
    item.zone
  ).join("\n");
}

function buildCareerNextMove_(diversification, cvPerformance, coldLeads) {
  if (diversification.isConcentrated && diversification.dominantSector) {
    return (
      "You have " +
      diversification.dominantSector.count +
      " " +
      escapeTelegramMarkdown(diversification.dominantSector.sector) +
      " applications. Add 2 wildcard applications outside that sector this week to spread risk."
    );
  }

  if (cvPerformance.bestTag && cvPerformance.bestTag.tag !== "General") {
    return (
      "Double down on " +
      escapeTelegramMarkdown(cvPerformance.bestTag.tag) +
      "-positioned roles for the next batch, while keeping 1-2 wildcard companies in a new sector."
    );
  }

  if (coldLeads.length > 0) {
    return "Follow up the oldest cold lead, then archive any Red Zone application with no strategic upside.";
  }

  return "Tag the next applications by CV Version first, then apply to 2 roles in a sector you have not tested yet.";
}

function formatApplicationFollowUpMemo_(item, index) {
  const daysSinceApplied = item.daysSinceApplied === null
    ? "Unknown"
    : item.daysSinceApplied + " days";
  const jobLinkLine = item.jobLink
    ? "*🔗 Job Link:* " + escapeTelegramMarkdown(item.jobLink) + "\n"
    : "";

  return (
    "*Memo " + (index + 1) + "*\n" +
    "*Company Name:* " + escapeTelegramMarkdown(item.companyName) + "\n" +
    "*Job Title:* " + escapeTelegramMarkdown(item.jobTitle) + "\n" +
    jobLinkLine +
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

function getTodayWibDate_() {
  return normalizeApplicationDate_(
    new Date(Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy/MM/dd"))
  );
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
