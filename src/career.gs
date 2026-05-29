// Career module: Application tracking DSS standardization.
const SYSTEM_STATE = {
  version: "V3.4",
  phase: "Agentic AI & Intelligence"
};
const VERSION_HISTORY = {
  "V3.0-V3.4": "Autonomous Career Coaching, CV Performance Gap Analysis (+46 pts), Automated Industry Mapping, and Executive Follow-up Memos."
};
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
  CATEGORY: 2,
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
  const analysis = buildCareerStrategyAnalysis_();
  const rows = analysis.rows;

  if (rows.length === 0) {
    return (
      "📊 *AI Career Advisor*\n" +
      "No application data found in `Applications` yet."
    );
  }

  return (
    "📊 *AI Career Advisor Memo*\n" +
    "Applications reviewed: " + rows.length + "\n\n" +
    "*🧭 Funnel Reality Check:*\n" +
    buildCareerFunnelMemo_(analysis.funnel) +
    "\n\n" +
    "*📊 Diversification Audit:*\n" +
    buildCareerDiversificationMemo_(analysis.diversification) +
    "\n\n*💡 CV & Positioning Insight:*\n" +
    buildCareerCvInsightMemo_(analysis.cvPerformance) +
    "\n\n*⏳ Action Queue:*\n" +
    buildCareerColdLeadMemo_(analysis.coldLeads) +
    "\n\n*🎯 Next Move:*\n" +
    analysis.nextMove
  );
}

function generateCareerMemo() {
  return analyzeCareerStrategy();
}

function generateLinkedInCampaign() {
  const analysis = buildCareerStrategyAnalysis_();
  const totalApplications = analysis.rows.length;
  const bestTag = analysis.cvPerformance.bestTag;
  const baselineTag = analysis.cvPerformance.baselineTag;
  const cvGap = bestTag && baselineTag
    ? bestTag.rate - baselineTag.rate
    : null;
  const bestCvLine = bestTag
    ? bestTag.tag + " CVs are leading with " + bestTag.rate + "% response signal across " + bestTag.total + " applications."
    : "CV performance data is still being built from the Applications sheet.";
  const cvGapLine = cvGap !== null && cvGap > 0
    ? "The strongest CV positioning is +" + cvGap + " pts above General CVs."
    : "The system is already structured to measure CV performance gaps as the dataset grows.";
  const applicationScale = totalApplications >= 27
    ? totalApplications + "+ apps"
    : totalApplications + " apps";

  return {
    hook: [
      "I built a personal infrastructure layer for career execution.",
      "V3 moved the system from manual tracking into an Agentic AI workflow: daily reliability, mobile data entry, and strategic coaching.",
      "This is where Management, Marketing, and AI Engineering meet in one operating system."
    ],
    carouselBlueprint: [
      "Slide 1: From V2 manual tracking to V3 Agentic execution - I stopped treating applications as scattered tasks and started treating them as an operating system.",
      "Slide 2: The Morning Routine (V3.0-V3.1) - Automated Executive Memos and Daily Briefs audit my pipeline every morning at 08:00 WIB for Reliability & Consistency.",
      "Slide 3: Zero-Friction Entry (V3.2) - Instant Telegram Data Entry lets me add job applications in about 5 seconds from mobile, preserving momentum instead of creating admin drag.",
      "Slide 4: The Brain (V3.3-V3.4) - /careercoach maps industries in real time and compares CV positioning. " + bestCvLine + " " + cvGapLine,
      "Slide 5: The Impact - Managing " + applicationScale + " with less mental fatigue because follow-ups, stagnancy, and strategy signals are surfaced automatically.",
      "Slide 6: Conclusion - I am ready to bring this Operational Intelligence mindset to a company: reliable execution, marketing-aware positioning, and practical AI Engineering."
    ],
    videoIdea: [
      "Start with the problem: manual career tracking creates follow-up debt, weak signal, and mental fatigue.",
      "Show the V3 foundation: Daily Briefs and Executive Memos supporting reliability at 08:00 WIB.",
      "Open Telegram and add an application with /careeradd to demonstrate the V3.2 mobile workflow.",
      "Run /careercoach to show the V3.3-V3.4 intelligence layer: industry mapping, CV performance gap, cold leads, and next move.",
      "Close with the message: I built personal infrastructure to turn execution, positioning, and analytics into one management system."
    ],
    githubReadmeDraft: generateGitHubREADME(),
    metrics: {
      totalApplications: totalApplications,
      bestCvTag: bestTag ? bestTag.tag : "Not enough data",
      bestCvRate: bestTag ? bestTag.rate : 0,
      cvPerformanceGap: cvGap,
      version: SYSTEM_STATE.version,
      phase: SYSTEM_STATE.phase,
      allowedFeatureScope: VERSION_HISTORY["V3.0-V3.4"],
      spreadsheetUrl: "Private spreadsheet - not included in public demo"
    }
  };
}

function prepLookerExport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lookerSheet = ss.getSheetByName("Looker_Data");
  const rows = getApplicationDataRows_();
  const exportRows = rows.map(row => {
    const statusLevel = getCareerStatusLevel_(row.status);

    return [
      row.companyName,
      classifyCareerSector_(row),
      extractCareerCvTags_(row).join(", "),
      statusLevel === null ? "" : statusLevel,
      row.daysSinceApplied === null ? "" : row.daysSinceApplied
    ];
  });
  const values = [[
    "Company",
    "Industry",
    "CV_Tag",
    "Status_Level",
    "Days_Since_Applied"
  ]].concat(exportRows);

  if (!lookerSheet) {
    lookerSheet = ss.insertSheet("Looker_Data");
  }

  lookerSheet.clearContents();
  lookerSheet
    .getRange(1, 1, values.length, values[0].length)
    .setValues(values);
  lookerSheet.hideSheet();

  writeSystemLog(
    "Career",
    "Prep Looker Export",
    "Success",
    "Looker_Data refreshed with " + exportRows.length + " rows."
  );

  return {
    sheetName: "Looker_Data",
    rowCount: exportRows.length,
    spreadsheetUrl: "Private spreadsheet - not included in public demo"
  };
}

function generateGitHubREADME() {
  const analysis = buildCareerStrategyAnalysis_();
  const bestTag = analysis.cvPerformance.bestTag;
  const totalApplications = analysis.rows.length;
  const cvGapText = bestTag
    ? "The strongest CV tag is " + bestTag.tag + " with a " + bestTag.rate + "% response signal."
    : "The CV performance layer is ready to surface response gaps once more data is available.";

  return [
    "# Portfolio User Productivity OS - V3",
    "",
    "## Project Overview",
    "",
    "This project is a personal infrastructure system for career execution, built on Google Apps Script, Google Sheets, Telegram, and Gemini API workflows. V3 transforms application tracking from a passive spreadsheet into an agentic operating layer for reliability, mobile execution, and strategic decision support.",
    "",
    "## Evolution from V2 to V3",
    "",
    "- V2 focused on manual tracking and operational visibility.",
    "- V3.0-V3.1 introduced Automated Executive Memos and Daily Briefs for reliability and consistency.",
    "- V3.2 added instant Telegram data entry for zero-friction mobile workflow.",
    "- V3.3-V3.4 added AI Career Advisor logic, CV performance analysis, industry mapping, evidence packaging, and Looker-ready exports.",
    "",
    "## Key Tech Stack (Gemini API, Apps Script)",
    "",
    "- Google Apps Script for automation, spreadsheet orchestration, triggers, and Telegram webhook handling.",
    "- Google Sheets as the operational database for applications, career metrics, and Looker Studio export data.",
    "- Telegram Bot API for fast mobile data entry, executive memos, and command-based decision support.",
    "- Gemini API for AI-assisted workflow intelligence and evidence generation.",
    "- Looker Studio-ready `Looker_Data` sheet for dashboarding and visual proof.",
    "",
    "## Real-world Impact (+46 pts CV Gap)",
    "",
    "The system manages " + totalApplications + " applications with reduced mental load by surfacing follow-ups, stagnancy, industry concentration, and CV performance signals. " + cvGapText + " The V3.0-V3.4 arc demonstrates Operational Excellence, Management & Marketing thinking, and practical AI Engineering in one working product.",
    "",
    "## System Architecture",
    "",
    "Telegram commands write and read structured data from Google Sheets. Apps Script modules transform the data into executive memos, career coaching, Looker Studio export tables, and LinkedIn evidence packages. The architecture keeps the workflow lightweight while preserving enough structure for analytics, storytelling, and technical proof."
  ].join("\n");
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
  const legacyContextMap = getLegacyApplicationContextMap_(ss, today);

  return values
    .slice(1)
    .filter(row => row.some(value => value !== "" && value !== null))
    .map(row => {
      const item = buildApplicationDataItem_(row, headerMap, today);
      const legacyContext = legacyContextMap[buildApplicationIdentityKey_(item.companyName, item.jobTitle)];

      if (legacyContext) {
        item.legacyContext = legacyContext;
      }

      return item;
    });
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

function getLegacyApplicationContextMap_(ss, today) {
  const legacySheet = ss.getSheetByName("Copy of Applications");
  const contextMap = {};

  if (!legacySheet || legacySheet.getLastRow() < 2) {
    return contextMap;
  }

  const values = legacySheet
    .getRange(1, 1, legacySheet.getLastRow(), legacySheet.getLastColumn())
    .getValues();
  const headerMap = buildApplicationsHeaderMap_(values[0]);

  values
    .slice(1)
    .filter(row => row.some(value => value !== "" && value !== null))
    .forEach(row => {
      const companyName = getApplicationField_(row, headerMap, ["Company Name", "Company"], 0) || "";
      const jobTitle = getApplicationField_(row, headerMap, ["Job Title", "Position"], 1) || "";

      if (!companyName && !jobTitle) {
        return;
      }

      const dateApplied = parseApplicationDate_(
        getApplicationField_(row, headerMap, ["Date Applied", "Date"], 3)
      );
      const nextStepDate = parseApplicationDate_(
        getApplicationField_(row, headerMap, ["Next Step Date", "Target Date"], 9)
      );
      const followUpDate = parseApplicationDate_(
        getApplicationField_(row, headerMap, ["Follow Up Date", "Date FU"], 18)
      );

      contextMap[buildApplicationIdentityKey_(companyName, jobTitle)] = {
        salaryRange: getApplicationField_(row, headerMap, ["Salary Range"], 4) || "",
        location: getApplicationField_(row, headerMap, ["Location"], 5) || "",
        workMode: getApplicationField_(row, headerMap, ["Work Mode"], 6) || "",
        source: getApplicationField_(row, headerMap, ["Source"], 13) || "",
        nextStep: getApplicationField_(row, headerMap, ["Next Step", "Next Action"], 8) || "",
        nextStepDate: nextStepDate,
        followUpRequired: getApplicationField_(row, headerMap, ["Follow Up Required?", "FU Required?", "FU"], 17) || "",
        followUpDate: followUpDate,
        daysSinceApplied: dateApplied
          ? Math.floor((today.getTime() - dateApplied.getTime()) / 86400000)
          : null
      };
    });

  return contextMap;
}

function buildApplicationIdentityKey_(companyName, jobTitle) {
  return normalizeApplicationsHeader_(companyName) + "|" + normalizeApplicationsHeader_(jobTitle);
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
    category: getApplicationField_(row, headerMap, ["Category", "Industry"], APPLICATIONS_COL.CATEGORY) || "",
    status: String(getApplicationField_(row, headerMap, ["Status"], APPLICATIONS_COL.STATUS) || "").trim().toLowerCase(),
    dateApplied: dateApplied,
    daysSinceApplied: daysSinceApplied,
    cvVersion: getApplicationField_(row, headerMap, ["CV VERSION", "CV", "CV Version"], APPLICATIONS_COL.CV_VERSION) || "",
    jobLink: getApplicationField_(row, headerMap, ["Link Job posting", "Link"], APPLICATIONS_COL.LINK_JOB_POSTING) || "",
    notes: getApplicationField_(row, headerMap, ["Notes"], APPLICATIONS_COL.NOTES) || ""
  };
}

function buildCareerStrategyAnalysis_() {
  const rows = getApplicationDataRows_();
  const funnel = analyzeCareerFunnel_(rows);
  const diversification = analyzeCareerDiversification_(rows);
  const cvPerformance = analyzeCareerCvPerformance_(rows);
  const coldLeads = getCareerColdLeads_(rows, 5);

  return {
    rows: rows,
    funnel: funnel,
    diversification: diversification,
    cvPerformance: cvPerformance,
    coldLeads: coldLeads,
    nextMove: buildCareerNextMove_(diversification, cvPerformance, coldLeads, funnel)
  };
}

function analyzeCareerFunnel_(rows) {
  const funnel = {
    total: rows.length,
    saved: 0,
    rejected: 0,
    activeApplied: 0,
    assessmentStage: 0,
    interviewStage: 0,
    eligible: 0
  };

  rows.forEach(row => {
    const normalizedStatus = normalizeCareerStatusLabel_(row.status);
    const statusLevel = getCareerStatusLevel_(row.status);

    if (normalizedStatus === "saved") {
      funnel.saved++;
      return;
    }

    if (normalizedStatus === "rejected") {
      funnel.rejected++;
      return;
    }

    if (statusLevel === null) {
      return;
    }

    funnel.eligible++;

    if (statusLevel === 0) {
      funnel.activeApplied++;
    }

    if (statusLevel >= 1) {
      funnel.assessmentStage++;
    }

    if (statusLevel === 2) {
      funnel.interviewStage++;
    }
  });

  funnel.assessmentRate = calculateApplicationRate_(funnel.assessmentStage, funnel.eligible);
  funnel.interviewRate = calculateApplicationRate_(funnel.interviewStage, funnel.eligible);

  return funnel;
}

function analyzeCareerDiversification_(rows) {
  const sectorCounts = {};

  rows.forEach(row => {
    const sector = classifyCareerSector_(row);
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

function classifyCareerSector_(application) {
  const company = String(application && application.companyName ? application.companyName : "");
  const jobTitle = String(application && application.jobTitle ? application.jobTitle : "");
  const notes = String(application && application.notes ? application.notes : "");
  const category = String(application && application.category ? application.category : "");
  const cvVersion = String(application && application.cvVersion ? application.cvVersion : "");
  const categoryText = category.trim();
  const categoryLower = categoryText.toLowerCase();
  const companyText = company.toLowerCase();
  const cvText = cvVersion.toLowerCase();
  const contextText = (company + " " + jobTitle + " " + notes + " " + cvVersion).toLowerCase();

  if (categoryText !== "" && categoryLower !== "other") {
    return normalizeCareerIndustryName_(categoryText);
  }

  if (/(bank|mandiri|ocbc|idx|indonesia\s*stock\s*exchange|stock\s*exchange|fif\s*group|bca|bri|bni|btn|cimb|danamon|permata|maybank|uob|hsbc)/i.test(companyText) ||
      /(^|[^a-z])(odp|bdp|mdp)([^a-z]|$)/i.test(cvText)) {
    return "Banking/Finance";
  }

  if (/(astra|auto2000|toyota|daihatsu|isuzu|honda|mitsubishi|adira|astra\s*credit|\bacc\b)/i.test(companyText)) {
    return "Automotive/Conglomerate";
  }

  if (/(nielsen|abeam|adecoo|accenture|deloitte|pwc|kpmg|ey|mckinsey|bcg|bain|data|analytics)/i.test(companyText)) {
    return "Consulting/Data";
  }

  if (/(ccep|coca[\s-]*cola|map|philip\s*morris|philip\s*moris|unilever|ulfp|nestle|indofood|wings|mayora|danone|sampoerna|korea\s*tomorrow|tobacco|fmcg|p&g|procter)/i.test(companyText) ||
      /(^|[^a-z])(ulfp|ccep)([^a-z]|$)/i.test(cvText)) {
    return "FMCG/Retail";
  }

  if (/garena/i.test(companyText)) {
    return "Tech/Gaming";
  }

  if (/(huawei|huawe|jakarta\s*digital\s*nusantara|siemens|telkom|gojek|tokopedia|shopee|grab|traveloka|bukalapak|tech|digital)/i.test(companyText)) {
    return "Tech/Digital";
  }

  if (/(transjakarta|deliveree|logistics|transport|shipping|supply\s*chain|warehouse)/i.test(companyText)) {
    return "Logistics/Transport";
  }

  if (/(tower\s*bersama|\btbg\b|telecom|telecommunication|infrastructure)/i.test(contextText)) {
    return "Telecom/Infrastructure";
  }

  if (/(eterna|manufacturing|factory|industrial|consumer\s*goods)/i.test(contextText)) {
    return "Manufacturing/Consumer Goods";
  }

  if (/(analyst|analytics|data|business\s*intelligence|bi\b)/i.test(contextText)) {
    return "Role-Inferred: Analyst/Data";
  }

  if (/(marketing|brand|growth|campaign|commercial|sales)/i.test(contextText)) {
    return "Role-Inferred: Marketing/Commercial";
  }

  if (/(engineer|engineering|developer|software|automation|technical|it\b)/i.test(contextText)) {
    return "Role-Inferred: Engineering/Tech";
  }

  return "Other";
}

function normalizeCareerIndustryName_(industry) {
  const normalizedIndustry = String(industry || "").trim();
  const lowerIndustry = normalizedIndustry.toLowerCase();

  if (/bank|finance|financial|idx|ocbc|fif/i.test(lowerIndustry)) {
    return "Banking/Finance";
  }

  if (/fmcg|retail|consumer|unilever|coca|ccep|map/i.test(lowerIndustry)) {
    return "FMCG/Retail";
  }

  if (/gaming/i.test(lowerIndustry)) {
    return "Tech/Gaming";
  }

  if (/tech|digital|software|it/i.test(lowerIndustry)) {
    return "Tech/Digital";
  }

  if (/telecom|telecommunication|tower|infrastructure|tbg/i.test(lowerIndustry)) {
    return "Telecom/Infrastructure";
  }

  if (/auto|automotive|conglomerate|astra/i.test(lowerIndustry)) {
    return "Automotive/Conglomerate";
  }

  if (/manufacturing|factory|industrial|consumer goods/i.test(lowerIndustry)) {
    return "Manufacturing/Consumer Goods";
  }

  if (/consult|data|analytics|research/i.test(lowerIndustry)) {
    return "Consulting/Data";
  }

  if (/logistic|transport|delivery|supply/i.test(lowerIndustry)) {
    return "Logistics/Transport";
  }

  return normalizedIndustry;
}

function analyzeCareerCvPerformance_(rows) {
  const tagStats = {};

  rows
    .filter(row => isCareerCvPerformanceEligible_(row.status))
    .forEach(row => {
    const tags = extractCareerCvTags_(row);
    const reachedAssessmentStage = hasReachedCareerAssessmentStage_(row.status);

    tags.forEach(tag => {
      if (!tagStats[tag]) {
        tagStats[tag] = {
          tag: tag,
          total: 0,
          assessmentStage: 0
        };
      }

      tagStats[tag].total++;

      if (reachedAssessmentStage) {
        tagStats[tag].assessmentStage++;
      }
    });
  });

  const tags = Object.keys(tagStats)
    .map(tag => ({
      tag: tag,
      total: tagStats[tag].total,
      assessmentStage: tagStats[tag].assessmentStage,
      rate: tagStats[tag].total === 0
        ? 0
        : Math.round((tagStats[tag].assessmentStage / tagStats[tag].total) * 100)
    }))
    .sort((a, b) => b.rate - a.rate || b.total - a.total || a.tag.localeCompare(b.tag));

  return {
    tags: tags,
    bestTag: tags.length > 0 ? tags[0] : null,
    baselineTag: tags.filter(item => item.tag === "General")[0] || null
  };
}

function extractCareerCvTags_(applicationOrCvVersion) {
  const isApplicationObject = applicationOrCvVersion &&
    typeof applicationOrCvVersion === "object";
  const cvVersion = isApplicationObject
    ? applicationOrCvVersion.cvVersion
    : applicationOrCvVersion;
  const cvText = String(cvVersion || "");
  const cvTextLower = cvText.trim().toLowerCase();
  const fallbackText = isApplicationObject &&
    (cvTextLower === "" || cvTextLower === "general")
    ? String(applicationOrCvVersion.jobTitle || "") + " " + String(applicationOrCvVersion.companyName || "")
    : "";
  const directTags = [];
  const fallbackTags = [];

  if (matchesCareerMtTag_(cvText)) {
    directTags.push("MT");
  }

  if (matchesCareerAutomationTag_(cvText)) {
    directTags.push("Automation");
  } else if (matchesCareerDataTag_(cvText)) {
    directTags.push("Data");
  }

  if (matchesCareerBusinessTag_(cvText) && directTags.indexOf("Data") === -1) {
    directTags.push("Business");
  }

  if (fallbackText && matchesCareerMtTag_(fallbackText)) {
    fallbackTags.push("MT");
  }

  const tags = directTags.length > 0 ? directTags : fallbackTags;

  return tags.length > 0 ? tags : ["General"];
}

function matchesCareerMtTag_(text) {
  return /(^|[^a-z])(mt|odp|ulfp|bdp|mdp|gtp)([^a-z]|$)|management[\s_-]*trainee|officer[\s_-]*development|future[\s_-]*program|global[\s_-]*trainee[\s_-]*program/i.test(String(text || ""));
}

function matchesCareerAutomationTag_(text) {
  return /automation/i.test(String(text || ""));
}

function matchesCareerDataTag_(text) {
  return /data|analyst|analytics|(^|[^a-z])bi([^a-z]|$)|business\s*analyst/i.test(String(text || ""));
}

function matchesCareerBusinessTag_(text) {
  return /business/i.test(String(text || ""));
}

function isCareerCvPerformanceEligible_(status) {
  const normalizedStatus = normalizeCareerStatusLabel_(status);

  return normalizedStatus !== "saved" &&
    normalizedStatus !== "rejected" &&
    getCareerStatusLevel_(status) !== null;
}

function hasReachedCareerAssessmentStage_(status) {
  const statusLevel = getCareerStatusLevel_(status);

  return statusLevel === 1 || statusLevel === 2;
}

function getCareerStatusLevel_(status) {
  const normalizedStatus = normalizeCareerStatusLabel_(status);

  if (normalizedStatus === "applied" || normalizedStatus === "in review") {
    return 0;
  }

  if (normalizedStatus === "assessment" || normalizedStatus === "online test") {
    return 1;
  }

  if (normalizedStatus === "interview" ||
      (/interview/.test(normalizedStatus) && /(user|hr|final)/.test(normalizedStatus))) {
    return 2;
  }

  return null;
}

function normalizeCareerStatusLabel_(status) {
  return String(status || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function getCareerColdLeads_(rows, limit) {
  return rows
    .filter(row =>
      getCareerStatusLevel_(row.status) === 0 &&
      row.daysSinceApplied !== null &&
      row.daysSinceApplied >= 7
    )
    .map(row => ({
      companyName: row.companyName,
      jobTitle: row.jobTitle,
      sector: classifyCareerSector_(row),
      cvTags: extractCareerCvTags_(row),
      daysSinceApplied: row.daysSinceApplied,
      zone: row.daysSinceApplied > 14 ? "Red Zone" : "Yellow Zone",
      legacyContext: row.legacyContext || null
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

function buildCareerFunnelMemo_(funnel) {
  const interviewText = funnel.interviewStage === 0
    ? "No interview-stage signal yet. Current movement is Assessment Stage only."
    : "Interview Stage: " + funnel.interviewStage + " (" + funnel.interviewRate + "%).";

  return (
    "Tracked: " + funnel.total + " total | " +
    funnel.eligible + " active funnel | " +
    funnel.rejected + " rejected | " +
    funnel.saved + " saved.\n" +
    "Applied/In Review: " + funnel.activeApplied + "\n" +
    "Assessment Stage: " + funnel.assessmentStage + " (" + funnel.assessmentRate + "% of active funnel)\n" +
    interviewText
  );
}

function buildCareerCvInsightMemo_(analysis) {
  if (!analysis.bestTag) {
    return "No CV version signal detected yet. Start labeling CV Version with tags like Data, Automation, Analyst, MT, or ODP.";
  }

  const topTags = analysis.tags
    .slice(0, 3)
    .map(item =>
      "- " + escapeTelegramMarkdown(item.tag) + ": " +
      item.assessmentStage + "/" + item.total +
      " reached Assessment Stage (" + item.rate + "%)"
    )
    .join("\n");
  const sampleWarning = analysis.bestTag.total < 3
    ? "\nSignal quality: promising, but sample size is still small. Validate with more applications before overcommitting."
    : "\nSignal quality: enough to guide the next application batch.";
  const bestTagText = (
    escapeTelegramMarkdown(analysis.bestTag.tag) +
    " is currently strongest for Assessment Stage: " +
    analysis.bestTag.assessmentStage + "/" +
    analysis.bestTag.total +
    " reached Assessment/Interview stage (" +
    analysis.bestTag.rate +
    "%)."
  );

  if (analysis.baselineTag && analysis.bestTag.tag !== "General") {
    const delta = analysis.bestTag.rate - analysis.baselineTag.rate;

    if (delta > 0) {
      return bestTagText + "\nThis is +" + delta + " pts above General CVs.\n\n*Top CV Signals:*\n" + topTags + sampleWarning;
    }
  }

  return bestTagText + "\n\n*Top CV Signals:*\n" + topTags + sampleWarning;
}

function buildCareerColdLeadMemo_(coldLeads) {
  if (coldLeads.length === 0) {
    return "No cold leads older than 7 days in Applied/In Review status.";
  }

  return coldLeads.map(item =>
    "- " +
    escapeTelegramMarkdown(item.companyName) +
    " | " +
    escapeTelegramMarkdown(item.jobTitle) +
    " | " +
    escapeTelegramMarkdown(item.sector) +
    " | CV: " +
    escapeTelegramMarkdown(item.cvTags.join(", ")) +
    " | " +
    item.daysSinceApplied +
    " days | " +
    item.zone +
    buildCareerColdLeadContext_(item)
  ).join("\n");
}

function buildCareerColdLeadContext_(item) {
  const context = item.legacyContext;

  if (!context) {
    return "";
  }

  const details = [];

  if (context.source) {
    details.push("source: " + context.source);
  }

  if (context.workMode) {
    details.push("mode: " + context.workMode);
  }

  if (context.nextStep) {
    details.push("next: " + context.nextStep);
  }

  if (details.length === 0) {
    return "";
  }

  return " (" + escapeTelegramMarkdown(details.join(", ")) + ")";
}

function buildCareerNextMove_(diversification, cvPerformance, coldLeads, funnel) {
  const automationSignal = cvPerformance.tags.filter(item => item.tag === "Automation")[0];

  if (automationSignal && automationSignal.assessmentStage > 0 && automationSignal.total < 5) {
    return (
      "Scale the strongest evidence: send 3-5 more Data/Automation or BI Analyst applications this week. " +
      "Automation has reached Assessment Stage, but the sample is still too small to conclude."
    );
  }

  if (funnel && funnel.interviewStage === 0 && funnel.assessmentStage > 0) {
    return (
      "Treat the next goal as Assessment-to-Interview conversion, not more raw applications. " +
      "Prepare interview stories for the companies already in Assessment Stage while maintaining 1-2 high-fit applications per day."
    );
  }

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
