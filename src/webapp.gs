function doGet() {
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Personal Productivity OS")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


const WEB_APP_MT_CV_FOLDER_PROPERTY = "WEB_APP_MT_CV_FOLDER_ID";
const WEB_APP_DATA_CV_FOLDER_PROPERTY = "WEB_APP_DATA_CV_FOLDER_ID";


function getWebDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = getWebAppToday_();


  return {
    generatedAt: formatWebAppDateTime_(new Date()),
    spreadsheetUrl: ss.getUrl(),
    habit: buildWebAppHabitSummary_(ss, today),
    career: buildWebAppCareerSummary_(ss, today),
    trading: buildWebAppTradingSummary_(ss, today),
    audit: buildWebAppAuditSummary_(ss)
  };
}


function getWebTradingDataForRange(startDateValue, endDateValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startDate = parseWebAppDate_(startDateValue);
  const endDate = parseWebAppDate_(endDateValue);


  if (!startDate || !endDate) {
    throw new Error("Start date and end date are required.");
  }


  if (startDate.getTime() > endDate.getTime()) {
    throw new Error("Start date must be before end date.");
  }


  return buildWebAppTradingSummaryForRange_(ss, startDate, endDate);
}


function updateWebHabitStatus(activityName, isDone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Daily_DB");
  const today = getWebAppToday_();
  const match = findTodayWebHabitRow_(dbSheet, today, activityName);


  dbSheet.getRange(match.rowIndex, 4).setValue(isDone === true);
  writeSystemLog(
    "Web App",
    "Update Habit Status",
    "Success",
    "Updated " + match.activity + " to " + (isDone === true ? "done" : "open")
  );


  return getWebDashboardData();
}


function updateWebHabitNote(activityName, noteText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Daily_DB");
  const today = getWebAppToday_();
  const match = findTodayWebHabitRow_(dbSheet, today, activityName);
  const cleanNote = noteText === null || noteText === undefined
    ? ""
    : noteText.toString().trim();


  dbSheet.getRange(match.rowIndex, 5).setValue(cleanNote);
  writeSystemLog(
    "Web App",
    "Update Habit Note",
    "Success",
    "Updated note for " + match.activity
  );


  return getWebDashboardData();
}


function addWebApplication(companyName, jobTitle, status, notes, jobPostingLink, cvFile, dateAppliedValue) {
  const safeNotes = notes && notes.toString().trim() !== ""
    ? notes
    : "Added from web app";
  const input = [
    companyName || "",
    jobTitle || "",
    status || "",
    safeNotes
  ].join(" | ");
  const parsed = parseCareerAddInput_(input);


  if (!parsed.isValid) {
    throw new Error(parsed.error);
  }


  const cvUrl = uploadWebApplicationCv_(cvFile);
  addApplicationFromTelegram(input);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Applications");
  const targetRow = findLatestWebApplicationRow_(sheet, parsed.company, parsed.role);
  const dateApplied = parseWebAppDate_(dateAppliedValue);


  if (!dateApplied) {
    throw new Error("Date applied is required.");
  }


  sheet.getRange(targetRow, APPLICATIONS_COL.DATE_APPLIED + 1).setValue(dateApplied);
  sheet.getRange(targetRow, APPLICATIONS_COL.LINK_JOB_POSTING + 1).setValue(jobPostingLink || "");
  sheet.getRange(targetRow, APPLICATIONS_COL.CV_VERSION + 1).setValue(cvUrl);


  writeSystemLog(
    "Web App",
    "Add Application",
    "Success",
    "Added application from web app with CV upload: " + (companyName || "") + " - " + (jobTitle || "")
  );


  return getWebDashboardData();
}


function uploadWebApplicationCv_(cvFile) {
  if (!cvFile || !cvFile.name || !cvFile.base64) {
    throw new Error("CV PDF file is required.");
  }


  const fileName = cvFile.name.toString();
  const lowerName = fileName.toLowerCase();
  const isPdf = lowerName.endsWith(".pdf") ||
    cvFile.mimeType === "application/pdf";


  if (!isPdf) {
    throw new Error("CV file must be a PDF.");
  }


  const folderId = getWebApplicationCvFolderId_(fileName);
  const bytes = Utilities.base64Decode(cvFile.base64);
  const blob = Utilities.newBlob(
    bytes,
    cvFile.mimeType || "application/pdf",
    fileName
  );
  const file = DriveApp
    .getFolderById(folderId)
    .createFile(blob)
    .setName(fileName);

  // Portfolio safety: keep uploaded CVs private by default.
  // If a reviewer needs access, share a sanitized demo file manually instead of
  // making every uploaded CV public with ANYONE_WITH_LINK.


  return file.getUrl();
}


function getWebApplicationCvFolderId_(fileName) {
  const lowerName = fileName.toLowerCase();


  if (lowerName.indexOf("mt") !== -1) {
    return getRequiredProperty(WEB_APP_MT_CV_FOLDER_PROPERTY);
  }


  if (lowerName.indexOf("data") !== -1) {
    return getRequiredProperty(WEB_APP_DATA_CV_FOLDER_PROPERTY);
  }


  return getRequiredProperty(WEB_APP_DATA_CV_FOLDER_PROPERTY);
}


function findLatestWebApplicationRow_(sheet, companyName, jobTitle) {
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error("Could not find newly added application row.");
  }


  const cleanCompany = String(companyName || "").trim();
  const cleanJobTitle = String(jobTitle || "").trim();
  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, APPLICATIONS_V3_HEADERS.length)
    .getValues();


  for (let i = values.length - 1; i >= 0; i--) {
    const rowCompany = String(values[i][APPLICATIONS_COL.COMPANY_NAME] || "").trim();
    const rowJobTitle = String(values[i][APPLICATIONS_COL.JOB_TITLE] || "").trim();


    if (rowCompany === cleanCompany && rowJobTitle === cleanJobTitle) {
      return i + 2;
    }
  }


  throw new Error("Could not find newly added application row.");
}


function updateWebApplicationStatus(rowIndex, companyName, jobTitle, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Applications");


  if (!sheet) {
    throw new Error("Sheet Applications tidak ditemukan.");
  }


  const targetRow = Number(rowIndex);
  const cleanCompany = String(companyName || "").trim();
  const cleanJobTitle = String(jobTitle || "").trim();
  const normalizedStatus = normalizeCareerStatus_(status);


  if (!targetRow || targetRow < 2 || targetRow > sheet.getLastRow()) {
    throw new Error("Application row is invalid.");
  }


  if (!normalizedStatus) {
    throw new Error("Status must be one of: " + APPLICATIONS_STATUS_OPTIONS.join(", "));
  }


  const rowValues = sheet
    .getRange(targetRow, 1, 1, APPLICATIONS_V3_HEADERS.length)
    .getValues()[0];
  const currentCompany = String(rowValues[APPLICATIONS_COL.COMPANY_NAME] || "").trim();
  const currentJobTitle = String(rowValues[APPLICATIONS_COL.JOB_TITLE] || "").trim();


  if (currentCompany !== cleanCompany || currentJobTitle !== cleanJobTitle) {
    throw new Error("Application row no longer matches the selected company and role.");
  }


  sheet.getRange(targetRow, APPLICATIONS_COL.STATUS + 1).setValue(normalizedStatus);
  writeSystemLog(
    "Web App",
    "Update Application Status",
    "Success",
    "Updated " + cleanCompany + " - " + cleanJobTitle + " to " + normalizedStatus
  );


  return getWebDashboardData();
}


function findTodayWebHabitRow_(dbSheet, today, activityName) {
  if (!dbSheet) {
    throw new Error("Sheet Daily_DB tidak ditemukan.");
  }


  const activity = activityName ? activityName.toString().trim() : "";


  if (activity === "") {
    throw new Error("Habit activity name is required.");
  }


  const targetTime = normalizeDateOnly_(today).getTime();
  const data = dbSheet.getDataRange().getValues();
  const matches = [];


  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0];
    const rowActivity = data[i][2] ? data[i][2].toString().trim() : "";


    if (!(rowDate instanceof Date)) continue;
    if (normalizeDateOnly_(rowDate).getTime() !== targetTime) continue;
    if (rowActivity !== activity) continue;


    matches.push({
      rowIndex: i + 1,
      activity: rowActivity
    });
  }


  if (matches.length === 0) {
    throw new Error("Habit hari ini tidak ditemukan: " + activity);
  }


  if (matches.length > 1) {
    throw new Error("Ada beberapa habit hari ini dengan nama yang sama: " + activity);
  }


  return matches[0];
}


function buildWebAppHabitSummary_(ss, today) {
  const sheet = ss.getSheetByName("Daily_DB");


  if (!sheet || sheet.getLastRow() < 2) {
    return {
      total: 0,
      done: 0,
      percentage: 0,
      missing: [],
      completed: [],
      notesCount: 0,
      weeklyDone: 0,
      weeklyTotal: 0,
      weeklyPercentage: 0
    };
  }


  const data = sheet.getDataRange().getValues();
  const todayTime = today.getTime();
  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(today.getDate() - 6);
  const todayRows = [];
  const weekRows = [];


  data.slice(1).forEach(row => {
    if (!row || !(row[0] instanceof Date)) return;


    const rowDate = normalizeDateOnly_(row[0]);
    const rowTime = rowDate.getTime();


    if (rowTime === todayTime) {
      todayRows.push(row);
    }


    if (rowTime >= sevenDaysAgo.getTime() && rowTime <= todayTime) {
      weekRows.push(row);
    }
  });


  const doneRows = todayRows.filter(row => row[3] === true);
  const missingRows = todayRows.filter(row => row[3] !== true);
  const weeklyDone = weekRows.filter(row => row[3] === true).length;


  return {
    total: todayRows.length,
    done: doneRows.length,
    percentage: calculateWebAppPercentage_(doneRows.length, todayRows.length),
    missing: missingRows.map(row => buildWebAppHabitItem_(row)),
    completed: doneRows.map(row => buildWebAppHabitItem_(row)),
    notesCount: todayRows.filter(row => row[4] && row[4].toString().trim() !== "").length,
    weeklyDone: weeklyDone,
    weeklyTotal: weekRows.length,
    weeklyPercentage: calculateWebAppPercentage_(weeklyDone, weekRows.length)
  };
}


function buildWebAppHabitItem_(row) {
  return {
    category: row[1] || "General",
    activity: row[2] || "Unnamed habit",
    note: row[4] || "",
    done: row[3] === true
  };
}


function buildWebAppCareerSummary_(ss, today) {
  const sheet = ss.getSheetByName("Applications");


  if (!sheet || sheet.getLastRow() < 2) {
    return {
      totalApplications: 0,
      interviewPipeline: 0,
      followUpsDue: 0,
      rejectionRate: 0,
      successRate: 0,
      recent: [],
      sectors: []
    };
  }


  const lastRow = sheet.getLastRow();
  const lastColumn = Math.min(Math.max(sheet.getLastColumn(), 12), 12);
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = buildWebAppHeaderMap_(values[0]);
  const rows = values
    .slice(1)
    .filter(row => row.some(value => value !== "" && value !== null))
    .map((row, index) => buildWebAppApplicationItem_(row, headers, today, index + 2));
  const totalApplications = rows.length;
  const offerCount = rows.filter(row => row.statusLower === "offer").length;
  const rejectedCount = rows.filter(row => row.statusLower === "rejected").length;
  const interviewPipeline = rows.filter(row => row.statusLower.indexOf("interview") !== -1).length;
  const followUpsDue = rows.filter(row => row.followUpDue).length;
  const recent = rows
    .slice()
    .sort((a, b) => {
      const dateA = a.dateApplied ? a.dateApplied.getTime() : 0;
      const dateB = b.dateApplied ? b.dateApplied.getTime() : 0;
      return dateB - dateA;
    })
    .slice(0, 5)
    .map(row => buildWebAppApplicationPayload_(row));


  return {
    totalApplications: totalApplications,
    interviewPipeline: interviewPipeline,
    followUpsDue: followUpsDue,
    rejectionRate: calculateWebAppPercentage_(rejectedCount, totalApplications),
    successRate: calculateWebAppPercentage_(offerCount, totalApplications),
    recent: recent,
    allApplications: rows
      .slice()
      .sort((a, b) => {
        const dateA = a.dateApplied ? a.dateApplied.getTime() : 0;
        const dateB = b.dateApplied ? b.dateApplied.getTime() : 0;
        return dateB - dateA;
      })
      .map(row => buildWebAppApplicationPayload_(row)),
    sectors: buildWebAppTopSectors_(rows)
  };
}


function buildWebAppApplicationPayload_(row) {
  return {
    companyName: row.companyName,
    jobTitle: row.jobTitle,
    rowIndex: row.rowIndex,
    status: formatWebAppStatus_(row.status),
    category: row.category || "Unmapped",
    daysSinceApplied: row.daysSinceApplied,
    cvVersion: row.cvVersion || "General",
    jobLink: row.jobLink || ""
  };
}


function buildWebAppApplicationItem_(row, headers, today, rowIndex) {
  const dateApplied = parseWebAppDate_(getWebAppField_(row, headers, ["Date Applied", "Date"], 3));
  const followUpDate = parseWebAppDate_(getWebAppField_(row, headers, ["Follow Up Date", "Date FU"], 8));
  const followUpRequired = String(getWebAppField_(row, headers, ["FU Required?", "FU"], 7) || "").trim().toLowerCase();
  const status = String(getWebAppField_(row, headers, ["Status"], 4) || "").trim();
  const statusLower = status.toLowerCase();


  return {
    rowIndex: rowIndex,
    companyName: getWebAppField_(row, headers, ["Company Name", "Company"], 0) || "-",
    jobTitle: getWebAppField_(row, headers, ["Job Title", "Position"], 1) || "-",
    category: getWebAppField_(row, headers, ["Category", "Industry"], 2) || "",
    status: status || "Unknown",
    statusLower: statusLower,
    dateApplied: dateApplied,
    daysSinceApplied: dateApplied ? Math.floor((today.getTime() - dateApplied.getTime()) / 86400000) : null,
    cvVersion: getWebAppField_(row, headers, ["CV VERSION", "CV", "CV Version"], 10) || "",
    jobLink: getWebAppField_(row, headers, ["Link Job posting", "Link"], 9) || "",
    notes: getWebAppField_(row, headers, ["Notes"], 11) || "",
    followUpDue: followUpRequired === "yes" && followUpDate && followUpDate.getTime() <= today.getTime()
  };
}


function buildWebAppHeaderMap_(headers) {
  const map = {};


  headers.forEach((header, index) => {
    const key = normalizeWebAppHeader_(header);
    if (key) map[key] = index;
  });


  return map;
}


function getWebAppField_(row, headers, candidates, fallbackIndex) {
  for (let i = 0; i < candidates.length; i++) {
    const index = headers[normalizeWebAppHeader_(candidates[i])];


    if (index !== undefined) {
      return row[index];
    }
  }


  return row[fallbackIndex];
}


function normalizeWebAppHeader_(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}


function buildWebAppTopSectors_(rows) {
  const counts = {};


  rows.forEach(row => {
    const sector = inferWebAppSector_(row);
    counts[sector] = (counts[sector] || 0) + 1;
  });


  return Object.keys(counts)
    .map(sector => ({
      label: sector,
      count: counts[sector]
    }))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label))
    .slice(0, 8);
}


function inferWebAppSector_(row) {
  const category = String(row.category || "").trim();
  const categoryLower = category.toLowerCase();
  const text = [
    category,
    row.companyName,
    row.jobTitle,
    row.cvVersion,
    row.notes
  ].join(" ").toLowerCase();


  if (category && !/^(other|unmapped|mt|management trainee|graduate trainee|project management)$/i.test(categoryLower)) {
    const normalizedCategory = normalizeWebAppSectorFromCategory_(category);


    if (normalizedCategory) {
      return normalizedCategory;
    }
  }


  if (/astra\s*credit|\bacc\b|adira|fif\s*group|bfi\s*finance|oto\s*finance|mega\s*finance|mandiri\s*tunas|automotive\s*finance|car\s*finance/.test(text)) return "Automotive Finance";
  if (/bank|ocbc|bca|bri|bni|btn|mandiri|cimb|danamon|permata|maybank|uob|hsbc|idx|stock\s*exchange|securities|broker|investment|asset\s*management|wealth|financial|finance/.test(text)) return "Banking/Capital Markets";
  if (/sampoerna|philip\s*morris|philip\s*moris|tobacco|rokok|hm\s*sampoerna|pmi\s*career|djarum|gudang\s*garam|korea\s*tomorrow/.test(text)) return "Tobacco/FMCG";
  if (/unilever|coca|ccep|nestle|indofood|mayora|wings|danone|p&g|procter|consumer\s*goods|fmcg|beverage|food/.test(text)) return "FMCG/Food & Beverage";
  if (/map|retail|store|merchandis|fashion|apparel|mall|ecommerce|commerce|shop|marketplace/.test(text)) return "Retail/E-Commerce";
  if (/adecco|recruitment|headhunter|staffing|human\s*resource|hr\b|talent\s*acquisition|outsourcing/.test(text)) return "Recruitment/HR Services";
  if (/data|analyst|analytics|business\s*intelligence|\bbi\b|automation|machine\s*learning|ai\b|dashboard|reporting/.test(text)) return "Data/Analytics";
  if (/software|developer|engineer|engineering|programmer|frontend|backend|fullstack|cloud|devops|it\b|information\s*technology|digital|tech|huawei|siemens|telkom|gojek|tokopedia|shopee|grab|traveloka|bukalapak/.test(text)) return "Technology/Digital";
  if (/marketing|brand|growth|campaign|commercial|sales|business\s*development|account\s*executive|partnership|customer\s*success/.test(text)) return "Marketing/Commercial";
  if (/logistic|transport|shipping|supply\s*chain|warehouse|procurement|purchasing|inventory|deliveree|transjakarta|freight|distribution/.test(text)) return "Logistics/Supply Chain";
  if (/mining|coal|nickel|oil|gas|energy|renewable|geothermal|petro|pertamina|pln|adaro|vale|freeport/.test(text)) return "Energy/Mining";
  if (/property|real\s*estate|construction|developer|building|civil|architecture|contractor|infrastructure|tower\s*bersama|\btbg\b|telecom\s*infrastructure/.test(text)) return "Property/Infrastructure";
  if (/hospital|health|medical|pharma|clinic|medicine|healthcare|kimia\s*farma|kalbe|biofarma/.test(text)) return "Healthcare/Pharma";
  if (/manufactur|factory|industrial|plant|production|quality\s*control|qa\b|qc\b|operations|operator/.test(text)) return "Manufacturing/Industrial";
  if (/consult|advisory|research|nielsen|abeam|accenture|deloitte|pwc|kpmg|ey|mckinsey|bcg|bain/.test(text)) return "Consulting/Research";
  if (/government|ministry|kementerian|bumn|state\s*owned|public\s*sector|ojk|bi\b|bank\s*indonesia/.test(text)) return "Government/Public Sector";
  if (/education|school|university|campus|academy|learning|training|course/.test(text)) return "Education/Training";
  if (/hotel|tourism|travel|hospitality|restaurant|f&b|cafe|resort/.test(text)) return "Hospitality/Travel";
  if (/media|creative|content|design|advertising|agency|production|social\s*media/.test(text)) return "Media/Creative";


  return "Other";
}


function normalizeWebAppSectorFromCategory_(category) {
  const lower = String(category || "").trim().toLowerCase();


  if (/bank|finance|financial|capital|investment|securities/.test(lower)) return "Banking/Capital Markets";
  if (/automotive|auto|car/.test(lower)) return "Automotive Finance";
  if (/fmcg|consumer|food|beverage|tobacco/.test(lower)) return "FMCG/Food & Beverage";
  if (/retail|commerce|marketplace/.test(lower)) return "Retail/E-Commerce";
  if (/data|analytics|business\s*intelligence|automation/.test(lower)) return "Data/Analytics";
  if (/tech|digital|software|it|engineering/.test(lower)) return "Technology/Digital";
  if (/marketing|sales|commercial|business\s*development/.test(lower)) return "Marketing/Commercial";
  if (/logistic|supply|procurement|transport/.test(lower)) return "Logistics/Supply Chain";
  if (/energy|mining|oil|gas/.test(lower)) return "Energy/Mining";
  if (/property|construction|infrastructure|telecom/.test(lower)) return "Property/Infrastructure";
  if (/health|pharma|medical/.test(lower)) return "Healthcare/Pharma";
  if (/manufactur|industrial|operations/.test(lower)) return "Manufacturing/Industrial";
  if (/consult|research|advisory/.test(lower)) return "Consulting/Research";
  if (/government|public|bumn/.test(lower)) return "Government/Public Sector";
  if (/education|training/.test(lower)) return "Education/Training";


  return "";
}


function buildWebAppTradingSummary_(ss, today) {
  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(today.getDate() - 6);


  return buildWebAppTradingSummaryForRange_(ss, sevenDaysAgo, today);
}


function buildWebAppTradingSummaryForRange_(ss, startDate, endDate) {
  const sheet = ss.getSheetByName("Log");


  if (!sheet || sheet.getLastRow() < 20) {
    return {
      startDate: formatWebAppDateInput_(startDate),
      endDate: formatWebAppDateInput_(endDate),
      closedTrades: 0,
      wins: 0,
      losses: 0,
      winRate: 0,
      totalPnl: 0,
      trades: [],
      recent: []
    };
  }


  const data = sheet.getRange(20, 2, sheet.getLastRow() - 19, 18).getValues();
  const trades = [];


  data.forEach(row => {
    const symbol = row[0] || "";
    const side = row[1] || "";
    const exitDate = row[8] ? new Date(row[8]) : null;
    const pL = Number(row[13]) || 0;
    const status = row[15] ? row[15].toString().trim() : "";


    if (!symbol || !(exitDate instanceof Date) || isNaN(exitDate.getTime())) return;
    const normalizedExitDate = normalizeDateOnly_(exitDate);
    if (normalizedExitDate < startDate || normalizedExitDate > endDate) return;
    if (status !== "Closed") return;


    trades.push({
      symbol: symbol,
      side: side,
      pL: pL,
      outcome: pL > 0 ? "Win" : pL < 0 ? "Loss" : "Flat",
      exitDate: formatWebAppDate_(exitDate),
      remarks: row[17] || ""
    });
  });


  const wins = trades.filter(trade => trade.pL > 0).length;
  const losses = trades.filter(trade => trade.pL < 0).length;
  const totalPnl = trades.reduce((sum, trade) => sum + trade.pL, 0);


  return {
    startDate: formatWebAppDateInput_(startDate),
    endDate: formatWebAppDateInput_(endDate),
    closedTrades: trades.length,
    wins: wins,
    losses: losses,
    winRate: calculateWebAppPercentage_(wins, trades.length),
    totalPnl: totalPnl,
    trades: trades.slice().reverse(),
    recent: trades.slice(-5).reverse()
  };
}


function buildWebAppAuditSummary_(ss) {
  const sheet = ss.getSheetByName("AI_Audit");


  if (!sheet || sheet.getLastRow() < 1) {
    return {
      hasAudit: false,
      row: null,
      preview: "No audit available yet."
    };
  }


  const latest = getLatestSavedAudit(sheet);


  if (!latest) {
    return {
      hasAudit: false,
      row: null,
      preview: "No valid audit saved yet."
    };
  }


  return {
    hasAudit: true,
    row: latest.row,
    preview: latest.text.substring(0, 520)
  };
}


function getWebAppToday_() {
  if (typeof getTodayWibDate_ === "function") {
    return getTodayWibDate_();
  }


  return normalizeDateOnly_(new Date());
}


function calculateWebAppPercentage_(count, total) {
  if (!total) return 0;
  return Math.round((count / total) * 100);
}


function formatWebAppStatus_(status) {
  const text = String(status || "").trim();
  if (!text) return "Unknown";


  return text
    .split(" ")
    .map(part => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}


function formatWebAppDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd MMM");
}


function formatWebAppDateInput_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}


function formatWebAppDateTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd MMM yyyy HH:mm");
}


function parseWebAppDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return normalizeDateOnly_(value);
  }


  if (value === "" || value === null || value === undefined) {
    return null;
  }


  const parsed = new Date(value);


  if (isNaN(parsed.getTime())) {
    return null;
  }


  return normalizeDateOnly_(parsed);
}
