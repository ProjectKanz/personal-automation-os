function doGet() {
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Personal Productivity OS")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


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


function addWebApplication(companyName, jobTitle, status, notes) {
  const input = [
    companyName || "",
    jobTitle || "",
    status || "",
    notes || ""
  ].join(" | ");
  const parsed = parseCareerAddInput_(input);


  if (!parsed.isValid) {
    throw new Error(parsed.error);
  }


  addApplicationFromTelegram(input);


  writeSystemLog(
    "Web App",
    "Add Application",
    "Success",
    "Added application from web app: " + (companyName || "") + " - " + (jobTitle || "")
  );


  return getWebDashboardData();
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
    .map(row => ({
      companyName: row.companyName,
      jobTitle: row.jobTitle,
      rowIndex: row.rowIndex,
      status: formatWebAppStatus_(row.status),
      category: row.category || "Unmapped",
      daysSinceApplied: row.daysSinceApplied,
      cvVersion: row.cvVersion || "General"
    }));


  return {
    totalApplications: totalApplications,
    interviewPipeline: interviewPipeline,
    followUpsDue: followUpsDue,
    rejectionRate: calculateWebAppPercentage_(rejectedCount, totalApplications),
    successRate: calculateWebAppPercentage_(offerCount, totalApplications),
    recent: recent,
    sectors: buildWebAppTopSectors_(rows)
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
    const sector = row.category || inferWebAppSector_(row);
    counts[sector] = (counts[sector] || 0) + 1;
  });


  return Object.keys(counts)
    .map(sector => ({
      label: sector,
      count: counts[sector]
    }))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label))
    .slice(0, 4);
}


function inferWebAppSector_(row) {
  const text = [
    row.companyName,
    row.jobTitle,
    row.cvVersion,
    row.notes
  ].join(" ").toLowerCase();


  if (/bank|finance|financial|idx|ocbc|bca|bri|bni|mandiri/.test(text)) return "Banking/Finance";
  if (/data|analyst|analytics|bi\b|automation/.test(text)) return "Data/Automation";
  if (/marketing|brand|sales|commercial|growth/.test(text)) return "Marketing/Commercial";
  if (/tech|digital|software|engineer|developer|it\b/.test(text)) return "Tech/Digital";
  if (/fmcg|retail|consumer|unilever|coca|nestle|indofood/.test(text)) return "FMCG/Retail";
  if (/logistic|transport|supply|warehouse/.test(text)) return "Logistics/Transport";


  return "Other";
}


function buildWebAppTradingSummary_(ss, today) {
  const sheet = ss.getSheetByName("Log");


  if (!sheet || sheet.getLastRow() < 20) {
    return {
      closedTrades: 0,
      wins: 0,
      losses: 0,
      winRate: 0,
      recent: []
    };
  }


  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(today.getDate() - 6);
  const data = sheet.getRange(20, 2, sheet.getLastRow() - 19, 18).getValues();
  const trades = [];


  data.forEach(row => {
    const symbol = row[0] || "";
    const side = row[1] || "";
    const exitDate = row[8] ? new Date(row[8]) : null;
    const pL = Number(row[13]) || 0;
    const status = row[15] ? row[15].toString().trim() : "";


    if (!symbol || !(exitDate instanceof Date) || isNaN(exitDate.getTime())) return;
    if (exitDate < sevenDaysAgo || exitDate > today) return;
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


  return {
    closedTrades: trades.length,
    wins: wins,
    losses: losses,
    winRate: calculateWebAppPercentage_(wins, trades.length),
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
