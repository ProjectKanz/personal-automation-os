# Source Code Bundle for Gemini

Generated from the `src` folder. Each section preserves the original file path.

## File Structure

```text
src/
  appsscript.json
  brief.gs
  career.gs
  habit.gs
  main.gs
  telegram.gs
  tradingParser.gs
  utils.gs
  weeklyAudit.gs
```

## src/appsscript.json

```json
{
    "timeZone": "Asia/Jakarta",
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8"
  }
  
```

## src/brief.gs

```javascript
// Daily operator brief for V3.1 Career Intelligence & Reliability.
function sendDailyOperatorBrief() {
  try {
    const message = buildDailyOperatorBriefMessage_();
    sendText(MY_ID, message);
    writeSystemLog(
      "Brief",
      "Send Daily Operator Brief",
      "Success",
      "Daily operator brief sent."
    );
  } catch (error) {
    writeSystemLog(
      "Brief",
      "Send Daily Operator Brief",
      "Error",
      getErrorMessage_(error)
    );
    sendText(MY_ID, "⚠️ Daily operator brief failed: " + escapeTelegramMarkdown(getErrorMessage_(error)));
  }
}

function buildDailyOperatorBriefMessage_() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return (
    "🧭 *Daily Operator Brief*\n" +
    "Date: " + escapeTelegramMarkdown(formatApplicationDate_(today)) + "\n\n" +
    buildBriefHabitSection_(today) + "\n\n" +
    buildBriefCareerSection_() + "\n\n" +
    buildBriefStatusAgingSection_()
  );
}

function buildBriefHabitSection_(today) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Daily_DB");

  if (!dbSheet || dbSheet.getLastRow() < 2) {
    return "✅ *Habit Check*\nNo habit records found for today.";
  }

  const data = dbSheet.getDataRange().getValues();
  const todayRows = data.slice(1).filter(row =>
    row && row[0] instanceof Date && normalizeApplicationDate_(row[0]).getTime() === today.getTime()
  );
  const incompleteRows = todayRows.filter(row => row[3] !== true);

  if (todayRows.length === 0) {
    return "✅ *Habit Check*\nNo habit checklist generated for today yet.";
  }

  if (incompleteRows.length === 0) {
    return "✅ *Habit Check*\nAll habits are complete for today.";
  }

  const incompleteList = incompleteRows
    .slice(0, 8)
    .map(row => "- " + escapeTelegramMarkdown(row[2] || "(Unnamed habit)"))
    .join("\n");
  const overflow = incompleteRows.length > 8
    ? "\n+" + (incompleteRows.length - 8) + " more"
    : "";

  return (
    "🟡 *Habit Check*\n" +
    incompleteRows.length + "/" + todayRows.length + " habits still open:\n" +
    incompleteList +
    overflow
  );
}

function buildBriefCareerSection_() {
  const followUpCount = getCareerFollowUpCount();

  if (followUpCount === 0) {
    return "✅ *Career Check*\nNo applications require follow-up today.";
  }

  return (
    "📌 *Career Check*\n" +
    followUpCount + " application follow-up" + (followUpCount === 1 ? "" : "s") + " due today."
  );
}

function buildBriefStatusAgingSection_() {
  const oldestApplied = getOldestAppliedApplications(3);

  if (oldestApplied.length === 0) {
    return "✅ *Status Aging*\nNo active Applied records to age.";
  }

  const agingList = oldestApplied.map(item =>
    "- " + escapeTelegramMarkdown(item.companyName) + ": " +
    item.daysSinceApplied + " days"
  );

  return "⏳ *Status Aging*\n" + agingList.join("\n");
}

function setupDailyOperatorBriefTrigger() {
  deleteDailyOperatorBriefTriggers();

  ScriptApp
    .newTrigger("sendDailyOperatorBrief")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  writeSystemLog(
    "Brief",
    "Setup Daily Operator Brief Trigger",
    "Success",
    "Daily operator brief trigger scheduled for 08:00 WIB."
  );
}

function deleteDailyOperatorBriefTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendDailyOperatorBrief") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
```

## src/career.gs

```javascript
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
```

## src/habit.gs

```javascript
// Habit module (Phase 1-2 split only).
// TODO(v2.2): Centralize date normalization helper to reduce repeated logic.
function generateDailyChecklist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master_Habit");
  const dbSheet = ss.getSheetByName("Daily_DB");
  const today = new Date();
  today.setHours(0,0,0,0);


  // 1. Ambil apa saja yang sudah ada di database hari ini
  const dbData = dbSheet.getDataRange().getValues();
  let existingActivities = [];
  dbData.forEach(row => {
    if (row[0] instanceof Date && row[0].getTime() === today.getTime()) {
      existingActivities.push(row[2]); // Simpan nama aktivitas yang sudah ada
    }
  });


  // 2. Ambil daftar dari Master Habit
  const masterHabits = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 2).getValues();
  let dataBaru = [];


  masterHabits.forEach(habit => {
    // HANYA tambah jika habit belum ada di database hari ini
    if (!existingActivities.includes(habit[1]) && habit[1] !== "") {
      dataBaru.push([today, habit[0], habit[1], false, ""]);
    }
  });


  // 3. Masukkan ke database
  if (dataBaru.length > 0) {
    dbSheet.getRange(dbSheet.getLastRow() + 1, 1, dataBaru.length, 5).setValues(dataBaru);
    dbSheet.getRange(dbSheet.getLastRow() - dataBaru.length + 1, 4, dataBaru.length, 1).insertCheckboxes();
    const message = "Berhasil menambah " + dataBaru.length + " habit ke database.";
    console.log(message);
    writeSystemLog("Habit", "Generate Daily Checklist", "Success", message);
  }
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() === "Dashboard" && range.getColumn() === 4 && range.getRow() > 1) {
    const dbSheet = e.source.getSheetByName("Daily_DB");
    const row = range.getRow();
   
    // Ambil Kunci Identifikasi: Tanggal (A) dan Aktivitas (C)
    const tglDashboard = sheet.getRange(row, 1).getValue();
    const aktDashboard = sheet.getRange(row, 3).getValue();
    const valBaru = range.getValue();
   
    if (!(tglDashboard instanceof Date)) return;
   
    const dbData = dbSheet.getDataRange().getValues();
    const tglTime = tglDashboard.getTime();


    for (let i = 1; i < dbData.length; i++) {
      if (dbData[i][0] instanceof Date &&
          dbData[i][0].getTime() === tglTime &&
          dbData[i][2] === aktDashboard) {
        dbSheet.getRange(i + 1, 4).setValue(valBaru);
        break;
      }
    }
  }
}

function logDailyToDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName("Dashboard");
  const dbSheet = ss.getSheetByName("Daily_DB");
 
  // 1. Ambil Tanggal Utama dari sel A2
  const tgl = dashSheet.getRange("A2").getValue();
  if (!(tgl instanceof Date)) {
    const message = "Peringatan: Sel A2 harus berisi format tanggal yang benar.";
    console.log(message);
    writeSystemLog("Habit", "Log Daily To DB", "Error", message);
    return;
  }
 
  const targetDateString = tgl.toDateString();


  // 2. Deteksi baris terakhir secara otomatis (berdasarkan Kolom C / Aktivitas)
  const lastRow = dashSheet.getRange("C:C").getValues().filter(String).length;
 
  // 3. Ambil data dari baris 2 sampai baris terakhir (Kolom B sampai G)
  const dataRange = dashSheet.getRange(2, 2, lastRow - 1, 6);
  const data = dataRange.getValues();


  // 4. Hapus data lama di Daily_DB untuk tanggal yang sama (Anti-Duplikat)
  const dbData = dbSheet.getDataRange().getValues();
  for (let i = dbData.length - 1; i >= 1; i--) {
    if (dbData[i][0] instanceof Date && dbData[i][0].toDateString() === targetDateString) {
      dbSheet.deleteRow(i + 1);
    }
  }


  // 5. Masukkan data habit ke Daily_DB
  data.forEach(row => {
    const kategori = row[0];  // Kolom B
    const aktivitas = row[1]; // Kolom C
    const status = row[2];    // Kolom D
    const note = row[5];      // Kolom G
   
    if (aktivitas) {
      dbSheet.appendRow([tgl, kategori, aktivitas, status, note]);
    }
  });


  const message = "Sinkronisasi Selesai! Data tanggal " + targetDateString + " sudah masuk ke Daily_DB.";
  console.log(message);
  writeSystemLog("Habit", "Log Daily To DB", "Success", message);
}


function sendDailyHabitReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Daily_DB");


  if (!dbSheet) {
    writeSystemLog(
      "Habit Reminder",
      "Send Daily Habit Reminder",
      "Error",
      "Sheet Daily_DB tidak ditemukan."
    );
    sendText(MY_ID, "⚠️ Reminder habit gagal: sheet `Daily_DB` tidak ditemukan.");
    return;
  }


  const data = dbSheet.getDataRange().getValues();
  const today = new Date().setHours(0, 0, 0, 0);
  const message = buildDailyHabitReminderMessage(data, today);


  sendText(MY_ID, message);
}


function buildDailyHabitReminderMessage(data, today) {
  const todayRows = getTodayHabitRows(data, today);
  const total = todayRows.length;
  const done = todayRows.filter(row => row[3] === true).length;
  const missingRows = todayRows.filter(row => row[3] !== true);


  if (total === 0) {
    return "📋 Reminder habit: belum ada habit untuk hari ini di `Daily_DB`.";
  }


  if (missingRows.length === 0) {
    return "✅ Reminder habit: semua habit hari ini sudah selesai (" + done + "/" + total + "). Mantap.";
  }


  const missingList = missingRows.map(row =>
    "⬜ " + escapeTelegramMarkdown(row[2] || "(Tanpa nama habit)")
  );


  return (
    "⏰ *Reminder Habit Harian*\n" +
    "Progress: " + done + "/" + total + " selesai.\n\n" +
    "*Belum selesai:*\n" +
    missingList.join("\n")
  );
}


function setupDailyHabitReminderTrigger() {
  deleteDailyHabitReminderTriggers();


  ScriptApp
    .newTrigger("sendDailyHabitReminder")
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .create();
}


function deleteDailyHabitReminderTriggers() {
  const triggers = ScriptApp.getProjectTriggers();


  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendDailyHabitReminder") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
```

## src/main.gs

```javascript
// Functions moved into module files:
// - src/brief.gs
// - src/career.gs
// - src/habit.gs
// - src/telegram.gs
// - src/tradingParser.gs
// - src/weeklyAudit.gs
// - src/utils.gs

```

## src/telegram.gs

```javascript
// Telegram module (Phase 1-2 split only).
// TODO(v2.2): Split doPost into smaller handlers for easier testing/maintenance.
const TOKEN = getRequiredProperty("TELEGRAM_TOKEN");


const MY_ID = getRequiredProperty("TELEGRAM_USER_ID"); 

const WEB_APP_URL = getRequiredProperty("WEB_APP_URL");


// --- 2. AKTIVASI WEBHOOK ---
function setWebhook() {
  const telegramUrl =
    "https://api.telegram.org/bot" +
    TOKEN +
    "/setWebhook?url=" +
    encodeURIComponent(WEB_APP_URL);

  const response = UrlFetchApp.fetch(telegramUrl);
  Logger.log(response.getContentText());
}


// --- 3. MESIN UTAMA (doPost) ---
function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);


    if (!contents.message) return;


    const msg = contents.message;
    const chatId = msg.from.id;


    if (chatId != MY_ID) return;


    const ss = SpreadsheetApp.getActiveSpreadsheet();


    // =========================
    // 1. HANDLE PHOTO / MT5 LOG
    // =========================
    if (msg.photo) {
      sendText(chatId, "🔍 Sinkronisasi Tanggal & Field MT5... Tunggu sebentar.");


      const fileId = msg.photo[msg.photo.length - 1].file_id;
      const fileUrl = getTelegramFileUrl(fileId);


      const caption = msg.caption || "|||";
      const capParts = caption.split("|").map(p => p.trim());


      const prompt =
        "Extract trade data from this MT5 history screenshot. " +
        "CRITICAL INSTRUCTION: The actual exitPrice is ALWAYS the number on the right side of the arrow (->). " +
        "Example: If '4718.914 -> 4728.590', then entry is 4718.914 and exitPrice is 4728.590. " +
        "DO NOT use the T/P or S/L values as exitPrice. " +
        "Return ONLY a JSON ARRAY of objects with keys: symbol, side, entry, qty, openDate, sl, exitDate, exitPrice.";


      let aiDataList = callGeminiVision(fileUrl, prompt);


      if (aiDataList && !Array.isArray(aiDataList)) {
        aiDataList = [aiDataList];
      }


      if (!aiDataList || !Array.isArray(aiDataList) || aiDataList.length === 0) {
        sendText(chatId, "⚠️ AI tidak menemukan data trade yang valid dari screenshot.");
        return;
      }


      const logSheet = ss.getSheetByName("Log");


      if (!logSheet) {
        sendText(chatId, "⚠️ Error: Sheet 'Log' tidak ditemukan!");
        return;
      }


      // Cari row kosong pertama di kolom B, mulai dari row 21
      const startRow = 21;
      const columnB = logSheet.getRange("B:B").getValues();


      let targetRow = logSheet.getLastRow() + 1;


      for (let i = startRow - 1; i < columnB.length; i++) {
        if (columnB[i][0] === "") {
          targetRow = i + 1;
          break;
        }
      }


      const cleanDate = (str) => {
        if (!str) return "";
        return str.toString().replace(/\./g, "/");
      };


      aiDataList.forEach((trade, index) => {
  const currentRow = targetRow + index;


  // Normalisasi Side
  const sideRaw = (trade.side || trade.type || "").toString().toLowerCase();


  const sideFinal = sideRaw.includes("buy")
    ? "Long"
    : sideRaw.includes("sell")
      ? "Short"
      : sideRaw;


  // Normalisasi Qty / Lot
  const rawQty = parseFloat(trade.qty || trade.volume || 0);
  const finalQty = rawQty * 100;


  const rowData = [
    trade.symbol || trade.Symbol || "",
    sideFinal,
    trade.entry || trade.Entry || trade.openPrice || "",
    finalQty,
    cleanDate(trade.openDate || trade.openTime || trade.Open || trade.open),
    capParts[0] || "",
    capParts[1] || "",
    trade.sl || trade.stopLoss || trade["S / L"] || trade["S/L"] || "",
    cleanDate(trade.exitDate || trade.closeTime || trade.Exit || ""),
    trade.exitPrice || trade.closePrice || trade["T / P"] || trade["T/P"] || ""
  ];


  logSheet.getRange(currentRow, 2, 1, 10).setValues([rowData]); // B-K
  logSheet.getRange(currentRow, 18).setValue(capParts[2] || ""); // R
  logSheet.getRange(currentRow, 19).setValue(capParts[3] || ""); // S


  const requiredFields = [
    trade.symbol || trade.Symbol || "",
    sideFinal,
    trade.entry || trade.Entry || trade.openPrice || "",
    finalQty,
    cleanDate(trade.openDate || trade.openTime || trade.Open || trade.open),
    trade.sl || trade.stopLoss || trade["S / L"] || trade["S/L"] || "",
    cleanDate(trade.exitDate || trade.closeTime || trade.Exit || ""),
    trade.exitPrice || trade.closePrice || trade["T / P"] || trade["T/P"] || ""
  ];


  const hasMissingField = requiredFields.some(field =>
    field === "" || field === null || field === undefined || field === 0
  );


  const validationStatus = hasMissingField ? "Needs Review" : "OK";


  logSheet.getRange(currentRow, 20).setValue(validationStatus); // T: Validation Status


  writeSystemLog(
    "Trading Log",
    "Validate AI Extracted Trade",
    validationStatus,
    "Trade row " + currentRow + " validation status: " + validationStatus
  );
});
sendText(
  chatId,
  "✅ Selesai! Baris " + targetRow + " sampai " + (targetRow + aiDataList.length - 1) + " sudah terisi lengkap (B-K, R, S, T)."
);


return;
}
    // =========================
    // 2. HANDLE TEXT / HABIT BOT
    // =========================
    if (msg.text) {
      const rawText = msg.text.trim();
      const text = rawText.toLowerCase();


      if (text === "/help" || text === "/list") {
        sendText(chatId, buildTelegramHelpMessage());
        return;
      }


      if (text === "/followup") {
        sendText(chatId, getFollowUpList());
        return;
      }


      if (text === "/career") {
        sendText(chatId, buildCareerDashboardMessage());
        return;
      }


      const dbSheet = ss.getSheetByName("Daily_DB");


      if (!dbSheet) {
        sendText(chatId, "⚠️ Error: Sheet 'Daily_DB' tidak ditemukan!");
        return;
      }


      const data = dbSheet.getDataRange().getValues();
      const today = new Date().setHours(0, 0, 0, 0);


      if (text === "/today") {
        sendText(chatId, buildTodayHabitChecklist(data, today));
        return;
      }


      if (text === "/missing") {
        sendText(chatId, buildMissingHabitsMessage(data, today));
        return;
      }


      if (text === "/status") {
        sendText(chatId, buildHabitStatusMessage(data, today));
        return;
      }


      if (text === "/daily") {
        sendText(chatId, buildDailyVisibilityMessage(data, today));
        return;
      }


      if (text === "/close") {
        sendText(chatId, buildDailyClosingReviewMessage(data, today));
        return;
      }


      if (text === "/week") {
        sendText(chatId, buildWeeklyVisibilityMessage(data, today));
        return;
      }


      if (text === "/reminderstatus") {
        sendText(chatId, buildReminderStatusMessage());
        return;
      }


      if (text === "/note" || text.indexOf("/note ") === 0) {
        sendText(chatId, handleHabitNoteCommand(dbSheet, data, today, rawText));
        return;
      }


      if (text === "/lastaudit") {
        sendText(chatId, buildLastAuditMessage(ss));
        return;
      }


      if (text === "/audit") {
        sendText(chatId, "⏳ Menjalankan Audit... Tunggu sebentar.");
        const auditResult = runWeeklyAIAudit();


        if (auditResult && auditResult.success) {
          sendText(chatId, buildAuditSuccessMessage(auditResult));
        } else {
          sendText(chatId, buildAuditErrorMessage(auditResult));
        }


        return;
      }


      // Checklist habit dari teks natural
      let found = false;


      for (let i = 1; i < data.length; i++) {
        if (data[i] && data[i][0] instanceof Date && data[i][0].getTime() === today) {
          const habitName = data[i][2] ? data[i][2].toString().toLowerCase() : "";


          if (habitName.includes(text)) {
            dbSheet.getRange(i + 1, 4).setValue(true);
            sendText(chatId, "✅ Habit '" + data[i][2] + "' dicentang!");
            found = true;
            break;
          }
        }
      }


      if (!found) {
        sendText(chatId, "❌ Habit tidak ditemukan atau sudah lewat hari.");
      }


      return;
    }


  } catch (err) {
    sendText(MY_ID, "⚠️ Error: " + err.message);
  }
}

function getTelegramFileUrl(fileId) {
  const res = UrlFetchApp.fetch("https://api.telegram.org/bot" + TOKEN + "/getFile?file_id=" + fileId);
  return "https://api.telegram.org/file/bot" + TOKEN + "/" + JSON.parse(res.getContentText()).result.file_path;
}


function buildTelegramHelpMessage() {
  return (
    "🤖 *Menu Kanzan*:\n" +
    "1. Kirim Foto (MT5 Log)\n" +
    "2. `/today` - Checklist habit hari ini\n" +
    "3. `/missing` - Habit yang belum selesai\n" +
    "4. `/status` - Ringkasan progres habit\n" +
    "5. `/daily` - Ringkasan habit hari ini\n" +
    "6. `/close` - Review penutup hari\n" +
    "7. `/week` - Ringkasan habit 7 hari\n" +
    "8. `/reminderstatus` - Status reminder\n" +
    "9. `/audit` - AI Audit mingguan\n" +
    "10. `/lastaudit` - Preview audit terakhir\n" +
    "11. `/followup` - Memo follow-up lamaran kerja\n" +
    "12. `/career` - Executive career dashboard\n" +
    "13. `/note habit | alasan` - Tambah catatan habit\n" +
    "14. Ketik nama habit untuk mencentang\n\n" +
    "`/list` tetap bisa dipakai sebagai alias `/help`."
  );
}


function getTodayHabitRows(data, today) {
  if (!data || data.length <= 1) return [];


  return data.slice(1).filter(row =>
    row && row[0] instanceof Date && row[0].getTime() === today
  );
}


function buildTodayHabitChecklist(data, today) {
  const todayRows = getTodayHabitRows(data, today);


  if (todayRows.length === 0) {
    return "📋 Belum ada habit untuk hari ini di `Daily_DB`.";
  }


  const checklist = todayRows.map(row => {
    const marker = row[3] === true ? "✅" : "⬜";
    const habitName = row[2] || "(Tanpa nama habit)";
    return marker + " " + habitName;
  });


  return "📋 *Habit Hari Ini:*\n" + checklist.join("\n");
}


function buildMissingHabitsMessage(data, today) {
  const todayRows = getTodayHabitRows(data, today);
  const missingRows = todayRows.filter(row => row[3] !== true);


  if (todayRows.length === 0) {
    return "📋 Belum ada habit untuk hari ini di `Daily_DB`.";
  }


  if (missingRows.length === 0) {
    return "✅ Semua habit hari ini sudah selesai. Mantap.";
  }


  const missingList = missingRows.map(row =>
    "⬜ " + escapeTelegramMarkdown(row[2] || "(Tanpa nama habit)")
  );


  return "⏳ *Habit Belum Selesai Hari Ini:*\n" + missingList.join("\n");
}


function buildHabitStatusMessage(data, today) {
  const todayRows = getTodayHabitRows(data, today);
  const total = todayRows.length;
  let done = 0;


  todayRows.forEach(row => {
    if (row[3] === true) {
      done++;
    }
  });


  const percentage = total === 0 ? 0 : Math.round((done / total) * 100);


  return "📊 Habit Hari Ini: " + done + "/" + total + " selesai (" + percentage + "%).";
}


function buildDailyVisibilityMessage(data, today) {
  const todayRows = getTodayHabitRows(data, today);
  const total = todayRows.length;
  const done = todayRows.filter(row => row[3] === true).length;
  const percentage = calculateCompletionPercentage(done, total);
  const missingRows = todayRows.filter(row => row[3] !== true);
  const notesCount = countRowsWithNotes(todayRows);


  if (total === 0) {
    return "📅 *Daily Summary*\nBelum ada habit untuk hari ini di `Daily_DB`.";
  }


  const missingText = missingRows.length === 0
    ? "✅ Tidak ada. Semua habit selesai."
    : missingRows.map(row =>
      "⬜ " + escapeTelegramMarkdown(row[2] || "(Tanpa nama habit)")
    ).join("\n");


  return (
    "📅 *Daily Summary*\n" +
    "Progress: " + done + "/" + total + " selesai (" + percentage + "%)\n" +
    "Catatan hari ini: " + notesCount + "\n\n" +
    "*Belum selesai:*\n" +
    missingText
  );
}


function buildDailyClosingReviewMessage(data, today) {
  const todayRows = getTodayHabitRows(data, today);
  const total = todayRows.length;
  const done = todayRows.filter(row => row[3] === true).length;
  const percentage = calculateCompletionPercentage(done, total);
  const unfinishedRows = todayRows.filter(row => row[3] !== true);


  if (total === 0) {
    return "🌙 *Closing Review*\nBelum ada habit untuk hari ini di `Daily_DB`.";
  }


  if (unfinishedRows.length === 0) {
    return (
      "🌙 *Closing Review*\n" +
      "Progress: " + done + "/" + total + " selesai (" + percentage + "%)\n\n" +
      "✅ Semua habit hari ini sudah selesai. Hari ini bisa ditutup dengan rapi."
    );
  }


  const unfinishedWithNotes = [];
  const unfinishedWithoutNotes = [];


  unfinishedRows.forEach(row => {
    const habitName = row[2] ? row[2].toString() : "(Tanpa nama habit)";
    const note = row[4] ? row[4].toString().trim() : "";


    if (note === "") {
      unfinishedWithoutNotes.push(habitName);
    } else {
      unfinishedWithNotes.push({
        habitName: habitName,
        note: note
      });
    }
  });


  const missingNoteText = unfinishedWithoutNotes.length === 0
    ? "✅ Semua habit yang belum selesai sudah punya catatan."
    : unfinishedWithoutNotes.map(habitName =>
      "⚠️ " + escapeTelegramMarkdown(habitName)
    ).join("\n");


  const notedText = unfinishedWithNotes.length === 0
    ? "Belum ada catatan untuk habit yang belum selesai."
    : unfinishedWithNotes.map(item =>
      "📝 " + escapeTelegramMarkdown(item.habitName) + ": " + buildShortNotePreview(item.note)
    ).join("\n");


  return (
    "🌙 *Closing Review*\n" +
    "Progress: " + done + "/" + total + " selesai (" + percentage + "%)\n\n" +
    "*Belum selesai tanpa catatan:*\n" +
    missingNoteText +
    "\n\n*Catatan yang sudah ada:*\n" +
    notedText +
    "\n\nTambah catatan dengan:\n" +
    "`/note habit name | reason`"
  );
}


function buildWeeklyVisibilityMessage(data, today) {
  const weekRows = getLastSevenDayHabitRows(data, today);
  const total = weekRows.length;
  const done = weekRows.filter(row => row[3] === true).length;
  const percentage = calculateCompletionPercentage(done, total);
  const notesCount = countRowsWithNotes(weekRows);
  const topIncomplete = getTopIncompleteHabits(weekRows, 5);


  if (total === 0) {
    return "📆 *Weekly Summary*\nBelum ada habit dalam 7 hari terakhir di `Daily_DB`.";
  }


  const incompleteText = topIncomplete.length === 0
    ? "✅ Tidak ada. Semua habit dalam periode ini selesai."
    : topIncomplete.map(item =>
      "- " + escapeTelegramMarkdown(item.habitName) + ": " + item.count + " belum selesai"
    ).join("\n");


  return (
    "📆 *Weekly Summary \\(7 Hari\\)*\n" +
    "Progress: " + done + "/" + total + " selesai (" + percentage + "%)\n" +
    "Catatan/reason minggu ini: " + notesCount + "\n\n" +
    "*Top incomplete:*\n" +
    incompleteText
  );
}


function getLastSevenDayHabitRows(data, today) {
  if (!data || data.length <= 1) return [];


  const todayDate = new Date(today);
  const sevenDaysAgo = new Date(todayDate);
  sevenDaysAgo.setDate(todayDate.getDate() - 6);


  return data.slice(1).filter(row => {
    if (!row || !(row[0] instanceof Date)) return false;


    const rowDate = new Date(row[0]);
    rowDate.setHours(0, 0, 0, 0);


    return rowDate.getTime() >= sevenDaysAgo.getTime() &&
      rowDate.getTime() <= todayDate.getTime();
  });
}


function calculateCompletionPercentage(done, total) {
  return total === 0 ? 0 : Math.round((done / total) * 100);
}


function countRowsWithNotes(rows) {
  return rows.filter(row =>
    row[4] !== null &&
    row[4] !== undefined &&
    row[4].toString().trim() !== ""
  ).length;
}


function getTopIncompleteHabits(rows, limit) {
  const counts = {};


  rows.forEach(row => {
    if (row[3] === true) return;


    const habitName = row[2] ? row[2].toString() : "(Tanpa nama habit)";
    counts[habitName] = (counts[habitName] || 0) + 1;
  });


  return Object.keys(counts)
    .map(habitName => ({
      habitName: habitName,
      count: counts[habitName]
    }))
    .sort((a, b) => b.count - a.count || a.habitName.localeCompare(b.habitName))
    .slice(0, limit);
}


function buildShortNotePreview(note) {
  const maxLength = 140;
  const cleanNote = note ? note.toString().trim() : "";


  if (cleanNote.length <= maxLength) {
    return escapeTelegramMarkdown(cleanNote);
  }


  return escapeTelegramMarkdown(cleanNote.substring(0, maxLength).trim()) + "...";
}


function buildReminderStatusMessage() {
  const triggers = ScriptApp.getProjectTriggers();
  const reminderTriggers = triggers.filter(trigger => {
    if (trigger.getHandlerFunction() !== "sendDailyHabitReminder") return false;


    if (ScriptApp.EventType && ScriptApp.EventType.CLOCK) {
      return trigger.getEventType() === ScriptApp.EventType.CLOCK;
    }


    return true;
  });


  if (reminderTriggers.length === 0) {
    return "⏰ Daily habit reminder: belum aktif.";
  }


  return "⏰ Daily habit reminder: aktif (" + reminderTriggers.length + " trigger ditemukan).";
}


function handleHabitNoteCommand(dbSheet, data, today, text) {
  const notePrefix = "/note";
  const rawCommand = text.substring(notePrefix.length).trim();
  const separatorIndex = rawCommand.indexOf("|");


  if (rawCommand === "" || separatorIndex === -1) {
    return buildHabitNoteUsageMessage();
  }


  const habitQuery = rawCommand.substring(0, separatorIndex).trim();
  const noteText = rawCommand.substring(separatorIndex + 1).trim();


  if (habitQuery === "" || noteText === "") {
    return buildHabitNoteUsageMessage();
  }


  const queryLower = habitQuery.toLowerCase();
  const matches = [];


  for (let i = 1; i < data.length; i++) {
    if (data[i] && data[i][0] instanceof Date && data[i][0].getTime() === today) {
      const habitName = data[i][2] ? data[i][2].toString() : "";


      if (habitName.toLowerCase().includes(queryLower)) {
        matches.push({
          rowIndex: i + 1,
          habitName: habitName
        });
      }
    }
  }


  if (matches.length === 0) {
    return (
      "❌ Habit hari ini tidak ditemukan untuk: `" +
      escapeTelegramMarkdown(habitQuery) +
      "`.\nCoba cek `/today` atau `/missing`."
    );
  }


  if (matches.length > 1) {
    const matchList = matches.map(match =>
      "- " + escapeTelegramMarkdown(match.habitName || "(Tanpa nama habit)")
    );


    return (
      "⚠️ Ada beberapa habit yang cocok. Tolong lebih spesifik:\n" +
      matchList.join("\n")
    );
  }


  dbSheet.getRange(matches[0].rowIndex, 5).setValue(noteText);


  return (
    "📝 Catatan habit `" +
    escapeTelegramMarkdown(matches[0].habitName || "(Tanpa nama habit)") +
    "` diperbarui."
  );
}


function buildHabitNoteUsageMessage() {
  return (
    "Format catatan belum sesuai.\n" +
    "Gunakan:\n" +
    "`/note habit name | reason`"
  );
}


function buildLastAuditMessage(ss) {
  const auditSheet = ss.getSheetByName("AI_Audit");


  if (!auditSheet) {
    return "📄 Sheet `AI_Audit` tidak ditemukan.";
  }


  const lastAudit = getLatestSavedAudit(auditSheet);


  if (!lastAudit) {
    return "📄 Belum ada laporan audit valid yang tersimpan di `AI_Audit`.";
  }


  return (
    "📄 *Audit Terakhir*\n" +
    "Laporan lengkap tersedia di sheet `AI_Audit` row " + lastAudit.row + ".\n\n" +
    "*Preview:*\n" +
    buildAuditPreview(lastAudit.text)
  );
}


function getLatestSavedAudit(auditSheet) {
  const data = auditSheet.getDataRange().getValues();


  for (let i = data.length - 1; i >= 0; i--) {
    const auditText = data[i] && data[i][1] ? data[i][1].toString().trim() : "";


    if (isValidSavedAuditText(auditText)) {
      return {
        row: i + 1,
        text: auditText
      };
    }
  }


  return null;
}


function isValidSavedAuditText(text) {
  if (!text) return false;


  const lowerText = text.toLowerCase();
  const legacyErrorLabels = [
    "gemini response error",
    "gemini format error",
    "script error"
  ];


  if (legacyErrorLabels.some(label => lowerText === label)) {
    return false;
  }


  return lowerText.indexOf("weekly ai audit") !== -1;
}


function buildAuditSuccessMessage(auditResult) {
  const preview = buildAuditPreview(auditResult.text || "");


  return (
    "✅ Audit mingguan selesai.\n" +
    "Laporan lengkap sudah disimpan di sheet `AI_Audit`" +
    (auditResult.row ? " row " + auditResult.row : "") +
    ".\n\n" +
    "*Preview:*\n" +
    preview
  );
}


function buildAuditErrorMessage(auditResult) {
  const errorMessage = auditResult && auditResult.error
    ? auditResult.error
    : "Audit gagal dijalankan. Silakan coba lagi nanti.";


  return (
    "⚠️ Audit mingguan gagal dibuat.\n" +
    escapeTelegramMarkdown(errorMessage) +
    "\nDetail teknis sudah dicatat di `System_Log`."
  );
}


function buildAuditPreview(text) {
  const maxLength = 1200;
  const cleanText = text ? text.toString().trim() : "";


  if (cleanText === "") {
    return "(Tidak ada preview yang tersedia.)";
  }


  if (cleanText.length <= maxLength) {
    return escapeTelegramMarkdown(cleanText);
  }


  return escapeTelegramMarkdown(cleanText.substring(0, maxLength).trim()) +
    "\n\n...preview dipotong. Laporan lengkap ada di `AI_Audit`.";
}


function escapeTelegramMarkdown(text) {
  return text.toString().replace(/([_*`\[])/g, "\\$1");
}


function sendText(id, text) {
  const url = "https://api.telegram.org/bot" + TOKEN + "/sendMessage?chat_id=" + id + "&text=" + encodeURIComponent(text) + "&parse_mode=Markdown";
  UrlFetchApp.fetch(url);
}
```

## src/tradingParser.gs

```javascript
// Trading parser module (Phase 1-2 split only).
// TODO(v2.2): Harden Gemini response parsing and validation paths.
function callGeminiVision(imageUrl, prompt) {
  const imgBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  const base64Img = Utilities.base64Encode(imgBlob.getBytes());
  const payload = { "contents": [{ "parts": [{ "text": prompt }, { "inline_data": { "mime_type": "image/jpeg", "data": base64Img } }] }] };
  const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
 
  const response = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + GEMINI_API_KEY, options);
  const rawResponse = response.getContentText();
  const parsed = JSON.parse(rawResponse);

  if (!parsed.candidates || !parsed.candidates[0]) {
    throw new Error("Gemini Vision response missing candidates.");
  }

  if (!parsed.candidates[0].content) {
    throw new Error("Gemini Vision response missing content.");
  }

  if (!parsed.candidates[0].content.parts || !parsed.candidates[0].content.parts[0]) {
    throw new Error("Gemini Vision response missing content parts.");
  }

  if (!parsed.candidates[0].content.parts[0].text) {
    throw new Error("Gemini Vision response missing text output.");
  }

  const resText = parsed.candidates[0].content.parts[0].text;
 
  const start = resText.indexOf('[');
  const end = resText.lastIndexOf(']') + 1;
  if (start !== -1 && end !== -1) return JSON.parse(resText.substring(start, end));
  throw new Error("Format JSON tidak valid.");
}
```

## src/utils.gs

```javascript
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
        " because of a typed column restriction: " + getErrorMessage_(error)
      );
      return;
    }

    throw error;
  }
}


function isTypedColumnsError_(error) {
  return getErrorMessage_(error).toLowerCase().indexOf("typed column") !== -1;
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
```

## src/weeklyAudit.gs

```javascript
// Weekly audit module (Phase 1-2 split only).
// TODO(v2.2): Extract prompt builder and data collectors into smaller helpers.
function runWeeklyAIAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();


  const dbSheet = ss.getSheetByName("Daily_DB");
  const logSheet = ss.getSheetByName("Log");
  const appSheet = ss.getSheetByName("Applications");
  const auditSheet = ss.getSheetByName("AI_Audit");


  // =========================
  // 0. SHEET VALIDATION
  // =========================
  if (!dbSheet) {
    Logger.log("Sheet Daily_DB tidak ditemukan.");
    writeSystemLog(
      "Weekly AI Audit",
      "Validate Required Sheets",
      "Error",
      "Missing required sheet: Daily_DB"
    );
    return {
      success: false,
      error: "Sheet Daily_DB tidak ditemukan."
    };
  }


  if (!logSheet) {
    Logger.log("Sheet Log tidak ditemukan.");
    writeSystemLog(
      "Weekly AI Audit",
      "Validate Required Sheets",
      "Error",
      "Missing required sheet: Log"
    );
    return {
      success: false,
      error: "Sheet Log tidak ditemukan."
    };
  }


  if (!appSheet) {
    Logger.log("Sheet Applications tidak ditemukan.");
    writeSystemLog(
      "Weekly AI Audit",
      "Validate Required Sheets",
      "Error",
      "Missing required sheet: Applications"
    );
    return {
      success: false,
      error: "Sheet Applications tidak ditemukan."
    };
  }


  if (!auditSheet) {
    Logger.log("Sheet AI_Audit tidak ditemukan.");
    writeSystemLog(
      "Weekly AI Audit",
      "Validate Required Sheets",
      "Error",
      "Missing required sheet: AI_Audit"
    );
    return {
      success: false,
      error: "Sheet AI_Audit tidak ditemukan."
    };
  }


  const today = new Date();
today.setHours(0, 0, 0, 0);


const sevenDaysAgo = new Date(today);
sevenDaysAgo.setDate(today.getDate() - 6);


  // =========================
  // HELPER: CONVERT ANY VALUE TO READABLE TEXT
  // =========================
  const toReadableText = (value) => {
    if (value === null || value === undefined) return "";


    if (typeof value === "string") return value;


    if (Array.isArray(value)) {
      return value.map(item => {
        if (typeof item === "string") {
          return "• " + item;
        }


        if (typeof item === "object" && item !== null) {
          return "• " + Object.entries(item)
            .map(([key, val]) => `${key}: ${val}`)
            .join(" | ");
        }


        return "• " + String(item);
      }).join("\n");
    }


    if (typeof value === "object") {
      return Object.entries(value)
        .map(([key, val]) => `${key}: ${val}`)
        .join("\n");
    }


    return String(value);
  };


  // =========================
  // 1. HABIT SUMMARY
  // =========================
  const habitData = dbSheet.getDataRange().getValues();
  let habitSummary = {};
  let habitNotesContext = [];


  habitData.slice(1).forEach(row => {
    const habitDate = row[0] ? new Date(row[0]) : null;
    const habitName = row[2] || "Unnamed Habit";
    const isDone = row[3] === true;
    const note = row[4] ? row[4].toString().trim() : "";


    if (habitDate && habitDate >= sevenDaysAgo) {
      if (!habitSummary[habitName]) {
        habitSummary[habitName] = {
          done: 0,
          total: 0
        };
      }


      habitSummary[habitName].total++;


      if (isDone) {
        habitSummary[habitName].done++;
      }


      if (note !== "") {
        habitNotesContext.push({
          date: habitDate,
          habit: habitName,
          completed: isDone,
          note: note
        });
      }
    }
  });


  // =========================
  // 2. TRADING SUMMARY
  // =========================
  const lastRowLog = logSheet.getLastRow();
  let tradeHistory = [];


  if (lastRowLog >= 20) {
    /*
      Range:
      Start row: 20
      Start col: 2 = B
      Number of columns: 18 = B sampai S


      Index mapping:
      row[0]  = B
      row[1]  = C
      row[2]  = D
      row[3]  = E
      row[4]  = F
      row[5]  = G
      row[6]  = H
      row[7]  = I
      row[8]  = J
      row[9]  = K
      row[10] = L
      row[11] = M
      row[12] = N
      row[13] = O
      row[14] = P
      row[15] = Q
      row[16] = R
      row[17] = S
    */
    const logData = logSheet.getRange(20, 2, lastRowLog - 19, 18).getValues();


    logData.forEach(row => {
      const exitDate = row[8] ? new Date(row[8]) : null;
      const status = row[15];
      const pL = row[13];
      const remarks = row[17];


      if (exitDate && exitDate >= sevenDaysAgo && status === "Closed") {
        tradeHistory.push({
          p_l: pL,
          remarks: remarks || ""
        });
      }
    });
  }


 // =========================
// 3. CAREER / APPLICATION SUMMARY
// =========================
const appData = appSheet.getDataRange().getValues();


let weeklyApplications = [];


appData.slice(1).forEach(row => {
  const dateApplied = row[3] ? new Date(row[3]) : null;
  const status = row[7] ? row[7].toString().trim() : "";
  const company = row[0];
  const jobTitle = row[1];


  if (dateApplied && dateApplied >= sevenDaysAgo && dateApplied <= today) {
    weeklyApplications.push({
      company: company || "",
      job_title: jobTitle || "",
      date_applied: dateApplied,
      status: status || ""
    });
  }
});


// Build career summary directly from weeklyApplications
const careerSummary = {
  total_records: weeklyApplications.length,
  applied: weeklyApplications.filter(app =>
    app.status.toLowerCase() === "applied"
  ).length,
  saved: weeklyApplications.filter(app =>
    app.status.toLowerCase() === "saved"
  ).length,
  rejected: weeklyApplications.filter(app =>
    app.status.toLowerCase() === "rejected"
  ).length,
  interview: weeklyApplications.filter(app =>
    app.status.toLowerCase().includes("interview")
  ).length
};


const weeklyData = `
Habit Summary:
${JSON.stringify(habitSummary, null, 2)}


Habit Notes / Reasons Context:
${JSON.stringify(habitNotesContext, null, 2)}


Trading History:
${JSON.stringify(tradeHistory, null, 2)}


Career / Application Summary:
${JSON.stringify(careerSummary, null, 2)}


Weekly Applications:
${JSON.stringify(weeklyApplications, null, 2)}
`;
  // =========================
  // 4. BUILD PROMPT
  // =========================
 const prompt = `
You are an AI performance analyst for my Personal Productivity & Trading Automation OS.


Analyze the weekly data from:
1. Trading_Log
2. Daily_DB
3. Applications


Your task is to generate a structured weekly audit.


Use this exact format:


WEEKLY AI AUDIT


1. WEEKLY SUMMARY
Summarize the overall weekly performance in 3-5 sentences.


2. TRADING PERFORMANCE REVIEW
Analyze trading activity, win/loss pattern, risk behavior, and any visible issue from the trading data.


3. HABIT CONSISTENCY REVIEW
Analyze habit consistency, including gym, learning, journaling, sleep, or other available habit data.
Use Habit Notes / Reasons Context before judging discipline or consistency.
Pay special attention to notes for incomplete habits. If an incomplete habit has a valid reason such as "no valid trading setup", illness, schedule conflict, or another clear constraint, do not automatically treat it as poor discipline. Distinguish between avoidable misses and justified skips.


4. JOB APPLICATION PROGRESS REVIEW
Analyze job application progress, consistency, and follow-up needs.


5. KEY PROBLEMS
List the 3 most important problems or weak points from this week.


6. RECOMMENDED ACTION PLAN
Give 3-5 specific actions for next week. Make them practical and measurable.


7. PRIORITY FOR NEXT WEEK
State the single most important focus for next week.


Important rules:
- Be specific and practical.
- Do not give generic motivation.
- Use only the data provided in Weekly data.
- Consider Daily_DB notes/reasons before judging habit consistency.
- Use the exact numbers provided in Career / Application Summary.
- Do not invent, estimate, or assume application totals.
- If Weekly Applications is empty, say no application records were found in the selected period.
- If Trading History is empty, say no closed trade data was found in the selected period.
- Do not conclude that no trading activity happened unless the data explicitly shows that.
- If data is missing, say which data is missing.
- Do not invent numbers that are not available in the data.
- Focus on action and improvement.
- In Career / Application Summary, total_records means all records found in the selected period, while applied means applications actually submitted.


Weekly data:
${weeklyData}
`;
  // =========================
  // 5. CALL GEMINI API
  // =========================
  const url =
  "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" +
  GEMINI_API_KEY;


  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt
          }
        ]
      }
    ]
  };


  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };


  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const resText = response.getContentText();
    let json;


    try {
      json = JSON.parse(resText);
    } catch (parseError) {
      Logger.log("Gemini response bukan JSON valid. Response: " + resText);
      writeSystemLog(
        "Weekly AI Audit",
        "Parse Gemini Response",
        "Error",
        "Gemini response was not valid JSON. HTTP status: " + responseCode + ". Parse error: " + parseError.message + ". Raw response: " + resText
      );
      return {
        success: false,
        error: "Audit gagal dibuat karena respons AI belum valid. Silakan coba lagi nanti."
      };
    }


    if (responseCode < 200 || responseCode >= 300) {
      const apiMessage = json && json.error && json.error.message
        ? json.error.message
        : "No API error message returned.";


      Logger.log("Gemini API error. HTTP status: " + responseCode + ". Response: " + resText);
      writeSystemLog(
        "Weekly AI Audit",
        "Call Gemini API",
        "Error",
        "Gemini API returned HTTP status " + responseCode + ". API message: " + apiMessage + ". Raw response: " + resText
      );
      return {
        success: false,
        error: "Audit gagal dibuat karena layanan AI belum merespons dengan benar. Silakan coba lagi nanti."
      };
    }


    // =========================
    // 6. GEMINI RESPONSE VALIDATION
    // =========================
    if (!json.candidates || !json.candidates[0]) {
      Logger.log("Gemini tidak mengembalikan candidates. Response: " + resText);
      writeSystemLog(
        "Weekly AI Audit",
        "Validate Gemini Response",
        "Error",
        "Gemini returned no candidates. Raw response: " + resText
      );
      return {
        success: false,
        error: "Audit gagal dibuat karena respons AI belum valid. Silakan coba lagi nanti."
      };
    }


    if (
      !json.candidates[0].content ||
      !json.candidates[0].content.parts ||
      !json.candidates[0].content.parts[0] ||
      !json.candidates[0].content.parts[0].text
    ) {
      Logger.log("Format response Gemini tidak sesuai. Response: " + resText);
      writeSystemLog(
        "Weekly AI Audit",
        "Validate Gemini Response",
        "Error",
        "Gemini response format was malformed. Raw response: " + resText
      );
      return {
        success: false,
        error: "Audit gagal dibuat karena format respons AI belum valid. Silakan coba lagi nanti."
      };
    }


    const aiResponse = json.candidates[0].content.parts[0].text;


// =========================
// 7. SAVE AI RESPONSE TO AI_AUDIT SHEET
// =========================
auditSheet.appendRow([
  new Date(),
  aiResponse
]);


    const lastRow = auditSheet.getLastRow();


    auditSheet
  .getRange(lastRow, 2, 1, 1)
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setVerticalAlignment("top");


    Logger.log("AI Audit berhasil disimpan di row " + lastRow);
writeSystemLog(
  "Weekly AI Audit",
  "Generate Weekly Audit",
  "Success",
  "Weekly AI Audit generated successfully at row " + lastRow
);
    return {
      success: true,
      text: aiResponse,
      row: lastRow
    };
   } catch (e) {
    Logger.log("Terjadi kesalahan saat menjalankan AI Audit: " + e.message);


    writeSystemLog(
      "Weekly AI Audit",
      "Generate Weekly Audit",
      "Error",
      "Exception: " + e.message + (e.stack ? " | Stack: " + e.stack : "")
    );


    return {
      success: false,
      error: "Audit gagal dijalankan karena terjadi error sistem. Silakan coba lagi nanti."
    };
  }
}
```

