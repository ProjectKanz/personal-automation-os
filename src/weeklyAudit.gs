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
    return;
  }


  if (!logSheet) {
    Logger.log("Sheet Log tidak ditemukan.");
    return;
  }


  if (!appSheet) {
    Logger.log("Sheet Applications tidak ditemukan.");
    return;
  }


  if (!auditSheet) {
    Logger.log("Sheet AI_Audit tidak ditemukan.");
    return;
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


  habitData.slice(1).forEach(row => {
    const habitDate = row[0] ? new Date(row[0]) : null;
    const habitName = row[2] || "Unnamed Habit";
    const isDone = row[3] === true;


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
    const resText = response.getContentText();
    const json = JSON.parse(resText);


    // =========================
    // 6. GEMINI RESPONSE VALIDATION
    // =========================
    if (!json.candidates || !json.candidates[0]) {
      Logger.log("Gemini tidak mengembalikan candidates. Response: " + resText);
      auditSheet.appendRow([
        new Date(),
        "Gemini response error",
        "Tidak ada candidates dari Gemini.",
        resText
      ]);
      return;
    }


    if (
      !json.candidates[0].content ||
      !json.candidates[0].content.parts ||
      !json.candidates[0].content.parts[0] ||
      !json.candidates[0].content.parts[0].text
    ) {
      Logger.log("Format response Gemini tidak sesuai. Response: " + resText);
      auditSheet.appendRow([
        new Date(),
        "Gemini format error",
        "Format response Gemini tidak sesuai.",
        resText
      ]);
      return;
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
   } catch (e) {
    Logger.log("Terjadi kesalahan saat menjalankan AI Audit: " + e.message);


    auditSheet.appendRow([
      new Date(),
      "Script error",
      e.message,
      "Cek Apps Script Logs untuk detail error."
    ]);


    writeSystemLog(
      "Weekly AI Audit",
      "Generate Weekly Audit",
      "Error",
      e.message
    );
  }
}
