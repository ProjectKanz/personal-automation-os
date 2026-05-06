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


      const dbSheet = ss.getSheetByName("Daily_DB");


      if (!dbSheet) {
        sendText(chatId, "⚠️ Error: Sheet 'Daily_DB' tidak ditemukan!");
        return;
      }


      const data = dbSheet.getDataRange().getValues();
      const today = new Date().setHours(0, 0, 0, 0);


      if (text === "/help" || text === "/list") {
        sendText(chatId, buildTelegramHelpMessage());


        return;
      }


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
    "6. `/week` - Ringkasan habit 7 hari\n" +
    "7. `/reminderstatus` - Status reminder\n" +
    "8. `/audit` - AI Audit mingguan\n" +
    "9. `/lastaudit` - Preview audit terakhir\n" +
    "10. `/note habit | alasan` - Tambah catatan habit\n" +
    "11. Ketik nama habit untuk mencentang\n\n" +
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
