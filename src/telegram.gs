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
      const text = msg.text.toLowerCase().trim();


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


      if (text === "/status") {
        sendText(chatId, buildHabitStatusMessage(data, today));
        return;
      }


      if (text === "/audit") {
        sendText(chatId, "⏳ Menjalankan Audit... Tunggu sebentar.");
        runWeeklyAIAudit();
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
    "3. `/status` - Ringkasan progres habit\n" +
    "4. `/audit` - AI Audit mingguan\n" +
    "5. Ketik nama habit untuk mencentang\n\n" +
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
    return "📋 Belum ada habit untuk hari ini di Daily_DB.";
  }


  const checklist = todayRows.map(row => {
    const marker = row[3] === true ? "✅" : "⬜";
    const habitName = row[2] || "(Tanpa nama habit)";
    return marker + " " + habitName;
  });


  return "📋 *Habit Hari Ini:*\n" + checklist.join("\n");
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


function sendText(id, text) {
  const url = "https://api.telegram.org/bot" + TOKEN + "/sendMessage?chat_id=" + id + "&text=" + encodeURIComponent(text) + "&parse_mode=Markdown";
  UrlFetchApp.fetch(url);
}
