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
    const tglTime = normalizeDateOnly_(tglDashboard).getTime();


    for (let i = 1; i < dbData.length; i++) {
      if (dbData[i][0] instanceof Date &&
          normalizeDateOnly_(dbData[i][0]).getTime() === tglTime &&
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

  const targetDate = normalizeDateOnly_(tgl);
  const targetDateString = targetDate.toDateString();


  // 2. Deteksi baris terakhir secara otomatis (berdasarkan Kolom C / Aktivitas)
  const lastRow = dashSheet.getRange("C:C").getValues().filter(String).length;

  if (lastRow < 2) {
    const message = "Sinkronisasi dibatalkan: tidak ada aktivitas di Dashboard untuk tanggal " + targetDateString + ".";
    console.log(message);
    writeSystemLog("Habit", "Log Daily To DB", "Error", message);
    return;
  }
 
  // 3. Ambil data dari baris 2 sampai baris terakhir (Kolom B sampai G)
  const dataRange = dashSheet.getRange(2, 2, lastRow - 1, 6);
  const data = dataRange.getValues();
  const rowsToSave = data
    .map(row => {
      return {
        kategori: row[0],
        aktivitas: row[1],
        status: row[2],
        note: row[5]
      };
    })
    .filter(row => row.aktivitas);

  if (rowsToSave.length === 0) {
    const message = "Sinkronisasi dibatalkan: Dashboard tidak punya aktivitas valid untuk tanggal " + targetDateString + ".";
    console.log(message);
    writeSystemLog("Habit", "Log Daily To DB", "Error", message);
    return;
  }


  // 4. Upsert ke Daily_DB tanpa menghapus riwayat checklist yang sudah tersimpan.
  const existingRowsByActivity = getDailyDbRowsByActivity_(dbSheet, targetDate);
  const rowsToAppend = [];

  rowsToSave.forEach(row => {
    const existing = existingRowsByActivity[row.aktivitas];

    if (existing) {
      const savedStatus = existing.values[3] === true;
      const dashboardStatus = row.status === true;
      const savedNote = existing.values[4] ? existing.values[4].toString().trim() : "";
      const dashboardNote = row.note ? row.note.toString().trim() : "";

      dbSheet
        .getRange(existing.rowIndex, 1, 1, 5)
        .setValues([[
          targetDate,
          row.kategori,
          row.aktivitas,
          savedStatus || dashboardStatus,
          dashboardNote || savedNote
        ]]);
      dbSheet.getRange(existing.rowIndex, 4).insertCheckboxes();
      return;
    }

    rowsToAppend.push([targetDate, row.kategori, row.aktivitas, row.status === true, row.note || ""]);
  });

  if (rowsToAppend.length > 0) {
    const startRow = dbSheet.getLastRow() + 1;
    dbSheet
      .getRange(startRow, 1, rowsToAppend.length, 5)
      .setValues(rowsToAppend);
    dbSheet
      .getRange(startRow, 4, rowsToAppend.length, 1)
      .insertCheckboxes();
  }

  const message = "Sinkronisasi Selesai! Data tanggal " + targetDateString + " sudah di-upsert ke Daily_DB tanpa menghapus checklist lama.";
  console.log(message);
  writeSystemLog("Habit", "Log Daily To DB", "Success", message);
}


function normalizeDateOnly_(dateValue) {
  const normalizedDate = new Date(dateValue);
  normalizedDate.setHours(0, 0, 0, 0);
  return normalizedDate;
}


function getDailyDbRowsByActivity_(dbSheet, targetDate) {
  const targetTime = normalizeDateOnly_(targetDate).getTime();
  const dbData = dbSheet.getDataRange().getValues();
  const rowsByActivity = {};


  for (let i = 1; i < dbData.length; i++) {
    const rowDate = dbData[i][0];
    const activity = dbData[i][2];


    if (!(rowDate instanceof Date) || !activity) continue;
    if (normalizeDateOnly_(rowDate).getTime() !== targetTime) continue;


    rowsByActivity[activity] = {
      rowIndex: i + 1,
      values: dbData[i]
    };
  }


  return rowsByActivity;
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
