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
