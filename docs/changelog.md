# Changelog

All notable changes to this project are documented here.

---

## V3.4 (May 2026) - Web Command Center

### Added

- Apps Script web dashboard with sidebar navigation.
- Home, Habit, Career, Trading, and AI Audit views.
- Career pipeline dashboard with sector mix visibility.
- Trading custom date range filter.
- Trading money condition cumulative P/L chart.
- AI Audit decision board with Weekly Diagnosis, Current Priority, Next 3 Actions, and Audit Evidence.
- `SYSTEM_STATE` metadata for the active product version and phase.
- `VERSION_HISTORY` scope for V3.0-V3.4 evidence packaging.
- `generateLinkedInCampaign()` to create LinkedIn hook, carousel blueprint, video idea, and data signal from career intelligence metrics.
- `prepLookerExport()` to refresh a hidden `Looker_Data` sheet for Looker Studio.
- `generateGitHubREADME()` to draft technical proof documentation for the V3 architecture.
- `/share` Telegram command for generating a Management & Marketing style campaign package.

### Notes

- V3.4 Web Command Center is working as a personal MVP/prototype dashboard.
- Telegram + Google Sheets workflow remains part of the system.
- Web dashboard layer does not change the positioning of this project as a single-operator prototype.
- `/share` only highlights V3.0-V3.4 career intelligence features.
- Output emphasizes Operational Excellence and AI Engineering.
- No legacy V2 features are included in the campaign package.
- Campaign story follows the V3 evolution arc: Foundation, Workflow, and Intelligence.
- `/share` refreshes portfolio-facing summary data and includes documentation prompts for README proof.
- CV tagging now supports inferred MT tags, Automation-over-Data priority, and more robust case-insensitive filename matching.
- Corporate intelligence now decodes recruitment abbreviations such as ULFP, ODP, BDP, MDP, CCEP, GTP, and MT for industry and CV analysis.
- Career Advisor status logic now separates Applied, Assessment Stage, and Interview while excluding Saved/Rejected from CV performance rates.
- Career Advisor memo now includes funnel reality, CV signal quality, richer cold-lead actions, and legacy context from `Copy of Applications`.

---

## V3.3 (May 2026) - AI Career Advisor

### Added

- `/careercoach` Telegram command for an executive career strategy memo.
- CV analytics engine that tags `CV VERSION` with MT, ODP, Data, Automation, Analyst, Business, or General.
- Conversion-rate analysis for CV tags using Assessment/Interview as positive response signals.
- Stagnancy detection for Applied applications older than 7 days, with Yellow Zone and Red Zone labels.
- Diversification audit by sector, including Banking, Automotive/Conglomerate, Consulting/Data, FMCG, Technology/Digital, and Other.
- Expanded industry recognition for Tech/Digital, FMCG/Retail, Banking/Finance, Consulting/Data, and Logistics/Transport.
- Fallback role-based categorization using Job Title and Notes when company names do not match known sectors.

### Notes

- Robust when `CV VERSION` is empty by assigning the General tag.
- No new sheet schema changes.
- No new AI API call introduced; this is rules-based advisory logic.

---

## V2.9 (May 2026) - Daily Closing Review

### Added

- `/close` Telegram command for end-of-day habit review.
- Closing review shows daily completion count, total count, and completion percentage.
- Closing review highlights unfinished habits that still need notes in `Daily_DB` column E.
- Closing review shows brief existing notes for unfinished habits.

### Notes

- Supports better weekly audit context by encouraging reasons before the day ends.
- No Google Sheets schema changes.
- No weekly audit logic changes.
- No MT5 screenshot parsing changes.
- No new AI calls introduced.

---

## V2.8 (May 2026) - Spreadsheet UI / Visual Dashboard Layer

### Added

- `applyPortfolioSheetTheme()` utility for applying a clean visual theme to selected spreadsheet interface sheets.
- Dashboard, Applications, and AI_Audit styling support for portfolio/demo readability.

### Notes

- Operational sheets such as `Log`, `Daily_DB`, `System_Log`, `Master_Habit`, `Helper`, `Lists`, and `Setup` are intentionally excluded.
- No data, formula, schema, or Telegram command changes.
- No new AI calls introduced.

---

## V2.7 (May 2026) - Daily and Weekly Visibility Pack

### Added

- `/daily` command for today's habit completion summary, unfinished habits, and note count.
- `/week` command for last-7-days habit completion summary, top incomplete habits, and weekly note count.
- `/reminderstatus` command to check whether the daily habit reminder trigger is active.

### Notes

- Commands read from `Daily_DB` and Apps Script trigger metadata only.
- No AI_Audit reads in `/week`.
- No new AI calls introduced.

---

## V2.6 (May 2026) - Last Audit Access

### Added

- `/lastaudit` command to retrieve a safe-length preview of the latest saved weekly audit from `AI_Audit`.

### Notes

- Does not call Gemini.
- Reads only saved audit output from the sheet.
- Full report remains available in `AI_Audit`.

---

## V2.5 (May 2026) - Habit Note Capture

### Added

- `/note habit name | reason` command for adding or updating today's habit note in `Daily_DB` column E.
- Partial habit-name matching, with clear no-match and multiple-match responses.

### Notes

- Habit notes improve weekly audit context without changing the sheet schema.
- No new AI calls introduced.

---

## V2.4 (May 2026) - Daily Habit Reminder System

### Added

- `/missing` command to show today's unfinished habits.
- `sendDailyHabitReminder()` to send a daily Telegram reminder for unfinished habits.
- `setupDailyHabitReminderTrigger()` to create a daily time-driven reminder trigger.
- `deleteDailyHabitReminderTriggers()` to avoid duplicate reminder triggers.

### Notes

- Reminder reads from `Daily_DB` only.
- No Google Sheets schema changes.
- No AI calls introduced.

---

## V2.3 (May 2026) - Weekly Audit UX + Context Upgrade

### Added

- Telegram `/audit` now sends a completion message after successful weekly audit generation.
- Telegram `/audit` now sends a safe-length audit preview after success.
- Telegram `/audit` now sends a clear error message if audit generation fails.
- Weekly audit context now includes `Daily_DB` column E notes/reasons.

### Changed

- `runWeeklyAIAudit()` now returns a structured success/error result for Telegram notification handling.
- Gemini weekly audit prompt now considers habit notes before judging habit consistency or discipline.

### Notes

- Full weekly audit reports are still saved to `AI_Audit`.
- No Google Sheets schema changes.
- No MT5 screenshot parsing changes.
- No reminder trigger added.

---

## V2.2 (May 2026) - Telegram Usability Upgrade

### Added

- `/help` command for Telegram bot usage guidance.
- `/today` command to show the current day's habit checklist from `Daily_DB`.
- `/list` remains available as an alias for `/help`.

### Changed

- Improved `/status` response to include completed count, total count, and completion percentage.

### Notes

- No Google Sheets schema changes.
- No MT5 screenshot parsing changes.
- No weekly audit logic changes.
- No new AI calls introduced.

---

## V2.1-C (May 2026) - Portfolio Documentation Packaging

### Changed

- Updated portfolio-facing documentation:
  - `README.md`
  - `docs/project-summary.md`
  - `docs/changelog.md`
- Reframed project positioning as an AI-assisted workflow automation system.
- Clarified MVP scope language for recruiter/interviewer readability.

### Notes

- Documentation-only release.
- No source code changes in this version.

---

## V2.1-B (May 2026) - Security and Error-Handling Cleanup

### Added

- `getRequiredProperty(propertyName)` helper in `src/utils.gs` for required Script Property validation.

### Changed

- Script properties now resolved through required-property helper for:
  - `GEMINI_API_KEY`
  - `TELEGRAM_TOKEN`
  - `TELEGRAM_USER_ID`
  - `WEB_APP_URL`
- `setWebhook()` now uses `WEB_APP_URL` from Script Properties (removed hardcoded placeholder dependency).
- `callGeminiVision()` now includes defensive checks for missing Gemini response paths (`candidates`, `content`, `parts`, `text`) and throws clear errors.

### Notes

- No sheet schema changes.
- No Telegram command behavior changes.
- No feature expansion; this phase focused only on safer configuration and clearer failure paths.

---

## V2.1-A (May 2026) - Modular Split

### Changed

- Split monolithic `src/main.gs` into responsibility-based modules:
  - `src/telegram.gs`
  - `src/habit.gs`
  - `src/tradingParser.gs`
  - `src/weeklyAudit.gs`
  - `src/utils.gs`
- Preserved existing function names and runtime behavior for compatibility with Google Apps Script global scope.

### Notes

- Refactor objective: improve maintainability and readability without changing output behavior.
- No import/export/module system introduced (Apps Script compatible).

---

## V1 (Apr 2026) - Initial Integrated MVP

### Added

- End-to-end Google Apps Script workflow integrating:
  - Telegram bot webhook handling
  - Daily activity tracking in Google Sheets
  - Screenshot-based data extraction using Gemini Vision, tested with MT5 history screenshots
  - Weekly AI audit generation via Gemini
  - Job application tracking integration in weekly reporting
- Core helper functions for messaging, file retrieval, and system logging.

### Notes

- Implemented as a personal MVP/prototype to validate automation flow and AI-assisted reporting in a real daily workflow.
