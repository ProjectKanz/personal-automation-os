# Changelog

All notable changes to this project are documented here.

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
