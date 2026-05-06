# Demo Flow - AI Workflow Automation OS (Personal MVP)

Simple walkthrough of the end-to-end flow:

## 1) Telegram Input

- User sends either:
  - text command/update (daily activity, notes, closing review, reminders, visibility, and audit flow), or
  - screenshot-based record (tested with MT5 history screenshots).

## 2) Google Apps Script Processing

- Webhook receives the message via `doPost(e)`.
- Script routes input by type:
  - text path for daily activity/status actions (`/help`, `/list`, `/today`, `/missing`, `/status`, `/daily`, `/week`, `/close`, `/note`, `/audit`, `/lastaudit`, or habit name),
  - image path for Gemini Vision extraction.
- Basic validation and error handling are applied before write operations.

## 3) Google Sheets as Structured Store

- Parsed/normalized records are written to the relevant sheets.
- Daily activity, extracted records, and application tracking data are maintained in structured tabular format.
- Habit notes/reasons are stored in `Daily_DB` column E for weekly review context.

## 4) Gemini / Gemini Vision Layer

- **Gemini Vision:** converts screenshot content into structured fields.
- **Gemini (text):** analyzes weekly aggregated data and generates a structured weekly report.

## 5) Weekly AI Report Output

- Weekly summary is saved to the audit sheet for review.
- When triggered from Telegram, `/audit` returns a completion/error message and a short preview while keeping the full report in `AI_Audit`.
- `/lastaudit` retrieves the latest saved audit preview without calling Gemini again.
- Output supports reflection and planning across:
  - daily activity consistency,
  - habit notes/reasons from `Daily_DB`,
  - screenshot-based record quality/completeness,
  - job application progress.

## 6) Daily-to-Weekly Loop

The current MVP supports this loop:

`daily tracking → reminder → note capture → closing review → weekly snapshot → AI audit → audit retrieval`

## Demo Talk Track (60-90 seconds)

"This personal MVP shows an AI workflow automation loop. Input starts in Telegram, Apps Script validates and routes it, then structured data is stored in Google Sheets. Daily habit tracking now includes reminders, note capture, and a closing review so the weekly audit has better context. Gemini Vision handles screenshot-to-structured-data extraction for tested MT5 history screenshots, and Gemini generates a weekly report from aggregated records. The focus is practical automation and reporting consistency, not enterprise-scale deployment."

## V2.2 Telegram Demo Notes

- `/help` or `/list`: shows the command menu.
- `/today`: displays today's checklist from `Daily_DB`.
- `/status`: shows completed habits, total habits, and completion percentage.
- Typing a habit name still checks off the matching habit for today.

## V2.3 Weekly Audit Demo Notes

- `/audit`: starts the weekly AI audit, then sends success/error feedback in Telegram.
- Successful audits include a short Telegram preview and save the full report to `AI_Audit`.
- Habit notes from `Daily_DB` column E are included as context, especially for incomplete habits.

## V2.4-V2.9 Demo Notes

- `/missing`: shows today's unfinished habits.
- Daily reminder trigger can send unfinished habit reminders around the end of the day.
- `/note habit | reason`: adds context to today's habit record.
- `/close`: checks unfinished habits and highlights missing notes before the day ends.
- `/daily` and `/week`: provide non-AI snapshots of habit progress.
- `/lastaudit`: retrieves the latest saved audit preview without generating a new one.
- `applyPortfolioSheetTheme()`: applies a light visual layer to selected portfolio-facing sheets.
