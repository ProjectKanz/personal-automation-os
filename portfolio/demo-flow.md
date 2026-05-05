# Demo Flow - AI Workflow Automation OS (Personal MVP)

Simple walkthrough of the end-to-end flow:

## 1) Telegram Input

- User sends either:
  - text command/update (daily activity, help, checklist, and status flow), or
  - screenshot-based record (tested with MT5 history screenshots).

## 2) Google Apps Script Processing

- Webhook receives the message via `doPost(e)`.
- Script routes input by type:
  - text path for daily activity/status actions (`/help`, `/list`, `/today`, `/status`, `/audit`, or habit name),
  - image path for Gemini Vision extraction.
- Basic validation and error handling are applied before write operations.

## 3) Google Sheets as Structured Store

- Parsed/normalized records are written to the relevant sheets.
- Daily activity, extracted records, and application tracking data are maintained in structured tabular format.

## 4) Gemini / Gemini Vision Layer

- **Gemini Vision:** converts screenshot content into structured fields.
- **Gemini (text):** analyzes weekly aggregated data and generates a structured weekly report.

## 5) Weekly AI Report Output

- Weekly summary is saved to the audit sheet for review.
- When triggered from Telegram, `/audit` returns a completion/error message and a short preview while keeping the full report in `AI_Audit`.
- Output supports reflection and planning across:
  - daily activity consistency,
  - habit notes/reasons from `Daily_DB`,
  - screenshot-based record quality/completeness,
  - job application progress.

## Demo Talk Track (60-90 seconds)

"This personal MVP shows an AI-assisted workflow automation loop. Input starts in Telegram, Apps Script validates and routes it, then structured data is stored in Google Sheets. Gemini Vision handles screenshot-to-structured-data extraction, and Gemini generates a weekly report from aggregated records. The focus is practical automation and reporting consistency, not enterprise-scale deployment."

## V2.2 Telegram Demo Notes

- `/help` or `/list`: shows the command menu.
- `/today`: displays today's checklist from `Daily_DB`.
- `/status`: shows completed habits, total habits, and completion percentage.
- Typing a habit name still checks off the matching habit for today.

## V2.3 Weekly Audit Demo Notes

- `/audit`: starts the weekly AI audit, then sends success/error feedback in Telegram.
- Successful audits include a short Telegram preview and save the full report to `AI_Audit`.
- Habit notes from `Daily_DB` column E are included as context, especially for incomplete habits.
