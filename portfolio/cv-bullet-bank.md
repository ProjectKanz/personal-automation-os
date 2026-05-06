# CV Bullet Bank - AI Workflow Automation OS (Personal MVP)

Use and adapt these bullets based on role focus.

## Data Analyst / Business Analyst

- Built a personal MVP workflow automation system using Google Apps Script + Google Sheets to consolidate daily activity, screenshot-based records, and job application tracking into a structured reporting pipeline.
- Designed sheet-based data flows that reduced manual logging steps and improved consistency of weekly review data.
- Implemented a full daily-to-weekly review loop: daily tracking, reminders, note capture, closing review, weekly snapshots, AI audit generation, and saved audit retrieval.
- Added note-aware weekly reporting so incomplete habits can include context/reasons instead of being treated as raw misses.
- Created a modular script architecture (`telegram`, `habit`, `tradingParser`, `weeklyAudit`, `utils`) to improve maintainability and change control.
- Established Git-based checkpoint workflow and documentation standards to support reproducible iteration and portfolio-ready project communication.

## Operations / Process Improvement

- Mapped fragmented personal workflows (chat input, screenshots, spreadsheet logs) into one repeatable automation process.
- Integrated Telegram input with Apps Script handlers to capture updates faster and reduce process friction in daily execution.
- Implemented screenshot-to-database extraction using Gemini Vision (tested with MT5 history screenshots) to reduce manual transcription effort.
- Added daily reminder, closing review, and note capture commands to improve data completeness before weekly review.
- Added weekly and daily visibility commands to surface progress without triggering AI calls.
- Improved operational reliability through required configuration checks and clearer error handling in the automation runtime.

## AI Automation / Digital Transformation

- Developed an AI-assisted automation MVP combining Telegram Bot API, Google Apps Script, Google Sheets, Gemini API, and Gemini Vision.
- Built an end-to-end flow from input capture to structured storage to AI-generated reporting without introducing complex infrastructure.
- Applied Gemini Vision for image-to-structured-data extraction and Gemini text generation for weekly report synthesis.
- Iterated through usability and context upgrades from V2.2-V2.9 while preserving schema compatibility and existing functional behavior.
- Positioned the project as a practical personal prototype for workflow automation, not enterprise production software.

## Optional One-Line Project Entry

- **AI Workflow Automation OS (Personal MVP):** Built a Google Apps Script-based automation system that consolidates daily tracking, reminders, note capture, screenshot-based extraction, job application monitoring, and weekly AI-generated review into a structured Sheets workflow.
