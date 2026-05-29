# AI Workflow Automation OS

Personal MVP/prototype for daily activity tracking, screenshot-based data extraction, job application monitoring, and AI-generated weekly reporting.

## Overview

This repository contains a personal prototype focused on workflow reliability and practical automation.  
The system reduces manual logging by combining:

- Telegram-driven input for quick updates
- Sheet-based structured storage for traceability
- Gemini-powered analysis for weekly review output
- Local Git workflow for iterative development and refactoring

The goal is not enterprise deployment. The goal is to demonstrate end-to-end automation design, modular scripting practices, and AI-assisted reporting in a realistic single-operator context.

## Problem It Solves

Many individual workflows are split across chat messages, screenshots, and multiple spreadsheets.  
This project provides one automation layer to:

- capture data from different sources,
- normalize and write it into structured sheets,
- and generate recurring, actionable summaries.

## Core Capabilities

- **V3.4 Web Command Center**
  - Apps Script web dashboard with sidebar navigation
  - Home, Habit, Career, Trading, and AI Audit views
  - Career pipeline dashboard with sector mix visibility
  - Trading dashboard with custom date range filtering
  - Trading money condition cumulative P/L chart
  - AI Audit decision board for weekly diagnosis, current priority, next actions, and audit evidence

- **Daily activity tracking**
  - Generate daily checklist entries from a master habit sheet
  - Sync checklist status between dashboard and database sheets
  - Update status, notes, and daily review context from Telegram commands
  - Review unfinished habits, missing notes, and daily completion status before closing the day
  - Send optional daily reminders for unfinished habits

- **Screenshot-to-database extraction module**
  - Accept screenshot-based records from Telegram
  - Use Gemini Vision to extract key trade fields
  - Write normalized trade data into a structured log sheet
  - Mark extraction status for manual review when needed
  - MT5 history screenshots are used as one test/source example

- **Weekly AI audit generation**
  - Aggregate recent habit, trading, and application data
  - Include habit notes/reasons from `Daily_DB` so skipped or incomplete habits have context
  - Build a structured weekly prompt
  - Generate and store an AI performance audit in a dedicated sheet
  - Retrieve the latest saved audit from Telegram without calling Gemini again

- **Job application tracking support**
  - Include application volume and status distribution in weekly reporting
  - Surface progress and follow-up context in generated audits

## Workflow Loop

The current MVP supports a practical review loop:

`daily tracking -> reminder -> note capture -> closing review -> weekly snapshot -> AI audit -> web dashboard review`

This keeps the daily habit record useful for the weekly audit without adding a complex product layer.

## Technology Stack

- Google Apps Script (execution runtime and orchestration)
- Google Sheets (data store and operational interface)
- Telegram Bot API (chat-based interaction layer)
- Gemini API (text analysis for weekly audits)
- Gemini Vision (image-to-structured-data extraction, tested with MT5 history screenshots)
- Git (local version control and modular refactor workflow)

## V3.4 Feature Summary

V3.4 adds a working Apps Script web dashboard layer on top of the existing Telegram and Google Sheets workflow. The web dashboard is a personal command center for reviewing habits, career applications, trading performance, and AI audit output in one place.

The AI Audit tab now works as a weekly decision board instead of a raw text preview, with Weekly Diagnosis, Current Priority, Next 3 Actions, and Audit Evidence panels.

## Current Architecture (V3.4 Web Command Center)

Apps Script modules are organized by responsibility:

- `src/telegram.gs` - Telegram webhook, message handling, bot responses
- `src/habit.gs` - daily checklist generation and habit sync logic
- `src/career.gs` - career pipeline data support and application intelligence
- `src/webapp.gs` - web dashboard data support, including trading dashboard calculations
- `src/tradingParser.gs` - Gemini Vision trade extraction
- `src/weeklyAudit.gs` - weekly data aggregation and audit generation
- `src/utils.gs` - shared utilities and script property helpers
- `src/Index.html` - Apps Script web dashboard UI for the V3.4 command center

## Skills Demonstrated

- End-to-end automation workflow design
- Telegram + Google Apps Script integration
- Apps Script HTML web dashboard development
- Gemini Vision image-to-structured-data extraction
- Weekly AI-generated reporting
- Modular Apps Script refactor
- Script Properties security cleanup
- Telegram command design for daily/weekly visibility
- Lightweight reminder and review workflow design
- Git checkpoint workflow

## MVP Status and Scope

This project is a **personal MVP/prototype**.

- It is intentionally lightweight and optimized for one-user workflow automation.
- It is suitable for demonstrating architecture, integration, and iteration discipline.
- It is not positioned as enterprise-grade production software.
- It does not claim guaranteed AI accuracy.

## Configuration Notes

Set required Script Properties in Apps Script:

- `GEMINI_API_KEY`
- `TELEGRAM_TOKEN`
- `TELEGRAM_USER_ID`
- `WEB_APP_URL`
- `WEB_APP_MT_CV_FOLDER_ID`
- `WEB_APP_DATA_CV_FOLDER_ID`

If a required property is missing, the script now fails with a clear error message (V2.1-B hardening).

## Setup Notes

1. Copy `.clasp.example.json` to `.clasp.json` locally and add your own Apps Script project ID.
2. Keep `.clasp.json` untracked because it identifies a real Apps Script deployment.
3. Configure the required Script Properties in the Apps Script editor.
4. Prepare the expected Google Sheets tabs, including `Daily_DB`, `Master_Habit`, `Log`, `Applications`, `AI_Audit`, and `System_Log`.
5. Deploy the Apps Script web app for personal access and configure the Telegram webhook only in your private environment.

## Screenshots and Demo

Screenshots should use anonymized or mock data only.

- Web command center home
- Habit checklist view
- Career pipeline view with mock company names
- Trading dashboard with sample/non-private records
- AI Audit view with sanitized generated text
- Telegram command demo with fake messages

See `portfolio/mock-demo-checklist.md` for a simple publication demo checklist.

## Security and Privacy Notes

- Do not commit API keys, Telegram tokens, webhook URLs, spreadsheet URLs, Drive folder IDs, CV links, or private operational data.
- Store secrets and deployment-specific IDs in Apps Script Properties, not in source files.
- Uploaded CV files are kept private by default. Share only sanitized demo files manually when needed.
- Public screenshots and videos should use fake names, fake applications, fake trading values, and sanitized audit text.

## Disclaimer

This project is a personal MVP and portfolio case study. It is not financial advice, not trading automation, and not enterprise production software.

## Repository Documents

- `docs/project-summary.md` - concise portfolio summary
- `docs/changelog.md` - version history across V1 through V3.4

## License

No license file is currently included. Add one before public distribution if needed.
