# AI Workflow Automation OS (Personal MVP)

This project is a personal workflow automation MVP and portfolio case study that consolidates fragmented activity, image-based records, job application tracking, and weekly review data into one structured reporting system.
Built with Google Apps Script and Google Sheets, it integrates Telegram for input, Gemini for AI-assisted reporting, and a V3.4 web command center for review.

### What it does

- Tracks daily activities in a structured checklist database
- Converts screenshot-based records into structured sheet logs using Gemini Vision, tested with MT5 history screenshots
- Aggregates weekly activity/trading/application data
- Generates a structured weekly AI audit report
- Provides a V3.4 web command center for Habit, Career, Trading, and AI Audit review

### Why it matters

The project demonstrates practical automation design for real-world workflows where data is spread across chats, images, and spreadsheets.  
It reduces manual copy-paste, improves consistency, and creates a repeatable review process.

### Technical highlights

- Modular Google Apps Script architecture (`telegram`, `habit`, `career`, `webapp`, `tradingParser`, `weeklyAudit`, `utils`)
- Script Property validation for safer configuration handling
- Defensive API response handling for Gemini Vision parsing
- Git/clasp-based local development with incremental refactor and hardening phases

### Scope statement

This is a portfolio-ready **personal MVP/prototype**, not enterprise production software.  
It is intended to show implementation capability, system thinking, and disciplined iteration. Public demos should use mock or sanitized data only.
