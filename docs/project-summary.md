# AI Workflow Automation OS (Personal MVP)

This project is a personal workflow automation MVP that consolidates fragmented activity, image-based records, and job application tracking data into one structured reporting system.  
Built with Google Apps Script and Google Sheets, it integrates Telegram for input and Gemini for AI-assisted reporting.

### What it does

- Tracks daily activities in a structured checklist database
- Converts screenshot-based records into structured sheet logs using Gemini Vision, tested with MT5 history screenshots
- Aggregates weekly activity/trading/application data
- Generates a structured weekly AI audit report

### Why it matters

The project demonstrates practical automation design for real-world workflows where data is spread across chats, images, and spreadsheets.  
It reduces manual copy-paste, improves consistency, and creates a repeatable review process.

### Technical highlights

- Modular Google Apps Script architecture (`telegram`, `habit`, `tradingParser`, `weeklyAudit`, `utils`)
- Script Property validation for safer configuration handling
- Defensive API response handling for Gemini Vision parsing
- Git-based local development with incremental refactor and hardening phases

### Scope statement

This is a portfolio-ready **personal MVP/prototype**, not enterprise production software.  
It is intended to show implementation capability, system thinking, and disciplined iteration.
