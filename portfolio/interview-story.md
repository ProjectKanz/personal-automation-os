# Interview Story (STAR) - AI Workflow Automation OS

## Situation

I had multiple personal workflows running in parallel: daily activity tracking, screenshot-based records, and job application updates.  
The data lived across chat messages, screenshots, and separate sheets, which made weekly review slow and inconsistent.

## Task

My goal was to build a single automation layer that could:

- capture inputs quickly,
- standardize data into Google Sheets,
- and generate a weekly AI-assisted report.

I also wanted the project to be portfolio-ready and explainable in interviews, while staying honest as a personal MVP/prototype.

## Action

- Built the core system using Google Apps Script + Google Sheets, with Telegram as the input channel.
- Implemented screenshot-based extraction using Gemini Vision (tested with MT5 history screenshots) and wrote normalized results into logs.
- Added a weekly aggregation + Gemini reporting function to generate structured weekly review output.
- Refactored from a monolithic script into modular files (`telegram`, `habit`, `tradingParser`, `weeklyAudit`, `utils`) in V2.1-A.
- Added small hardening changes in V2.1-B:
  - required Script Property helper for key configuration values,
  - defensive Gemini response checks for clearer failure handling.
- Packaged project-facing documentation in V2.1-C for portfolio and interview clarity.

## Result

- I now have a working personal automation MVP that connects input capture, data structuring, and AI-generated weekly reporting in one flow.
- The codebase is cleaner to maintain after modularization and easier to explain to non-technical stakeholders.
- The project demonstrates practical workflow automation, integration design, and incremental improvement discipline.
- I do not position it as enterprise-ready software; I position it as a realistic prototype with clear next-step pathways.

## What I Learned

- Fast MVPs benefit from early modular boundaries, even in small Apps Script projects.
- AI-assisted extraction/reporting is useful when paired with structured logging and clear validation points.
- Documentation quality significantly improves portfolio communication and interview outcomes.
