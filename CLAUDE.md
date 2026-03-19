# InboxForge

## Overview
Gmail inbox organizer with AI-powered categorization. PyQt6 GUI that connects via IMAP, scans headers, auto-categorizes by domain/pattern/newsletter detection, and optionally uses Claude Haiku for ambiguous emails.

## Tech Stack
- Python 3, PyQt6, imaplib, anthropic SDK (optional)
- Single-file: `inboxforge.py` with `_bootstrap()` auto-install
- Catppuccin Mocha dark theme

## Version
- v0.1.0 — Initial release

## Architecture
- **4-page stacked widget flow**: Connect → Analyze → Review → Execute
- **CategoryEngine**: Domain mapping (150+ known domains), List-Unsubscribe detection, subject pattern matching, automated sender detection, domain grouping for unknowns
- **Workers (QThread)**: ImapScanWorker (header fetch), ImapLabelWorker (label creation + COPY), AiClassifyWorker (Claude Haiku batch classification)
- **Save/Load**: JSON export of scan state for resume capability

## Key Details
- Gmail IMAP: `imap.gmail.com:993`, requires App Password
- Batch fetches headers (FROM, SUBJECT, DATE, LIST-UNSUBSCRIBE) in groups of 200
- Labels created via IMAP CREATE, applied via UID COPY
- Archive = STORE +FLAGS \Deleted + EXPUNGE on INBOX
- AI classification groups by domain, sends 30 domains per API call to Haiku
- Email table capped at 2000 rows for UI performance
- Nested categories use `/` separator (Gmail hierarchy delimiter)

## Run
```bash
python inboxforge.py
```
