# InboxForge

## Overview
Full Gmail mailbox downloader & AI-powered organizer. PyQt6 GUI that connects via IMAP, downloads all folders as .eml files preserving original structure, auto-categorizes by domain/pattern/newsletter detection, and optionally uses Claude Haiku for ambiguous emails. Organizes locally without touching the live mailbox.

## Tech Stack
- Python 3, PyQt6, imaplib, anthropic SDK (optional)
- Single-file: `inboxforge.py` with `_bootstrap()` auto-install
- Catppuccin Mocha dark theme

## Version
- v0.3.0 — Full mailbox download with folder preservation, dedup, resume

## Architecture
- **5-page stacked widget flow**: Connect -> Download -> Analyze -> Review -> Execute
- **CategoryEngine**: Domain mapping (150+ known domains), List-Unsubscribe detection, subject pattern matching, automated sender detection, domain grouping for unknowns
- **Workers (QThread)**: ImapScanWorker (header-only fast scan), ImapDownloadWorker (full .eml download), ImapLabelWorker (Gmail label mode), LocalOrganizeWorker (file organization), AiClassifyWorker (Claude Haiku)
- **Save/Load**: JSON manifest for download resume, JSON export of categorization state

## Key Details
- Gmail IMAP: `imap.gmail.com:993`, requires App Password
- Downloads ALL folders: INBOX, Sent Mail, Drafts, Starred, custom labels
- Skips by default: [Gmail]/All Mail (duplicates), Spam, Trash, Important
- Deduplicates via Message-ID across Gmail labels (same email stored once)
- Local structure: `output/folders/<FolderName>/<uid>.eml` (original) + `output/organized/<Category>/` (AI-sorted)
- UIDs stored as `folder:uid` format for cross-folder uniqueness
- manifest.json tracks per-folder downloads for resume
- Batch fetches: 50 emails per IMAP FETCH for downloads, 200 for header scans
- Email table capped at 2000 rows for UI performance
- AI classification groups by domain, sends 30 domains per API call to Haiku

## Run
```bash
python inboxforge.py
```

## Gotchas
- Gmail IMAP folder names may contain `[Gmail]/` prefix — sanitized for local paths
- imaplib requires folder names in quotes for SELECT/COPY
- IMAP UIDs are per-folder, not global — composite key needed
- Gmail COPY to label = applying that label (not moving)
- Archiving from Inbox = STORE +FLAGS \Deleted + EXPUNGE
