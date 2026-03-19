# InboxForge

## Overview
Full Gmail mailbox downloader, AI-powered organizer, and analytics suite. PyQt6 GUI with 10 power features: statistics dashboard, group-by views, subscription manager, feedback loop, auto clean rules, attachment extraction, CSV/JSON export, Gmail filter import, sensitive content scanner, and AI thread summarization.

## Tech Stack
- Python 3, PyQt6, imaplib, anthropic SDK (optional)
- Single-file: `inboxforge.py` with `_bootstrap()` auto-install
- Catppuccin Mocha dark theme

## Version
- v1.0.0 — Full power release with all 10 features

## Architecture
- **5-page flow**: Connect -> Download -> Analyze -> Review -> Execute
- **CategoryEngine**: Domain mapping (150+), List-Unsubscribe + newsletter platform detection, subject patterns, learned rules (feedback loop), clean rules engine, subscription detection, thread building, sensitive scanning
- **Workers (QThread)**: ImapScanWorker, ImapDownloadWorker, ImapLabelWorker, LocalOrganizeWorker, AttachmentExtractWorker, SensitiveScanWorker, ThreadSummaryWorker, AiClassifyWorker
- **Dialogs**: StatsDialog (charts + heatmap), SubscriptionDialog (unsubscribe links), RulesEditorDialog (CRUD + Gmail filter import), ThreadSummaryDialog
- **Custom Widgets**: HBarChart, ActivityHeatmap (QPainter-based, no external deps)

## Features
1. **Stats Dashboard** — Emails/month bar chart, activity heatmap (day x hour), top senders/domains, category distribution, storage breakdown
2. **Group-by Views** — Tree view switches between Category, Sender Domain, Sender, Source Folder grouping
3. **Subscription Scanner** — Detects newsletters via List-Unsubscribe + 25+ platform domains, shows frequency, one-click unsubscribe (opens URL in browser)
4. **Feedback Loop** — Moving emails teaches the engine (domain->category saved to learned_rules.json). Applied first on next run.
5. **Auto Clean Rules** — Persistent rules (domain, sender, subject, age, newsletter flag) -> categorize/flag/skip. Gmail filter XML import.
6. **Attachment Extraction** — Extracts from .eml files, deduplicates by SHA-256, organizes by category
7. **CSV/JSON Export** — Full metadata export (date, from, subject, category, confidence, folder, size, newsletter, sensitive flags)
8. **Gmail Filter Import** — Parse Gmail's Atom XML filter export into CleanRules
9. **Sensitive Content Scanner** — Regex detection of SSN, credit cards, passwords, API keys, GitHub tokens
10. **Thread Summarization** — Reconstructs threads via In-Reply-To/References headers, sends to Claude Haiku for 2-3 sentence summaries

## Key Files
- `inboxforge.py` — Single-file app
- `learned_rules.json` — Feedback loop persistence (in download dir)
- `clean_rules.json` — Auto clean rules (in download dir)
- `manifest.json` — Download resume tracking (in download dir)

## Run
```bash
python inboxforge.py
```

## Gotchas
- Gmail IMAP `[Gmail]/` prefix sanitized for local paths
- IMAP UIDs are per-folder — composite `folder:uid` key used
- Message-ID dedup prevents redundant .eml downloads across Gmail labels
- Chart widgets use QPainter (no matplotlib dependency)
- Sensitive scanner reads first 50KB of each .eml for performance
- Thread summarization limited to top 50 longest threads, 10 emails each
