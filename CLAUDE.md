# GmailDownloader

## Overview
Full Gmail mailbox downloader, AI-powered organizer, and analytics suite. PyQt6 GUI with 16+ features including email preview, search/filter, statistics dashboard, contact analysis, subscription manager, HTML archive export, and more.

## Tech Stack
- Python 3, PyQt6, imaplib, anthropic SDK (optional)
- Single-file: `gmaildownloader.py` with `_bootstrap()` auto-install
- Catppuccin Mocha dark theme

## Version
- v1.1.0 — Email preview, search/filter, contacts, HTML archive, large email finder, window persistence

## Architecture
- **5-page flow**: Connect -> Download -> Analyze -> Review -> Execute
- **CategoryEngine**: Domain mapping (150+), List-Unsubscribe + newsletter platform detection, subject patterns, learned rules (feedback loop), clean rules engine, subscription detection, thread building, sensitive scanning
- **Workers (QThread)**: ImapScanWorker, ImapDownloadWorker, ImapLabelWorker, LocalOrganizeWorker, AttachmentExtractWorker, SensitiveScanWorker, ThreadSummaryWorker, AiClassifyWorker, HtmlArchiveWorker
- **Dialogs**: StatsDialog (charts + heatmap + large emails + quota), SubscriptionDialog, RulesEditorDialog, ContactDialog, ThreadSummaryDialog
- **Custom Widgets**: HBarChart, ActivityHeatmap (QPainter-based, no external deps)

## Features (v1.1.0)
1. **Email Preview Panel** — Click any email to see rendered HTML/text body in a preview pane below the table
2. **Search & Filter** — Instant keyword search across subject/sender + date range pickers
3. **Statistics Dashboard** — Emails/month, activity heatmap, top senders/domains, category distribution, storage, large email finder, Gmail quota estimate
4. **Contact Analysis** — Sortable/filterable contact table with email counts, sent/received split, first/last seen, dormant contact highlighting
5. **Group-by Views** — Tree view switches between Category, Sender Domain, Sender, Source Folder
6. **Subscription Scanner** — Detects newsletters via List-Unsubscribe + 25+ platform domains, frequency, one-click unsubscribe
7. **Feedback Loop** — Moving emails teaches the engine (learned_rules.json). Applied first on next run
8. **Auto Clean Rules** — Persistent rules engine with CRUD editor + Gmail filter XML import
9. **Attachment Extraction** — SHA-256 deduplicated, organized by category
10. **CSV/JSON/HTML Export** — Full metadata export + static browseable HTML archive with per-email pages
11. **Gmail Filter Import** — Parse Gmail's Atom XML into GmailDownloader rules
12. **Sensitive Content Scanner** — SSN, credit cards, passwords, API keys, tokens
13. **Thread Summarization** — AI summaries of top 50 longest email threads
14. **Select All / Bulk Move** — Bulk operations on visible emails
15. **Window State Persistence** — Remembers size, position, last email via QSettings
16. **Large Email Finder** — Top 20 largest emails surfaced in stats for cleanup

## Key Files
- `gmaildownloader.py` — Single-file app (~2800 lines)
- `learned_rules.json` — Feedback loop persistence (in download dir)
- `clean_rules.json` — Auto clean rules (in download dir)
- `manifest.json` — Download resume tracking (in download dir)

## Run
```bash
python gmaildownloader.py
```

## Gotchas
- Gmail IMAP `[Gmail]/` prefix sanitized for local paths
- IMAP UIDs are per-folder — composite `folder:uid` key used
- Message-ID dedup prevents redundant .eml downloads across Gmail labels
- Chart widgets use QPainter (no matplotlib dependency)
- Sensitive scanner reads first 50KB of each .eml for performance
- Thread summarization limited to top 50 longest threads, 10 emails each
- HTML archive renders email bodies with basic HTML — complex emails may not render perfectly in QTextBrowser
- QSettings stores window geometry and last-used email address
