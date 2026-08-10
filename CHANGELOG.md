# Changelog

All notable changes to GmailDownloader will be documented in this file.

## [v1.2.0]

- Added OAuth2/XOAUTH2, Gmail REST, Microsoft Graph, generic IMAP, mbox, and Thunderbird source paths.
- Added resumable SHA-256 manifests, delta sync, attachments-only and encrypted archive workflows.
- Added headless sync/import/export tooling, Markdown/PDF/MBOX/graph exports, local Ollama classification, and analytics forecasts.
- Added PDF/image receipt vision, optional OCR, normalized receipt JSON, OFX export, and Received-header location timelines.
- Added reproducible Windows, macOS, and Linux packaging scripts with optional Authenticode signing.

## [v1.1.0]

- Rename project from InboxForge to GmailDownloader
- v1.1.0 — email preview, search/filter, contacts, HTML archive, large email finder, window persistence
- Added: Add README
- v1.0.0 — full power release: stats, subscriptions, rules, attachments, sensitive scan, thread summaries, exports, feedback loop
- v0.3.0 — full mailbox download with folder preservation, dedup, and resume
- InboxForge v0.1.0 — Gmail inbox organizer with AI categorization

## Roadmap archive — 2026-08-10 — ROADMAP.md

<details>
<summary>Original roadmap snapshot</summary>

```markdown
# ROADMAP

Backlog for GmailDownloader. Local-first Gmail mailbox downloader, organizer, and analytics.
Stays read-only against the mailbox by default; organize-and-ship without mutating the server.

## Planned Features

### Protocol / source

### Other mail sources

### AI / classification

### Analytics

### Organize / export

### Safety

### Distribution

## Competitive Research

- **Thunderbird + IMAP** — free baseline; creates local mail store but no AI / analytics. Offer
  Thunderbird profile import as a migration path.
- **MailStore Home** — free archiving, strong search. GmailDownloader's edge: analytics + AI
  categorization; integrate by exporting to MailStore-compatible formats.
- **Google Takeout** — official `.mbox` export. Keep as an import source so users with
  pre-existing Takeout don't re-fetch.
- **IMAP Downloader / IMAPSize** — minimal free tools; not competitors, just reminders to keep
  the download step robust.
- **UpTrends / Mailbrew / Notion Mail** — newsletter-digest tools; reference for the
  subscription-management UX.

## Nice-to-Haves


## Open-Source Research (Round 2)

### Related OSS Projects
- **gmvault** — https://github.com/gaubert/gmvault — Mature Gmail backup CLI; XOAuth2; incremental + quick sync; encrypted storage; restore-to-Gmail is its signature feature.
- **abjennings/gmail-backup** — https://github.com/abjennings/gmail-backup — Minimalist IMAP → `.eml` script; resumable; keyed by Gmail unique message id.
- **rosenloecher-it/mail-backup** — https://github.com/rosenloecher-it/mail-backup — Configurable IMAP backup with path/filename templates, date-range filter, dedup.
- **rjmoggach/python-gmail-export** — https://github.com/rjmoggach/python-gmail-export — Gmail API + label filter → `.eml`; optional PDF/HTML + attachment extraction.
- **TSTP-Enterprises/TSTP-GMail_Backup** — https://github.com/TSTP-Enterprises/TSTP-GMail_Backup — Desktop GUI; multi-format export (.txt/.eml/.csv/.pdf); built-in EML viewer.
- **mcaceresb/gmail-download** — https://github.com/mcaceresb/gmail-download — Date-bucket sort + attachment size caps.
- **got-your-back (GYB)** — https://github.com/GAM-team/got-your-back — Google Workspace-grade Gmail backup; works around Gmail API quota the right way.
- **offlineimap** — https://github.com/OfflineIMAP/offlineimap — Mature IMAP sync engine; architecture lessons for two-way sync if ever needed.

### Features to Borrow
- Incremental + quick-sync modes from `gmvault` — current full-download-only; add "sync last 30 days" + "full sync then watch" for cron use.
- Restore-to-Gmail path (`gmvault`, `GYB`) — re-upload `.eml` back to Gmail as hidden-label archive; value-add few OSS tools have.
- Template-based file naming (`rosenloecher`) — `{date}/{from_domain}/{subject}.eml` with tokens; users vary wildly on how they want tree layout.
- Date-range + label filter (`rjmoggach`, `mcaceresb`) — first-class filter UI in GUI, not only as "download all then sort."
- EML viewer inside the GUI (`TSTP`) — PyQt `QTextBrowser` + `email` module render; users don't need external viewers.
- Attachment size cap (`mcaceresb`) — skip >N MB mails or detach-to-sidecar; saves disk on "big PDF newsletter" accounts.
- GYB-style quota-friendly batch + exponential backoff — Gmail API is aggressive with 429s; GYB's algorithm is the gold standard.

### Patterns & Architectures Worth Studying
- **Gmail API batchHttpRequest** (`rjmoggach`, `GYB`): 100 messages per batch call instead of 100 serial calls — dominant factor in total backup time.
- **Resumable state keyed by message id** (`gmvault`, `abjennings`): ids are stable forever; never re-download already-stored messages even across runs or account moves.
- **IMAP UIDVALIDITY + UIDNEXT bookkeeping** (`offlineimap`): how "last synced" is properly tracked across reconnects without missing or duplicating.
- **Encrypted archive storage** (`gmvault` GPG mode): optionally store `.eml` inside an encrypted tarball; integrates naturally with `restic` for off-site.
- **Rate-limit-aware producer/consumer** (`GYB`): one coroutine pulls message-ids, a worker pool fetches bodies with token-bucket throttling; handles Gmail's per-user 250 quota-units/sec ceiling cleanly.
```

</details>
