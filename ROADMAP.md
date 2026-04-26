# ROADMAP

Backlog for GmailDownloader. Local-first Gmail mailbox downloader, organizer, and analytics.
Stays read-only against the mailbox by default; organize-and-ship without mutating the server.

## Planned Features

### Protocol / source
- **OAuth2 / "Sign in with Google"** instead of App Passwords — Google has pushed away from IMAP
  App Passwords; support XOAUTH2 as primary, fall back to App Password.
- **Gmail API (REST) backend** alongside IMAP — native label awareness, better threading, less
  rate-limit friction on large mailboxes.
- **Incremental sync** — on re-run, only fetch UIDs newer than the last successful manifest
  position per folder.
- **Delta since date** option for targeted backups.
- **Multi-account mode** — back up several Gmail accounts in one project with side-by-side
  analytics.

### Other mail sources
- **Generic IMAP** — Fastmail, Proton Bridge, Yahoo, custom IMAP — same pipeline, different auth.
- **Microsoft 365 / Outlook via Graph API** — downloads as `.eml`, feeds the same AI pipeline.
- **Google Takeout .mbox import** — skip IMAP, ingest an existing Takeout archive.
- **Thunderbird profile import** — walk `ImapMail/` and ingest `.msf`-indexed folders.

### AI / classification
- **Local LLM option** via Ollama (parallel to the current Claude Haiku path) — same model
  choices as FileOrganizer.
- **Vision classifier for receipt / invoice attachments** — PDF → image → LLM → structured JSON.
- **Relationship graph** — sender-to-sender co-occurrence in threads, surface cliques.
- **Thread clustering** — group related threads by topic using embeddings (sentence-transformers)
  for "research a topic across your history" UX.
- **Confidence-calibrated batch mode** — LLM only where rule confidence < threshold.

### Analytics
- **Sender health scoring** — combine frequency, reply rate, unsubscribe signals to flag "this
  contact is effectively dead" or "never replied to any of your emails".
- **Thread reply latency histogram** per sender.
- **Storage forecast** — projected inbox size over 12 months if current-rate continues.
- **Location timeline** — extract IP/country from `Received:` headers for travel history audit.

### Organize / export
- **Export to Notion / Obsidian / Markdown vault** with per-email frontmatter.
- **PDF rendering** of email body + inline images per email or per category.
- **MBOX output** alongside `.eml` for Thunderbird / MailStore ingestion.
- **Attachments-only mode** — skip `.eml` storage, just grab + dedupe attachments.
- **Encrypted archive** — age/AES-256 encryption of the `organized/` tree with a passphrase.

### Safety
- **Read-only default** — keep the mailbox-modifying "Apply Gmail labels" mode gated behind a
  separate confirm + dry-run preview.
- **Resume manifest integrity check** — SHA-256 per downloaded `.eml` on manifest, verify before
  re-use.
- **Scrubber for sensitive findings** — already detects; add redact-in-place option that writes
  a redacted `.eml` alongside the original.

### Distribution
- **PyInstaller signed exe** with `multiprocessing.freeze_support()` guard; `anthropic` SDK is
  thread-safe but embeds greenlets that can interact badly with Qt if frozen without care.
- **macOS `.app`**, **Linux AppImage**.
- **Docker image** for headless / cron-scheduled backups.

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

- **Scheduled backups** via OS scheduler with rotation/retention.
- **Natural-language search** (Meilisearch / Tantivy index + embeddings) across the local archive.
- **Contact graph export** to CSV/vCard/Airtable.
- **"Inbox zero" assistant** — rule-based: flag X, unsubscribe from Y, suggest auto-archive
  for Z; everything preview-then-apply.
- **Receipt extraction** — OCR + LLM parse amounts/dates, export OFX for YNAB/Fidelity.
- **Google Photos attachment linking** — cross-reference photos sent via Gmail against Google
  Photos library (once a user-supplied OAuth token is wired).

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
