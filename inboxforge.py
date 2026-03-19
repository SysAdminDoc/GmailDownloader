#!/usr/bin/env python3
"""InboxForge v0.1.0 — Gmail Inbox Organizer with AI-Powered Categorization"""

VERSION = "0.1.0"

import sys, os, subprocess

def _bootstrap():
    """Auto-install dependencies."""
    deps = {'PyQt6': 'PyQt6', 'anthropic': 'anthropic'}
    for imp_name, pkg_name in deps.items():
        try:
            __import__(imp_name)
        except ImportError:
            for cmd in [
                [sys.executable, '-m', 'pip', 'install', pkg_name],
                [sys.executable, '-m', 'pip', 'install', '--user', pkg_name],
                [sys.executable, '-m', 'pip', 'install', '--break-system-packages', pkg_name],
            ]:
                if subprocess.run(cmd, capture_output=True).returncode == 0:
                    break

_bootstrap()

import imaplib
import email
import email.header
import email.utils
import re
import json
import traceback
from collections import Counter, defaultdict
from datetime import datetime
from dataclasses import dataclass, field
from typing import Optional

from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


# ─── Theme Colors (Catppuccin Mocha) ──────────────────────────────────────

class C:
    BASE = "#1e1e2e"
    MANTLE = "#181825"
    CRUST = "#11111b"
    SURFACE0 = "#313244"
    SURFACE1 = "#45475a"
    SURFACE2 = "#585b70"
    TEXT = "#cdd6f4"
    SUBTEXT0 = "#a6adc8"
    SUBTEXT1 = "#bac2de"
    BLUE = "#89b4fa"
    GREEN = "#a6e3a1"
    MAUVE = "#cba6f7"
    RED = "#f38ba8"
    PEACH = "#fab387"
    YELLOW = "#f9e2af"
    TEAL = "#94e2d5"
    LAVENDER = "#b4befe"
    OVERLAY0 = "#6c7086"


STYLESHEET = f"""
    QMainWindow, QWidget {{
        background-color: {C.BASE};
        color: {C.TEXT};
        font-family: 'Segoe UI', sans-serif;
        font-size: 13px;
    }}
    QLineEdit, QTextEdit, QPlainTextEdit {{
        background-color: {C.SURFACE0};
        color: {C.TEXT};
        border: 1px solid {C.SURFACE1};
        border-radius: 6px;
        padding: 8px;
        selection-background-color: {C.BLUE};
    }}
    QLineEdit:focus, QTextEdit:focus {{
        border: 1px solid {C.BLUE};
    }}
    QPushButton {{
        background-color: {C.BLUE};
        color: {C.CRUST};
        border: none;
        border-radius: 6px;
        padding: 8px 20px;
        font-weight: bold;
    }}
    QPushButton:hover {{
        background-color: {C.LAVENDER};
    }}
    QPushButton:disabled {{
        background-color: {C.SURFACE1};
        color: {C.OVERLAY0};
    }}
    QPushButton[secondary="true"] {{
        background-color: {C.SURFACE1};
        color: {C.TEXT};
    }}
    QPushButton[secondary="true"]:hover {{
        background-color: {C.SURFACE2};
    }}
    QPushButton[danger="true"] {{
        background-color: {C.RED};
        color: {C.CRUST};
    }}
    QProgressBar {{
        background-color: {C.SURFACE0};
        border: none;
        border-radius: 4px;
        height: 8px;
        text-align: center;
    }}
    QProgressBar::chunk {{
        background-color: {C.BLUE};
        border-radius: 4px;
    }}
    QTreeWidget, QTableWidget {{
        background-color: {C.MANTLE};
        color: {C.TEXT};
        border: 1px solid {C.SURFACE0};
        border-radius: 6px;
        outline: none;
    }}
    QTreeWidget::item, QTableWidget::item {{
        padding: 4px;
    }}
    QTreeWidget::item:selected, QTableWidget::item:selected {{
        background-color: {C.SURFACE1};
    }}
    QTreeWidget::item:hover, QTableWidget::item:hover {{
        background-color: {C.SURFACE0};
    }}
    QHeaderView::section {{
        background-color: {C.SURFACE0};
        color: {C.SUBTEXT1};
        border: none;
        padding: 6px;
        font-weight: bold;
    }}
    QScrollBar:vertical {{
        background-color: {C.MANTLE};
        width: 10px;
        border-radius: 5px;
    }}
    QScrollBar::handle:vertical {{
        background-color: {C.SURFACE1};
        border-radius: 5px;
        min-height: 30px;
    }}
    QScrollBar::handle:vertical:hover {{
        background-color: {C.SURFACE2};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0;
    }}
    QScrollBar:horizontal {{
        background-color: {C.MANTLE};
        height: 10px;
        border-radius: 5px;
    }}
    QScrollBar::handle:horizontal {{
        background-color: {C.SURFACE1};
        border-radius: 5px;
        min-width: 30px;
    }}
    QLabel {{
        color: {C.TEXT};
    }}
    QGroupBox {{
        color: {C.TEXT};
        border: 1px solid {C.SURFACE0};
        border-radius: 8px;
        margin-top: 12px;
        padding-top: 16px;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 6px;
    }}
    QSplitter::handle {{
        background-color: {C.SURFACE0};
    }}
    QComboBox {{
        background-color: {C.SURFACE0};
        color: {C.TEXT};
        border: 1px solid {C.SURFACE1};
        border-radius: 6px;
        padding: 6px 12px;
    }}
    QComboBox::drop-down {{
        border: none;
    }}
    QComboBox QAbstractItemView {{
        background-color: {C.SURFACE0};
        color: {C.TEXT};
        selection-background-color: {C.SURFACE1};
        border: 1px solid {C.SURFACE1};
    }}
    QCheckBox {{
        color: {C.TEXT};
        spacing: 8px;
    }}
    QCheckBox::indicator {{
        width: 18px;
        height: 18px;
        border-radius: 4px;
        border: 2px solid {C.SURFACE2};
        background-color: {C.SURFACE0};
    }}
    QCheckBox::indicator:checked {{
        background-color: {C.BLUE};
        border-color: {C.BLUE};
    }}
    QInputDialog, QMessageBox {{
        background-color: {C.BASE};
        color: {C.TEXT};
    }}
    QMenu {{
        background-color: {C.SURFACE0};
        color: {C.TEXT};
        border: 1px solid {C.SURFACE1};
        padding: 4px;
    }}
    QMenu::item:selected {{
        background-color: {C.SURFACE1};
    }}
"""


# ─── Known Domain Mappings ────────────────────────────────────────────────

DOMAIN_CATEGORIES = {
    # Shopping
    'amazon.com': 'Shopping', 'amazon.co.uk': 'Shopping', 'ebay.com': 'Shopping',
    'walmart.com': 'Shopping', 'target.com': 'Shopping', 'bestbuy.com': 'Shopping',
    'etsy.com': 'Shopping', 'shopify.com': 'Shopping', 'aliexpress.com': 'Shopping',
    'newegg.com': 'Shopping', 'costco.com': 'Shopping', 'homedepot.com': 'Shopping',
    'lowes.com': 'Shopping', 'macys.com': 'Shopping', 'nordstrom.com': 'Shopping',
    'wayfair.com': 'Shopping', 'chewy.com': 'Shopping', 'wish.com': 'Shopping',
    'kohls.com': 'Shopping', 'samsclub.com': 'Shopping', 'zappos.com': 'Shopping',
    'overstock.com': 'Shopping', 'bhphotovideo.com': 'Shopping',
    # Social Media
    'facebook.com': 'Social Media', 'facebookmail.com': 'Social Media',
    'twitter.com': 'Social Media', 'x.com': 'Social Media',
    'linkedin.com': 'Social Media', 'instagram.com': 'Social Media',
    'reddit.com': 'Social Media', 'redditmail.com': 'Social Media',
    'tiktok.com': 'Social Media', 'snapchat.com': 'Social Media',
    'pinterest.com': 'Social Media', 'nextdoor.com': 'Social Media',
    'discord.com': 'Social Media', 'discordapp.com': 'Social Media',
    'tumblr.com': 'Social Media', 'mastodon.social': 'Social Media',
    # Financial
    'chase.com': 'Financial', 'bankofamerica.com': 'Financial',
    'wellsfargo.com': 'Financial', 'citibank.com': 'Financial',
    'capitalone.com': 'Financial', 'paypal.com': 'Financial',
    'venmo.com': 'Financial', 'cashapp.com': 'Financial',
    'stripe.com': 'Financial', 'square.com': 'Financial',
    'mint.com': 'Financial', 'intuit.com': 'Financial',
    'turbotax.com': 'Financial', 'creditkarma.com': 'Financial',
    'discover.com': 'Financial', 'americanexpress.com': 'Financial',
    'synchrony.com': 'Financial', 'ally.com': 'Financial',
    'fidelity.com': 'Financial', 'schwab.com': 'Financial',
    'vanguard.com': 'Financial', 'robinhood.com': 'Financial',
    # Tech & Services
    'google.com': 'Tech & Services', 'microsoft.com': 'Tech & Services',
    'apple.com': 'Tech & Services', 'dropbox.com': 'Tech & Services',
    'zoom.us': 'Tech & Services', 'slack.com': 'Tech & Services',
    'github.com': 'Tech & Services', 'atlassian.com': 'Tech & Services',
    'cloudflare.com': 'Tech & Services', 'digitalocean.com': 'Tech & Services',
    'heroku.com': 'Tech & Services', 'notion.so': 'Tech & Services',
    'airtable.com': 'Tech & Services', 'canva.com': 'Tech & Services',
    'adobe.com': 'Tech & Services', 'jetbrains.com': 'Tech & Services',
    'docker.com': 'Tech & Services', 'npmjs.com': 'Tech & Services',
    'vercel.com': 'Tech & Services', 'netlify.com': 'Tech & Services',
    'godaddy.com': 'Tech & Services', 'namecheap.com': 'Tech & Services',
    # Travel
    'airbnb.com': 'Travel', 'booking.com': 'Travel',
    'expedia.com': 'Travel', 'hotels.com': 'Travel',
    'delta.com': 'Travel', 'united.com': 'Travel',
    'southwest.com': 'Travel', 'aa.com': 'Travel',
    'uber.com': 'Travel', 'lyft.com': 'Travel',
    'kayak.com': 'Travel', 'tripadvisor.com': 'Travel',
    'vrbo.com': 'Travel', 'hilton.com': 'Travel',
    'marriott.com': 'Travel', 'jetblue.com': 'Travel',
    # Food & Delivery
    'doordash.com': 'Food & Delivery', 'ubereats.com': 'Food & Delivery',
    'grubhub.com': 'Food & Delivery', 'instacart.com': 'Food & Delivery',
    'postmates.com': 'Food & Delivery', 'seamless.com': 'Food & Delivery',
    'chipotle.com': 'Food & Delivery', 'dominos.com': 'Food & Delivery',
    # Entertainment
    'netflix.com': 'Entertainment', 'spotify.com': 'Entertainment',
    'hulu.com': 'Entertainment', 'disneyplus.com': 'Entertainment',
    'twitch.tv': 'Entertainment', 'youtube.com': 'Entertainment',
    'steampowered.com': 'Entertainment', 'epicgames.com': 'Entertainment',
    'playstation.com': 'Entertainment', 'xbox.com': 'Entertainment',
    'crunchyroll.com': 'Entertainment', 'max.com': 'Entertainment',
    'paramountplus.com': 'Entertainment', 'peacocktv.com': 'Entertainment',
    'suno.com': 'Entertainment', 'audible.com': 'Entertainment',
    # News
    'nytimes.com': 'News', 'washingtonpost.com': 'News',
    'cnn.com': 'News', 'bbc.com': 'News', 'bbc.co.uk': 'News',
    'reuters.com': 'News', 'apnews.com': 'News',
    'theguardian.com': 'News', 'wsj.com': 'News',
    # Health
    'myfitnesspal.com': 'Health', 'fitbit.com': 'Health',
    'headspace.com': 'Health', 'calm.com': 'Health',
    'mychart.com': 'Health', 'zocdoc.com': 'Health',
    # Education
    'coursera.org': 'Education', 'udemy.com': 'Education',
    'edx.org': 'Education', 'khanacademy.org': 'Education',
    'skillshare.com': 'Education',
}

SUBJECT_PATTERNS = {
    'Shipping & Tracking': [
        r'(?i)\b(shipped|tracking|delivery|delivered|out for delivery|package|shipment)\b',
        r'(?i)\b(ups|fedex|usps|dhl)\b.*(?:tracking|delivery)',
    ],
    'Invoices & Billing': [
        r'(?i)\b(invoice|receipt|payment\s+(?:received|confirmed)|billing\s+statement)\b',
        r'(?i)\b(order\s+confirm|order\s+#|your\s+order)\b',
    ],
    'Security Alerts': [
        r'(?i)\b(security\s+alert|suspicious|unauthorized|password\s+reset|verify\s+your)\b',
        r'(?i)\b(two-factor|2fa|verification\s+code|login\s+attempt|sign-in)\b',
    ],
    'Calendar & Meetings': [
        r'(?i)\b(meeting\s+(?:invite|invitation|reminder)|calendar|rsvp|webinar)\b',
        r'(?i)\b(zoom\s+meeting|teams\s+meeting|google\s+meet)\b',
    ],
}


# ─── Data Classes ─────────────────────────────────────────────────────────

@dataclass
class EmailInfo:
    uid: str
    sender: str = ""
    sender_name: str = ""
    sender_domain: str = ""
    subject: str = ""
    date: str = ""
    date_parsed: Optional[datetime] = None
    has_list_unsubscribe: bool = False
    category: str = ""
    confidence: float = 0.0


# ─── Category Engine ──────────────────────────────────────────────────────

class CategoryEngine:
    def __init__(self, user_domain: str = ""):
        self.user_domain = user_domain.lower()
        self.emails: list[EmailInfo] = []
        self.categories: dict[str, list[EmailInfo]] = defaultdict(list)
        self.domain_stats: Counter = Counter()
        self.ambiguous: list[EmailInfo] = []

    def extract_domain(self, email_addr: str) -> str:
        match = re.search(r'@([\w.-]+)', email_addr.lower())
        if not match:
            return ""
        domain = match.group(1)
        parts = domain.split('.')
        if len(parts) > 2:
            country_slds = ('co', 'com', 'org', 'ac', 'gov', 'net', 'edu')
            if len(parts) >= 3 and parts[-2] in country_slds:
                domain = '.'.join(parts[-3:])
            else:
                domain = '.'.join(parts[-2:])
        return domain

    def categorize_email(self, em: EmailInfo) -> tuple[str, float]:
        domain = em.sender_domain

        # 1. Internal/Work emails
        if self.user_domain and domain == self.user_domain:
            return "Work/Internal", 0.95

        # 2. Known domain mapping
        if domain in DOMAIN_CATEGORIES:
            return DOMAIN_CATEGORIES[domain], 0.9

        # 3. Newsletter detection
        if em.has_list_unsubscribe:
            return "Newsletters", 0.85

        # 4. Subject pattern matching
        for cat, patterns in SUBJECT_PATTERNS.items():
            for pattern in patterns:
                if re.search(pattern, em.subject):
                    return cat, 0.7

        # 5. Automated sender detection
        sender_lower = em.sender.lower()
        if re.search(r'(?i)(no-?reply|noreply|notifications?@|alerts?@|mailer-daemon)', sender_lower):
            return "Automated/Notifications", 0.6

        # 6. Ambiguous
        return "", 0.0

    def process_all(self, emails: list[EmailInfo]):
        self.emails = emails
        self.categories.clear()
        self.ambiguous.clear()
        self.domain_stats.clear()

        for em in emails:
            em.sender_domain = self.extract_domain(em.sender)
            self.domain_stats[em.sender_domain] += 1
            cat, conf = self.categorize_email(em)
            em.category = cat
            em.confidence = conf
            if cat:
                self.categories[cat].append(em)
            else:
                self.ambiguous.append(em)

        # Second pass: group ambiguous by domain — if a domain has 5+ emails, auto-group
        ambiguous_domains = Counter(em.sender_domain for em in self.ambiguous)
        new_ambiguous = []
        for em in self.ambiguous:
            if ambiguous_domains[em.sender_domain] >= 5:
                cat = f"Other/{em.sender_domain}"
                em.category = cat
                em.confidence = 0.5
                self.categories[cat].append(em)
            else:
                new_ambiguous.append(em)
        self.ambiguous = new_ambiguous

        # Remaining go to Uncategorized
        if self.ambiguous:
            for em in self.ambiguous:
                em.category = "Uncategorized"
                em.confidence = 0.0
            self.categories["Uncategorized"] = list(self.ambiguous)

    def get_summary(self) -> dict:
        total = len(self.emails)
        categorized = sum(1 for em in self.emails if em.confidence > 0)
        dates = [em.date_parsed for em in self.emails if em.date_parsed]
        date_range = ("", "")
        if dates:
            date_range = (min(dates).strftime("%Y-%m-%d"), max(dates).strftime("%Y-%m-%d"))
        return {
            'total': total,
            'categorized': categorized,
            'uncategorized': total - categorized,
            'categories': {k: len(v) for k, v in sorted(self.categories.items(), key=lambda x: -len(x[1]))},
            'top_domains': self.domain_stats.most_common(20),
            'date_range': date_range,
        }

    def rename_category(self, old_name: str, new_name: str):
        if old_name in self.categories:
            emails = self.categories.pop(old_name)
            for em in emails:
                em.category = new_name
            self.categories[new_name].extend(emails)

    def merge_categories(self, sources: list[str], target: str):
        for src in sources:
            if src in self.categories and src != target:
                emails = self.categories.pop(src)
                for em in emails:
                    em.category = target
                self.categories[target].extend(emails)

    def move_emails(self, uids: list[str], target_category: str):
        uid_set = set(uids)
        for em in self.emails:
            if em.uid in uid_set:
                old_cat = em.category
                if old_cat in self.categories:
                    self.categories[old_cat] = [e for e in self.categories[old_cat] if e.uid != em.uid]
                    if not self.categories[old_cat]:
                        del self.categories[old_cat]
                em.category = target_category
                em.confidence = max(em.confidence, 0.5)
                self.categories[target_category].append(em)

    def delete_category(self, name: str):
        if name in self.categories:
            emails = self.categories.pop(name)
            for em in emails:
                em.category = "Uncategorized"
                em.confidence = 0.0
            self.categories["Uncategorized"].extend(emails)

    def save_state(self, path: str):
        """Save analysis state to JSON for resume capability."""
        data = {
            'version': VERSION,
            'user_domain': self.user_domain,
            'emails': [
                {
                    'uid': em.uid, 'sender': em.sender, 'sender_name': em.sender_name,
                    'sender_domain': em.sender_domain, 'subject': em.subject,
                    'date': em.date, 'has_list_unsubscribe': em.has_list_unsubscribe,
                    'category': em.category, 'confidence': em.confidence,
                }
                for em in self.emails
            ],
        }
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)

    def load_state(self, path: str) -> bool:
        """Load analysis state from JSON."""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.user_domain = data.get('user_domain', '')
            self.emails = []
            self.categories.clear()
            for ed in data.get('emails', []):
                em = EmailInfo(
                    uid=ed['uid'], sender=ed['sender'], sender_name=ed['sender_name'],
                    sender_domain=ed['sender_domain'], subject=ed['subject'],
                    date=ed['date'], has_list_unsubscribe=ed['has_list_unsubscribe'],
                    category=ed['category'], confidence=ed['confidence'],
                )
                self.emails.append(em)
                if em.category:
                    self.categories[em.category].append(em)
            self.domain_stats = Counter(em.sender_domain for em in self.emails)
            return True
        except Exception:
            return False


# ─── IMAP Workers ─────────────────────────────────────────────────────────

class ImapScanWorker(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    email_batch = pyqtSignal(list)
    finished_signal = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, host, email_addr, password):
        super().__init__()
        self.host = host
        self.email_addr = email_addr
        self.password = password
        self._stop = False

    def stop(self):
        self._stop = True

    def _decode_header(self, raw):
        if not raw:
            return ""
        try:
            parts = email.header.decode_header(raw)
            decoded = []
            for part, charset in parts:
                if isinstance(part, bytes):
                    decoded.append(part.decode(charset or 'utf-8', errors='replace'))
                else:
                    decoded.append(str(part))
            return ' '.join(decoded).strip()
        except Exception:
            return str(raw).strip()

    def _parse_date(self, date_str):
        if not date_str:
            return None
        try:
            parsed = email.utils.parsedate_to_datetime(date_str)
            return parsed.replace(tzinfo=None)
        except Exception:
            return None

    def run(self):
        try:
            self.status.emit("Connecting to Gmail IMAP...")
            imap = imaplib.IMAP4_SSL(self.host, 993)
            imap.login(self.email_addr, self.password)

            self.status.emit("Selecting INBOX...")
            imap.select('INBOX', readonly=True)

            self.status.emit("Fetching message list...")
            _, data = imap.uid('SEARCH', None, 'ALL')
            uids = data[0].split()
            total = len(uids)
            self.status.emit(f"Found {total:,} emails. Scanning headers...")

            all_emails = []
            batch_size = 200

            for i in range(0, total, batch_size):
                if self._stop:
                    break

                batch_uids = uids[i:i + batch_size]
                uid_range = b','.join(batch_uids)

                _, msg_data = imap.uid(
                    'FETCH', uid_range,
                    '(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE LIST-UNSUBSCRIBE)])'
                )

                batch = []
                current_uid_idx = 0
                for j in range(len(msg_data)):
                    item = msg_data[j]
                    if isinstance(item, tuple) and len(item) == 2:
                        uid_match = re.search(rb'UID (\d+)', item[0])
                        uid = uid_match.group(1).decode() if uid_match else batch_uids[current_uid_idx].decode()
                        current_uid_idx += 1

                        raw_headers = item[1]
                        if isinstance(raw_headers, bytes):
                            try:
                                msg = email.message_from_bytes(raw_headers)
                                from_raw = msg.get('From', '')
                                from_decoded = self._decode_header(from_raw)
                                name, addr = email.utils.parseaddr(from_decoded)

                                em = EmailInfo(
                                    uid=uid,
                                    sender=addr or from_decoded,
                                    sender_name=name or addr or from_decoded,
                                    subject=self._decode_header(msg.get('Subject', '(no subject)')),
                                    date=msg.get('Date', ''),
                                    date_parsed=self._parse_date(msg.get('Date', '')),
                                    has_list_unsubscribe=msg.get('List-Unsubscribe') is not None,
                                )
                                batch.append(em)
                                all_emails.append(em)
                            except Exception:
                                pass

                self.progress.emit(min(i + batch_size, total), total)
                if batch:
                    self.email_batch.emit(batch)

            try:
                imap.close()
                imap.logout()
            except Exception:
                pass

            self.finished_signal.emit(all_emails)

        except imaplib.IMAP4.error as e:
            self.error.emit(f"IMAP Error: {e}")
        except Exception as e:
            self.error.emit(f"Error: {e}\n{traceback.format_exc()}")


class ImapLabelWorker(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    log = pyqtSignal(str)
    finished_signal = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, host, email_addr, password, categories, label_prefix="", archive_after=False):
        super().__init__()
        self.host = host
        self.email_addr = email_addr
        self.password = password
        self.categories = categories
        self.label_prefix = label_prefix
        self.archive_after = archive_after
        self._stop = False

    def stop(self):
        self._stop = True

    def _imap_encode_label(self, label):
        """Encode label name for IMAP using modified UTF-7."""
        # Gmail accepts UTF-8 label names in quotes
        return f'"{label}"'

    def run(self):
        try:
            self.status.emit("Connecting to Gmail IMAP...")
            imap = imaplib.IMAP4_SSL(self.host, 993)
            imap.login(self.email_addr, self.password)

            total_emails = sum(len(v) for v in self.categories.values())
            processed = 0

            for cat_name, emails in self.categories.items():
                if self._stop:
                    break
                if not emails:
                    continue

                label = f"{self.label_prefix}/{cat_name}" if self.label_prefix else cat_name

                self.log.emit(f"Creating label: {label}")
                try:
                    imap.create(self._imap_encode_label(label))
                except Exception:
                    pass  # Already exists

                imap.select('INBOX')

                batch_size = 100
                for i in range(0, len(emails), batch_size):
                    if self._stop:
                        break

                    batch = emails[i:i + batch_size]
                    uid_list = ','.join(em.uid for em in batch)

                    try:
                        result = imap.uid('COPY', uid_list, self._imap_encode_label(label))
                        if result[0] == 'OK':
                            processed += len(batch)
                        else:
                            self.log.emit(f"  Warning: COPY returned {result[0]} for batch")
                            processed += len(batch)
                    except Exception as e:
                        self.log.emit(f"  Error labeling batch: {e}")

                    self.progress.emit(processed, total_emails)
                    self.log.emit(f"  Labeled {min(i + batch_size, len(emails))}/{len(emails)} as '{label}'")

                if self.archive_after and not self._stop:
                    self.log.emit(f"  Archiving {len(emails)} emails from Inbox...")
                    imap.select('INBOX')
                    for i in range(0, len(emails), batch_size):
                        if self._stop:
                            break
                        batch = emails[i:i + batch_size]
                        uid_list = ','.join(em.uid for em in batch)
                        try:
                            imap.uid('STORE', uid_list, '+FLAGS', '(\\Deleted)')
                        except Exception as e:
                            self.log.emit(f"  Error archiving batch: {e}")
                    try:
                        imap.expunge()
                    except Exception:
                        pass

            try:
                imap.logout()
            except Exception:
                pass

            self.finished_signal.emit()

        except Exception as e:
            self.error.emit(f"Error: {e}\n{traceback.format_exc()}")


class AiClassifyWorker(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    classified = pyqtSignal(dict)
    finished_signal = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, api_key, ambiguous_emails, existing_categories):
        super().__init__()
        self.api_key = api_key
        self.emails = ambiguous_emails
        self.existing_categories = existing_categories
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            client = anthropic.Anthropic(api_key=self.api_key)

            domain_groups = defaultdict(list)
            for em in self.emails:
                domain_groups[em.sender_domain].append(em)

            domains = list(domain_groups.keys())
            total = len(domains)
            batch_size = 30

            self.status.emit(f"Classifying {total} unknown domains with AI...")

            for i in range(0, total, batch_size):
                if self._stop:
                    break

                batch_domains = domains[i:i + batch_size]

                domain_info = []
                for d in batch_domains:
                    emails = domain_groups[d][:5]
                    domain_info.append({
                        'domain': d,
                        'count': len(domain_groups[d]),
                        'senders': list(set(em.sender_name for em in emails))[:3],
                        'sample_subjects': [em.subject for em in emails],
                    })

                prompt = (
                    f"Categorize these email sender domains into appropriate categories.\n"
                    f"Existing categories: {', '.join(self.existing_categories)}\n\n"
                    f"You may use existing categories or suggest new ones. "
                    f"Reply with ONLY a JSON object mapping domain to category.\n\n"
                    f"Domains to classify:\n{json.dumps(domain_info, indent=2)}\n\n"
                    f'Reply with ONLY a JSON object like: '
                    f'{{"domain1.com": "Category Name", "domain2.com": "Category Name"}}'
                )

                try:
                    response = client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=1024,
                        messages=[{"role": "user", "content": prompt}]
                    )

                    text = response.content[0].text
                    json_match = re.search(r'\{[^{}]*\}', text, re.DOTALL)
                    if json_match:
                        result = json.loads(json_match.group())
                        self.classified.emit(result)
                except Exception as e:
                    self.status.emit(f"AI batch error (continuing): {e}")

                self.progress.emit(min(i + batch_size, total), total)

            self.finished_signal.emit()

        except Exception as e:
            self.error.emit(f"AI Error: {e}")


# ─── Connection Tester ────────────────────────────────────────────────────

class ConnectionTester(QObject):
    success = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, host, email_addr, password):
        super().__init__()
        self.host = host
        self.email_addr = email_addr
        self.password = password

    def run(self):
        try:
            imap = imaplib.IMAP4_SSL(self.host, 993)
            imap.login(self.email_addr, self.password)
            _, data = imap.select('INBOX', readonly=True)
            count = int(data[0])
            imap.close()
            imap.logout()
            self.success.emit(count)
        except imaplib.IMAP4.error as e:
            self.error.emit(f"Login failed: {e}")
        except Exception as e:
            self.error.emit(f"Connection failed: {e}")


# ─── UI: Connect Page ────────────────────────────────────────────────────

class ConnectPage(QWidget):
    connected = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.imap_host = "imap.gmail.com"
        self.email_addr = ""
        self.password = ""
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(16)

        title = QLabel("InboxForge")
        title.setStyleSheet(f"font-size: 32px; color: {C.BLUE}; font-weight: bold;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel("Gmail Inbox Organizer with AI Categorization")
        subtitle.setStyleSheet(f"color: {C.SUBTEXT0}; font-size: 12px;")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(subtitle)

        layout.addSpacing(20)

        form = QWidget()
        form.setMaximumWidth(420)
        fl = QVBoxLayout(form)
        fl.setSpacing(12)

        fl.addWidget(QLabel("Gmail Address"))
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("you@gmail.com")
        fl.addWidget(self.email_input)

        fl.addWidget(QLabel("App Password"))
        self.pass_input = QLineEdit()
        self.pass_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.pass_input.setPlaceholderText("16-character app password")
        fl.addWidget(self.pass_input)

        hint = QLabel("Generate at: Google Account > Security > App Passwords")
        hint.setStyleSheet(f"color: {C.SUBTEXT0}; font-size: 11px;")
        hint.setWordWrap(True)
        fl.addWidget(hint)

        fl.addSpacing(8)

        btn_row = QHBoxLayout()
        self.connect_btn = QPushButton("Connect & Scan")
        self.connect_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.connect_btn.clicked.connect(self._on_connect)
        btn_row.addWidget(self.connect_btn)

        self.load_btn = QPushButton("Load Previous Scan")
        self.load_btn.setProperty("secondary", True)
        self.load_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.load_btn.clicked.connect(self._on_load)
        btn_row.addWidget(self.load_btn)
        fl.addLayout(btn_row)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet(f"color: {C.SUBTEXT0}; font-size: 12px;")
        fl.addWidget(self.status_label)

        layout.addWidget(form, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()

        # Store loaded engine if user loads from file
        self.loaded_engine = None

    def _on_connect(self):
        self.email_addr = self.email_input.text().strip()
        self.password = self.pass_input.text().strip()

        if not self.email_addr or not self.password:
            self.status_label.setText("Please enter email and app password")
            self.status_label.setStyleSheet(f"color: {C.RED};")
            return

        self.status_label.setText("Testing connection...")
        self.status_label.setStyleSheet(f"color: {C.YELLOW};")
        self.connect_btn.setEnabled(False)

        self._test_thread = QThread()
        self._test_worker = ConnectionTester(self.imap_host, self.email_addr, self.password)
        self._test_worker.moveToThread(self._test_thread)
        self._test_thread.started.connect(self._test_worker.run)
        self._test_worker.success.connect(self._on_test_success)
        self._test_worker.error.connect(self._on_test_error)
        self._test_thread.start()

    def _on_test_success(self, msg_count):
        self._test_thread.quit()
        self.status_label.setText(f"Connected! {msg_count:,} messages in Inbox")
        self.status_label.setStyleSheet(f"color: {C.GREEN};")
        self.loaded_engine = None
        self.connected.emit()

    def _on_test_error(self, err):
        self._test_thread.quit()
        self.status_label.setText(err)
        self.status_label.setStyleSheet(f"color: {C.RED};")
        self.connect_btn.setEnabled(True)

    def _on_load(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Scan File", "", "JSON Files (*.json)"
        )
        if not path:
            return

        engine = CategoryEngine()
        if engine.load_state(path):
            self.loaded_engine = engine
            self.email_addr = self.email_input.text().strip()
            self.password = self.pass_input.text().strip()
            self.status_label.setText(f"Loaded {len(engine.emails):,} emails from file")
            self.status_label.setStyleSheet(f"color: {C.GREEN};")
            self.connected.emit()
        else:
            self.status_label.setText("Failed to load scan file")
            self.status_label.setStyleSheet(f"color: {C.RED};")


# ─── UI: Analyze Page ────────────────────────────────────────────────────

class AnalyzePage(QWidget):
    analysis_complete = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.engine = None
        self.worker = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        title = QLabel("Analyzing Inbox")
        title.setStyleSheet(f"font-size: 20px; font-weight: bold;")
        layout.addWidget(title)

        self.status_label = QLabel("Preparing scan...")
        self.status_label.setStyleSheet(f"color: {C.SUBTEXT0};")
        layout.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(8)
        layout.addWidget(self.progress)

        self.percent_label = QLabel("0%")
        self.percent_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.percent_label)

        layout.addSpacing(8)

        stats_group = QGroupBox("Live Statistics")
        stats_layout = QVBoxLayout(stats_group)
        self.stats_text = QPlainTextEdit()
        self.stats_text.setReadOnly(True)
        self.stats_text.setMaximumHeight(350)
        self.stats_text.setStyleSheet(
            f"font-family: 'Cascadia Code', 'Consolas', monospace; font-size: 12px;"
        )
        stats_layout.addWidget(self.stats_text)
        layout.addWidget(stats_group)

        layout.addStretch()

        bottom = QHBoxLayout()
        self.save_btn = QPushButton("Save Scan")
        self.save_btn.setProperty("secondary", True)
        self.save_btn.setVisible(False)
        self.save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.save_btn.clicked.connect(self._on_save)
        bottom.addWidget(self.save_btn)

        bottom.addStretch()

        self.continue_btn = QPushButton("Review Categories")
        self.continue_btn.setVisible(False)
        self.continue_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.continue_btn.clicked.connect(self.analysis_complete.emit)
        bottom.addWidget(self.continue_btn)
        layout.addLayout(bottom)

    def start_scan(self, host, email_addr, password):
        user_domain = email_addr.split('@')[1] if '@' in email_addr else ""
        self.engine = CategoryEngine(user_domain)
        self._domain_counter = Counter()

        self.worker = ImapScanWorker(host, email_addr, password)
        self.worker.progress.connect(self._on_progress)
        self.worker.status.connect(self._on_status)
        self.worker.email_batch.connect(self._on_batch)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def set_preloaded(self, engine: CategoryEngine):
        """Skip scanning, use a pre-loaded engine."""
        self.engine = engine
        self._show_summary()

    def _on_progress(self, current, total):
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        pct = int(current / total * 100) if total > 0 else 0
        self.percent_label.setText(f"{pct}% ({current:,} / {total:,})")

    def _on_status(self, msg):
        self.status_label.setText(msg)

    def _on_batch(self, batch):
        for em in batch:
            domain = self.engine.extract_domain(em.sender)
            self._domain_counter[domain] += 1

        top = self._domain_counter.most_common(15)
        lines = [f"Emails scanned: {sum(self._domain_counter.values()):,}", ""]
        lines.append("Top Sender Domains:")
        for domain, count in top:
            lines.append(f"  {domain:40s} {count:>6,}")
        self.stats_text.setPlainText('\n'.join(lines))

    def _on_finished(self, all_emails):
        self.status_label.setText("Categorizing emails...")
        self.status_label.setStyleSheet(f"color: {C.YELLOW};")
        QApplication.processEvents()

        self.engine.process_all(all_emails)
        self._show_summary()

    def _show_summary(self):
        summary = self.engine.get_summary()
        lines = [
            f"Total emails: {summary['total']:,}",
            f"Auto-categorized: {summary['categorized']:,}",
            f"Uncategorized: {summary['uncategorized']:,}",
            f"Date range: {summary['date_range'][0]} to {summary['date_range'][1]}",
            "",
            "Categories:",
        ]
        for cat, count in summary['categories'].items():
            lines.append(f"  {cat:40s} {count:>6,}")

        self.stats_text.setPlainText('\n'.join(lines))
        self.progress.setMaximum(1)
        self.progress.setValue(1)
        self.percent_label.setText("100%")
        self.status_label.setText("Analysis complete!")
        self.status_label.setStyleSheet(f"color: {C.GREEN};")
        self.continue_btn.setVisible(True)
        self.save_btn.setVisible(True)

    def _on_error(self, err):
        self.status_label.setText(err)
        self.status_label.setStyleSheet(f"color: {C.RED};")

    def _on_save(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Scan", "inboxforge_scan.json", "JSON Files (*.json)"
        )
        if path and self.engine:
            self.engine.save_state(path)
            self.status_label.setText(f"Scan saved to {os.path.basename(path)}")


# ─── UI: Review Page ─────────────────────────────────────────────────────

class ReviewPage(QWidget):
    execute_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.engine: Optional[CategoryEngine] = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        # Top bar
        top_bar = QHBoxLayout()
        title = QLabel("Review Categories")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        top_bar.addWidget(title)
        top_bar.addStretch()

        self.summary_label = QLabel("")
        self.summary_label.setStyleSheet(f"color: {C.SUBTEXT0};")
        top_bar.addWidget(self.summary_label)

        self.ai_btn = QPushButton("Classify Uncategorized with AI")
        self.ai_btn.setProperty("secondary", True)
        self.ai_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.ai_btn.clicked.connect(self._on_ai_classify)
        top_bar.addWidget(self.ai_btn)

        layout.addLayout(top_bar)

        # Splitter: category tree | email table
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left: category tree
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.cat_tree = QTreeWidget()
        self.cat_tree.setHeaderLabels(["Category", "Count"])
        self.cat_tree.setColumnWidth(0, 260)
        self.cat_tree.itemClicked.connect(self._on_cat_selected)
        self.cat_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.cat_tree.customContextMenuRequested.connect(self._on_cat_context_menu)
        left_layout.addWidget(self.cat_tree)

        cat_btns = QHBoxLayout()
        for label, slot, prop in [
            ("Rename", self._on_rename, "secondary"),
            ("Merge", self._on_merge, "secondary"),
            ("Delete", self._on_delete, "danger"),
        ]:
            btn = QPushButton(label)
            btn.setProperty(prop, True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(slot)
            cat_btns.addWidget(btn)
        left_layout.addLayout(cat_btns)
        splitter.addWidget(left)

        # Right: email table
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.email_count_label = QLabel("")
        self.email_count_label.setStyleSheet(f"color: {C.SUBTEXT0}; font-size: 12px;")
        right_layout.addWidget(self.email_count_label)

        self.email_table = QTableWidget()
        self.email_table.setColumnCount(4)
        self.email_table.setHorizontalHeaderLabels(["From", "Subject", "Date", "Conf"])
        header = self.email_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.email_table.setColumnWidth(0, 200)
        self.email_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.email_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.email_table.verticalHeader().setVisible(False)
        right_layout.addWidget(self.email_table)

        move_layout = QHBoxLayout()
        move_layout.addWidget(QLabel("Move selected to:"))
        self.move_combo = QComboBox()
        self.move_combo.setMinimumWidth(200)
        move_layout.addWidget(self.move_combo, 1)
        move_btn = QPushButton("Move")
        move_btn.setProperty("secondary", True)
        move_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        move_btn.clicked.connect(self._on_move_emails)
        move_layout.addWidget(move_btn)
        right_layout.addLayout(move_layout)

        splitter.addWidget(right)
        splitter.setSizes([350, 650])
        layout.addWidget(splitter, 1)

        # Bottom: options + execute
        bottom = QHBoxLayout()

        bottom.addWidget(QLabel("Label prefix:"))
        self.prefix_input = QLineEdit()
        self.prefix_input.setPlaceholderText("optional, e.g. 'InboxForge'")
        self.prefix_input.setMaximumWidth(250)
        bottom.addWidget(self.prefix_input)

        self.archive_check = QCheckBox("Archive from Inbox after labeling")
        bottom.addWidget(self.archive_check)

        bottom.addStretch()

        self.execute_btn = QPushButton("Apply Labels")
        self.execute_btn.setStyleSheet(
            f"background-color: {C.GREEN}; font-size: 14px; padding: 10px 30px;"
        )
        self.execute_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.execute_btn.clicked.connect(self.execute_requested.emit)
        bottom.addWidget(self.execute_btn)

        layout.addLayout(bottom)

    def load_categories(self, engine: CategoryEngine):
        self.engine = engine
        self._refresh_tree()
        self._refresh_combo()
        s = engine.get_summary()
        self.summary_label.setText(
            f"{s['total']:,} emails | {s['categorized']:,} categorized | "
            f"{s['uncategorized']:,} uncategorized"
        )

    def _refresh_tree(self):
        self.cat_tree.clear()
        if not self.engine:
            return

        # Group by parent for nested categories
        parents = {}  # parent_name -> {child_name -> count}
        top_level = {}  # name -> count

        for cat_name, emails in sorted(self.engine.categories.items(), key=lambda x: -len(x[1])):
            if '/' in cat_name:
                parent, child = cat_name.split('/', 1)
                if parent not in parents:
                    parents[parent] = {}
                parents[parent][child] = (cat_name, len(emails))
            else:
                top_level[cat_name] = len(emails)

        # Add top-level items
        for name, count in sorted(top_level.items(), key=lambda x: -x[1]):
            item = QTreeWidgetItem([name, f"{count:,}"])
            item.setData(0, Qt.ItemDataRole.UserRole, name)
            item.setForeground(0, QColor(C.BLUE))
            self.cat_tree.addTopLevelItem(item)

        # Add parent groups with children
        for parent_name, children in sorted(parents.items()):
            total = sum(c[1] for c in children.values())
            parent_item = QTreeWidgetItem([parent_name, f"{total:,}"])
            parent_item.setData(0, Qt.ItemDataRole.UserRole, None)  # Not a real category
            parent_item.setForeground(0, QColor(C.MAUVE))
            self.cat_tree.addTopLevelItem(parent_item)

            for child_name, (full_name, count) in sorted(children.items(), key=lambda x: -x[1][1]):
                child_item = QTreeWidgetItem([child_name, f"{count:,}"])
                child_item.setData(0, Qt.ItemDataRole.UserRole, full_name)
                parent_item.addChild(child_item)

        self.cat_tree.expandAll()

    def _refresh_combo(self):
        self.move_combo.clear()
        if self.engine:
            for cat in sorted(self.engine.categories.keys()):
                self.move_combo.addItem(cat)

    def _get_selected_cat(self):
        item = self.cat_tree.currentItem()
        if not item:
            return None
        return item.data(0, Qt.ItemDataRole.UserRole)

    def _on_cat_selected(self, item):
        cat_name = item.data(0, Qt.ItemDataRole.UserRole)
        if not self.engine:
            return

        if cat_name and cat_name in self.engine.categories:
            emails = self.engine.categories[cat_name]
        else:
            # Parent group — collect all children
            emails = []
            for i in range(item.childCount()):
                child_cat = item.child(i).data(0, Qt.ItemDataRole.UserRole)
                if child_cat and child_cat in self.engine.categories:
                    emails.extend(self.engine.categories[child_cat])

        self._show_emails(emails)

    def _show_emails(self, emails):
        self.email_count_label.setText(f"{len(emails):,} emails")
        # Limit display to 2000 for performance, show newest first
        display = sorted(emails, key=lambda e: e.date_parsed or datetime.min, reverse=True)[:2000]

        self.email_table.setRowCount(len(display))
        for row, em in enumerate(display):
            self.email_table.setItem(row, 0, QTableWidgetItem(em.sender_name or em.sender))
            self.email_table.setItem(row, 1, QTableWidgetItem(em.subject))
            date_str = em.date_parsed.strftime("%Y-%m-%d") if em.date_parsed else ""
            self.email_table.setItem(row, 2, QTableWidgetItem(date_str))

            conf_text = f"{em.confidence:.0%}"
            conf_item = QTableWidgetItem(conf_text)
            if em.confidence >= 0.8:
                conf_item.setForeground(QColor(C.GREEN))
            elif em.confidence >= 0.5:
                conf_item.setForeground(QColor(C.YELLOW))
            else:
                conf_item.setForeground(QColor(C.RED))
            self.email_table.setItem(row, 3, conf_item)

            # Store UID
            self.email_table.item(row, 0).setData(Qt.ItemDataRole.UserRole, em.uid)

        if len(emails) > 2000:
            self.email_count_label.setText(f"{len(emails):,} emails (showing newest 2,000)")

    def _on_cat_context_menu(self, pos):
        item = self.cat_tree.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        menu.addAction("Rename", self._on_rename)
        menu.addAction("Merge into...", self._on_merge)
        menu.addSeparator()
        menu.addAction("Delete (move to Uncategorized)", self._on_delete)
        menu.exec(self.cat_tree.viewport().mapToGlobal(pos))

    def _on_rename(self):
        cat_name = self._get_selected_cat()
        if not cat_name:
            return
        new_name, ok = QInputDialog.getText(self, "Rename Category", "New name:", text=cat_name)
        if ok and new_name and new_name != cat_name:
            self.engine.rename_category(cat_name, new_name)
            self._refresh_tree()
            self._refresh_combo()

    def _on_merge(self):
        if not self.engine:
            return
        source = self._get_selected_cat()
        if not source:
            return
        cats = sorted(self.engine.categories.keys())
        target, ok = QInputDialog.getItem(
            self, "Merge Into", f"Merge '{source}' into:", cats, 0, False
        )
        if ok and target and target != source:
            self.engine.merge_categories([source], target)
            self._refresh_tree()
            self._refresh_combo()

    def _on_delete(self):
        cat_name = self._get_selected_cat()
        if not cat_name:
            return
        self.engine.delete_category(cat_name)
        self._refresh_tree()
        self._refresh_combo()

    def _on_move_emails(self):
        if not self.engine:
            return
        target = self.move_combo.currentText()
        if not target:
            return
        selected_rows = set(idx.row() for idx in self.email_table.selectedIndexes())
        uids = []
        for row in selected_rows:
            uid_item = self.email_table.item(row, 0)
            if uid_item:
                uid = uid_item.data(Qt.ItemDataRole.UserRole)
                if uid:
                    uids.append(uid)
        if uids:
            self.engine.move_emails(uids, target)
            self._refresh_tree()
            # Re-select current category to refresh table
            item = self.cat_tree.currentItem()
            if item:
                self._on_cat_selected(item)

    def _on_ai_classify(self):
        if not HAS_ANTHROPIC:
            QMessageBox.warning(self, "Missing", "anthropic package not installed.")
            return

        uncategorized = self.engine.categories.get("Uncategorized", [])
        if not uncategorized:
            QMessageBox.information(self, "Done", "No uncategorized emails remaining.")
            return

        api_key, ok = QInputDialog.getText(
            self, "Claude API Key",
            f"Enter your Anthropic API key to classify {len(uncategorized):,} emails.\n"
            f"Uses Claude Haiku (fast & cheap).",
            QLineEdit.EchoMode.Password
        )
        if not ok or not api_key:
            return

        existing_cats = [k for k in self.engine.categories.keys() if k != "Uncategorized"]

        self.ai_btn.setEnabled(False)
        self.ai_btn.setText("Classifying...")

        self._ai_worker = AiClassifyWorker(api_key, uncategorized, existing_cats)
        self._ai_worker.classified.connect(self._on_ai_result)
        self._ai_worker.finished_signal.connect(self._on_ai_done)
        self._ai_worker.error.connect(self._on_ai_error)
        self._ai_worker.start()

    def _on_ai_result(self, domain_map: dict):
        for domain, category in domain_map.items():
            to_move = [
                em for em in self.engine.categories.get("Uncategorized", [])
                if em.sender_domain == domain
            ]
            for em in to_move:
                self.engine.categories["Uncategorized"].remove(em)
                em.category = category
                em.confidence = 0.75
                self.engine.categories[category].append(em)

        if not self.engine.categories.get("Uncategorized"):
            if "Uncategorized" in self.engine.categories:
                del self.engine.categories["Uncategorized"]

        self._refresh_tree()
        self._refresh_combo()

    def _on_ai_done(self):
        self.ai_btn.setEnabled(True)
        self.ai_btn.setText("Classify Uncategorized with AI")
        s = self.engine.get_summary()
        self.summary_label.setText(
            f"{s['total']:,} emails | {s['categorized']:,} categorized | "
            f"{s['uncategorized']:,} uncategorized"
        )

    def _on_ai_error(self, err):
        self.ai_btn.setEnabled(True)
        self.ai_btn.setText("Classify Uncategorized with AI")
        QMessageBox.warning(self, "AI Error", str(err))


# ─── UI: Execute Page ────────────────────────────────────────────────────

class ExecutePage(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        title = QLabel("Applying Labels")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        layout.addWidget(title)

        self.status_label = QLabel("Starting...")
        self.status_label.setStyleSheet(f"color: {C.SUBTEXT0};")
        layout.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(8)
        layout.addWidget(self.progress)

        self.percent_label = QLabel("0%")
        self.percent_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.percent_label)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet(
            f"font-family: 'Cascadia Code', 'Consolas', monospace; font-size: 12px;"
        )
        layout.addWidget(self.log_text, 1)

        btn_layout = QHBoxLayout()
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setProperty("danger", True)
        self.stop_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.addStretch()

        self.done_label = QLabel("")
        btn_layout.addWidget(self.done_label)
        layout.addLayout(btn_layout)

    def start_labeling(self, host, email_addr, password, categories, prefix, archive):
        self.worker = ImapLabelWorker(host, email_addr, password, categories, prefix, archive)
        self.worker.progress.connect(self._on_progress)
        self.worker.status.connect(self._on_status)
        self.worker.log.connect(self._on_log)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.stop_btn.clicked.connect(self.worker.stop)
        self.worker.start()

    def _on_progress(self, current, total):
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        pct = int(current / total * 100) if total > 0 else 0
        self.percent_label.setText(f"{pct}% ({current:,} / {total:,})")

    def _on_status(self, msg):
        self.status_label.setText(msg)

    def _on_log(self, msg):
        self.log_text.appendPlainText(msg)

    def _on_finished(self):
        self.status_label.setText("All labels applied successfully!")
        self.status_label.setStyleSheet(f"color: {C.GREEN};")
        self.done_label.setText("Complete!")
        self.done_label.setStyleSheet(f"color: {C.GREEN}; font-size: 16px; font-weight: bold;")
        self.stop_btn.setEnabled(False)

    def _on_error(self, err):
        self.status_label.setText("Error occurred")
        self.status_label.setStyleSheet(f"color: {C.RED};")
        self._on_log(f"ERROR: {err}")


# ─── Main Window ──────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"InboxForge v{VERSION}")
        self.setMinimumSize(1000, 650)
        self.resize(1200, 800)

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.connect_page = ConnectPage()
        self.analyze_page = AnalyzePage()
        self.review_page = ReviewPage()
        self.execute_page = ExecutePage()

        self.stack.addWidget(self.connect_page)
        self.stack.addWidget(self.analyze_page)
        self.stack.addWidget(self.review_page)
        self.stack.addWidget(self.execute_page)

        self.connect_page.connected.connect(self._on_connected)
        self.analyze_page.analysis_complete.connect(self._show_review)
        self.review_page.execute_requested.connect(self._start_execute)

    def _on_connected(self):
        if self.connect_page.loaded_engine:
            # Loaded from file — skip scan, go straight to review
            self.analyze_page.set_preloaded(self.connect_page.loaded_engine)
            self.stack.setCurrentWidget(self.analyze_page)
        else:
            self.stack.setCurrentWidget(self.analyze_page)
            self.analyze_page.start_scan(
                self.connect_page.imap_host,
                self.connect_page.email_addr,
                self.connect_page.password,
            )

    def _show_review(self):
        self.review_page.load_categories(self.analyze_page.engine)
        self.stack.setCurrentWidget(self.review_page)

    def _start_execute(self):
        engine = self.review_page.engine
        if not engine:
            return

        categories = {k: v for k, v in engine.categories.items() if v and k != "Uncategorized"}

        if not categories:
            QMessageBox.warning(self, "Nothing to Apply", "No categories with emails to label.")
            return

        total = sum(len(v) for v in categories.values())

        # Check credentials
        if not self.connect_page.email_addr or not self.connect_page.password:
            QMessageBox.warning(
                self, "Credentials Needed",
                "Enter your Gmail address and app password on the Connect page to apply labels."
            )
            return

        self.stack.setCurrentWidget(self.execute_page)
        self.execute_page.start_labeling(
            self.connect_page.imap_host,
            self.connect_page.email_addr,
            self.connect_page.password,
            categories,
            self.review_page.prefix_input.text().strip(),
            self.review_page.archive_check.isChecked(),
        )


# ─── Entry Point ──────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
