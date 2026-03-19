#!/usr/bin/env python3
"""InboxForge v0.3.0 — Full Gmail Mailbox Downloader & AI-Powered Organizer"""

VERSION = "0.3.0"

import sys, os, subprocess

def _bootstrap():
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
import shutil
import hashlib
import traceback
from collections import Counter, defaultdict
from datetime import datetime
from dataclasses import dataclass
from typing import Optional
from pathlib import Path

from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

# Gmail IMAP folders to skip by default (All Mail duplicates everything)
GMAIL_SKIP_FOLDERS = {'[Gmail]/All Mail', '[Gmail]/Important', '[Gmail]/Spam', '[Gmail]/Trash'}


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
    QCheckBox, QRadioButton {{
        color: {C.TEXT};
        spacing: 8px;
    }}
    QCheckBox::indicator, QRadioButton::indicator {{
        width: 18px;
        height: 18px;
        border-radius: 4px;
        border: 2px solid {C.SURFACE2};
        background-color: {C.SURFACE0};
    }}
    QRadioButton::indicator {{
        border-radius: 9px;
    }}
    QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
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
    QListWidget {{
        background-color: {C.MANTLE};
        color: {C.TEXT};
        border: 1px solid {C.SURFACE0};
        border-radius: 6px;
        outline: none;
    }}
    QListWidget::item {{
        padding: 4px;
    }}
    QListWidget::item:selected {{
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
    local_path: str = ""       # Path to .eml file on disk
    source_folder: str = ""    # Original IMAP folder name
    message_id: str = ""       # For deduplication across folders


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
        if self.user_domain and domain == self.user_domain:
            return "Work/Internal", 0.95
        if domain in DOMAIN_CATEGORIES:
            return DOMAIN_CATEGORIES[domain], 0.9
        if em.has_list_unsubscribe:
            return "Newsletters", 0.85
        for cat, patterns in SUBJECT_PATTERNS.items():
            for pattern in patterns:
                if re.search(pattern, em.subject):
                    return cat, 0.7
        sender_lower = em.sender.lower()
        if re.search(r'(?i)(no-?reply|noreply|notifications?@|alerts?@|mailer-daemon)', sender_lower):
            return "Automated/Notifications", 0.6
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

        folder_counts = Counter(em.source_folder for em in self.emails if em.source_folder)
        return {
            'total': total,
            'categorized': categorized,
            'uncategorized': total - categorized,
            'categories': {k: len(v) for k, v in sorted(self.categories.items(), key=lambda x: -len(x[1]))},
            'top_domains': self.domain_stats.most_common(20),
            'date_range': date_range,
            'folder_counts': dict(folder_counts.most_common()),
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
        data = {
            'version': VERSION,
            'user_domain': self.user_domain,
            'emails': [
                {
                    'uid': em.uid, 'sender': em.sender, 'sender_name': em.sender_name,
                    'sender_domain': em.sender_domain, 'subject': em.subject,
                    'date': em.date, 'has_list_unsubscribe': em.has_list_unsubscribe,
                    'category': em.category, 'confidence': em.confidence,
                    'local_path': em.local_path, 'source_folder': em.source_folder,
                    'message_id': em.message_id,
                }
                for em in self.emails
            ],
        }
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)

    def load_state(self, path: str) -> bool:
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
                    local_path=ed.get('local_path', ''),
                    source_folder=ed.get('source_folder', ''),
                    message_id=ed.get('message_id', ''),
                )
                self.emails.append(em)
                if em.category:
                    self.categories[em.category].append(em)
            self.domain_stats = Counter(em.sender_domain for em in self.emails)
            return True
        except Exception:
            return False


# ─── Helpers ──────────────────────────────────────────────────────────────

def decode_header(raw):
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


def parse_date(date_str):
    if not date_str:
        return None
    try:
        parsed = email.utils.parsedate_to_datetime(date_str)
        return parsed.replace(tzinfo=None)
    except Exception:
        return None


def sanitize_filename(s: str, max_len: int = 60) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', s)
    s = re.sub(r'_+', '_', s).strip('_. ')
    return s[:max_len] if s else "untitled"


def sanitize_folder_name(imap_name: str) -> str:
    """Convert IMAP folder name to safe local directory name."""
    # Strip [Gmail]/ prefix
    name = re.sub(r'^\[Gmail\]/', '', imap_name)
    # Sanitize for filesystem
    name = re.sub(r'[<>:"|?*\x00-\x1f]', '_', name)
    return name.strip('_. ') or "Other"


def parse_imap_folder_list(line: bytes) -> Optional[str]:
    """Parse a folder name from IMAP LIST response."""
    # Format: (\\flags) "delimiter" "folder_name"
    match = re.match(rb'\(.*?\)\s+"(.?)"\s+"?(.+?)"?\s*$', line)
    if match:
        name = match.group(2)
        try:
            return name.decode('utf-7').replace('&', '+').encode('ascii').decode('utf-7')
        except Exception:
            pass
        try:
            return name.decode('utf-8')
        except Exception:
            return name.decode('ascii', errors='replace')
    # Fallback: try simple decode
    try:
        parts = line.decode('utf-8').split('"')
        if len(parts) >= 4:
            return parts[3] if parts[3] else parts[-2]
    except Exception:
        pass
    return None


# ─── IMAP Workers ─────────────────────────────────────────────────────────

class ImapScanWorker(QThread):
    """Scans headers only from INBOX (fast mode, no download)."""
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

    def run(self):
        try:
            self.status.emit("Connecting to Gmail IMAP...")
            imap = imaplib.IMAP4_SSL(self.host, 993)
            imap.login(self.email_addr, self.password)
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
                    '(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE LIST-UNSUBSCRIBE MESSAGE-ID)])'
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
                                from_decoded = decode_header(msg.get('From', ''))
                                name, addr = email.utils.parseaddr(from_decoded)
                                em = EmailInfo(
                                    uid=uid,
                                    sender=addr or from_decoded,
                                    sender_name=name or addr or from_decoded,
                                    subject=decode_header(msg.get('Subject', '(no subject)')),
                                    date=msg.get('Date', ''),
                                    date_parsed=parse_date(msg.get('Date', '')),
                                    has_list_unsubscribe=msg.get('List-Unsubscribe') is not None,
                                    source_folder='INBOX',
                                    message_id=msg.get('Message-ID', ''),
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

        except Exception as e:
            self.error.emit(f"Error: {e}\n{traceback.format_exc()}")


class ImapDownloadWorker(QThread):
    """Downloads full .eml files from ALL Gmail folders with resume + dedup."""
    progress = pyqtSignal(int, int)     # current, total
    status = pyqtSignal(str)
    log = pyqtSignal(str)
    folder_started = pyqtSignal(str, int)  # folder_name, count
    email_saved = pyqtSignal(object)    # EmailInfo
    finished_signal = pyqtSignal(list)  # all EmailInfo
    error = pyqtSignal(str)

    def __init__(self, host, email_addr, password, output_dir, skip_folders=None):
        super().__init__()
        self.host = host
        self.email_addr = email_addr
        self.password = password
        self.output_dir = Path(output_dir)
        self.skip_folders = skip_folders or GMAIL_SKIP_FOLDERS
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            folders_dir = self.output_dir / "folders"
            folders_dir.mkdir(parents=True, exist_ok=True)

            # Load manifest for resume
            manifest_path = self.output_dir / "manifest.json"
            manifest = {'folders': {}, 'message_ids': {}}
            if manifest_path.exists():
                try:
                    with open(manifest_path, 'r', encoding='utf-8') as f:
                        manifest = json.load(f)
                    if 'folders' not in manifest:
                        manifest = {'folders': {}, 'message_ids': {}}
                except Exception:
                    manifest = {'folders': {}, 'message_ids': {}}

            # message_id -> local_path for dedup
            seen_ids = manifest.get('message_ids', {})

            self.status.emit("Connecting to Gmail IMAP...")
            imap = imaplib.IMAP4_SSL(self.host, 993)
            imap.login(self.email_addr, self.password)

            # List all folders
            self.status.emit("Listing mailbox folders...")
            _, folder_data = imap.list()
            all_folders = []
            for line in folder_data:
                if isinstance(line, bytes):
                    name = parse_imap_folder_list(line)
                    if name:
                        all_folders.append(name)

            # Filter out skipped folders
            active_folders = [f for f in all_folders if f not in self.skip_folders]
            self.log.emit(f"Found {len(all_folders)} folders, downloading {len(active_folders)}:")
            for f in active_folders:
                self.log.emit(f"  {f}")
            if self.skip_folders:
                self.log.emit(f"Skipping: {', '.join(self.skip_folders)}")

            # First pass: count total emails across all folders
            self.status.emit("Counting emails across all folders...")
            folder_uids = {}
            total_count = 0
            for folder_name in active_folders:
                if self._stop:
                    break
                try:
                    result = imap.select(f'"{folder_name}"', readonly=True)
                    if result[0] != 'OK':
                        continue
                    _, data = imap.uid('SEARCH', None, 'ALL')
                    uids = data[0].split() if data[0] else []
                    folder_uids[folder_name] = uids
                    total_count += len(uids)
                    self.log.emit(f"  {folder_name}: {len(uids):,} messages")
                except Exception as e:
                    self.log.emit(f"  {folder_name}: error listing - {e}")

            self.log.emit(f"\nTotal messages across folders: {total_count:,}")
            self.log.emit("(Emails in multiple labels will be deduplicated)\n")

            all_emails = []
            global_processed = 0

            # Rebuild EmailInfo for already-downloaded emails from manifest
            for folder_name, folder_manifest in manifest.get('folders', {}).items():
                for uid_str, info in folder_manifest.items():
                    em = EmailInfo(
                        uid=f"{folder_name}:{uid_str}",
                        sender=info.get('sender', ''),
                        sender_name=info.get('sender_name', ''),
                        subject=info.get('subject', ''),
                        date=info.get('date', ''),
                        date_parsed=parse_date(info.get('date', '')),
                        has_list_unsubscribe=info.get('has_list_unsubscribe', False),
                        local_path=info.get('local_path', ''),
                        source_folder=folder_name,
                        message_id=info.get('message_id', ''),
                    )
                    all_emails.append(em)

            # Download from each folder
            for folder_name, uids in folder_uids.items():
                if self._stop:
                    break
                if not uids:
                    continue

                safe_name = sanitize_folder_name(folder_name)
                folder_dir = folders_dir / safe_name
                folder_dir.mkdir(parents=True, exist_ok=True)

                # Get already-downloaded UIDs for this folder
                folder_manifest = manifest['folders'].get(folder_name, {})
                remaining = [u for u in uids if u.decode() not in folder_manifest]

                already = len(uids) - len(remaining)
                self.folder_started.emit(folder_name, len(uids))
                if already > 0:
                    self.log.emit(
                        f"[{safe_name}] {already:,} already downloaded, "
                        f"{len(remaining):,} remaining"
                    )
                    global_processed += already
                else:
                    self.log.emit(f"[{safe_name}] Downloading {len(uids):,} emails...")

                imap.select(f'"{folder_name}"', readonly=True)

                batch_size = 50
                for i in range(0, len(remaining), batch_size):
                    if self._stop:
                        break

                    batch_uids = remaining[i:i + batch_size]
                    uid_range = b','.join(batch_uids)

                    try:
                        _, msg_data = imap.uid('FETCH', uid_range, '(RFC822)')
                    except Exception as e:
                        self.log.emit(f"  [{safe_name}] Fetch error: {e}")
                        global_processed += len(batch_uids)
                        continue

                    for j in range(len(msg_data)):
                        if self._stop:
                            break
                        item = msg_data[j]
                        if not isinstance(item, tuple) or len(item) != 2:
                            continue

                        uid_match = re.search(rb'UID (\d+)', item[0])
                        if not uid_match:
                            continue
                        uid = uid_match.group(1).decode()

                        raw = item[1]
                        if not isinstance(raw, bytes):
                            continue

                        # Parse headers for metadata + dedup
                        try:
                            msg = email.message_from_bytes(raw)
                            msg_id = msg.get('Message-ID', '')
                            from_decoded = decode_header(msg.get('From', ''))
                            name, addr = email.utils.parseaddr(from_decoded)

                            em_sender = addr or from_decoded
                            em_sender_name = name or addr or from_decoded
                            em_subject = decode_header(msg.get('Subject', '(no subject)'))
                            em_date = msg.get('Date', '')
                            em_has_unsub = msg.get('List-Unsubscribe') is not None
                        except Exception:
                            msg_id = ''
                            em_sender = em_sender_name = em_subject = em_date = ''
                            em_has_unsub = False

                        # Dedup: if we've seen this Message-ID, reference existing file
                        if msg_id and msg_id in seen_ids:
                            existing_path = seen_ids[msg_id]
                            eml_path = existing_path  # reference same file
                        else:
                            # Save .eml file
                            eml_path = str(folder_dir / f"{uid}.eml")
                            try:
                                with open(eml_path, 'wb') as f:
                                    f.write(raw)
                            except Exception:
                                continue

                            if msg_id:
                                seen_ids[msg_id] = eml_path

                        em = EmailInfo(
                            uid=f"{folder_name}:{uid}",
                            sender=em_sender,
                            sender_name=em_sender_name,
                            subject=em_subject,
                            date=em_date,
                            date_parsed=parse_date(em_date),
                            has_list_unsubscribe=em_has_unsub,
                            local_path=eml_path,
                            source_folder=folder_name,
                            message_id=msg_id,
                        )
                        all_emails.append(em)
                        self.email_saved.emit(em)

                        # Update folder manifest
                        if folder_name not in manifest['folders']:
                            manifest['folders'][folder_name] = {}
                        manifest['folders'][folder_name][uid] = {
                            'sender': em.sender, 'sender_name': em.sender_name,
                            'subject': em.subject, 'date': em.date,
                            'has_list_unsubscribe': em.has_list_unsubscribe,
                            'local_path': eml_path, 'message_id': msg_id,
                        }

                    global_processed += len(batch_uids)
                    self.progress.emit(global_processed, total_count)

                # Save manifest after each folder
                manifest['message_ids'] = seen_ids
                try:
                    with open(manifest_path, 'w', encoding='utf-8') as f:
                        json.dump(manifest, f, ensure_ascii=False)
                except Exception:
                    pass

                self.log.emit(f"  [{safe_name}] Done.")

            # Final manifest save
            manifest['message_ids'] = seen_ids
            try:
                with open(manifest_path, 'w', encoding='utf-8') as f:
                    json.dump(manifest, f, ensure_ascii=False)
            except Exception:
                pass

            try:
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
                label_q = f'"{label}"'

                self.log.emit(f"Creating label: {label}")
                try:
                    imap.create(label_q)
                except Exception:
                    pass

                # Group emails by source folder for correct UID context
                by_folder = defaultdict(list)
                for em in emails:
                    folder = em.source_folder or 'INBOX'
                    # UID is stored as "folder:uid" — extract the raw uid
                    raw_uid = em.uid.split(':', 1)[1] if ':' in em.uid else em.uid
                    by_folder[folder].append(raw_uid)

                for folder, uid_list in by_folder.items():
                    if self._stop:
                        break
                    imap.select(f'"{folder}"')

                    batch_size = 100
                    for i in range(0, len(uid_list), batch_size):
                        if self._stop:
                            break
                        batch = uid_list[i:i + batch_size]
                        uid_str = ','.join(batch)
                        try:
                            imap.uid('COPY', uid_str, label_q)
                        except Exception as e:
                            self.log.emit(f"  Error: {e}")

                        processed += len(batch)
                        self.progress.emit(processed, total_emails)

                self.log.emit(f"  Labeled {len(emails):,} as '{label}'")

                if self.archive_after and not self._stop:
                    inbox_uids = by_folder.get('INBOX', [])
                    if inbox_uids:
                        self.log.emit(f"  Archiving {len(inbox_uids)} from Inbox...")
                        imap.select('INBOX')
                        for i in range(0, len(inbox_uids), 100):
                            batch = inbox_uids[i:i + 100]
                            try:
                                imap.uid('STORE', ','.join(batch), '+FLAGS', '(\\Deleted)')
                            except Exception:
                                pass
                        imap.expunge()

            try:
                imap.logout()
            except Exception:
                pass
            self.finished_signal.emit()

        except Exception as e:
            self.error.emit(f"Error: {e}\n{traceback.format_exc()}")


class LocalOrganizeWorker(QThread):
    """Organizes downloaded .eml files: preserves original folders + creates categorized structure."""
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    log = pyqtSignal(str)
    finished_signal = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, categories, output_dir, copy_mode=True):
        super().__init__()
        self.categories = categories
        self.output_dir = Path(output_dir)
        self.copy_mode = copy_mode
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            organized_dir = self.output_dir / "organized"
            organized_dir.mkdir(parents=True, exist_ok=True)

            total = sum(len(v) for v in self.categories.values())
            processed = 0

            self.log.emit(f"Organizing {total:,} emails into {len(self.categories)} categories")
            self.log.emit(f"Output: {organized_dir}")
            self.log.emit(f"Mode: {'copy' if self.copy_mode else 'move'}")
            self.log.emit("")

            for cat_name, emails in self.categories.items():
                if self._stop:
                    break
                if not emails:
                    continue

                # Nested categories: Shopping/Amazon -> organized/Shopping/Amazon/
                cat_parts = cat_name.replace('/', os.sep)
                cat_folder = organized_dir / sanitize_filename(cat_parts, 120)
                cat_folder.mkdir(parents=True, exist_ok=True)
                self.log.emit(f"[{cat_name}] {len(emails):,} emails")

                for em in emails:
                    if self._stop:
                        break

                    src = Path(em.local_path) if em.local_path else None
                    if not src or not src.exists():
                        processed += 1
                        continue

                    # Readable filename: YYYY-MM-DD_sender_subject.eml
                    date_str = em.date_parsed.strftime("%Y-%m-%d") if em.date_parsed else "unknown"
                    sender_part = sanitize_filename(em.sender_domain or em.sender_name, 25)
                    subject_part = sanitize_filename(em.subject, 45)
                    filename = f"{date_str}_{sender_part}_{subject_part}.eml"

                    dst = cat_folder / filename
                    counter = 1
                    while dst.exists():
                        dst = cat_folder / f"{date_str}_{sender_part}_{subject_part}_{counter}.eml"
                        counter += 1

                    try:
                        if self.copy_mode:
                            shutil.copy2(str(src), str(dst))
                        else:
                            shutil.move(str(src), str(dst))
                    except Exception as e:
                        self.log.emit(f"  Error: {em.uid} - {e}")

                    processed += 1
                    if processed % 500 == 0:
                        self.progress.emit(processed, total)

                self.progress.emit(processed, total)

            self.log.emit("")
            self.log.emit(f"Organized {processed:,} emails into {organized_dir}")
            self.progress.emit(total, total)
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
    connected = pyqtSignal(str)  # mode: "scan", "download", "load"

    def __init__(self):
        super().__init__()
        self.imap_host = "imap.gmail.com"
        self.email_addr = ""
        self.password = ""
        self.download_dir = ""
        self.loaded_engine = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(16)

        title = QLabel("InboxForge")
        title.setStyleSheet(f"font-size: 32px; color: {C.BLUE}; font-weight: bold;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel(f"v{VERSION} — Full Gmail Mailbox Downloader & AI Organizer")
        subtitle.setStyleSheet(f"color: {C.SUBTEXT0}; font-size: 12px;")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(subtitle)
        layout.addSpacing(20)

        form = QWidget()
        form.setMaximumWidth(520)
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

        hint = QLabel("Generate at: Google Account > Security > 2-Step Verification > App Passwords")
        hint.setStyleSheet(f"color: {C.SUBTEXT0}; font-size: 11px;")
        hint.setWordWrap(True)
        fl.addWidget(hint)
        fl.addSpacing(12)

        # Action buttons
        row1 = QHBoxLayout()
        self.download_btn = QPushButton("Download Full Mailbox")
        self.download_btn.setToolTip(
            "Downloads ALL folders (Inbox, Sent, Drafts, Starred, labels) as .eml files.\n"
            "Preserves original folder structure. Resumable. Deduplicates across labels."
        )
        self.download_btn.setStyleSheet(f"background-color: {C.GREEN};")
        self.download_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.download_btn.clicked.connect(lambda: self._on_action("download"))
        row1.addWidget(self.download_btn)

        self.scan_btn = QPushButton("Scan Inbox Headers Only")
        self.scan_btn.setProperty("secondary", True)
        self.scan_btn.setToolTip("Fast scan of Inbox headers only. No files downloaded.")
        self.scan_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.scan_btn.clicked.connect(lambda: self._on_action("scan"))
        row1.addWidget(self.scan_btn)
        fl.addLayout(row1)

        row2 = QHBoxLayout()
        self.load_scan_btn = QPushButton("Load Previous Scan")
        self.load_scan_btn.setProperty("secondary", True)
        self.load_scan_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.load_scan_btn.clicked.connect(self._on_load_scan)
        row2.addWidget(self.load_scan_btn)

        self.load_local_btn = QPushButton("Load Downloaded Mailbox")
        self.load_local_btn.setProperty("secondary", True)
        self.load_local_btn.setToolTip("Resume from a previous download folder")
        self.load_local_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.load_local_btn.clicked.connect(self._on_load_local)
        row2.addWidget(self.load_local_btn)
        fl.addLayout(row2)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet(f"color: {C.SUBTEXT0}; font-size: 12px;")
        self.status_label.setWordWrap(True)
        fl.addWidget(self.status_label)

        layout.addWidget(form, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()

    def _set_buttons_enabled(self, enabled):
        for btn in (self.scan_btn, self.download_btn, self.load_scan_btn, self.load_local_btn):
            btn.setEnabled(enabled)

    def _on_action(self, mode):
        self.email_addr = self.email_input.text().strip()
        self.password = self.pass_input.text().strip()
        if not self.email_addr or not self.password:
            self.status_label.setText("Please enter email and app password")
            self.status_label.setStyleSheet(f"color: {C.RED};")
            return

        if mode == "download":
            default_dir = str(Path.home() / "Desktop" / "InboxForge")
            folder = QFileDialog.getExistingDirectory(self, "Choose Download Folder", default_dir)
            if not folder:
                return
            self.download_dir = folder

        self.status_label.setText("Testing connection...")
        self.status_label.setStyleSheet(f"color: {C.YELLOW};")
        self._set_buttons_enabled(False)
        self._pending_mode = mode

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
        self.connected.emit(self._pending_mode)

    def _on_test_error(self, err):
        self._test_thread.quit()
        self.status_label.setText(err)
        self.status_label.setStyleSheet(f"color: {C.RED};")
        self._set_buttons_enabled(True)

    def _on_load_scan(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load Scan File", "", "JSON (*.json)")
        if not path:
            return
        engine = CategoryEngine()
        if engine.load_state(path):
            self.loaded_engine = engine
            self.email_addr = self.email_input.text().strip()
            self.password = self.pass_input.text().strip()
            self.status_label.setText(f"Loaded {len(engine.emails):,} emails")
            self.status_label.setStyleSheet(f"color: {C.GREEN};")
            self.connected.emit("load")
        else:
            self.status_label.setText("Failed to load scan file")
            self.status_label.setStyleSheet(f"color: {C.RED};")

    def _on_load_local(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Download Folder")
        if not folder:
            return
        self.download_dir = folder
        manifest_path = Path(folder) / "manifest.json"
        if not manifest_path.exists():
            self.status_label.setText("No manifest.json found — not a valid InboxForge download.")
            self.status_label.setStyleSheet(f"color: {C.RED};")
            return

        self.status_label.setText("Loading downloaded emails...")
        self.status_label.setStyleSheet(f"color: {C.YELLOW};")
        QApplication.processEvents()

        try:
            with open(manifest_path, 'r', encoding='utf-8') as f:
                manifest = json.load(f)

            user_domain = self.email_input.text().strip()
            user_domain = user_domain.split('@')[1] if '@' in user_domain else ""
            engine = CategoryEngine(user_domain)

            emails = []
            seen_msg_ids = set()
            for folder_name, folder_data in manifest.get('folders', {}).items():
                for uid_str, info in folder_data.items():
                    msg_id = info.get('message_id', '')
                    # Dedup: only process unique Message-IDs for categorization
                    if msg_id and msg_id in seen_msg_ids:
                        continue
                    if msg_id:
                        seen_msg_ids.add(msg_id)

                    em = EmailInfo(
                        uid=f"{folder_name}:{uid_str}",
                        sender=info.get('sender', ''),
                        sender_name=info.get('sender_name', ''),
                        subject=info.get('subject', ''),
                        date=info.get('date', ''),
                        date_parsed=parse_date(info.get('date', '')),
                        has_list_unsubscribe=info.get('has_list_unsubscribe', False),
                        local_path=info.get('local_path', ''),
                        source_folder=folder_name,
                        message_id=msg_id,
                    )
                    emails.append(em)

            engine.process_all(emails)
            self.loaded_engine = engine
            self.status_label.setText(
                f"Loaded {len(emails):,} unique emails from "
                f"{len(manifest.get('folders', {}))} folders"
            )
            self.status_label.setStyleSheet(f"color: {C.GREEN};")
            self.connected.emit("load")

        except Exception as e:
            self.status_label.setText(f"Error: {e}")
            self.status_label.setStyleSheet(f"color: {C.RED};")


# ─── UI: Download Page ───────────────────────────────────────────────────

class DownloadPage(QWidget):
    download_complete = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.engine = None
        self.worker = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        title = QLabel("Downloading Full Mailbox")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        layout.addWidget(title)

        self.status_label = QLabel("Preparing...")
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

        bottom = QHBoxLayout()
        self.stop_btn = QPushButton("Stop (Resumable)")
        self.stop_btn.setProperty("danger", True)
        self.stop_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        bottom.addWidget(self.stop_btn)
        bottom.addStretch()

        self.size_label = QLabel("")
        self.size_label.setStyleSheet(f"color: {C.SUBTEXT0};")
        bottom.addWidget(self.size_label)

        self.continue_btn = QPushButton("Continue to Analysis")
        self.continue_btn.setVisible(False)
        self.continue_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.continue_btn.clicked.connect(self.download_complete.emit)
        bottom.addWidget(self.continue_btn)
        layout.addLayout(bottom)

    def start_download(self, host, email_addr, password, output_dir):
        user_domain = email_addr.split('@')[1] if '@' in email_addr else ""
        self.engine = CategoryEngine(user_domain)
        self._email_count = 0
        self._output_dir = output_dir

        self.log_text.setPlainText(
            f"Download folder: {output_dir}\n"
            f"Structure: {output_dir}/folders/<FolderName>/<uid>.eml\n"
            f"Manifest: {output_dir}/manifest.json\n"
            f"Skipping: {', '.join(GMAIL_SKIP_FOLDERS)}\n"
            f"Resume: stop anytime, already-downloaded emails are tracked.\n"
            f"{'='*60}\n"
        )

        self.worker = ImapDownloadWorker(host, email_addr, password, output_dir)
        self.worker.progress.connect(self._on_progress)
        self.worker.status.connect(self._on_status)
        self.worker.log.connect(self._on_log)
        self.worker.email_saved.connect(self._on_email_saved)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.stop_btn.clicked.connect(self._on_stop)
        self.worker.start()

    def _on_stop(self):
        if self.worker:
            self.worker.stop()
        self.status_label.setText("Stopping... manifest saved. Resume anytime.")
        self.status_label.setStyleSheet(f"color: {C.YELLOW};")
        self.continue_btn.setVisible(True)
        self.continue_btn.setText("Continue with partial download")

    def _on_progress(self, current, total):
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        pct = int(current / total * 100) if total > 0 else 0
        self.percent_label.setText(f"{pct}% ({current:,} / {total:,})")

    def _on_status(self, msg):
        self.status_label.setText(msg)

    def _on_log(self, msg):
        self.log_text.appendPlainText(msg)

    def _on_email_saved(self, em):
        self._email_count += 1
        if self._email_count % 200 == 0:
            # Show disk usage
            folders_dir = Path(self._output_dir) / "folders"
            if folders_dir.exists():
                try:
                    total_bytes = sum(
                        f.stat().st_size
                        for f in folders_dir.rglob('*.eml')
                    )
                    size_mb = total_bytes / (1024 * 1024)
                    self.size_label.setText(f"{size_mb:,.0f} MB downloaded")
                except Exception:
                    pass

    def _on_finished(self, all_emails):
        # Dedup for categorization
        seen = set()
        unique = []
        for em in all_emails:
            if em.message_id and em.message_id in seen:
                continue
            if em.message_id:
                seen.add(em.message_id)
            unique.append(em)

        self.engine.process_all(unique)

        summary = self.engine.get_summary()
        self.log_text.appendPlainText(f"\n{'='*60}")
        self.log_text.appendPlainText(f"Download complete!")
        self.log_text.appendPlainText(f"Total unique emails: {summary['total']:,}")
        self.log_text.appendPlainText(f"Date range: {summary['date_range'][0]} to {summary['date_range'][1]}")
        if summary.get('folder_counts'):
            self.log_text.appendPlainText(f"\nEmails per folder:")
            for fname, count in summary['folder_counts'].items():
                self.log_text.appendPlainText(f"  {fname:40s} {count:>6,}")

        self.status_label.setText("Download complete!")
        self.status_label.setStyleSheet(f"color: {C.GREEN};")
        self.continue_btn.setVisible(True)
        self.stop_btn.setEnabled(False)

        # Final disk size
        folders_dir = Path(self._output_dir) / "folders"
        if folders_dir.exists():
            try:
                total_bytes = sum(f.stat().st_size for f in folders_dir.rglob('*.eml'))
                self.size_label.setText(f"Total: {total_bytes / (1024*1024):,.0f} MB")
            except Exception:
                pass

    def _on_error(self, err):
        self.status_label.setText("Error occurred")
        self.status_label.setStyleSheet(f"color: {C.RED};")
        self.log_text.appendPlainText(f"ERROR: {err}")
        self.continue_btn.setVisible(True)
        self.continue_btn.setText("Continue with partial download")


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

        title = QLabel("Analyzing Mailbox")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        layout.addWidget(title)

        self.status_label = QLabel("Preparing...")
        self.status_label.setStyleSheet(f"color: {C.SUBTEXT0};")
        layout.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(8)
        layout.addWidget(self.progress)

        self.percent_label = QLabel("0%")
        self.percent_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.percent_label)

        stats_group = QGroupBox("Analysis Results")
        stats_layout = QVBoxLayout(stats_group)
        self.stats_text = QPlainTextEdit()
        self.stats_text.setReadOnly(True)
        self.stats_text.setStyleSheet(
            f"font-family: 'Cascadia Code', 'Consolas', monospace; font-size: 12px;"
        )
        stats_layout.addWidget(self.stats_text)
        layout.addWidget(stats_group)

        layout.addStretch()

        bottom = QHBoxLayout()
        self.save_btn = QPushButton("Save Analysis")
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
        self.status_label.setText("Categorizing...")
        self.status_label.setStyleSheet(f"color: {C.YELLOW};")
        QApplication.processEvents()
        self.engine.process_all(all_emails)
        self._show_summary()

    def _show_summary(self):
        s = self.engine.get_summary()
        lines = [
            f"Total unique emails: {s['total']:,}",
            f"Auto-categorized:    {s['categorized']:,}",
            f"Uncategorized:       {s['uncategorized']:,}",
            f"Date range:          {s['date_range'][0]} to {s['date_range'][1]}",
        ]
        if s.get('folder_counts'):
            lines.append("")
            lines.append("Source Folders:")
            for fname, count in s['folder_counts'].items():
                lines.append(f"  {fname:40s} {count:>6,}")
        lines.append("")
        lines.append("Categories:")
        for cat, count in s['categories'].items():
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
        path, _ = QFileDialog.getSaveFileName(self, "Save", "inboxforge_scan.json", "JSON (*.json)")
        if path and self.engine:
            self.engine.save_state(path)
            self.status_label.setText(f"Saved to {os.path.basename(path)}")


# ─── UI: Review Page ─────────────────────────────────────────────────────

class ReviewPage(QWidget):
    execute_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.engine: Optional[CategoryEngine] = None
        self.has_local_files = False
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

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

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left: tree
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
        self.email_table.setColumnCount(5)
        self.email_table.setHorizontalHeaderLabels(["From", "Subject", "Date", "Folder", "Conf"])
        header = self.email_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.email_table.setColumnWidth(0, 180)
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

        # Bottom: execution options
        bottom = QVBoxLayout()
        bottom.setSpacing(8)

        mode_row = QHBoxLayout()
        mode_label = QLabel("Execute mode:")
        mode_label.setStyleSheet("font-weight: bold;")
        mode_row.addWidget(mode_label)

        self.mode_local = QRadioButton("Organize Local Files (safe, no mailbox changes)")
        self.mode_gmail = QRadioButton("Apply Gmail Labels (modifies live mailbox)")
        self.mode_local.setChecked(True)
        mode_row.addWidget(self.mode_local)
        mode_row.addWidget(self.mode_gmail)
        mode_row.addStretch()
        bottom.addLayout(mode_row)

        opts_row = QHBoxLayout()

        # Gmail options
        self.gmail_opts = QWidget()
        gl = QHBoxLayout(self.gmail_opts)
        gl.setContentsMargins(0, 0, 0, 0)
        gl.addWidget(QLabel("Label prefix:"))
        self.prefix_input = QLineEdit()
        self.prefix_input.setPlaceholderText("optional")
        self.prefix_input.setMaximumWidth(180)
        gl.addWidget(self.prefix_input)
        self.archive_check = QCheckBox("Archive from Inbox after labeling")
        gl.addWidget(self.archive_check)
        opts_row.addWidget(self.gmail_opts)

        # Local options
        self.local_opts = QWidget()
        ll = QHBoxLayout(self.local_opts)
        ll.setContentsMargins(0, 0, 0, 0)
        self.copy_radio = QRadioButton("Copy .eml files")
        self.copy_radio.setChecked(True)
        self.move_radio = QRadioButton("Move .eml files")
        ll.addWidget(self.copy_radio)
        ll.addWidget(self.move_radio)
        ll.addWidget(QLabel("Into:"))
        self.local_dir_label = QLabel("organized/")
        self.local_dir_label.setStyleSheet(f"color: {C.SUBTEXT0};")
        ll.addWidget(self.local_dir_label, 1)
        opts_row.addWidget(self.local_opts)

        opts_row.addStretch()
        self.execute_btn = QPushButton("Organize Local Files")
        self.execute_btn.setStyleSheet(
            f"background-color: {C.GREEN}; font-size: 14px; padding: 10px 30px;"
        )
        self.execute_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.execute_btn.clicked.connect(self.execute_requested.emit)
        opts_row.addWidget(self.execute_btn)

        bottom.addLayout(opts_row)
        layout.addLayout(bottom)

        self.mode_gmail.toggled.connect(self._on_mode_changed)
        self.mode_local.toggled.connect(self._on_mode_changed)
        self._on_mode_changed()

    def _on_mode_changed(self):
        is_gmail = self.mode_gmail.isChecked()
        self.gmail_opts.setVisible(is_gmail)
        self.local_opts.setVisible(not is_gmail)
        self.execute_btn.setText("Apply Gmail Labels" if is_gmail else "Organize Local Files")
        self.execute_btn.setStyleSheet(
            f"background-color: {C.BLUE if is_gmail else C.GREEN}; "
            f"font-size: 14px; padding: 10px 30px;"
        )

    def load_categories(self, engine, has_local=False, local_dir=""):
        self.engine = engine
        self.has_local_files = has_local
        self._refresh_tree()
        self._refresh_combo()
        s = engine.get_summary()
        self.summary_label.setText(
            f"{s['total']:,} emails | {s['categorized']:,} categorized | "
            f"{s['uncategorized']:,} uncategorized"
        )
        if has_local:
            self.mode_local.setChecked(True)
            self.local_dir_label.setText(f"{local_dir}/organized/")
        else:
            self.mode_gmail.setChecked(True)
            self.mode_local.setEnabled(False)
            self.mode_local.setToolTip("Download mailbox first to enable")

    def _refresh_tree(self):
        self.cat_tree.clear()
        if not self.engine:
            return
        parents = {}
        top_level = {}
        for cat_name, emails in sorted(self.engine.categories.items(), key=lambda x: -len(x[1])):
            if '/' in cat_name:
                parent, child = cat_name.split('/', 1)
                if parent not in parents:
                    parents[parent] = {}
                parents[parent][child] = (cat_name, len(emails))
            else:
                top_level[cat_name] = len(emails)

        for name, count in sorted(top_level.items(), key=lambda x: -x[1]):
            item = QTreeWidgetItem([name, f"{count:,}"])
            item.setData(0, Qt.ItemDataRole.UserRole, name)
            item.setForeground(0, QColor(C.BLUE))
            self.cat_tree.addTopLevelItem(item)

        for parent_name, children in sorted(parents.items()):
            total = sum(c[1] for c in children.values())
            pi = QTreeWidgetItem([parent_name, f"{total:,}"])
            pi.setData(0, Qt.ItemDataRole.UserRole, None)
            pi.setForeground(0, QColor(C.MAUVE))
            self.cat_tree.addTopLevelItem(pi)
            for child_name, (full, count) in sorted(children.items(), key=lambda x: -x[1][1]):
                ci = QTreeWidgetItem([child_name, f"{count:,}"])
                ci.setData(0, Qt.ItemDataRole.UserRole, full)
                pi.addChild(ci)

        self.cat_tree.expandAll()

    def _refresh_combo(self):
        self.move_combo.clear()
        if self.engine:
            for cat in sorted(self.engine.categories.keys()):
                self.move_combo.addItem(cat)

    def _get_selected_cat(self):
        item = self.cat_tree.currentItem()
        return item.data(0, Qt.ItemDataRole.UserRole) if item else None

    def _on_cat_selected(self, item):
        cat_name = item.data(0, Qt.ItemDataRole.UserRole)
        if not self.engine:
            return
        if cat_name and cat_name in self.engine.categories:
            emails = self.engine.categories[cat_name]
        else:
            emails = []
            for i in range(item.childCount()):
                cc = item.child(i).data(0, Qt.ItemDataRole.UserRole)
                if cc and cc in self.engine.categories:
                    emails.extend(self.engine.categories[cc])
        self._show_emails(emails)

    def _show_emails(self, emails):
        self.email_count_label.setText(f"{len(emails):,} emails")
        display = sorted(emails, key=lambda e: e.date_parsed or datetime.min, reverse=True)[:2000]
        self.email_table.setRowCount(len(display))
        for row, em in enumerate(display):
            self.email_table.setItem(row, 0, QTableWidgetItem(em.sender_name or em.sender))
            self.email_table.setItem(row, 1, QTableWidgetItem(em.subject))
            ds = em.date_parsed.strftime("%Y-%m-%d") if em.date_parsed else ""
            self.email_table.setItem(row, 2, QTableWidgetItem(ds))
            folder_display = sanitize_folder_name(em.source_folder) if em.source_folder else ""
            self.email_table.setItem(row, 3, QTableWidgetItem(folder_display))

            ct = f"{em.confidence:.0%}"
            ci = QTableWidgetItem(ct)
            if em.confidence >= 0.8:
                ci.setForeground(QColor(C.GREEN))
            elif em.confidence >= 0.5:
                ci.setForeground(QColor(C.YELLOW))
            else:
                ci.setForeground(QColor(C.RED))
            self.email_table.setItem(row, 4, ci)
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
        menu.addAction("Delete (-> Uncategorized)", self._on_delete)
        menu.exec(self.cat_tree.viewport().mapToGlobal(pos))

    def _on_rename(self):
        cat = self._get_selected_cat()
        if not cat:
            return
        new, ok = QInputDialog.getText(self, "Rename", "New name:", text=cat)
        if ok and new and new != cat:
            self.engine.rename_category(cat, new)
            self._refresh_tree()
            self._refresh_combo()

    def _on_merge(self):
        if not self.engine:
            return
        src = self._get_selected_cat()
        if not src:
            return
        cats = sorted(self.engine.categories.keys())
        tgt, ok = QInputDialog.getItem(self, "Merge", f"Merge '{src}' into:", cats, 0, False)
        if ok and tgt and tgt != src:
            self.engine.merge_categories([src], tgt)
            self._refresh_tree()
            self._refresh_combo()

    def _on_delete(self):
        cat = self._get_selected_cat()
        if not cat:
            return
        self.engine.delete_category(cat)
        self._refresh_tree()
        self._refresh_combo()

    def _on_move_emails(self):
        if not self.engine:
            return
        tgt = self.move_combo.currentText()
        if not tgt:
            return
        rows = set(idx.row() for idx in self.email_table.selectedIndexes())
        uids = []
        for r in rows:
            it = self.email_table.item(r, 0)
            if it:
                uid = it.data(Qt.ItemDataRole.UserRole)
                if uid:
                    uids.append(uid)
        if uids:
            self.engine.move_emails(uids, tgt)
            self._refresh_tree()
            item = self.cat_tree.currentItem()
            if item:
                self._on_cat_selected(item)

    def _on_ai_classify(self):
        if not HAS_ANTHROPIC:
            QMessageBox.warning(self, "Missing", "anthropic package not installed.")
            return
        uncat = self.engine.categories.get("Uncategorized", [])
        if not uncat:
            QMessageBox.information(self, "Done", "No uncategorized emails.")
            return
        key, ok = QInputDialog.getText(
            self, "Claude API Key",
            f"Classify {len(uncat):,} emails via Claude Haiku (fast & cheap).",
            QLineEdit.EchoMode.Password
        )
        if not ok or not key:
            return
        existing = [k for k in self.engine.categories if k != "Uncategorized"]
        self.ai_btn.setEnabled(False)
        self.ai_btn.setText("Classifying...")
        self._ai_worker = AiClassifyWorker(key, uncat, existing)
        self._ai_worker.classified.connect(self._on_ai_result)
        self._ai_worker.finished_signal.connect(self._on_ai_done)
        self._ai_worker.error.connect(self._on_ai_error)
        self._ai_worker.start()

    def _on_ai_result(self, dmap):
        for domain, cat in dmap.items():
            to_move = [e for e in self.engine.categories.get("Uncategorized", [])
                       if e.sender_domain == domain]
            for em in to_move:
                self.engine.categories["Uncategorized"].remove(em)
                em.category = cat
                em.confidence = 0.75
                self.engine.categories[cat].append(em)
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
        self._output_dir = ""
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        self.title_label = QLabel("Executing...")
        self.title_label.setStyleSheet("font-size: 20px; font-weight: bold;")
        layout.addWidget(self.title_label)

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

        self.open_btn = QPushButton("Open Output Folder")
        self.open_btn.setProperty("secondary", True)
        self.open_btn.setVisible(False)
        self.open_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_layout.addWidget(self.open_btn)

        self.done_label = QLabel("")
        btn_layout.addWidget(self.done_label)
        layout.addLayout(btn_layout)

    def start_gmail_labeling(self, host, email_addr, password, categories, prefix, archive):
        self.title_label.setText("Applying Gmail Labels")
        self.worker = ImapLabelWorker(host, email_addr, password, categories, prefix, archive)
        self._connect_worker()

    def start_local_organize(self, categories, output_dir, copy_mode):
        self.title_label.setText("Organizing Local Files")
        self._output_dir = str(Path(output_dir) / "organized")
        self.worker = LocalOrganizeWorker(categories, output_dir, copy_mode)
        self._connect_worker()
        self.open_btn.clicked.connect(lambda: os.startfile(self._output_dir))

    def _connect_worker(self):
        self.worker.progress.connect(self._on_progress)
        self.worker.status.connect(self._on_status)
        self.worker.log.connect(self._on_log)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.stop_btn.clicked.connect(self.worker.stop)
        self.worker.start()

    def _on_progress(self, cur, tot):
        self.progress.setMaximum(tot)
        self.progress.setValue(cur)
        pct = int(cur / tot * 100) if tot > 0 else 0
        self.percent_label.setText(f"{pct}% ({cur:,} / {tot:,})")

    def _on_status(self, msg):
        self.status_label.setText(msg)

    def _on_log(self, msg):
        self.log_text.appendPlainText(msg)

    def _on_finished(self):
        self.status_label.setText("Complete!")
        self.status_label.setStyleSheet(f"color: {C.GREEN};")
        self.done_label.setText("Done!")
        self.done_label.setStyleSheet(f"color: {C.GREEN}; font-size: 16px; font-weight: bold;")
        self.stop_btn.setEnabled(False)
        if self._output_dir:
            self.open_btn.setVisible(True)

    def _on_error(self, err):
        self.status_label.setText("Error")
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
        self.download_page = DownloadPage()
        self.analyze_page = AnalyzePage()
        self.review_page = ReviewPage()
        self.execute_page = ExecutePage()

        self.stack.addWidget(self.connect_page)
        self.stack.addWidget(self.download_page)
        self.stack.addWidget(self.analyze_page)
        self.stack.addWidget(self.review_page)
        self.stack.addWidget(self.execute_page)

        self._download_dir = ""

        self.connect_page.connected.connect(self._on_connected)
        self.download_page.download_complete.connect(self._on_download_complete)
        self.analyze_page.analysis_complete.connect(self._show_review)
        self.review_page.execute_requested.connect(self._start_execute)

    def _on_connected(self, mode):
        if mode == "load":
            engine = self.connect_page.loaded_engine
            self._download_dir = self.connect_page.download_dir
            self.analyze_page.set_preloaded(engine)
            self.stack.setCurrentWidget(self.analyze_page)

        elif mode == "download":
            self._download_dir = self.connect_page.download_dir
            self.stack.setCurrentWidget(self.download_page)
            self.download_page.start_download(
                self.connect_page.imap_host,
                self.connect_page.email_addr,
                self.connect_page.password,
                self._download_dir,
            )

        else:  # scan
            self.stack.setCurrentWidget(self.analyze_page)
            self.analyze_page.start_scan(
                self.connect_page.imap_host,
                self.connect_page.email_addr,
                self.connect_page.password,
            )

    def _on_download_complete(self):
        self.analyze_page.set_preloaded(self.download_page.engine)
        self.stack.setCurrentWidget(self.analyze_page)

    def _show_review(self):
        engine = self.analyze_page.engine
        has_local = any(em.local_path for em in engine.emails)
        self.review_page.load_categories(engine, has_local, self._download_dir)
        self.stack.setCurrentWidget(self.review_page)

    def _start_execute(self):
        engine = self.review_page.engine
        if not engine:
            return
        categories = {k: v for k, v in engine.categories.items() if v and k != "Uncategorized"}
        if not categories:
            QMessageBox.warning(self, "Nothing", "No categories to apply.")
            return

        if self.review_page.mode_local.isChecked():
            has_files = any(
                em.local_path and Path(em.local_path).exists()
                for emails in categories.values() for em in emails
            )
            if not has_files:
                QMessageBox.warning(self, "No Files", "Download mailbox first.")
                return
            output_dir = self._download_dir or str(Path.home() / "Desktop" / "InboxForge")
            self.stack.setCurrentWidget(self.execute_page)
            self.execute_page.start_local_organize(
                categories, output_dir, self.review_page.copy_radio.isChecked()
            )
        else:
            if not self.connect_page.email_addr or not self.connect_page.password:
                QMessageBox.warning(self, "Credentials", "Enter Gmail credentials on Connect page.")
                return
            self.stack.setCurrentWidget(self.execute_page)
            self.execute_page.start_gmail_labeling(
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
