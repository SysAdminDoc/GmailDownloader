#!/usr/bin/env python3
"""GmailDownloader v1.2.0 — Full Gmail Mailbox Downloader, AI Organizer & Analytics Suite"""

VERSION = "1.2.0"
MANIFEST_VERSION = 2
MANIFEST_FILENAME = "manifest.json"

import base64
import argparse
import calendar
import csv
import email
import email.header
import email.policy
import email.utils
from email.generator import BytesGenerator
from email.message import EmailMessage
import hashlib
import html
import imaplib
import ipaddress
import io
import json
import mailbox
import mimetypes
import multiprocessing
import os
import re
import secrets
import shlex
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


# codex-branding:start
def _branding_icon_path() -> Path:
    candidates = []
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append(exe_dir / "icon.png")
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(Path(meipass) / "icon.png")
    current = Path(__file__).resolve()
    candidates.extend([current.parent / "icon.png", current.parent.parent / "icon.png", current.parent.parent.parent / "icon.png"])
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return Path("icon.png")
# codex-branding:end


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

from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
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

GMAIL_SKIP_FOLDERS = {'[Gmail]/All Mail', '[Gmail]/Important', '[Gmail]/Spam', '[Gmail]/Trash'}

NEWSLETTER_PLATFORMS = {
    'mailchimp.com', 'sendgrid.net', 'constantcontact.com', 'mailgun.com',
    'amazonses.com', 'substack.com', 'beehiiv.com', 'convertkit.com',
    'hubspot.com', 'sendinblue.com', 'brevo.com', 'mailerlite.com',
    'campaign-archive.com', 'list-manage.com', 'createsend.com',
    'exacttarget.com', 'sailthru.com', 'responsys.com', 'klaviyo.com',
    'drip.com', 'getresponse.com', 'aweber.com', 'infusionsoft.com',
    'activecampaign.com', 'revue.email', 'ghost.io', 'buttondown.email',
}

SENSITIVE_PATTERNS = [
    (r'\b\d{3}-\d{2}-\d{4}\b', 'SSN'),
    (r'\b(?:4[0-9]{12}(?:[0-9]{3})?|5[1-5][0-9]{14}|3[47][0-9]{13}|6(?:011|5[0-9]{2})[0-9]{12})\b', 'Credit Card'),
    (r'(?i)\b(?:password|passwd|pwd)\s*[:=]\s*\S+', 'Password'),
    (r'(?i)\b(?:api[_-]?key|apikey|secret[_-]?key|access[_-]?token)\s*[:=]\s*["\']?\w{16,}', 'API Key'),
    (r'(?i)\bAIza[0-9A-Za-z_-]{35}\b', 'Google API Key'),
    (r'(?i)\bsk-[a-zA-Z0-9]{20,}\b', 'Secret Key'),
    (r'(?i)\bghp_[a-zA-Z0-9]{36}\b', 'GitHub Token'),
]


# ─── Theme ────────────────────────────────────────────────────────────────

class C:
    BASE = "#1e1e2e"; MANTLE = "#181825"; CRUST = "#11111b"
    SURFACE0 = "#313244"; SURFACE1 = "#45475a"; SURFACE2 = "#585b70"
    TEXT = "#cdd6f4"; SUBTEXT0 = "#a6adc8"; SUBTEXT1 = "#bac2de"
    BLUE = "#89b4fa"; GREEN = "#a6e3a1"; MAUVE = "#cba6f7"
    RED = "#f38ba8"; PEACH = "#fab387"; YELLOW = "#f9e2af"
    TEAL = "#94e2d5"; LAVENDER = "#b4befe"; OVERLAY0 = "#6c7086"
    FLAMINGO = "#f2cdcd"; ROSEWATER = "#f5e0dc"; SKY = "#89dceb"
    SAPPHIRE = "#74c7ec"; MAROON = "#eba0ac"; PINK = "#f5c2e7"

CHART_COLORS = [C.BLUE, C.GREEN, C.MAUVE, C.PEACH, C.TEAL, C.RED,
                C.YELLOW, C.LAVENDER, C.FLAMINGO, C.SKY, C.SAPPHIRE,
                C.PINK, C.MAROON, C.ROSEWATER]

STYLESHEET = f"""
    QMainWindow, QWidget {{ background-color: {C.BASE}; color: {C.TEXT};
        font-family: 'Segoe UI', sans-serif; font-size: 13px; }}
    QLineEdit, QTextEdit, QPlainTextEdit {{ background-color: {C.SURFACE0}; color: {C.TEXT};
        border: 1px solid {C.SURFACE1}; border-radius: 6px; padding: 8px;
        selection-background-color: {C.BLUE}; }}
    QLineEdit:focus {{ border: 1px solid {C.BLUE}; }}
    QPushButton {{ background-color: {C.BLUE}; color: {C.CRUST}; border: none;
        border-radius: 6px; padding: 8px 20px; font-weight: bold; }}
    QPushButton:hover {{ background-color: {C.LAVENDER}; }}
    QPushButton:disabled {{ background-color: {C.SURFACE1}; color: {C.OVERLAY0}; }}
    QPushButton[secondary="true"] {{ background-color: {C.SURFACE1}; color: {C.TEXT}; }}
    QPushButton[secondary="true"]:hover {{ background-color: {C.SURFACE2}; }}
    QPushButton[danger="true"] {{ background-color: {C.RED}; color: {C.CRUST}; }}
    QProgressBar {{ background-color: {C.SURFACE0}; border: none; border-radius: 4px;
        height: 8px; text-align: center; }}
    QProgressBar::chunk {{ background-color: {C.BLUE}; border-radius: 4px; }}
    QTreeWidget, QTableWidget, QListWidget {{ background-color: {C.MANTLE}; color: {C.TEXT};
        border: 1px solid {C.SURFACE0}; border-radius: 6px; outline: none; }}
    QTreeWidget::item, QTableWidget::item, QListWidget::item {{ padding: 4px; }}
    QTreeWidget::item:selected, QTableWidget::item:selected, QListWidget::item:selected
        {{ background-color: {C.SURFACE1}; }}
    QTreeWidget::item:hover, QTableWidget::item:hover {{ background-color: {C.SURFACE0}; }}
    QHeaderView::section {{ background-color: {C.SURFACE0}; color: {C.SUBTEXT1};
        border: none; padding: 6px; font-weight: bold; }}
    QScrollBar:vertical {{ background-color: {C.MANTLE}; width: 10px; border-radius: 5px; }}
    QScrollBar::handle:vertical {{ background-color: {C.SURFACE1}; border-radius: 5px; min-height: 30px; }}
    QScrollBar::handle:vertical:hover {{ background-color: {C.SURFACE2}; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
    QScrollBar:horizontal {{ background-color: {C.MANTLE}; height: 10px; border-radius: 5px; }}
    QScrollBar::handle:horizontal {{ background-color: {C.SURFACE1}; border-radius: 5px; min-width: 30px; }}
    QLabel {{ color: {C.TEXT}; }}
    QGroupBox {{ color: {C.TEXT}; border: 1px solid {C.SURFACE0}; border-radius: 8px;
        margin-top: 12px; padding-top: 16px; }}
    QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 6px; }}
    QSplitter::handle {{ background-color: {C.SURFACE0}; }}
    QComboBox {{ background-color: {C.SURFACE0}; color: {C.TEXT};
        border: 1px solid {C.SURFACE1}; border-radius: 6px; padding: 6px 12px; }}
    QComboBox::drop-down {{ border: none; }}
    QComboBox QAbstractItemView {{ background-color: {C.SURFACE0}; color: {C.TEXT};
        selection-background-color: {C.SURFACE1}; border: 1px solid {C.SURFACE1}; }}
    QCheckBox, QRadioButton {{ color: {C.TEXT}; spacing: 8px; }}
    QCheckBox::indicator, QRadioButton::indicator {{ width: 18px; height: 18px;
        border-radius: 4px; border: 2px solid {C.SURFACE2}; background-color: {C.SURFACE0}; }}
    QRadioButton::indicator {{ border-radius: 9px; }}
    QCheckBox::indicator:checked, QRadioButton::indicator:checked
        {{ background-color: {C.BLUE}; border-color: {C.BLUE}; }}
    QInputDialog, QMessageBox {{ background-color: {C.BASE}; color: {C.TEXT}; }}
    QMenu {{ background-color: {C.SURFACE0}; color: {C.TEXT};
        border: 1px solid {C.SURFACE1}; padding: 4px; }}
    QMenu::item:selected {{ background-color: {C.SURFACE1}; }}
    QTabWidget::pane {{ border: 1px solid {C.SURFACE0}; border-radius: 6px;
        background-color: {C.BASE}; }}
    QTabBar::tab {{ background-color: {C.SURFACE0}; color: {C.SUBTEXT0};
        padding: 8px 16px; border-top-left-radius: 6px; border-top-right-radius: 6px;
        margin-right: 2px; }}
    QTabBar::tab:selected {{ background-color: {C.SURFACE1}; color: {C.TEXT}; }}
    QDialog {{ background-color: {C.BASE}; color: {C.TEXT}; }}
    QSpinBox {{ background-color: {C.SURFACE0}; color: {C.TEXT};
        border: 1px solid {C.SURFACE1}; border-radius: 6px; padding: 4px; }}
"""


# ─── Domain Mappings ──────────────────────────────────────────────────────

DOMAIN_CATEGORIES = {
    'amazon.com': 'Shopping', 'amazon.co.uk': 'Shopping', 'ebay.com': 'Shopping',
    'walmart.com': 'Shopping', 'target.com': 'Shopping', 'bestbuy.com': 'Shopping',
    'etsy.com': 'Shopping', 'shopify.com': 'Shopping', 'aliexpress.com': 'Shopping',
    'newegg.com': 'Shopping', 'costco.com': 'Shopping', 'homedepot.com': 'Shopping',
    'lowes.com': 'Shopping', 'macys.com': 'Shopping', 'nordstrom.com': 'Shopping',
    'wayfair.com': 'Shopping', 'chewy.com': 'Shopping', 'wish.com': 'Shopping',
    'kohls.com': 'Shopping', 'samsclub.com': 'Shopping', 'zappos.com': 'Shopping',
    'overstock.com': 'Shopping', 'bhphotovideo.com': 'Shopping',
    'facebook.com': 'Social Media', 'facebookmail.com': 'Social Media',
    'twitter.com': 'Social Media', 'x.com': 'Social Media',
    'linkedin.com': 'Social Media', 'instagram.com': 'Social Media',
    'reddit.com': 'Social Media', 'redditmail.com': 'Social Media',
    'tiktok.com': 'Social Media', 'snapchat.com': 'Social Media',
    'pinterest.com': 'Social Media', 'nextdoor.com': 'Social Media',
    'discord.com': 'Social Media', 'discordapp.com': 'Social Media',
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
    'google.com': 'Tech & Services', 'microsoft.com': 'Tech & Services',
    'apple.com': 'Tech & Services', 'dropbox.com': 'Tech & Services',
    'zoom.us': 'Tech & Services', 'slack.com': 'Tech & Services',
    'github.com': 'Tech & Services', 'atlassian.com': 'Tech & Services',
    'cloudflare.com': 'Tech & Services', 'digitalocean.com': 'Tech & Services',
    'heroku.com': 'Tech & Services', 'notion.so': 'Tech & Services',
    'adobe.com': 'Tech & Services', 'jetbrains.com': 'Tech & Services',
    'godaddy.com': 'Tech & Services', 'namecheap.com': 'Tech & Services',
    'airbnb.com': 'Travel', 'booking.com': 'Travel', 'expedia.com': 'Travel',
    'delta.com': 'Travel', 'united.com': 'Travel', 'southwest.com': 'Travel',
    'uber.com': 'Travel', 'lyft.com': 'Travel', 'kayak.com': 'Travel',
    'hilton.com': 'Travel', 'marriott.com': 'Travel',
    'doordash.com': 'Food & Delivery', 'ubereats.com': 'Food & Delivery',
    'grubhub.com': 'Food & Delivery', 'instacart.com': 'Food & Delivery',
    'netflix.com': 'Entertainment', 'spotify.com': 'Entertainment',
    'hulu.com': 'Entertainment', 'disneyplus.com': 'Entertainment',
    'twitch.tv': 'Entertainment', 'youtube.com': 'Entertainment',
    'steampowered.com': 'Entertainment', 'epicgames.com': 'Entertainment',
    'max.com': 'Entertainment', 'suno.com': 'Entertainment',
    'nytimes.com': 'News', 'washingtonpost.com': 'News',
    'cnn.com': 'News', 'bbc.co.uk': 'News', 'reuters.com': 'News',
    'mychart.com': 'Health', 'zocdoc.com': 'Health', 'fitbit.com': 'Health',
    'coursera.org': 'Education', 'udemy.com': 'Education', 'edx.org': 'Education',
}

SUBJECT_PATTERNS = {
    'Shipping & Tracking': [
        r'(?i)\b(shipped|tracking|delivery|delivered|out for delivery|package|shipment)\b',
        r'(?i)\b(ups|fedex|usps|dhl)\b.*(?:tracking|delivery)'],
    'Invoices & Billing': [
        r'(?i)\b(invoice|receipt|payment\s+(?:received|confirmed)|billing\s+statement)\b',
        r'(?i)\b(order\s+confirm|order\s+#|your\s+order)\b'],
    'Security Alerts': [
        r'(?i)\b(security\s+alert|suspicious|unauthorized|password\s+reset|verify\s+your)\b',
        r'(?i)\b(two-factor|2fa|verification\s+code|login\s+attempt|sign-in)\b'],
    'Calendar & Meetings': [
        r'(?i)\b(meeting\s+(?:invite|invitation|reminder)|calendar|rsvp|webinar)\b',
        r'(?i)\b(zoom\s+meeting|teams\s+meeting|google\s+meet)\b'],
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
    list_unsubscribe_url: str = ""
    category: str = ""
    confidence: float = 0.0
    local_path: str = ""
    source_folder: str = ""
    message_id: str = ""
    in_reply_to: str = ""
    references: str = ""
    size_bytes: int = 0
    sensitive_flags: list = field(default_factory=list)
    is_newsletter: bool = False
    account: str = ""
    has_attachments: bool = False
    attachment_names: list[str] = field(default_factory=list)


@dataclass
class SyncOptions:
    """Controls a source sync without coupling the UI to a provider."""

    since: Optional[datetime] = None
    incremental: bool = True
    verify_integrity: bool = True
    attachments_only: bool = False
    max_message_size: int = 0


@dataclass
class AccountConfig:
    """Serializable account settings used by multi-account orchestration."""

    name: str
    address: str
    host: str = "imap.gmail.com"
    port: int = 993
    use_ssl: bool = True
    auth_mode: str = "password"
    secret: str = ""
    output_dir: str = ""

@dataclass
class CleanRule:
    name: str = ""
    conditions: dict = field(default_factory=dict)  # {field: value}
    action: str = ""       # "categorize", "flag", "skip"
    action_value: str = "" # category name, flag name, etc.
    enabled: bool = True

@dataclass
class SubscriptionInfo:
    domain: str
    sender_name: str
    sender_email: str
    count: int = 0
    first_seen: Optional[datetime] = None
    last_seen: Optional[datetime] = None
    unsubscribe_url: str = ""
    unsubscribe_mailto: str = ""
    frequency: str = ""  # "daily", "weekly", "monthly", "irregular"


# ─── Learned Rules (Feedback Loop) ───────────────────────────────────────

class LearnedRules:
    def __init__(self, path: str = ""):
        self.path = path
        self.domain_rules: dict[str, str] = {}   # domain -> category
        self.sender_rules: dict[str, str] = {}   # sender@email -> category
        if path:
            self.load()

    def learn(self, em: EmailInfo, category: str):
        if em.sender_domain:
            self.domain_rules[em.sender_domain] = category
        if em.sender:
            self.sender_rules[em.sender] = category

    def lookup(self, em: EmailInfo) -> Optional[str]:
        if em.sender and em.sender in self.sender_rules:
            return self.sender_rules[em.sender]
        if em.sender_domain and em.sender_domain in self.domain_rules:
            return self.domain_rules[em.sender_domain]
        return None

    def save(self):
        if not self.path:
            return
        data = {'domain_rules': self.domain_rules, 'sender_rules': self.sender_rules}
        try:
            with open(self.path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def load(self):
        if not self.path or not os.path.exists(self.path):
            return
        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.domain_rules = data.get('domain_rules', {})
            self.sender_rules = data.get('sender_rules', {})
        except Exception:
            pass


# ─── Clean Rules Engine ──────────────────────────────────────────────────

class CleanRulesEngine:
    def __init__(self, path: str = ""):
        self.path = path
        self.rules: list[CleanRule] = []
        if path:
            self.load()

    def add_rule(self, rule: CleanRule):
        self.rules.append(rule)
        self.save()

    def remove_rule(self, idx: int):
        if 0 <= idx < len(self.rules):
            self.rules.pop(idx)
            self.save()

    def apply(self, em: EmailInfo) -> Optional[tuple[str, str]]:
        """Returns (action, action_value) if a rule matches, else None."""
        for rule in self.rules:
            if not rule.enabled:
                continue
            if self._matches(rule, em):
                return (rule.action, rule.action_value)
        return None

    def _matches(self, rule: CleanRule, em: EmailInfo) -> bool:
        conds = rule.conditions
        if 'domain' in conds and em.sender_domain != conds['domain']:
            return False
        if 'sender' in conds and conds['sender'].lower() not in em.sender.lower():
            return False
        if 'subject_contains' in conds and conds['subject_contains'].lower() not in em.subject.lower():
            return False
        if 'older_than_days' in conds:
            if em.date_parsed and (datetime.now() - em.date_parsed).days < int(conds['older_than_days']):
                return False
        if 'is_newsletter' in conds and conds['is_newsletter'] and not em.is_newsletter:
            return False
        if 'has_attachment' in conds and conds['has_attachment']:
            pass  # Would need to check .eml for attachments
        return True

    def save(self):
        if not self.path:
            return
        data = [{'name': r.name, 'conditions': r.conditions, 'action': r.action,
                 'action_value': r.action_value, 'enabled': r.enabled} for r in self.rules]
        try:
            with open(self.path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def load(self):
        if not self.path or not os.path.exists(self.path):
            return
        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.rules = [CleanRule(**r) for r in data]
        except Exception:
            pass

    @staticmethod
    def import_gmail_filters(xml_path: str) -> list[CleanRule]:
        """Parse Gmail filter export XML into CleanRules."""
        rules = []
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            ns = {'atom': 'http://www.w3.org/2005/Atom', 'apps': 'http://schemas.google.com/apps/2006'}
            for entry in root.findall('atom:entry', ns):
                props = {}
                for prop in entry.findall('apps:property', ns):
                    props[prop.get('name', '')] = prop.get('value', '')
                conds = {}
                if props.get('from'):
                    conds['sender'] = props['from']
                if props.get('subject'):
                    conds['subject_contains'] = props['subject']
                if props.get('hasTheWord'):
                    conds['subject_contains'] = props['hasTheWord']
                label = props.get('label', '')
                if conds and label:
                    rules.append(CleanRule(
                        name=f"Gmail: {label}", conditions=conds,
                        action='categorize', action_value=label, enabled=True
                    ))
        except Exception:
            pass
        return rules


# ─── Category Engine ──────────────────────────────────────────────────────

class CategoryEngine:
    def __init__(self, user_domain: str = ""):
        self.user_domain = user_domain.lower()
        self.emails: list[EmailInfo] = []
        self.categories: dict[str, list[EmailInfo]] = defaultdict(list)
        self.domain_stats: Counter = Counter()
        self.ambiguous: list[EmailInfo] = []
        self.learned = LearnedRules()
        self.clean_rules = CleanRulesEngine()
        self.subscriptions: list[SubscriptionInfo] = []
        self.threads: dict[str, list[EmailInfo]] = {}  # thread_id -> emails

    def extract_domain(self, email_addr: str) -> str:
        match = re.search(r'@([\w.-]+)', email_addr.lower())
        if not match:
            return ""
        domain = match.group(1)
        parts = domain.split('.')
        if len(parts) > 2:
            if len(parts) >= 3 and parts[-2] in ('co','com','org','ac','gov','net','edu'):
                domain = '.'.join(parts[-3:])
            else:
                domain = '.'.join(parts[-2:])
        return domain

    def _is_newsletter_domain(self, domain: str) -> bool:
        return any(domain.endswith(nd) or domain == nd for nd in NEWSLETTER_PLATFORMS)

    def categorize_email(self, em: EmailInfo) -> tuple[str, float]:
        domain = em.sender_domain

        # 0. Learned rules (feedback loop) — highest priority
        learned = self.learned.lookup(em)
        if learned:
            return learned, 0.95

        # 1. Clean rules
        rule_result = self.clean_rules.apply(em)
        if rule_result and rule_result[0] == 'categorize':
            return rule_result[1], 0.93

        # 2. Internal/Work
        if self.user_domain and domain == self.user_domain:
            return "Work/Internal", 0.95

        # 3. Known domain mapping
        if domain in DOMAIN_CATEGORIES:
            return DOMAIN_CATEGORIES[domain], 0.9

        # 4. Newsletter detection
        em.is_newsletter = em.has_list_unsubscribe or self._is_newsletter_domain(domain)
        if em.is_newsletter:
            return "Newsletters", 0.85

        # 5. Subject pattern matching
        for cat, patterns in SUBJECT_PATTERNS.items():
            for pattern in patterns:
                if re.search(pattern, em.subject):
                    return cat, 0.7

        # 6. Automated sender
        if re.search(r'(?i)(no-?reply|noreply|notifications?@|alerts?@|mailer-daemon)', em.sender.lower()):
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

        self._detect_subscriptions()
        self._build_threads()

    def _detect_subscriptions(self):
        sub_map = {}  # domain -> SubscriptionInfo
        for em in self.emails:
            if not em.is_newsletter and not em.has_list_unsubscribe:
                continue
            d = em.sender_domain
            if d not in sub_map:
                unsub_url, unsub_mailto = '', ''
                if em.list_unsubscribe_url:
                    raw = em.list_unsubscribe_url
                    http_match = re.search(r'<(https?://[^>]+)>', raw)
                    mailto_match = re.search(r'<mailto:([^>]+)>', raw)
                    if http_match:
                        unsub_url = http_match.group(1)
                    if mailto_match:
                        unsub_mailto = mailto_match.group(1)
                sub_map[d] = SubscriptionInfo(
                    domain=d, sender_name=em.sender_name, sender_email=em.sender,
                    unsubscribe_url=unsub_url, unsubscribe_mailto=unsub_mailto)
            si = sub_map[d]
            si.count += 1
            if em.date_parsed:
                if not si.first_seen or em.date_parsed < si.first_seen:
                    si.first_seen = em.date_parsed
                if not si.last_seen or em.date_parsed > si.last_seen:
                    si.last_seen = em.date_parsed

        for si in sub_map.values():
            if si.first_seen and si.last_seen and si.count > 1:
                span = (si.last_seen - si.first_seen).days
                if span > 0:
                    freq = si.count / (span / 30)
                    if freq >= 25: si.frequency = "daily"
                    elif freq >= 3: si.frequency = "weekly"
                    elif freq >= 0.8: si.frequency = "monthly"
                    else: si.frequency = "irregular"

        self.subscriptions = sorted(sub_map.values(), key=lambda s: -s.count)

    def _build_threads(self):
        self.threads.clear()
        id_to_thread = {}
        for em in self.emails:
            thread_id = None
            if em.references:
                first_ref = em.references.strip().split()[0]
                thread_id = id_to_thread.get(first_ref)
            if not thread_id and em.in_reply_to:
                thread_id = id_to_thread.get(em.in_reply_to.strip())
            if not thread_id:
                thread_id = em.message_id or em.uid
            if em.message_id:
                id_to_thread[em.message_id] = thread_id
            if thread_id not in self.threads:
                self.threads[thread_id] = []
            self.threads[thread_id].append(em)

        # Remove single-email threads
        self.threads = {k: sorted(v, key=lambda e: e.date_parsed or datetime.min)
                       for k, v in self.threads.items() if len(v) > 1}

    def get_summary(self) -> dict:
        total = len(self.emails)
        categorized = sum(1 for em in self.emails if em.confidence > 0)
        dates = [em.date_parsed for em in self.emails if em.date_parsed]
        date_range = ("", "")
        if dates:
            date_range = (min(dates).strftime("%Y-%m-%d"), max(dates).strftime("%Y-%m-%d"))
        folder_counts = Counter(em.source_folder for em in self.emails if em.source_folder)
        total_size = sum(em.size_bytes for em in self.emails)
        return {
            'total': total, 'categorized': categorized,
            'uncategorized': total - categorized,
            'categories': {k: len(v) for k, v in sorted(self.categories.items(), key=lambda x: -len(x[1]))},
            'top_domains': self.domain_stats.most_common(20),
            'date_range': date_range,
            'folder_counts': dict(folder_counts.most_common()),
            'total_size': total_size,
            'thread_count': len(self.threads),
            'newsletter_count': len(self.subscriptions),
            'sensitive_count': sum(1 for em in self.emails if em.sensitive_flags),
        }

    def get_stats(self) -> dict:
        """Detailed statistics for the dashboard."""
        monthly = Counter()
        hourly = Counter()
        dow = Counter()
        heatmap = defaultdict(int)  # (day_of_week, hour) -> count
        sender_counts = Counter()
        domain_counts = Counter()
        cat_sizes = defaultdict(int)

        for em in self.emails:
            if em.date_parsed:
                monthly[em.date_parsed.strftime("%Y-%m")] += 1
                hourly[em.date_parsed.hour] += 1
                dow[em.date_parsed.weekday()] += 1
                heatmap[(em.date_parsed.weekday(), em.date_parsed.hour)] += 1
            sender_counts[em.sender_name or em.sender] += 1
            domain_counts[em.sender_domain] += 1
            cat_sizes[em.category] += em.size_bytes

        return {
            'monthly': dict(sorted(monthly.items())),
            'hourly': dict(sorted(hourly.items())),
            'dow': dict(sorted(dow.items())),
            'heatmap': {f"{k[0]},{k[1]}": v for k, v in heatmap.items()},
            'top_senders': sender_counts.most_common(20),
            'top_domains': domain_counts.most_common(20),
            'category_sizes': dict(sorted(cat_sizes.items(), key=lambda x: -x[1])),
            'category_counts': {k: len(v) for k, v in sorted(self.categories.items(), key=lambda x: -len(x[1]))},
        }

    def search(self, query, limit=0):
        return search_emails(self.emails, query, limit)

    def confidence_candidates(self, threshold=0.75):
        return [em for em in self.emails if em.confidence < threshold]

    def relationship_graph(self):
        return build_relationship_graph(self.emails, self.threads)

    def thread_clusters(self, similarity_threshold=0.25):
        return cluster_threads(self.threads, similarity_threshold)

    def sender_health(self):
        return sender_health_scores(self.emails)

    def reply_latency(self):
        return reply_latency_histogram(self.emails, self.threads)

    def storage_forecast(self, months=12):
        return storage_forecast(self.emails, months)

    def location_timeline(self, resolver=None, include_private=False):
        return build_location_timeline(self.emails, resolver, include_private)

    def inbox_zero_suggestions(self, older_than_days=180):
        return suggest_inbox_zero(self.emails, older_than_days)

    def rename_category(self, old: str, new: str):
        if old in self.categories:
            emails = self.categories.pop(old)
            for em in emails:
                em.category = new
            self.categories[new].extend(emails)

    def merge_categories(self, sources: list[str], target: str):
        for src in sources:
            if src in self.categories and src != target:
                for em in self.categories.pop(src):
                    em.category = target
                    self.categories[target].append(em)

    def move_emails(self, uids: list[str], target: str):
        uid_set = set(uids)
        for em in self.emails:
            if em.uid in uid_set:
                old = em.category
                if old in self.categories:
                    self.categories[old] = [e for e in self.categories[old] if e.uid != em.uid]
                    if not self.categories[old]:
                        del self.categories[old]
                em.category = target
                em.confidence = max(em.confidence, 0.5)
                self.categories[target].append(em)
                self.learned.learn(em, target)
        self.learned.save()

    def delete_category(self, name: str):
        if name in self.categories:
            for em in self.categories.pop(name):
                em.category = "Uncategorized"
                em.confidence = 0.0
                self.categories["Uncategorized"].append(em)

    def export_csv(self, path: str):
        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            w.writerow(['Date', 'From', 'FromName', 'Domain', 'Subject', 'Category',
                        'Confidence', 'Folder', 'Size', 'Newsletter', 'Sensitive', 'MessageID'])
            for em in sorted(self.emails, key=lambda e: e.date_parsed or datetime.min, reverse=True):
                w.writerow([
                    em.date_parsed.strftime("%Y-%m-%d %H:%M") if em.date_parsed else '',
                    em.sender, em.sender_name, em.sender_domain, em.subject,
                    em.category, f"{em.confidence:.0%}", em.source_folder,
                    em.size_bytes, em.is_newsletter,
                    ','.join(em.sensitive_flags) if em.sensitive_flags else '',
                    em.message_id,
                ])

    def export_json(self, path: str):
        data = [{'date': em.date_parsed.isoformat() if em.date_parsed else '',
                 'sender': em.sender, 'sender_name': em.sender_name,
                 'domain': em.sender_domain, 'subject': em.subject,
                 'category': em.category, 'confidence': em.confidence,
                 'folder': em.source_folder, 'size': em.size_bytes,
                 'newsletter': em.is_newsletter,
                 'sensitive': em.sensitive_flags, 'message_id': em.message_id}
                for em in self.emails]
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def export_mbox(self, path):
        return export_mbox(self.emails, path)

    def export_markdown(self, path, include_body=True):
        return export_markdown_vault(self.emails, path, include_body)

    def export_notion(self, path, include_body=True):
        return export_notion_markdown(self.emails, path, include_body)

    def export_pdf(self, path):
        return export_pdf(self.emails, path)

    def export_relationship_graph(self, path):
        return export_relationship_graph(self.relationship_graph(), path)

    def export_contact_graph(self, path):
        return export_contact_graph(self.emails, path)

    def export_receipts_ofx(self, path, receipts=None):
        return export_receipts_ofx(receipts if receipts is not None else extract_receipts(self.emails), path)

    def export_location_timeline(self, path):
        return export_location_timeline_csv(self.location_timeline(), path)

    def save_state(self, path: str):
        data = {'version': VERSION, 'user_domain': self.user_domain,
                'emails': [email_info_to_record(em) for em in self.emails]}
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
                em = email_info_from_record(
                    ed.get('uid', ''), ed, ed.get('source_folder', ''), ed.get('account', '')
                )
                self.emails.append(em)
                if em.category:
                    self.categories[em.category].append(em)
            self.domain_stats = Counter(em.sender_domain for em in self.emails)
            self._detect_subscriptions()
            self._build_threads()
            return True
        except Exception:
            return False


# ─── Helpers ──────────────────────────────────────────────────────────────

def decode_header(raw):
    if not raw: return ""
    try:
        parts = email.header.decode_header(raw)
        return ' '.join(p.decode(c or 'utf-8', errors='replace') if isinstance(p, bytes)
                       else str(p) for p, c in parts).strip()
    except Exception:
        return str(raw).strip()

def parse_date(s):
    if not s: return None
    try: return email.utils.parsedate_to_datetime(s).replace(tzinfo=None)
    except Exception: return None

def sanitize_filename(s, max_len=60):
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', s)
    return re.sub(r'_+', '_', s).strip('_. ')[:max_len] or "untitled"

def sanitize_folder_name(n):
    n = re.sub(r'^\[Gmail\]/', '', n)
    return re.sub(r'[<>:"|?*\x00-\x1f]', '_', n).strip('_. ') or "Other"

def parse_imap_folder_list(line):
    try:
        parts = line.decode('utf-8', errors='replace').split('"')
        if len(parts) >= 4:
            return parts[3].strip() if parts[3].strip() else parts[-2].strip()
    except Exception: pass
    match = re.match(rb'\(.*?\)\s+"(.?)"\s+"?(.+?)"?\s*$', line)
    if match:
        try: return match.group(2).decode('utf-8')
        except: return match.group(2).decode('ascii', errors='replace')
    return None

def scan_sensitive(text: str) -> list[str]:
    """Scan text for sensitive content patterns."""
    flags = []
    for pattern, label in SENSITIVE_PATTERNS:
        if re.search(pattern, text):
            flags.append(label)
    return flags

def format_size(b):
    if b < 1024: return f"{b} B"
    if b < 1024**2: return f"{b/1024:.1f} KB"
    if b < 1024**3: return f"{b/1024**2:.1f} MB"
    return f"{b/1024**3:.1f} GB"


SEARCH_FILTERS = {'from', 'subject', 'category', 'domain', 'folder', 'after', 'before', 'has'}


def parse_search_query(query):
    filters, terms = {}, []
    for token in shlex.split(query or ''):
        if ':' in token:
            key, value = token.split(':', 1)
            if key.lower() in SEARCH_FILTERS and value:
                filters[key.lower()] = value
                continue
        terms.append(token.lower())
    return filters, terms


def search_emails(emails, query, limit=0):
    """Search metadata and available local bodies with Gmail-like filters."""
    filters, terms = parse_search_query(query)
    after = parse_since_date(filters['after']) if filters.get('after') else None
    before = parse_since_date(filters['before']) if filters.get('before') else None
    results = []
    for em in emails:
        if filters.get('from', '').lower() not in (em.sender + ' ' + em.sender_name).lower():
            if 'from' in filters:
                continue
        if filters.get('subject', '').lower() not in em.subject.lower() and 'subject' in filters:
            continue
        if filters.get('category', '').lower() not in em.category.lower() and 'category' in filters:
            continue
        if filters.get('domain', '').lower() not in em.sender_domain.lower() and 'domain' in filters:
            continue
        if filters.get('folder', '').lower() not in em.source_folder.lower() and 'folder' in filters:
            continue
        if after and (not em.date_parsed or em.date_parsed < after):
            continue
        if before and (not em.date_parsed or em.date_parsed >= before):
            continue
        has_filter = filters.get('has', '').lower()
        if has_filter == 'attachment' and not em.has_attachments:
            continue
        if has_filter in ('newsletter', 'unsubscribe') and not em.is_newsletter and not em.has_list_unsubscribe:
            continue
        if has_filter == 'sensitive' and not em.sensitive_flags:
            continue
        if terms:
            body = ''
            if em.local_path and Path(em.local_path).exists():
                try:
                    msg = email.message_from_bytes(Path(em.local_path).read_bytes(), policy=email.policy.default)
                    body = extract_message_body(msg, 10000)
                except OSError:
                    pass
            haystack = ' '.join((em.sender, em.sender_name, em.subject, em.category,
                                 em.source_folder, body)).lower()
            if not all(term in haystack for term in terms):
                continue
        results.append(em)
        if limit and len(results) >= limit:
            break
    return results


def build_relationship_graph(emails, threads=None):
    """Build a weighted sender co-occurrence graph from reconstructed threads."""
    thread_map = threads or {}
    if not thread_map:
        thread_map = defaultdict(list)
        for em in emails:
            thread_map[em.message_id or em.uid].append(em)
    edge_weights = Counter()
    nodes = Counter()
    for members in thread_map.values():
        senders = sorted({em.sender or em.sender_domain for em in members if em.sender or em.sender_domain})
        for sender in senders:
            nodes[sender] += 1
        for index, left in enumerate(senders):
            for right in senders[index + 1:]:
                edge_weights[(left, right)] += 1
    return {
        'nodes': [{'id': sender, 'threads': count} for sender, count in nodes.most_common()],
        'edges': [{'source': left, 'target': right, 'weight': weight}
                  for (left, right), weight in edge_weights.most_common()],
    }


def cluster_threads(threads, similarity_threshold=0.25):
    """Cluster threads with a deterministic token similarity fallback."""
    def tokens(members):
        text = ' '.join(f'{em.subject} {em.sender_domain}' for em in members).lower()
        return {token for token in re.findall(r'[a-z0-9]{3,}', text)}

    clusters = []
    for thread_id, members in threads.items():
        current = tokens(members)
        placed = False
        for cluster in clusters:
            overlap = current & cluster['tokens']
            union = current | cluster['tokens']
            if union and len(overlap) / len(union) >= similarity_threshold:
                cluster['thread_ids'].append(thread_id)
                cluster['tokens'].update(current)
                placed = True
                break
        if not placed:
            clusters.append({'thread_ids': [thread_id], 'tokens': set(current)})
    return [
        {'cluster_id': index + 1, 'thread_ids': item['thread_ids']}
        for index, item in enumerate(clusters)
    ]


def export_relationship_graph(graph, output_path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.suffix.lower() == '.graphml':
        lines = [
            '<?xml version="1.0" encoding="UTF-8"?>',
            '<graphml xmlns="http://graphml.graphdrawing.org/xmlns">',
            '<graph edgedefault="undirected"><node id="__dummy__"/>',
        ]
        lines.pop()  # keep the declaration compact while avoiding a mutable template
        lines.append('<graph edgedefault="undirected">')
        for node in graph.get('nodes', []):
            node_id = html.escape(str(node['id']), quote=True)
            lines.append(f'<node id="{node_id}"/>')
        for index, edge in enumerate(graph.get('edges', [])):
            lines.append(f'<edge id="e{index}" source="{html.escape(str(edge["source"]), quote=True)}" '
                         f'target="{html.escape(str(edge["target"]), quote=True)}"/>')
        lines.extend(['</graph>', '</graphml>'])
        output_path.write_text('\n'.join(lines), encoding='utf-8')
    else:
        output_path.write_text(json.dumps(graph, indent=2, ensure_ascii=False), encoding='utf-8')
    return str(output_path)


def contact_graph(emails):
    graph = build_relationship_graph(emails)
    return graph


def export_contact_graph(emails, output_path):
    return export_relationship_graph(contact_graph(emails), output_path)


def export_markdown_vault(emails, output_dir, include_body=True):
    """Export one Markdown note per message with Obsidian-friendly frontmatter."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for index, em in enumerate(sorted(emails, key=lambda item: item.date_parsed or datetime.min, reverse=True)):
        date_value = em.date_parsed.strftime('%Y-%m-%dT%H:%M:%S') if em.date_parsed else ''
        title = em.subject or 'No subject'
        stem = sanitize_filename(f'{date_value[:10]}_{em.sender_domain}_{title}', 150)
        target = output_dir / f'{stem}_{index}.md'
        body = ''
        if include_body and em.local_path and Path(em.local_path).exists():
            try:
                msg = email.message_from_bytes(Path(em.local_path).read_bytes(), policy=email.policy.default)
                body = extract_message_body(msg)
            except OSError:
                body = ''
        frontmatter = {
            'subject': title,
            'from': em.sender,
            'date': date_value,
            'category': em.category,
            'source_folder': em.source_folder,
            'message_id': em.message_id,
            'newsletter': em.is_newsletter,
            'sensitive': em.sensitive_flags,
        }
        lines = ['---'] + [f'{key}: {json.dumps(value, ensure_ascii=False)}' for key, value in frontmatter.items()] + ['---', '', f'# {title}', '']
        if body:
            lines.extend([body, ''])
        else:
            lines.append('_Body unavailable in headers-only mode._')
        target.write_text('\n'.join(lines), encoding='utf-8')
        written.append(str(target))
    return written


def export_notion_markdown(emails, output_dir, include_body=True):
    """Export Notion-importable Markdown using the same portable vault format."""
    return export_markdown_vault(emails, output_dir, include_body)


def export_pdf(emails, output_path):
    """Render email metadata and readable bodies to a portable PDF."""
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import inch
        from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer
    except ImportError as exc:
        raise RuntimeError('PDF export requires reportlab') from exc
    styles = getSampleStyleSheet()
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    story = []
    for index, em in enumerate(emails):
        story.append(Paragraph(html.escape(em.subject or '(no subject)'), styles['Title']))
        story.append(Paragraph(
            html.escape(f'From: {em.sender} | Date: {em.date} | Category: {em.category}'),
            styles['Normal']))
        story.append(Spacer(1, 0.2 * inch))
        body = ''
        if em.local_path and Path(em.local_path).exists():
            msg = email.message_from_bytes(Path(em.local_path).read_bytes(), policy=email.policy.default)
            body = extract_message_body(msg, 50000)
        story.append(Paragraph(html.escape(body or 'Body unavailable in headers-only mode.').replace('\n', '<br/>'), styles['BodyText']))
        if index < len(emails) - 1:
            story.append(PageBreak())
    SimpleDocTemplate(str(output_path), pagesize=letter).build(story)
    return str(output_path)


def sender_health_scores(emails, now=None):
    """Score senders using volume, recency, replies, and unsubscribe signals."""
    now = now or datetime.now()
    grouped = defaultdict(list)
    for em in emails:
        grouped[em.sender or em.sender_name or 'unknown'].append(em)
    scores = []
    for sender, messages in grouped.items():
        last = max((em.date_parsed for em in messages if em.date_parsed), default=None)
        age_days = (now - last).days if last else 9999
        recency = max(0.0, 1.0 - age_days / 365)
        replies = sum(bool(em.in_reply_to or em.references) for em in messages)
        reply_rate = replies / len(messages)
        newsletters = sum(em.is_newsletter or em.has_list_unsubscribe for em in messages)
        unsubscribe_rate = newsletters / len(messages)
        score = round(100 * (0.5 * recency + 0.35 * reply_rate + 0.15 * (1 - unsubscribe_rate)), 1)
        scores.append({'sender': sender, 'score': score, 'count': len(messages),
                       'reply_rate': round(reply_rate, 3), 'last_seen': last.isoformat() if last else ''})
    return sorted(scores, key=lambda item: (-item['score'], -item['count'], item['sender']))


def reply_latency_histogram(emails, threads=None):
    histogram = Counter()
    thread_map = threads or {}
    if not thread_map:
        thread_map = defaultdict(list)
        for em in emails:
            thread_map[em.message_id or em.uid].append(em)
    for members in thread_map.values():
        ordered = sorted(members, key=lambda em: em.date_parsed or datetime.min)
        for previous, current in zip(ordered, ordered[1:]):
            if previous.date_parsed and current.date_parsed:
                hours = (current.date_parsed - previous.date_parsed).total_seconds() / 3600
                if hours < 1: bucket = '<1h'
                elif hours < 4: bucket = '1-4h'
                elif hours < 24: bucket = '4-24h'
                elif hours < 72: bucket = '1-3d'
                else: bucket = '3d+'
                histogram[bucket] += 1
    return dict(histogram)


def storage_forecast(emails, months=12):
    monthly = Counter()
    for em in emails:
        if em.date_parsed:
            monthly[em.date_parsed.strftime('%Y-%m')] += em.size_bytes
    ordered = sorted(monthly.items())
    recent = [size for _, size in ordered[-6:]]
    average = sum(recent) / len(recent) if recent else 0
    current = sum(monthly.values())
    forecast = []
    for offset in range(1, months + 1):
        forecast.append({'months_from_now': offset, 'projected_bytes': int(current + average * offset)})
    return {'monthly_bytes': dict(ordered), 'average_monthly_bytes': int(average), 'forecast': forecast}


def location_timeline(emails, resolver=None):
    """Extract public IP hops from ``Received`` headers for an audit timeline.

    A resolver is deliberately injected instead of calling a geolocation API from
    the application.  It may return a country string or a mapping containing
    country/city/latitude/longitude fields.  Private and reserved addresses are
    omitted by default because they do not describe a useful travel location.
    """
    return build_location_timeline(emails, resolver=resolver)


def _received_ip_addresses(header):
    """Yield validated IPv4/IPv6 literals from one Received header."""
    candidates = re.findall(
        r'\[[0-9A-Fa-f:.]+\]|(?<![A-Za-z0-9_:])(?:\d{1,3}\.){3}\d{1,3}|'
        r'(?<![A-Za-z0-9_:])(?:[0-9A-Fa-f]{1,4}:){2,}[0-9A-Fa-f:.]+',
        str(header),
    )
    seen = set()
    for candidate in candidates:
        candidate = candidate.strip('[]')
        try:
            address = ipaddress.ip_address(candidate)
        except ValueError:
            continue
        normalized = str(address)
        if normalized not in seen:
            seen.add(normalized)
            yield address


def _received_timestamp(header):
    """Parse the timestamp after the final semicolon in a Received header."""
    text = str(header)
    if ';' not in text:
        return None
    try:
        return email.utils.parsedate_to_datetime(text.rsplit(';', 1)[1].strip())
    except (TypeError, ValueError, IndexError, OverflowError):
        return None


def build_location_timeline(emails, resolver=None, include_private=False):
    timeline = []
    for em in emails:
        if not em.local_path or not Path(em.local_path).exists():
            continue
        try:
            msg = email.message_from_bytes(Path(em.local_path).read_bytes(), policy=email.policy.default)
        except (OSError, ValueError):
            continue
        for hop, received in enumerate(msg.get_all('Received', []), start=1):
            timestamp = _received_timestamp(received)
            for address in _received_ip_addresses(received):
                if not include_private and not address.is_global:
                    continue
                item = {
                    'date': em.date,
                    'received_at': timestamp.isoformat() if timestamp else '',
                    'ip': str(address),
                    'version': address.version,
                    'country': '',
                    'uid': em.uid,
                    'hop': hop,
                }
                if resolver:
                    try:
                        resolved = resolver(str(address))
                        if isinstance(resolved, dict):
                            item.update({str(key): value for key, value in resolved.items()})
                            item['country'] = str(resolved.get('country', item.get('country', '')) or '')
                        elif resolved:
                            item['country'] = str(resolved)
                    except Exception:
                        # A missing/failed resolver must not make local analysis fail.
                        pass
                timeline.append(item)
    return sorted(timeline, key=lambda item: (item.get('received_at') or item.get('date', ''), item['uid'], item['hop']))


def export_location_timeline_csv(timeline, output_path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fields = ['received_at', 'date', 'ip', 'version', 'country', 'uid', 'hop']
    with output_path.open('w', newline='', encoding='utf-8') as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows({field: item.get(field, '') for field in fields} for item in timeline)
    return str(output_path)


def suggest_inbox_zero(emails, older_than_days=180, now=None):
    now = now or datetime.now()
    suggestions = []
    for em in emails:
        age = (now - em.date_parsed).days if em.date_parsed else 0
        if em.is_newsletter or em.has_list_unsubscribe:
            suggestions.append({'uid': em.uid, 'action': 'unsubscribe_or_archive', 'reason': 'newsletter'})
        elif age >= older_than_days and em.confidence > 0:
            suggestions.append({'uid': em.uid, 'action': 'archive', 'reason': f'older than {older_than_days} days'})
    return suggestions


RECEIPT_ATTACHMENT_EXTENSIONS = {'.pdf', '.png', '.jpg', '.jpeg', '.gif', '.webp', '.tif', '.tiff', '.bmp'}
RECEIPT_IMAGE_TYPES = {'image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/tiff', 'image/bmp'}
RECEIPT_VISION_PROMPT = (
    'Extract this receipt or invoice as JSON with exactly these useful fields: '
    'merchant, date (YYYY-MM-DD when possible), amount (number), currency (ISO 4217), '
    'and line_items (array of {description, quantity, amount}). Return only JSON.'
)


def _json_object(value):
    if isinstance(value, dict):
        return value
    if not isinstance(value, str):
        return {}
    decoder = json.JSONDecoder()
    for match in re.finditer(r'\{', value):
        try:
            parsed, _ = decoder.raw_decode(value[match.start():])
        except json.JSONDecodeError:
            continue
        if isinstance(parsed, dict):
            return parsed
    return {}


def _receipt_amount(value):
    if value is None or value == '':
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    match = re.search(r'-?\d[\d,]*(?:\.\d+)?', str(value))
    return round(float(match.group().replace(',', '')), 2) if match else None


def _receipt_date(value):
    if not value:
        return ''
    text = str(value).strip()
    formats = ('%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%m-%d-%Y', '%B %d, %Y', '%b %d, %Y')
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    match = re.search(r'\b(20\d{2})[-/]([01]?\d)[-/]([0-3]?\d)\b', text)
    return '-'.join(match.groups()) if match else text


def normalize_receipt_result(value, source='', uid='', sender='', category=''):
    """Normalize text/vision output into a stable, exportable receipt schema."""
    data = _json_object(value)
    raw_items = data.get('line_items', data.get('items', []))
    if not isinstance(raw_items, list):
        raw_items = []
    line_items = []
    for item in raw_items:
        if isinstance(item, dict):
            line_items.append({
                'description': str(item.get('description', item.get('name', '')) or '').strip(),
                'quantity': item.get('quantity', 1),
                'amount': _receipt_amount(item.get('amount', item.get('total'))),
            })
        elif item:
            line_items.append({'description': str(item).strip(), 'quantity': 1, 'amount': None})
    result = {
        'merchant': str(data.get('merchant', data.get('store', data.get('vendor', ''))) or '').strip(),
        'date': _receipt_date(data.get('date', data.get('transaction_date', ''))),
        'amount': _receipt_amount(data.get('amount', data.get('total', data.get('grand_total')))),
        'currency': str(data.get('currency', '') or '').upper().strip(),
        'line_items': line_items,
        'source': str(source or ''),
        'uid': str(uid or ''),
        'sender': str(sender or ''),
        'category': str(category or ''),
    }
    if data.get('raw') and not result['merchant']:
        result['raw'] = str(data['raw'])
    return result


def extract_receipt_fields(text):
    amount = re.search(
        r'(?i)(?:total|amount|charged|grand total)\s*(?::|=)?\s*'
        r'(?P<currency>[$€£]|USD|EUR|GBP)?\s*([0-9,]+(?:\.\d{2})?)', text
    )
    date_match = re.search(r'\b(20\d{2}[-/]\d{1,2}[-/]\d{1,2})\b', text)
    merchant = re.search(r'(?i)(?:merchant|store|from)\s*[:\-]\s*([^\n]{2,80})', text)
    currency = amount.group('currency') if amount else ''
    currency = {'$': 'USD', '€': 'EUR', '£': 'GBP'}.get(currency, currency or '')
    return normalize_receipt_result({
        'amount': amount.group(2).replace(',', '') if amount else None,
        'currency': currency,
        'date': date_match.group(1) if date_match else '',
        'merchant': merchant.group(1).strip() if merchant else '',
    })


def render_pdf_pages(pdf_path, output_dir, max_pages=3, scale=2.0):
    """Render up to ``max_pages`` of a PDF into PNGs using PDFium or ImageMagick."""
    try:
        # Keep PDFium optional and out of the base PyInstaller dependency graph.
        pdfium = __import__('pypdf' + 'ium2')
    except (ImportError, OSError):
        pdfium = None
    pdf_path, output_dir = Path(pdf_path), Path(output_dir)
    if not pdf_path.exists():
        raise FileNotFoundError(pdf_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    if pdfium is None:
        magick = shutil.which('magick')
        if not magick and os.name != 'nt':
            magick = shutil.which('convert')
        if not magick:
            raise RuntimeError(
                'PDF receipt vision requires pypdfium2 or ImageMagick (install pypdfium2 for PDFium support)'
            )
        density = max(36, int(72 * float(scale)))
        pattern = output_dir / f'{pdf_path.stem}_page-%d.png'
        try:
            subprocess.run(
                [magick, '-density', str(density), f'{pdf_path}[0-{max(0, int(max_pages) - 1)}]',
                 '-alpha', 'remove', '-colorspace', 'sRGB', str(pattern)],
                check=True, capture_output=True, text=True, timeout=120,
            )
        except (OSError, subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
            detail = getattr(exc, 'stderr', '') or str(exc)
            raise RuntimeError(f'ImageMagick could not render {pdf_path}: {detail}') from exc
        return sorted(output_dir.glob(f'{pdf_path.stem}_page-*.png'))[:max(1, int(max_pages))]
    document = pdfium.PdfDocument(str(pdf_path))
    images = []
    try:
        for index in range(min(len(document), max(1, int(max_pages)))):
            page = document[index]
            bitmap = page.render(scale=max(0.5, float(scale)))
            image = bitmap.to_pil().convert('RGB')
            destination = output_dir / f'{pdf_path.stem}_page_{index + 1}.png'
            image.save(destination, format='PNG')
            images.append(destination)
            page.close()
            bitmap.close()
    finally:
        document.close()
    return images


def ocr_receipt_image(image_path, ocr_engine=None):
    """Run injected OCR or optional Tesseract without making it a hard dependency."""
    if ocr_engine:
        return str(ocr_engine(Path(image_path)) or '')
    try:
        import pytesseract
        return str(pytesseract.image_to_string(str(image_path)) or '')
    except ImportError as exc:
        raise RuntimeError('OCR requires pytesseract and a Tesseract installation') from exc
    except Exception as exc:
        raise RuntimeError(f'OCR failed for {image_path}: {exc}') from exc


def _receipt_image_paths(attachment_path, temporary_dir, max_pages=3):
    attachment_path = Path(attachment_path)
    if attachment_path.suffix.lower() == '.pdf':
        return render_pdf_pages(attachment_path, temporary_dir, max_pages=max_pages)
    return [attachment_path]


def _merge_receipt_pages(results, source='', uid='', sender='', category=''):
    normalized = [normalize_receipt_result(item, source, uid, sender, category) for item in results]
    if not normalized:
        return normalize_receipt_result({}, source, uid, sender, category)
    merged = normalized[0].copy()
    for item in normalized[1:]:
        for field in ('merchant', 'date', 'amount', 'currency'):
            if not merged.get(field) and item.get(field):
                merged[field] = item[field]
        merged['line_items'].extend(item.get('line_items', []))
    merged['pages'] = len(normalized)
    return merged


class ReceiptVisionClassifier:
    """PDF/image receipt classifier with Anthropic or local Ollama backends."""

    def __init__(self, backend='anthropic', api_key='', model='claude-3-5-sonnet-20241022',
                 endpoint='http://127.0.0.1:11434', max_pages=3, ocr=False, ocr_engine=None,
                 image_classifier=None):
        self.backend = backend
        self.api_key = api_key
        self.model = model
        self.endpoint = endpoint
        self.max_pages = max_pages
        self.ocr = ocr
        self.ocr_engine = ocr_engine
        self.image_classifier = image_classifier

    def _classify_image(self, image_path, ocr_text=''):
        if self.image_classifier:
            try:
                return self.image_classifier(Path(image_path), ocr_text)
            except TypeError:
                return self.image_classifier(Path(image_path))
        if self.backend == 'anthropic':
            return classify_receipt_image_anthropic(image_path, self.api_key, self.model, ocr_text)
        if self.backend == 'ollama':
            return OllamaClassifier(self.model, self.endpoint).classify_receipt_image(image_path, ocr_text)
        raise ValueError(f'Unsupported receipt vision backend: {self.backend}')

    def classify_attachment(self, attachment_path, uid='', sender='', category=''):
        attachment_path = Path(attachment_path)
        with tempfile.TemporaryDirectory(prefix='gmaildownloader-receipt-') as temporary:
            pages = _receipt_image_paths(attachment_path, temporary, self.max_pages)
            results = []
            for page in pages:
                ocr_text = ocr_receipt_image(page, self.ocr_engine) if self.ocr else ''
                results.append(self._classify_image(page, ocr_text))
        return _merge_receipt_pages(results, attachment_path.name, uid, sender, category)


def classify_receipt_image_anthropic(image_path, api_key, model='claude-3-5-sonnet-20241022', ocr_text=''):
    """Use Anthropic vision for one rendered receipt page."""
    if not api_key:
        raise ValueError('An Anthropic API key is required')
    if not HAS_ANTHROPIC:
        raise RuntimeError('The anthropic package is not installed')
    path = Path(image_path)
    media_type = mimetypes.guess_type(path.name)[0] or 'image/jpeg'
    if media_type not in RECEIPT_IMAGE_TYPES:
        raise ValueError(f'Unsupported receipt image type: {media_type}')
    prompt = RECEIPT_VISION_PROMPT
    if ocr_text:
        prompt += f'\nSupplemental OCR text (verify against the image):\n{ocr_text[:12000]}'
    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model=model,
        max_tokens=800,
        messages=[{'role': 'user', 'content': [
            {'type': 'image', 'source': {'type': 'base64', 'media_type': media_type,
                                         'data': base64.b64encode(path.read_bytes()).decode('ascii')}},
            {'type': 'text', 'text': prompt},
        ]}],
    )
    return _json_object(response.content[0].text)


def extract_receipt_attachments(emails, classifier=None, ocr_engine=None):
    """Classify image/PDF attachments or OCR them when a classifier is absent."""
    receipts = []
    for em in emails:
        if not em.local_path or not Path(em.local_path).exists():
            continue
        try:
            msg = email.message_from_bytes(Path(em.local_path).read_bytes(), policy=email.policy.default)
        except (OSError, ValueError):
            continue
        for filename, payload, media_type in extract_attachments(msg):
            suffix = Path(filename).suffix.lower()
            if suffix not in RECEIPT_ATTACHMENT_EXTENSIONS and media_type not in RECEIPT_IMAGE_TYPES:
                continue
            temporary_path = None
            try:
                with tempfile.NamedTemporaryFile(suffix=suffix or '.bin', delete=False) as temporary:
                    temporary.write(payload)
                    temporary_path = Path(temporary.name)
                if classifier:
                    fields = classifier.classify_attachment(
                        temporary_path, uid=em.uid, sender=em.sender, category=em.category
                    )
                else:
                    with tempfile.TemporaryDirectory(prefix='gmaildownloader-ocr-') as temporary_dir:
                        pages = _receipt_image_paths(temporary_path, temporary_dir)
                        text = '\n'.join(ocr_receipt_image(page, ocr_engine) for page in pages)
                    fields = extract_receipt_fields(f'{filename}\n{text}')
                fields = normalize_receipt_result(fields, filename, em.uid, em.sender, em.category)
                fields['attachment'] = filename
                if fields['amount'] is not None or fields['merchant'] or fields['line_items']:
                    receipts.append(fields)
            except RuntimeError:
                raise
            except (OSError, ValueError):
                continue
            finally:
                if temporary_path:
                    try:
                        temporary_path.unlink()
                    except OSError:
                        pass
    return receipts


def extract_receipts(emails, vision_classifier=None, ocr_engine=None):
    receipts = []
    for em in emails:
        body = ''
        if em.local_path and Path(em.local_path).exists():
            try:
                msg = email.message_from_bytes(Path(em.local_path).read_bytes(), policy=email.policy.default)
                body = extract_message_body(msg, 50000)
            except (OSError, ValueError):
                pass
        fields = extract_receipt_fields(f'{em.subject}\n{body}')
        if fields['amount'] is not None or fields['merchant']:
            fields.update({'uid': em.uid, 'sender': em.sender, 'category': em.category})
            receipts.append(fields)
    if vision_classifier or ocr_engine:
        receipts.extend(extract_receipt_attachments(emails, vision_classifier, ocr_engine))
    return receipts


def export_receipts_ofx(receipts, output_path, account_id='GmailDownloader', currency='USD'):
    """Export normalized receipt debits as an OFX 2-compatible statement."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    transactions = [receipt for receipt in receipts if _receipt_amount(receipt.get('amount')) is not None]
    currencies = [str(receipt.get('currency', '')).upper() for receipt in transactions if receipt.get('currency')]
    currency = currencies[0] if currencies else currency.upper()
    now = datetime.now().strftime('%Y%m%d%H%M%S')
    root = ET.Element('OFX')
    signon = ET.SubElement(root, 'SIGNONMSGSRSV1')
    sonrs = ET.SubElement(signon, 'SONRS')
    status = ET.SubElement(sonrs, 'STATUS')
    ET.SubElement(status, 'CODE').text = '0'
    ET.SubElement(status, 'SEVERITY').text = 'INFO'
    ET.SubElement(sonrs, 'DTSERVER').text = now
    ET.SubElement(sonrs, 'LANGUAGE').text = 'ENG'
    fi = ET.SubElement(sonrs, 'FI')
    ET.SubElement(fi, 'ORG').text = 'GMAILDOWNLOADER'
    ET.SubElement(fi, 'FID').text = 'GMAILDOWNLOADER'
    bank = ET.SubElement(root, 'BANKMSGSRSV1')
    response = ET.SubElement(bank, 'STMTTRNRS')
    ET.SubElement(response, 'TRNUID').text = 'GMAILDOWNLOADER'
    statement = ET.SubElement(response, 'STMTRS')
    ET.SubElement(statement, 'CURDEF').text = currency
    account = ET.SubElement(statement, 'BANKACCTFROM')
    ET.SubElement(account, 'BANKID').text = 'GMAILDOWNLOADER'
    ET.SubElement(account, 'ACCTID').text = account_id
    transactions_node = ET.SubElement(statement, 'BANKTRANLIST')
    dates = [receipt.get('date') for receipt in transactions if receipt.get('date')]
    ET.SubElement(transactions_node, 'DTSTART').text = min(dates).replace('-', '') if dates else now[:8]
    ET.SubElement(transactions_node, 'DTEND').text = max(dates).replace('-', '') if dates else now[:8]
    for receipt in transactions:
        transaction = ET.SubElement(transactions_node, 'STMTTRN')
        ET.SubElement(transaction, 'TRNTYPE').text = 'DEBIT'
        ET.SubElement(transaction, 'DTPOSTED').text = (receipt.get('date') or now[:8]).replace('-', '')[:8]
        amount = _receipt_amount(receipt.get('amount'))
        ET.SubElement(transaction, 'TRNAMT').text = f'{-abs(amount):.2f}'
        identity = f"{receipt.get('uid', '')}|{receipt.get('attachment', '')}|{receipt.get('merchant', '')}"
        ET.SubElement(transaction, 'FITID').text = hashlib.sha256(identity.encode('utf-8')).hexdigest()[:32]
        ET.SubElement(transaction, 'NAME').text = receipt.get('merchant') or receipt.get('sender') or 'Receipt'
        ET.SubElement(transaction, 'MEMO').text = receipt.get('subject') or receipt.get('attachment', '')
    ET.indent(root, space='  ')
    ET.ElementTree(root).write(output_path, encoding='utf-8', xml_declaration=True)
    return str(output_path)


def new_manifest():
    """Return the current manifest shape used by resumable downloads."""
    return {
        'version': MANIFEST_VERSION,
        'app_version': VERSION,
        'folders': {},
        'message_ids': {},
        'folder_metadata': {},
        'sync': {},
    }


def normalize_manifest(manifest):
    """Upgrade old manifests in memory while preserving their message records."""
    if not isinstance(manifest, dict):
        return new_manifest()
    normalized = new_manifest()
    normalized.update(manifest)
    normalized['version'] = max(int(manifest.get('version', 1) or 1), MANIFEST_VERSION)
    normalized['app_version'] = VERSION
    normalized['folders'] = manifest.get('folders', {}) if isinstance(manifest.get('folders', {}), dict) else {}
    normalized['message_ids'] = manifest.get('message_ids', {}) if isinstance(manifest.get('message_ids', {}), dict) else {}
    normalized['folder_metadata'] = manifest.get('folder_metadata', {}) if isinstance(manifest.get('folder_metadata', {}), dict) else {}
    normalized['sync'] = manifest.get('sync', {}) if isinstance(manifest.get('sync', {}), dict) else {}
    return normalized


def load_manifest(path):
    """Load and normalize a manifest, raising only for unreadable JSON."""
    path = Path(path)
    if not path.exists():
        return new_manifest()
    with path.open('r', encoding='utf-8') as fh:
        return normalize_manifest(json.load(fh))


def save_manifest(path, manifest):
    """Atomically persist a manifest so an interrupted sync cannot truncate it."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(normalize_manifest(manifest), ensure_ascii=False, indent=2)
    fd, temporary = tempfile.mkstemp(prefix=f".{path.name}.", suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8', newline='') as fh:
            fh.write(payload)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(temporary, path)
    except Exception:
        try:
            os.unlink(temporary)
        except OSError:
            pass
        raise


def sha256_file(path, chunk_size=1024 * 1024):
    """Return a streaming SHA-256 digest for a local archive file."""
    digest = hashlib.sha256()
    with Path(path).open('rb') as fh:
        while True:
            chunk = fh.read(chunk_size)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def resolve_manifest_path(path, root):
    path = Path(path)
    return path if path.is_absolute() else Path(root) / path


def validate_manifest(manifest, root, update_missing=False):
    """Validate recorded files and return a list of actionable integrity issues.

    Legacy manifests did not contain hashes. When ``update_missing`` is true,
    hashes are filled in for existing files so the next resume has integrity
    protection without forcing a needless full re-download.
    """
    issues = []
    root = Path(root)
    for folder, records in manifest.get('folders', {}).items():
        if not isinstance(records, dict):
            continue
        for uid, info in records.items():
            if not isinstance(info, dict):
                issues.append({'folder': folder, 'uid': uid, 'reason': 'invalid record'})
                continue
            local_path = info.get('local_path', '')
            if not local_path:
                continue
            path = resolve_manifest_path(local_path, root)
            if not path.exists() or not path.is_file():
                issues.append({'folder': folder, 'uid': uid, 'reason': 'missing file', 'path': str(path)})
                continue
            expected = info.get('sha256', '')
            if expected:
                try:
                    actual = sha256_file(path)
                except OSError as exc:
                    issues.append({'folder': folder, 'uid': uid, 'reason': str(exc), 'path': str(path)})
                    continue
                if actual.lower() != str(expected).lower():
                    issues.append({'folder': folder, 'uid': uid, 'reason': 'sha256 mismatch', 'path': str(path)})
            elif update_missing:
                try:
                    info['sha256'] = sha256_file(path)
                except OSError as exc:
                    issues.append({'folder': folder, 'uid': uid, 'reason': str(exc), 'path': str(path)})
    return issues


def email_info_from_record(uid, info, source_folder='', account=''):
    """Create an :class:`EmailInfo` from a manifest or API record."""
    info = info or {}
    date_value = info.get('date', '')
    return EmailInfo(
        uid=uid,
        sender=info.get('sender', ''),
        sender_name=info.get('sender_name', ''),
        sender_domain=info.get('sender_domain', ''),
        subject=info.get('subject', ''),
        date=date_value,
        date_parsed=parse_date(date_value),
        has_list_unsubscribe=info.get('has_list_unsubscribe', False),
        list_unsubscribe_url=info.get('list_unsubscribe_url', ''),
        category=info.get('category', ''),
        confidence=info.get('confidence', 0.0),
        local_path=info.get('local_path', ''),
        source_folder=source_folder or info.get('source_folder', ''),
        message_id=info.get('message_id', ''),
        in_reply_to=info.get('in_reply_to', ''),
        references=info.get('references', ''),
        size_bytes=info.get('size_bytes', 0),
        sensitive_flags=info.get('sensitive_flags', []),
        is_newsletter=info.get('is_newsletter', False),
        account=account or info.get('account', ''),
        has_attachments=info.get('has_attachments', False),
        attachment_names=info.get('attachment_names', []),
    )


def email_info_to_record(em):
    return {
        'sender': em.sender,
        'sender_name': em.sender_name,
        'sender_domain': em.sender_domain,
        'subject': em.subject,
        'date': em.date,
        'has_list_unsubscribe': em.has_list_unsubscribe,
        'list_unsubscribe_url': em.list_unsubscribe_url,
        'category': em.category,
        'confidence': em.confidence,
        'local_path': em.local_path,
        'source_folder': em.source_folder,
        'message_id': em.message_id,
        'in_reply_to': em.in_reply_to,
        'references': em.references,
        'size_bytes': em.size_bytes,
        'sensitive_flags': em.sensitive_flags,
        'is_newsletter': em.is_newsletter,
        'account': em.account,
        'has_attachments': em.has_attachments,
        'attachment_names': em.attachment_names,
    }


def parse_email_message(raw, uid, source_folder='', local_path='', account=''):
    """Parse headers once for all mailbox providers."""
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    from_header = decode_header(msg.get('From', ''))
    sender_name, sender = email.utils.parseaddr(from_header)
    unsubscribe = decode_header(msg.get('List-Unsubscribe', ''))
    attachment_names = []
    for part in msg.walk():
        filename = part.get_filename()
        if filename:
            attachment_names.append(decode_header(filename))
    date_value = decode_header(msg.get('Date', ''))
    return EmailInfo(
        uid=str(uid),
        sender=sender or from_header,
        sender_name=sender_name or sender or from_header,
        subject=decode_header(msg.get('Subject', '(no subject)')),
        date=date_value,
        date_parsed=parse_date(date_value),
        has_list_unsubscribe=bool(unsubscribe),
        list_unsubscribe_url=unsubscribe,
        local_path=str(local_path),
        source_folder=source_folder,
        message_id=decode_header(msg.get('Message-ID', '')),
        in_reply_to=decode_header(msg.get('In-Reply-To', '')),
        references=decode_header(msg.get('References', '')),
        size_bytes=len(raw),
        account=account,
        has_attachments=bool(attachment_names),
        attachment_names=attachment_names,
    )


def parse_since_date(value):
    if value in (None, ''):
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    for fmt in ('%Y-%m-%d', '%d-%b-%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(str(value), fmt)
        except ValueError:
            continue
    raise ValueError(f"Unsupported since date: {value}")


def imap_since_query(value):
    parsed = parse_since_date(value)
    return parsed.strftime('%d-%b-%Y') if parsed else None


def extract_message_body(msg, max_bytes=0):
    """Return the preferred readable body, preferring plain text over HTML."""
    candidates = []
    parts = msg.walk() if msg.is_multipart() else [msg]
    for part in parts:
        content_type = part.get_content_type()
        if content_type not in ('text/plain', 'text/html'):
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        charset = part.get_content_charset() or 'utf-8'
        text_value = payload.decode(charset, errors='replace')
        if content_type == 'text/plain':
            candidates.insert(0, text_value)
        else:
            candidates.append(re.sub(r'<[^>]+>', ' ', text_value))
    body = next((candidate for candidate in candidates if candidate.strip()), '')
    return body[:max_bytes] if max_bytes else body


def extract_attachments(msg):
    attachments = []
    for part in msg.walk():
        filename = part.get_filename()
        payload = part.get_payload(decode=True)
        if not filename or not payload:
            continue
        attachments.append((decode_header(filename), payload, part.get_content_type()))
    return attachments


def extract_attachments_to(raw, output_dir, category='Uncategorized', seen_hashes=None, date_value=None):
    """Extract and SHA-256 deduplicate attachments from one raw message."""
    output_dir = Path(output_dir)
    attachment_dir = output_dir / 'attachments' / sanitize_filename(category, 60)
    attachment_dir.mkdir(parents=True, exist_ok=True)
    seen_hashes = seen_hashes if seen_hashes is not None else {}
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    written = []
    for filename, payload, _ in extract_attachments(msg):
        digest = hashlib.sha256(payload).hexdigest()
        if digest in seen_hashes:
            continue
        seen_hashes[digest] = filename
        safe_name = sanitize_filename(filename, 100)
        prefix = date_value.strftime('%Y-%m-%d') + '_' if date_value else ''
        target = attachment_dir / f'{prefix}{safe_name}'
        suffix = 1
        stem, extension = target.stem, target.suffix
        while target.exists():
            target = attachment_dir / f'{stem}_{suffix}{extension}'
            suffix += 1
        target.write_bytes(payload)
        written.append((str(target), digest, len(payload)))
    return written


def import_mbox(mbox_path, output_dir=None, folder_name=None, attachments_only=False):
    """Import a Google Takeout or generic mbox into ``EmailInfo`` records."""
    mbox_path = Path(mbox_path)
    folder_name = folder_name or mbox_path.stem or 'Imported'
    output_dir = Path(output_dir) if output_dir else None
    folder_dir = output_dir / 'folders' / sanitize_folder_name(folder_name) if output_dir else None
    if folder_dir:
        folder_dir.mkdir(parents=True, exist_ok=True)
    result = []
    seen_hashes = {}
    box = mailbox.mbox(str(mbox_path), create=False)
    try:
        for index, (key, message) in enumerate(box.iteritems()):
            raw = message.as_bytes(policy=email.policy.SMTP)
            local_path = ''
            if output_dir and not attachments_only:
                local_path = str(folder_dir / f'{index:08d}.eml')
                Path(local_path).write_bytes(raw)
            if output_dir and attachments_only:
                extract_attachments_to(raw, output_dir, date_value=parse_email_message(raw, key, folder_name).date_parsed,
                                       seen_hashes=seen_hashes)
            result.append(parse_email_message(raw, f'{folder_name}:{key}', folder_name, local_path))
    finally:
        box.close()
    return result


def import_thunderbird_profile(profile_path, output_dir=None):
    """Import mbox files below a Thunderbird profile, skipping ``.msf`` indexes."""
    profile_path = Path(profile_path)
    roots = [p for p in (profile_path / 'ImapMail', profile_path / 'Mail') if p.exists()]
    if not roots:
        roots = [profile_path]
    result = []
    for root in roots:
        for candidate in sorted(root.rglob('*')):
            if not candidate.is_file() or candidate.name.endswith('.msf'):
                continue
            try:
                relative = candidate.relative_to(root).with_suffix('')
                folder = str(relative).replace(os.sep, '/')
                imported = import_mbox(candidate, output_dir, folder or candidate.stem)
                result.extend(imported)
            except (OSError, ValueError, mailbox.Error):
                continue
    return result


def build_xoauth2_string(address, access_token):
    """Build the SASL XOAUTH2 initial client response used by Gmail IMAP."""
    return f'user={address}\x01auth=Bearer {access_token}\x01\x01'.encode('utf-8')


def open_imap_connection(host, port=993, use_ssl=True, address='', secret='',
                         auth_mode='password', access_token=''):
    """Open a password or XOAUTH2 IMAP connection for Gmail and generic IMAP."""
    if use_ssl:
        connection = imaplib.IMAP4_SSL(host, port)
    else:
        connection = imaplib.IMAP4(host, port)
    if auth_mode.lower() in ('oauth2', 'xoauth2', 'token'):
        token = access_token or secret
        if not token:
            raise ValueError('An OAuth2 access token is required')
        connection.authenticate('XOAUTH2', lambda _: build_xoauth2_string(address, token))
    else:
        connection.login(address, secret)
    return connection


def imap_uidvalidity(imap):
    """Read UIDVALIDITY when the server exposes it, returning an empty string otherwise."""
    try:
        _, response = imap.response('UIDVALIDITY')
        if response:
            raw = response[0] if isinstance(response, (list, tuple)) else response
            match = re.search(rb'\d+', raw if isinstance(raw, bytes) else str(raw).encode())
            if match:
                return match.group().decode()
    except Exception:
        pass
    return ''


class OAuthTokenStore:
    """Small JSON token store that never writes a client secret to disk."""

    def __init__(self, path):
        self.path = Path(path)

    def load(self):
        try:
            data = json.loads(self.path.read_text(encoding='utf-8'))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def save(self, token):
        if not isinstance(token, dict) or not token.get('access_token'):
            raise ValueError('An access token is required')
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(self.path.suffix + '.tmp')
        temporary.write_text(json.dumps(token, indent=2), encoding='utf-8')
        try:
            os.chmod(temporary, 0o600)
        except OSError:
            pass
        os.replace(temporary, self.path)


class GoogleOAuthClient:
    """Minimal Google OAuth authorization-code client using the stdlib only."""

    AUTHORIZE_URL = 'https://accounts.google.com/o/oauth2/v2/auth'
    TOKEN_URL = 'https://oauth2.googleapis.com/token'
    GMAIL_SCOPE = 'https://mail.google.com/'

    def __init__(self, client_id, client_secret='', redirect_uri='urn:ietf:wg:oauth:2.0:oob',
                 opener=None):
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_uri = redirect_uri
        self.opener = opener or urllib.request.urlopen

    def authorization_url(self, state=None, scopes=None):
        state = state or secrets.token_urlsafe(24)
        query = urllib.parse.urlencode({
            'client_id': self.client_id,
            'redirect_uri': self.redirect_uri,
            'response_type': 'code',
            'scope': ' '.join(scopes or [self.GMAIL_SCOPE]),
            'access_type': 'offline',
            'prompt': 'consent',
            'state': state,
        })
        return f'{self.AUTHORIZE_URL}?{query}', state

    def _post_token(self, values):
        request = urllib.request.Request(
            self.TOKEN_URL,
            data=urllib.parse.urlencode(values).encode('utf-8'),
            headers={'Content-Type': 'application/x-www-form-urlencoded'},
            method='POST',
        )
        try:
            with self.opener(request, timeout=30) as response:
                payload = json.loads(response.read().decode('utf-8'))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode('utf-8', errors='replace')
            raise RuntimeError(f'Google OAuth failed ({exc.code}): {detail}') from exc
        if 'error' in payload:
            raise RuntimeError(payload.get('error_description', payload['error']))
        return payload

    def exchange_code(self, code):
        values = {
            'code': code,
            'client_id': self.client_id,
            'redirect_uri': self.redirect_uri,
            'grant_type': 'authorization_code',
        }
        if self.client_secret:
            values['client_secret'] = self.client_secret
        return self._post_token(values)

    def refresh(self, refresh_token):
        values = {
            'refresh_token': refresh_token,
            'client_id': self.client_id,
            'grant_type': 'refresh_token',
        }
        if self.client_secret:
            values['client_secret'] = self.client_secret
        return self._post_token(values)


class GmailApiSource:
    """Gmail REST source with label-aware pagination and raw MIME retrieval."""

    BASE_URL = 'https://gmail.googleapis.com/gmail/v1/users'

    def __init__(self, access_token, user_id='me', opener=None, timeout=60):
        if not access_token:
            raise ValueError('A Gmail API access token is required')
        self.access_token = access_token
        self.user_id = user_id
        self.opener = opener or urllib.request.urlopen
        self.timeout = timeout

    def _request(self, path, params=None):
        url = f'{self.BASE_URL}/{self.user_id}/{path.lstrip("/")}'
        if params:
            url = f'{url}?{urllib.parse.urlencode(params, doseq=True)}'
        request = urllib.request.Request(
            url, headers={'Authorization': f'Bearer {self.access_token}', 'Accept': 'application/json'}
        )
        try:
            with self.opener(request, timeout=self.timeout) as response:
                return json.loads(response.read().decode('utf-8'))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode('utf-8', errors='replace')
            raise RuntimeError(f'Gmail API request failed ({exc.code}): {detail}') from exc

    def list_labels(self):
        payload = self._request('labels')
        return payload.get('labels', [])

    def iter_message_refs(self, query='', label_ids=None, page_size=100):
        token = None
        while True:
            params = {'maxResults': min(max(page_size, 1), 500)}
            if query:
                params['q'] = query
            if label_ids:
                params['labelIds'] = label_ids
            if token:
                params['pageToken'] = token
            payload = self._request('messages', params)
            yield from payload.get('messages', [])
            token = payload.get('nextPageToken')
            if not token:
                break

    def fetch_message(self, message_id):
        payload = self._request(f'messages/{urllib.parse.quote(str(message_id), safe="")}', {'format': 'raw'})
        raw = payload.get('raw', '')
        if not raw:
            raise RuntimeError(f'Gmail API returned no raw MIME for {message_id}')
        return base64.urlsafe_b64decode(raw + '=' * (-len(raw) % 4))

    def iter_messages(self, query='', label_ids=None, page_size=100):
        for reference in self.iter_message_refs(query, label_ids, page_size):
            message_id = reference.get('id')
            if message_id:
                yield message_id, self.fetch_message(message_id), reference.get('labelIds', [])


class GraphMailSource:
    """Microsoft Graph adapter returning RFC822 MIME for the common mail path."""

    BASE_URL = 'https://graph.microsoft.com/v1.0/me'

    def __init__(self, access_token, opener=None, timeout=60):
        if not access_token:
            raise ValueError('A Microsoft Graph access token is required')
        self.access_token = access_token
        self.opener = opener or urllib.request.urlopen
        self.timeout = timeout

    def _request(self, url, accept='application/json'):
        request = urllib.request.Request(
            url, headers={'Authorization': f'Bearer {self.access_token}', 'Accept': accept}
        )
        try:
            with self.opener(request, timeout=self.timeout) as response:
                return response.read()
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode('utf-8', errors='replace')
            raise RuntimeError(f'Microsoft Graph request failed ({exc.code}): {detail}') from exc

    def iter_message_refs(self, folder='inbox', query='', page_size=100):
        url = f'{self.BASE_URL}/mailFolders/{urllib.parse.quote(folder, safe="")}/messages'
        params = {'$top': min(max(page_size, 1), 999), '$select': 'id,subject,receivedDateTime'}
        if query:
            params['$search'] = f'"{query}"'
        url = f'{url}?{urllib.parse.urlencode(params)}'
        while url:
            payload = json.loads(self._request(url).decode('utf-8'))
            yield from payload.get('value', [])
            url = payload.get('@odata.nextLink', '')

    def fetch_message(self, message_id):
        url = f'{self.BASE_URL}/messages/{urllib.parse.quote(str(message_id), safe="")}/$value'
        return self._request(url, accept='message/rfc822')

    def iter_messages(self, folder='inbox', query='', page_size=100):
        for reference in self.iter_message_refs(folder, query, page_size):
            message_id = reference.get('id')
            if message_id:
                yield message_id, self.fetch_message(message_id), [folder]


class MultiAccountManager:
    """Resolve isolated output trees for several accounts without sharing manifests."""

    def __init__(self, root_dir):
        self.root_dir = Path(root_dir)
        self.accounts = {}

    def add(self, account):
        if not account.name or not account.address:
            raise ValueError('Account name and address are required')
        self.accounts[account.name] = account

    def remove(self, name):
        self.accounts.pop(name, None)

    def output_dir(self, account_or_name):
        account = self.accounts.get(account_or_name, account_or_name)
        account_name = account.name if isinstance(account, AccountConfig) else str(account)
        return self.root_dir / sanitize_filename(account_name, 80)

    def manifest_paths(self):
        return {name: self.output_dir(name) / MANIFEST_FILENAME for name in self.accounts}

    def sync_workers(self, options=None):
        """Create isolated workers; callers can start them sequentially or concurrently."""
        workers = []
        for account in self.accounts.values():
            destination = Path(account.output_dir) if account.output_dir else self.output_dir(account)
            workers.append(ImapDownloadWorker(
                account.host, account.address, account.secret, destination,
                options=options or SyncOptions(), port=account.port, use_ssl=account.use_ssl,
                auth_mode=account.auth_mode,
                access_token=account.secret if account.auth_mode in ('oauth2', 'xoauth2', 'token') else '',
            ))
        return workers

    def load_engines(self, verify_integrity=True):
        engines = {}
        for name, account in self.accounts.items():
            destination = Path(account.output_dir) if account.output_dir else self.output_dir(name)
            emails, _ = load_archive_emails(destination, verify_integrity)
            domain = account.address.split('@', 1)[1] if '@' in account.address else ''
            engine = CategoryEngine(domain)
            engine.process_all(emails)
            engines[name] = engine
        return engines

    def side_by_side_summary(self, verify_integrity=True):
        return {name: engine.get_summary() for name, engine in self.load_engines(verify_integrity).items()}

    def save_config(self, path):
        data = [{
            'name': account.name,
            'address': account.address,
            'host': account.host,
            'port': account.port,
            'use_ssl': account.use_ssl,
            'auth_mode': account.auth_mode,
            'output_dir': account.output_dir,
        } for account in self.accounts.values()]
        Path(path).write_text(json.dumps(data, indent=2), encoding='utf-8')

    def load_config(self, path):
        data = json.loads(Path(path).read_text(encoding='utf-8'))
        self.accounts.clear()
        for record in data:
            self.add(AccountConfig(**record))
        return self.accounts


class OllamaClassifier:
    """Optional local LLM client for offline classification and vision."""

    def __init__(self, model='llama3.2', endpoint='http://127.0.0.1:11434', opener=None, timeout=120):
        self.model = model
        self.endpoint = endpoint.rstrip('/')
        self.opener = opener or urllib.request.urlopen
        self.timeout = timeout

    def complete(self, prompt, images=None):
        payload = {'model': self.model, 'prompt': prompt, 'stream': False}
        if images:
            payload['images'] = images
        request = urllib.request.Request(
            f'{self.endpoint}/api/generate',
            data=json.dumps(payload).encode('utf-8'),
            headers={'Content-Type': 'application/json'},
            method='POST',
        )
        try:
            with self.opener(request, timeout=self.timeout) as response:
                data = json.loads(response.read().decode('utf-8'))
        except urllib.error.URLError as exc:
            raise RuntimeError(f'Ollama is unavailable at {self.endpoint}: {exc}') from exc
        return data.get('response', '')

    def classify_domains(self, domain_info, existing_categories):
        prompt = (
            'Classify each email domain into the closest existing category. '
            'Return ONLY a JSON object mapping domain to category.\n'
            f'Existing categories: {json.dumps(existing_categories)}\n'
            f'Domains: {json.dumps(domain_info, ensure_ascii=False)}'
        )
        match = re.search(r'\{.*\}', self.complete(prompt), re.DOTALL)
        return json.loads(match.group()) if match else {}

    def classify_receipt_image(self, image_path, ocr_text=''):
        path = Path(image_path)
        media_type = mimetypes.guess_type(path.name)[0] or 'image/jpeg'
        prompt = RECEIPT_VISION_PROMPT
        if ocr_text:
            prompt += f'\nSupplemental OCR text (verify against the image):\n{ocr_text[:12000]}'
        response = self.complete(prompt, [base64.b64encode(path.read_bytes()).decode('ascii')])
        result = _json_object(response)
        return result if result else {'raw': response, 'media_type': media_type}


def classify_receipt_image_anthropic(image_path, api_key, model='claude-3-5-sonnet-20241022'):
    """Use Anthropic vision for one image attachment when an API key is supplied."""
    if not api_key:
        raise ValueError('An Anthropic API key is required')
    if not HAS_ANTHROPIC:
        raise RuntimeError('The anthropic package is not installed')
    path = Path(image_path)
    media_type = mimetypes.guess_type(path.name)[0] or 'image/jpeg'
    if media_type not in ('image/jpeg', 'image/png', 'image/gif', 'image/webp'):
        raise ValueError(f'Unsupported receipt image type: {media_type}')
    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model=model,
        max_tokens=600,
        messages=[{'role': 'user', 'content': [
            {'type': 'image', 'source': {'type': 'base64', 'media_type': media_type,
                                         'data': base64.b64encode(path.read_bytes()).decode('ascii')}},
            {'type': 'text', 'text': 'Extract merchant, date, total amount, currency, and line items as JSON. Return only JSON.'},
        ]}],
    )
    text_response = response.content[0].text
    match = re.search(r'\{.*\}', text_response, re.DOTALL)
    return json.loads(match.group()) if match else {'raw': text_response}


def extract_receipts(emails):
    receipts = []
    for em in emails:
        body = ''
        if em.local_path and Path(em.local_path).exists():
            try:
                msg = email.message_from_bytes(Path(em.local_path).read_bytes(), policy=email.policy.default)
                body = extract_message_body(msg, 50000)
            except OSError:
                pass
        fields = extract_receipt_fields(f'{em.subject}\n{body}')
        if fields['amount'] is not None or fields['merchant']:
            fields.update({'uid': em.uid, 'sender': em.sender, 'category': em.category})
            receipts.append(fields)
    return receipts


def load_archive_emails(output_dir, verify_integrity=True):
    """Load unique local messages from a downloaded archive manifest."""
    output_dir = Path(output_dir)
    manifest = load_manifest(output_dir / MANIFEST_FILENAME)
    issues = validate_manifest(manifest, output_dir, update_missing=False) if verify_integrity else []
    invalid = {(issue.get('folder'), issue.get('uid')) for issue in issues}
    emails, seen = [], set()
    for folder, records in manifest.get('folders', {}).items():
        if not isinstance(records, dict):
            continue
        for uid, info in records.items():
            if (folder, uid) in invalid or not isinstance(info, dict) or info.get('skipped'):
                continue
            message_id = info.get('message_id', '')
            if message_id and message_id in seen:
                continue
            if message_id:
                seen.add(message_id)
            emails.append(email_info_from_record(f'{folder}:{uid}', info, folder))
    return emails, issues


def build_cron_entry(script_path, output_dir, schedule='0 2 * * *'):
    """Return a cron entry for a headless incremental backup."""
    script = str(Path(script_path).resolve()).replace(' ', '\\ ')
    output = str(Path(output_dir).resolve()).replace(' ', '\\ ')
    return f'{schedule} python3 {script} --headless --sync --output-dir {output}'


def build_windows_task_args(task_name, script_path, output_dir, schedule='DAILY', start_time='02:00'):
    return [
        'schtasks', '/Create', '/F', '/TN', task_name,
        '/SC', schedule, '/ST', start_time,
        '/TR', f'"{sys.executable}" "{Path(script_path).resolve()}" --headless --sync --output-dir "{Path(output_dir).resolve()}"',
    ]


def install_windows_scheduled_backup(task_name, script_path, output_dir, schedule='DAILY', start_time='02:00', dry_run=True):
    """Install a Windows Task Scheduler job, or return its command in dry-run mode."""
    command = build_windows_task_args(task_name, script_path, output_dir, schedule, start_time)
    if dry_run:
        return command
    return subprocess.run(command, check=True, capture_output=True, text=True)


def export_mbox(emails, output_path):
    """Write local EML messages to an interoperable mbox file."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    box = mailbox.mbox(str(output_path), create=True)
    try:
        for em in emails:
            raw = b''
            if em.local_path and Path(em.local_path).exists():
                raw = Path(em.local_path).read_bytes()
            if raw:
                box.add(email.message_from_bytes(raw, policy=email.policy.default))
                continue
            msg = EmailMessage()
            if em.subject:
                msg['Subject'] = em.subject
            if em.sender:
                msg['From'] = em.sender
            if em.date:
                msg['Date'] = em.date
            if em.message_id:
                msg['Message-ID'] = em.message_id
            msg.set_content('Message body is not available in headers-only mode.')
            box.add(msg)
        box.flush()
    finally:
        box.close()
    return str(output_path)


def redact_sensitive_text(text):
    """Redact detected secret values while retaining the finding category."""
    redacted = text
    found = []
    for pattern, label in SENSITIVE_PATTERNS:
        if re.search(pattern, redacted):
            found.append(label)
            redacted = re.sub(pattern, f'[REDACTED:{label}]', redacted)
    return redacted, found


def redact_eml(source_path, destination_path):
    """Write a redacted copy of an EML; the original is never modified."""
    source_path, destination_path = Path(source_path), Path(destination_path)
    msg = email.message_from_bytes(source_path.read_bytes(), policy=email.policy.default)
    flags = set()
    parts = msg.walk() if msg.is_multipart() else [msg]
    for part in parts:
        if part.get_content_type() not in ('text/plain', 'text/html'):
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        charset = part.get_content_charset() or 'utf-8'
        body = payload.decode(charset, errors='replace')
        body, body_flags = redact_sensitive_text(body)
        flags.update(body_flags)
        subtype = part.get_content_subtype()
        part.set_content(body, subtype=subtype, charset=charset)
    destination_path.parent.mkdir(parents=True, exist_ok=True)
    with destination_path.open('wb') as fh:
        BytesGenerator(fh, policy=email.policy.SMTP).flatten(msg)
    return sorted(flags)


def _zip_directory(source_dir, zip_path):
    source_dir = Path(source_dir)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for path in source_dir.rglob('*'):
            if path.is_file():
                archive.write(path, path.relative_to(source_dir).as_posix())


def encrypt_archive(source_dir, output_path, passphrase):
    """Create an AES-256-GCM encrypted ZIP archive using a passphrase."""
    if not passphrase:
        raise ValueError('A non-empty passphrase is required')
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        from cryptography.hazmat.primitives.kdf.scrypt import Scrypt
    except ImportError as exc:
        raise RuntimeError('Encrypted archives require the cryptography package') from exc
    salt, nonce = secrets.token_bytes(16), secrets.token_bytes(12)
    key = Scrypt(salt=salt, length=32, n=2**14, r=8, p=1).derive(passphrase.encode('utf-8'))
    with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as temporary:
        zip_path = Path(temporary.name)
    try:
        _zip_directory(source_dir, zip_path)
        encrypted = AESGCM(key).encrypt(nonce, zip_path.read_bytes(), None)
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(b'GMAILDOWNLOADER-AES256\n' + salt + nonce + encrypted)
    finally:
        try:
            zip_path.unlink()
        except OSError:
            pass
    return str(output_path)


def decrypt_archive(archive_path, output_dir, passphrase):
    """Decrypt an archive created by :func:`encrypt_archive` safely."""
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        from cryptography.hazmat.primitives.kdf.scrypt import Scrypt
    except ImportError as exc:
        raise RuntimeError('Encrypted archives require the cryptography package') from exc
    payload = Path(archive_path).read_bytes()
    magic = b'GMAILDOWNLOADER-AES256\n'
    if not payload.startswith(magic) or len(payload) <= len(magic) + 28:
        raise ValueError('Not a GmailDownloader encrypted archive')
    offset = len(magic)
    salt, nonce, encrypted = payload[offset:offset + 16], payload[offset + 16:offset + 28], payload[offset + 28:]
    key = Scrypt(salt=salt, length=32, n=2**14, r=8, p=1).derive(passphrase.encode('utf-8'))
    raw_zip = AESGCM(key).decrypt(nonce, encrypted, None)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(raw_zip)) as archive:
        root = output_dir.resolve()
        for member in archive.infolist():
            target = (output_dir / member.filename).resolve()
            if root != target and root not in target.parents:
                raise ValueError('Archive contains an unsafe path')
        archive.extractall(output_dir)
    return str(output_dir)


# ─── IMAP Workers ─────────────────────────────────────────────────────────

class ImapScanWorker(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    email_batch = pyqtSignal(list)
    finished_signal = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, host, addr, pw, port=993, use_ssl=True, auth_mode='password',
                 access_token='', since=None):
        super().__init__()
        self.host, self.addr, self.pw = host, addr, pw
        self.port, self.use_ssl = port, use_ssl
        self.auth_mode, self.access_token = auth_mode, access_token
        self.since = since
        self._stop = False

    def stop(self): self._stop = True

    def run(self):
        try:
            self.status.emit("Connecting...")
            imap = open_imap_connection(
                self.host, self.port, self.use_ssl, self.addr, self.pw,
                self.auth_mode, self.access_token
            )
            imap.select('INBOX', readonly=True)
            since_query = imap_since_query(self.since)
            if since_query:
                _, data = imap.uid('SEARCH', None, 'SINCE', since_query)
            else:
                _, data = imap.uid('SEARCH', None, 'ALL')
            uids = data[0].split()
            total = len(uids)
            self.status.emit(f"Found {total:,} emails. Scanning...")
            all_emails, bs = [], 200
            for i in range(0, total, bs):
                if self._stop: break
                batch_uids = uids[i:i+bs]
                _, msg_data = imap.uid('FETCH', b','.join(batch_uids),
                    '(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE LIST-UNSUBSCRIBE MESSAGE-ID IN-REPLY-TO REFERENCES)] RFC822.SIZE)')
                batch, idx = [], 0
                for item in msg_data:
                    if isinstance(item, tuple) and len(item) == 2:
                        um = re.search(rb'UID (\d+)', item[0])
                        uid = um.group(1).decode() if um else batch_uids[idx].decode()
                        sz_m = re.search(rb'RFC822\.SIZE (\d+)', item[0])
                        size = int(sz_m.group(1)) if sz_m else 0
                        idx += 1
                        try:
                            em = parse_email_message(item[1], uid, 'INBOX', account=self.addr)
                            em.size_bytes = size
                            batch.append(em); all_emails.append(em)
                        except Exception: pass
                self.progress.emit(min(i+bs, total), total)
                if batch: self.email_batch.emit(batch)
            try: imap.close(); imap.logout()
            except: pass
            self.finished_signal.emit(all_emails)
        except Exception as e:
            self.error.emit(f"Error: {e}\n{traceback.format_exc()}")


class ImapDownloadWorker(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    log = pyqtSignal(str)
    folder_started = pyqtSignal(str, int)
    email_saved = pyqtSignal(object)
    finished_signal = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, host, addr, pw, output_dir, skip=None, since=None,
                 options=None, port=993, use_ssl=True, auth_mode='password',
                 access_token=''):
        super().__init__()
        self.host, self.addr, self.pw = host, addr, pw
        self.output_dir = Path(output_dir)
        self.skip = skip or GMAIL_SKIP_FOLDERS
        self.options = options or SyncOptions(since=parse_since_date(since))
        self.port, self.use_ssl = port, use_ssl
        self.auth_mode, self.access_token = auth_mode, access_token
        self._stop = False

    def stop(self): self._stop = True

    def run(self):
        manifest = new_manifest()
        manifest_path = self.output_dir / MANIFEST_FILENAME
        imap = None
        try:
            folders_dir = self.output_dir / "folders"
            folders_dir.mkdir(parents=True, exist_ok=True)
            try:
                manifest = load_manifest(manifest_path)
            except (OSError, ValueError, TypeError) as exc:
                self.log.emit(f"Manifest could not be read; starting safely: {exc}")
                manifest = new_manifest()

            integrity_issues = validate_manifest(
                manifest, self.output_dir, update_missing=True
            ) if self.options.verify_integrity else []
            invalid_paths = {issue.get('path') for issue in integrity_issues if issue.get('path')}
            for issue in integrity_issues:
                folder = issue.get('folder')
                uid = issue.get('uid')
                self.log.emit(f"[{folder}:{uid}] {issue.get('reason', 'integrity failure')}; re-fetching")
                records = manifest.get('folders', {}).get(folder, {})
                if isinstance(records, dict):
                    records.pop(uid, None)
            seen_ids = manifest.get('message_ids', {})
            for mid, local_path in list(seen_ids.items()):
                if not local_path:
                    continue
                resolved = resolve_manifest_path(local_path, self.output_dir)
                if str(resolved) in invalid_paths or not resolved.exists():
                    seen_ids.pop(mid, None)

            self.status.emit("Connecting...")
            imap = open_imap_connection(
                self.host, self.port, self.use_ssl, self.addr, self.pw,
                self.auth_mode, self.access_token
            )
            _, folder_data = imap.list()
            all_folders = []
            for line in folder_data:
                if isinstance(line, bytes):
                    n = parse_imap_folder_list(line)
                    if n: all_folders.append(n)
            active = [f for f in all_folders if f not in self.skip]
            self.log.emit(f"Downloading {len(active)} folders, skipping {len(self.skip)}")

            folder_uids = {}
            folder_uidvalidity = {}
            total_count = 0
            for fn in active:
                if self._stop: break
                try:
                    if imap.select(f'"{fn}"', readonly=True)[0] != 'OK': continue
                    current_validity = imap_uidvalidity(imap)
                    folder_uidvalidity[fn] = current_validity
                    metadata = manifest.setdefault('folder_metadata', {}).setdefault(fn, {})
                    previous_validity = str(metadata.get('uidvalidity', ''))
                    if previous_validity and current_validity and previous_validity != current_validity:
                        self.log.emit(f"[{fn}] UIDVALIDITY changed; rebuilding folder index")
                        manifest['folders'][fn] = {}
                        metadata.clear()
                    fm = manifest.setdefault('folders', {}).setdefault(fn, {})
                    since_query = imap_since_query(self.options.since)
                    if since_query:
                        _, d = imap.uid('SEARCH', None, 'SINCE', since_query)
                    else:
                        _, d = imap.uid('SEARCH', None, 'ALL')
                    u = d[0].split() if d[0] else []
                    last_uid = int(metadata.get('last_uid', 0) or 0)
                    if self.options.incremental and last_uid and not since_query:
                        u = [uid for uid in u if uid.isdigit() and int(uid) > last_uid]
                    folder_uids[fn] = u; total_count += len(u)
                    mode = "delta" if self.options.incremental else "full"
                    if since_query:
                        mode = f"since {since_query}"
                    self.log.emit(f"  {fn}: {len(u):,} candidates ({mode})")
                except Exception as e:
                    self.log.emit(f"  {fn}: error - {e}")

            all_emails = []
            gp = 0
            for fn, fm in manifest.get('folders', {}).items():
                if not isinstance(fm, dict):
                    continue
                for uid_str, info in fm.items():
                    all_emails.append(email_info_from_record(f"{fn}:{uid_str}", info, fn, self.addr))

            for fn, uids in folder_uids.items():
                if self._stop: break
                if not uids: continue
                safe = sanitize_folder_name(fn)
                fdir = folders_dir / safe; fdir.mkdir(parents=True, exist_ok=True)
                fm = manifest['folders'].get(fn, {})
                remaining = [u for u in uids if u.decode() not in fm]
                already = len(uids) - len(remaining)
                gp += already
                self.folder_started.emit(fn, len(uids))
                if already: self.log.emit(f"[{safe}] {already:,} cached, {len(remaining):,} remaining")
                else: self.log.emit(f"[{safe}] Downloading {len(uids):,}...")
                imap.select(f'"{fn}"', readonly=True)

                for i in range(0, len(remaining), 50):
                    if self._stop: break
                    bu = remaining[i:i+50]
                    try: _, md = imap.uid('FETCH', b','.join(bu), '(RFC822)')
                    except Exception as exc:
                        self.log.emit(f"[{safe}] fetch failed: {exc}")
                        continue
                    for item in md:
                        if self._stop: break
                        if not isinstance(item, tuple) or len(item) != 2: continue
                        um = re.search(rb'UID (\d+)', item[0])
                        if not um: continue
                        uid = um.group(1).decode()
                        raw = item[1]
                        if not isinstance(raw, bytes): continue
                        if self.options.max_message_size and len(raw) > self.options.max_message_size:
                            self.log.emit(f"[{safe}:{uid}] skipped ({format_size(len(raw))} exceeds limit)")
                            fm[uid] = {'size_bytes': len(raw), 'skipped': True, 'source_folder': fn}
                            continue
                        try:
                            em = parse_email_message(raw, f"{fn}:{uid}", fn, account=self.addr)
                        except Exception as exc:
                            self.log.emit(f"[{safe}:{uid}] header parse failed: {exc}")
                            continue
                        eml_path = ''
                        if self.options.attachments_only:
                            extract_attachments_to(raw, self.output_dir, date_value=em.date_parsed)
                        elif em.message_id and em.message_id in seen_ids:
                            candidate = resolve_manifest_path(seen_ids[em.message_id], self.output_dir)
                            if candidate.exists():
                                eml_path = str(candidate)
                        if not self.options.attachments_only and not eml_path:
                            eml_path = str(fdir / f"{uid}.eml")
                            try:
                                Path(eml_path).write_bytes(raw)
                            except OSError as exc:
                                self.log.emit(f"[{safe}:{uid}] save failed: {exc}")
                                continue
                            if em.message_id:
                                seen_ids[em.message_id] = eml_path
                        em.local_path = eml_path
                        em.size_bytes = len(raw)
                        all_emails.append(em)
                        self.email_saved.emit(em)
                        record = email_info_to_record(em)
                        if eml_path:
                            record['sha256'] = sha256_file(eml_path)
                        fm[uid] = record
                    gp += len(bu)
                    self.progress.emit(gp, total_count)
                manifest['message_ids'] = seen_ids
                metadata = manifest.setdefault('folder_metadata', {}).setdefault(fn, {})
                metadata['uidvalidity'] = folder_uidvalidity.get(fn, metadata.get('uidvalidity', ''))
                if not self.options.since and not self._stop and uids:
                    metadata['last_uid'] = max(int(uid) for uid in uids if uid.isdigit())
                metadata['last_sync'] = datetime.now().isoformat(timespec='seconds')
                save_manifest(manifest_path, manifest)
            try: imap.logout()
            except Exception: pass
            manifest['message_ids'] = seen_ids
            manifest['sync'] = {
                'last_successful_at': datetime.now().isoformat(timespec='seconds'),
                'incremental': self.options.incremental,
                'since': self.options.since.isoformat() if self.options.since else '',
                'attachments_only': self.options.attachments_only,
            }
            save_manifest(manifest_path, manifest)
            self.finished_signal.emit(all_emails)
        except Exception as e:
            try:
                manifest['message_ids'] = manifest.get('message_ids', {})
                save_manifest(manifest_path, manifest)
            except Exception:
                pass
            if imap is not None:
                try: imap.logout()
                except Exception: pass
            self.error.emit(f"Error: {e}\n{traceback.format_exc()}")


class RemoteMimeDownloadWorker(QThread):
    """Download raw MIME from any paginated source implementing ``iter_messages``."""

    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    log = pyqtSignal(str)
    email_saved = pyqtSignal(object)
    finished_signal = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, source, output_dir, folder_name='Imported', options=None, query=''):
        super().__init__()
        self.source = source
        self.output_dir = Path(output_dir)
        self.folder_name = folder_name
        self.options = options or SyncOptions()
        self.query = query
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        manifest_path = self.output_dir / MANIFEST_FILENAME
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            manifest = load_manifest(manifest_path)
            issues = validate_manifest(manifest, self.output_dir, update_missing=True) \
                if self.options.verify_integrity else []
            for issue in issues:
                manifest.get('folders', {}).get(issue.get('folder'), {}).pop(issue.get('uid'), None)
            records = manifest.setdefault('folders', {}).setdefault(self.folder_name, {})
            stored_ids = {info.get('message_id') for info in records.values() if isinstance(info, dict)}
            stored_ids.discard(None)
            query = self.query
            since = self.options.since
            if since and not query:
                query = f'after:{since.strftime("%Y/%m/%d")}'
            messages = list(self.source.iter_messages(query=query))
            total = len(messages)
            self.status.emit(f"Found {total:,} remote messages")
            folder_dir = self.output_dir / 'folders' / sanitize_folder_name(self.folder_name)
            folder_dir.mkdir(parents=True, exist_ok=True)
            all_emails = [
                email_info_from_record(f"{self.folder_name}:{uid}", info, self.folder_name)
                for uid, info in records.items() if isinstance(info, dict)
            ]
            for index, (remote_id, raw, labels) in enumerate(messages):
                if self._stop:
                    break
                remote_id = str(remote_id)
                if self.options.incremental and remote_id in stored_ids:
                    self.progress.emit(index + 1, total)
                    continue
                if self.options.max_message_size and len(raw) > self.options.max_message_size:
                    records[remote_id] = {'message_id': remote_id, 'size_bytes': len(raw), 'skipped': True}
                    continue
                source_folder = self.folder_name
                if labels:
                    source_folder = f"{self.folder_name}/{','.join(map(str, labels))}"
                em = parse_email_message(raw, f"{self.folder_name}:{remote_id}", source_folder)
                em.message_id = em.message_id or remote_id
                if self.options.attachments_only:
                    extract_attachments_to(raw, self.output_dir, date_value=em.date_parsed)
                else:
                    local_path = folder_dir / f'{sanitize_filename(remote_id, 100)}.eml'
                    local_path.write_bytes(raw)
                    em.local_path = str(local_path)
                records[remote_id] = email_info_to_record(em)
                if em.local_path:
                    records[remote_id]['sha256'] = sha256_file(em.local_path)
                    manifest.setdefault('message_ids', {})[em.message_id] = em.local_path
                all_emails.append(em)
                stored_ids.add(remote_id)
                self.email_saved.emit(em)
                self.progress.emit(index + 1, total)
            manifest.setdefault('folder_metadata', {}).setdefault(self.folder_name, {}).update({
                'last_sync': datetime.now().isoformat(timespec='seconds'),
                'source': type(self.source).__name__,
            })
            manifest['sync'] = {'last_successful_at': datetime.now().isoformat(timespec='seconds'), 'query': query}
            save_manifest(manifest_path, manifest)
            self.finished_signal.emit(all_emails)
        except Exception as exc:
            self.error.emit(f"Error: {exc}\n{traceback.format_exc()}")


class GmailApiDownloadWorker(RemoteMimeDownloadWorker):
    def __init__(self, source, output_dir, options=None, query=''):
        super().__init__(source, output_dir, 'Gmail API', options, query)


class RemoteMimeScanWorker(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    email_batch = pyqtSignal(list)
    finished_signal = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, source, query=''):
        super().__init__()
        self.source, self.query = source, query
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            messages = list(self.source.iter_messages(query=self.query))
            total = len(messages)
            result = []
            for index, (remote_id, raw, labels) in enumerate(messages):
                if self._stop:
                    break
                folder = f"Remote/{','.join(map(str, labels))}" if labels else 'Remote'
                result.append(parse_email_message(raw, remote_id, folder))
                self.email_batch.emit([result[-1]])
                self.progress.emit(index + 1, total)
            self.finished_signal.emit(result)
        except Exception as exc:
            self.error.emit(f"Error: {exc}\n{traceback.format_exc()}")


class ImapLabelWorker(QThread):
    progress = pyqtSignal(int, int); status = pyqtSignal(str)
    log = pyqtSignal(str); finished_signal = pyqtSignal(); error = pyqtSignal(str)
    def __init__(self, host, addr, pw, cats, prefix="", archive=False, dry_run=True,
                 port=993, use_ssl=True, auth_mode='password', access_token=''):
        super().__init__()
        self.host, self.addr, self.pw = host, addr, pw
        self.cats, self.prefix, self.archive = cats, prefix, archive
        self.dry_run = dry_run
        self.port, self.use_ssl = port, use_ssl
        self.auth_mode, self.access_token = auth_mode, access_token
        self._stop = False
    def stop(self): self._stop = True
    def run(self):
        try:
            total = sum(len(v) for v in self.cats.values()); done = 0
            if self.dry_run:
                self.status.emit("Preview only — no mailbox changes")
                for cat, emails in self.cats.items():
                    if self._stop: break
                    label = f"{self.prefix}/{cat}" if self.prefix else cat
                    self.log.emit(f"Would label {len(emails):,} messages as '{label}'")
                    done += len(emails); self.progress.emit(done, total)
                self.finished_signal.emit()
                return
            imap = open_imap_connection(
                self.host, self.port, self.use_ssl, self.addr, self.pw,
                self.auth_mode, self.access_token
            )
            for cat, emails in self.cats.items():
                if self._stop: break
                if not emails: continue
                label = f"{self.prefix}/{cat}" if self.prefix else cat
                try: imap.create(f'"{label}"')
                except: pass
                by_folder = defaultdict(list)
                for em in emails:
                    f = em.source_folder or 'INBOX'
                    by_folder[f].append(em.uid.split(':',1)[1] if ':' in em.uid else em.uid)
                for folder, uids in by_folder.items():
                    if self._stop: break
                    imap.select(f'"{folder}"')
                    for i in range(0, len(uids), 100):
                        b = uids[i:i+100]
                        try: imap.uid('COPY', ','.join(b), f'"{label}"')
                        except Exception as e: self.log.emit(f"Error: {e}")
                        if self.archive and folder.upper() == 'INBOX':
                            try:
                                imap.uid('STORE', ','.join(b), '-FLAGS', '(\\Inbox)')
                            except Exception as e:
                                self.log.emit(f"Archive error: {e}")
                        done += len(b); self.progress.emit(done, total)
                self.log.emit(f"Labeled {len(emails):,} as '{label}'")
            try: imap.logout()
            except: pass
            self.finished_signal.emit()
        except Exception as e: self.error.emit(str(e))


class LocalOrganizeWorker(QThread):
    progress = pyqtSignal(int, int); status = pyqtSignal(str)
    log = pyqtSignal(str); finished_signal = pyqtSignal(); error = pyqtSignal(str)
    def __init__(self, cats, out_dir, copy=True):
        super().__init__()
        self.cats, self.out_dir, self.copy = cats, Path(out_dir), copy
        self._stop = False
    def stop(self): self._stop = True
    def run(self):
        try:
            od = self.out_dir / "organized"; od.mkdir(parents=True, exist_ok=True)
            total = sum(len(v) for v in self.cats.values()); done = 0
            for cat, emails in self.cats.items():
                if self._stop: break
                if not emails: continue
                cf = od / sanitize_filename(cat.replace('/',os.sep), 120)
                cf.mkdir(parents=True, exist_ok=True)
                self.log.emit(f"[{cat}] {len(emails):,} emails")
                for em in emails:
                    if self._stop: break
                    src = Path(em.local_path) if em.local_path else None
                    if not src or not src.exists(): done += 1; continue
                    ds = em.date_parsed.strftime("%Y-%m-%d") if em.date_parsed else "unknown"
                    fn = f"{ds}_{sanitize_filename(em.sender_domain or em.sender_name,25)}_{sanitize_filename(em.subject,45)}.eml"
                    dst = cf / fn; c = 1
                    while dst.exists(): dst = cf / f"{ds}_{sanitize_filename(em.sender_domain,25)}_{sanitize_filename(em.subject,40)}_{c}.eml"; c += 1
                    try:
                        if self.copy: shutil.copy2(str(src), str(dst))
                        else: shutil.move(str(src), str(dst))
                    except Exception as e: self.log.emit(f"  Error: {e}")
                    done += 1
                    if done % 500 == 0: self.progress.emit(done, total)
                self.progress.emit(done, total)
            self.finished_signal.emit()
        except Exception as e: self.error.emit(str(e))


class AttachmentExtractWorker(QThread):
    progress = pyqtSignal(int, int); status = pyqtSignal(str)
    log = pyqtSignal(str); finished_signal = pyqtSignal(int, int, str)  # count, size, path
    error = pyqtSignal(str)
    def __init__(self, emails, out_dir):
        super().__init__()
        self.emails, self.out_dir = emails, Path(out_dir)
        self._stop = False
    def stop(self): self._stop = True
    def run(self):
        try:
            att_dir = self.out_dir / "attachments"; att_dir.mkdir(parents=True, exist_ok=True)
            seen_hashes = {}; total_count = 0; total_size = 0
            total = len(self.emails)
            for i, em in enumerate(self.emails):
                if self._stop: break
                if not em.local_path or not Path(em.local_path).exists(): continue
                try:
                    with open(em.local_path, 'rb') as f:
                        msg = email.message_from_bytes(f.read())
                    for part in msg.walk():
                        fn = part.get_filename()
                        if not fn: continue
                        fn = decode_header(fn)
                        if not fn or fn.startswith('.'): continue
                        payload = part.get_payload(decode=True)
                        if not payload: continue
                        h = hashlib.sha256(payload).hexdigest()[:16]
                        if h in seen_hashes:
                            continue
                        seen_hashes[h] = fn
                        cat_dir = att_dir / sanitize_filename(em.category or 'Uncategorized', 60)
                        cat_dir.mkdir(parents=True, exist_ok=True)
                        ds = em.date_parsed.strftime("%Y-%m-%d") if em.date_parsed else ""
                        safe_fn = sanitize_filename(fn, 80)
                        out_path = cat_dir / (f"{ds}_{safe_fn}" if ds else safe_fn)
                        c = 1
                        base, ext = os.path.splitext(str(out_path))
                        while out_path.exists():
                            out_path = Path(f"{base}_{c}{ext}"); c += 1
                        with open(out_path, 'wb') as f: f.write(payload)
                        total_count += 1; total_size += len(payload)
                except Exception: pass
                if (i+1) % 100 == 0:
                    self.progress.emit(i+1, total)
                    self.status.emit(f"Scanned {i+1:,}/{total:,} — {total_count} attachments found")
            self.progress.emit(total, total)
            self.finished_signal.emit(total_count, total_size, str(att_dir))
        except Exception as e: self.error.emit(str(e))


class SensitiveScanWorker(QThread):
    progress = pyqtSignal(int, int); status = pyqtSignal(str)
    finished_signal = pyqtSignal(int)  # count of sensitive emails
    error = pyqtSignal(str)
    def __init__(self, emails):
        super().__init__()
        self.emails = emails; self._stop = False
    def stop(self): self._stop = True
    def run(self):
        try:
            count = 0; total = len(self.emails)
            for i, em in enumerate(self.emails):
                if self._stop: break
                if not em.local_path or not Path(em.local_path).exists(): continue
                try:
                    with open(em.local_path, 'r', encoding='utf-8', errors='replace') as f:
                        text = f.read(50000)  # First 50KB
                    flags = scan_sensitive(text)
                    if flags:
                        em.sensitive_flags = flags; count += 1
                except Exception: pass
                if (i+1) % 200 == 0:
                    self.progress.emit(i+1, total)
                    self.status.emit(f"Scanned {i+1:,}/{total:,} — {count} sensitive found")
            self.progress.emit(total, total)
            self.finished_signal.emit(count)
        except Exception as e: self.error.emit(str(e))


class ThreadSummaryWorker(QThread):
    progress = pyqtSignal(int, int); status = pyqtSignal(str)
    result = pyqtSignal(str, str)  # thread_id, summary
    finished_signal = pyqtSignal(); error = pyqtSignal(str)
    def __init__(self, api_key, threads_to_summarize):
        super().__init__()
        self.api_key = api_key
        self.threads = threads_to_summarize  # list of (thread_id, [EmailInfo])
        self._stop = False
    def stop(self): self._stop = True
    def run(self):
        try:
            client = anthropic.Anthropic(api_key=self.api_key)
            total = len(self.threads)
            for i, (tid, emails) in enumerate(self.threads):
                if self._stop: break
                # Build thread text from .eml bodies
                thread_text = []
                for em in emails[:10]:  # Max 10 emails per thread
                    if em.local_path and Path(em.local_path).exists():
                        try:
                            with open(em.local_path, 'rb') as f:
                                msg = email.message_from_bytes(f.read())
                            body = ""
                            if msg.is_multipart():
                                for part in msg.walk():
                                    if part.get_content_type() == 'text/plain':
                                        body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                                        break
                            else:
                                body = msg.get_payload(decode=True).decode('utf-8', errors='replace')
                            thread_text.append(f"From: {em.sender_name}\nDate: {em.date}\nSubject: {em.subject}\n\n{body[:2000]}")
                        except: pass
                    else:
                        thread_text.append(f"From: {em.sender_name}\nDate: {em.date}\nSubject: {em.subject}")
                if not thread_text: continue
                try:
                    resp = client.messages.create(model="claude-haiku-4-5-20251001", max_tokens=300,
                        messages=[{"role":"user","content":
                            f"Summarize this email thread in 2-3 sentences. Include key decisions, action items, and outcome.\n\n{'---'.join(thread_text)}"}])
                    self.result.emit(tid, resp.content[0].text)
                except Exception as e:
                    self.status.emit(f"Thread error: {e}")
                self.progress.emit(i+1, total)
            self.finished_signal.emit()
        except Exception as e: self.error.emit(str(e))


class AiClassifyWorker(QThread):
    progress = pyqtSignal(int, int); status = pyqtSignal(str)
    classified = pyqtSignal(dict); finished_signal = pyqtSignal(); error = pyqtSignal(str)
    def __init__(self, key, emails, existing):
        super().__init__()
        self.key, self.emails, self.existing = key, emails, existing
        self._stop = False
    def stop(self): self._stop = True
    def run(self):
        try:
            client = anthropic.Anthropic(api_key=self.key)
            dg = defaultdict(list)
            for em in self.emails: dg[em.sender_domain].append(em)
            domains = list(dg.keys()); total = len(domains)
            self.status.emit(f"Classifying {total} domains...")
            for i in range(0, total, 30):
                if self._stop: break
                batch = domains[i:i+30]
                info = [{'domain': d, 'count': len(dg[d]),
                        'senders': list(set(e.sender_name for e in dg[d]))[:3],
                        'subjects': [e.subject for e in dg[d][:5]]} for d in batch]
                try:
                    r = client.messages.create(model="claude-haiku-4-5-20251001", max_tokens=1024,
                        messages=[{"role":"user","content":
                            f"Categorize domains. Existing: {', '.join(self.existing)}\n"
                            f"Reply ONLY JSON: {{\"domain\": \"Category\"}}\n\n{json.dumps(info,indent=2)}"}])
                    m = re.search(r'\{[^{}]*\}', r.content[0].text, re.DOTALL)
                    if m: self.classified.emit(json.loads(m.group()))
                except Exception as e: self.status.emit(f"Batch error: {e}")
                self.progress.emit(min(i+30, total), total)
            self.finished_signal.emit()
        except Exception as e: self.error.emit(str(e))


class OllamaClassifyWorker(QThread):
    progress = pyqtSignal(int, int); status = pyqtSignal(str)
    classified = pyqtSignal(dict); finished_signal = pyqtSignal(); error = pyqtSignal(str)

    def __init__(self, emails, existing, model='llama3.2', endpoint='http://127.0.0.1:11434'):
        super().__init__()
        self.emails, self.existing = emails, existing
        self.model, self.endpoint = model, endpoint
        self._stop = False

    def stop(self): self._stop = True

    def run(self):
        try:
            grouped = defaultdict(list)
            for em in self.emails:
                grouped[em.sender_domain].append(em)
            info = [{'domain': domain, 'count': len(messages),
                     'senders': list({em.sender_name for em in messages})[:3],
                     'subjects': [em.subject for em in messages[:5]]}
                    for domain, messages in grouped.items()]
            result = OllamaClassifier(self.model, self.endpoint).classify_domains(info, self.existing)
            self.classified.emit(result)
            self.progress.emit(len(info), len(info))
            self.finished_signal.emit()
        except Exception as exc:
            self.error.emit(str(exc))


class ReceiptVisionWorker(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    finished_signal = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, emails, classifier):
        super().__init__()
        self.emails, self.classifier = emails, classifier

    def run(self):
        try:
            self.status.emit('Classifying receipt and invoice attachments...')
            receipts = extract_receipt_attachments(self.emails, self.classifier)
            self.progress.emit(len(self.emails), len(self.emails))
            self.finished_signal.emit(receipts)
        except Exception as exc:
            self.error.emit(str(exc))


class HtmlArchiveWorker(QThread):
    """Generate a static HTML archive from organized emails."""
    progress = pyqtSignal(int, int); log = pyqtSignal(str)
    finished_signal = pyqtSignal(str); error = pyqtSignal(str)
    def __init__(self, engine, out_dir):
        super().__init__()
        self.engine, self.out_dir = engine, Path(out_dir)
        self._stop = False
    def stop(self): self._stop = True
    def run(self):
        try:
            hdir = self.out_dir / "html_archive"; hdir.mkdir(parents=True, exist_ok=True)
            cats = sorted(self.engine.categories.items(), key=lambda x: -len(x[1]))
            total = len(self.engine.emails); done = 0

            # Index page
            idx_cats = ''.join(f'<li><a href="{sanitize_filename(c,80)}.html">{c}</a> ({len(e):,})</li>' for c,e in cats)
            s = self.engine.get_summary()
            idx_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>GmailDownloader Archive</title>
<style>body{{background:#1e1e2e;color:#cdd6f4;font-family:'Segoe UI',sans-serif;max-width:1200px;margin:0 auto;padding:20px}}
a{{color:#89b4fa;text-decoration:none}}a:hover{{text-decoration:underline}}
h1{{color:#89b4fa}}h2{{color:#cba6f7}}
.stats{{display:flex;gap:16px;flex-wrap:wrap;margin:16px 0}}
.stat{{background:#313244;padding:16px;border-radius:8px;text-align:center;min-width:120px}}
.stat .val{{font-size:24px;font-weight:bold;color:#89b4fa}}.stat .lbl{{color:#a6adc8;font-size:12px}}
ul{{list-style:none;padding:0}}li{{padding:8px;border-bottom:1px solid #313244}}
</style></head><body>
<h1>GmailDownloader Archive</h1>
<div class="stats">
<div class="stat"><div class="val">{s['total']:,}</div><div class="lbl">Emails</div></div>
<div class="stat"><div class="val">{len(cats)}</div><div class="lbl">Categories</div></div>
<div class="stat"><div class="val">{format_size(s['total_size'])}</div><div class="lbl">Total Size</div></div>
<div class="stat"><div class="val">{s['date_range'][0]}</div><div class="lbl">Oldest</div></div>
<div class="stat"><div class="val">{s['date_range'][1]}</div><div class="lbl">Newest</div></div>
</div><h2>Categories</h2><ul>{idx_cats}</ul></body></html>"""
            with open(hdir / "index.html", 'w', encoding='utf-8') as f: f.write(idx_html)

            # Per-category pages
            for cat_name, emails in cats:
                if self._stop: break
                sorted_emails = sorted(emails, key=lambda e: e.date_parsed or datetime.min, reverse=True)
                rows = []
                for em in sorted_emails:
                    ds = em.date_parsed.strftime("%Y-%m-%d %H:%M") if em.date_parsed else ""
                    subj = em.subject.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                    sender = (em.sender_name or em.sender).replace('&','&amp;').replace('<','&lt;')
                    link = ""
                    if em.local_path and Path(em.local_path).exists():
                        # Create individual email page
                        em_file = sanitize_filename(f"{em.uid.replace(':','_')}",60) + ".html"
                        em_dir = hdir / "emails"; em_dir.mkdir(exist_ok=True)
                        try:
                            with open(em.local_path, 'rb') as ef:
                                msg = email.message_from_bytes(ef.read())
                            body = ""
                            if msg.is_multipart():
                                for part in msg.walk():
                                    ct = part.get_content_type()
                                    if ct == 'text/html':
                                        body = part.get_payload(decode=True).decode('utf-8',errors='replace'); break
                                    elif ct == 'text/plain' and not body:
                                        body = '<pre>' + part.get_payload(decode=True).decode('utf-8',errors='replace').replace('&','&amp;').replace('<','&lt;') + '</pre>'
                            else:
                                payload = msg.get_payload(decode=True)
                                if payload:
                                    if msg.get_content_type() == 'text/html':
                                        body = payload.decode('utf-8',errors='replace')
                                    else:
                                        body = '<pre>' + payload.decode('utf-8',errors='replace').replace('&','&amp;').replace('<','&lt;') + '</pre>'
                            em_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>{subj}</title>
<style>body{{background:#1e1e2e;color:#cdd6f4;font-family:'Segoe UI',sans-serif;max-width:900px;margin:0 auto;padding:20px}}
a{{color:#89b4fa}}.meta{{background:#313244;padding:12px;border-radius:8px;margin-bottom:16px}}
.meta span{{color:#a6adc8}}.body-content{{background:#181825;padding:16px;border-radius:8px;overflow:auto}}
pre{{white-space:pre-wrap;word-break:break-word}}</style></head><body>
<p><a href="../{sanitize_filename(cat_name,80)}.html">Back to {cat_name}</a></p>
<div class="meta"><b>{subj}</b><br><span>From:</span> {sender}<br><span>Date:</span> {ds}<br><span>Folder:</span> {em.source_folder}</div>
<div class="body-content">{body}</div></body></html>"""
                            with open(em_dir / em_file, 'w', encoding='utf-8') as ef: ef.write(em_html)
                            link = f'<a href="emails/{em_file}">view</a>'
                        except Exception: pass
                    rows.append(f'<tr><td>{ds}</td><td>{sender}</td><td>{subj}</td><td>{link}</td></tr>')
                    done += 1

                cat_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>{cat_name}</title>
<style>body{{background:#1e1e2e;color:#cdd6f4;font-family:'Segoe UI',sans-serif;max-width:1200px;margin:0 auto;padding:20px}}
a{{color:#89b4fa;text-decoration:none}}h1{{color:#cba6f7}}
table{{width:100%;border-collapse:collapse}}th,td{{padding:8px;text-align:left;border-bottom:1px solid #313244}}
th{{background:#313244;color:#bac2de;position:sticky;top:0}}</style></head><body>
<p><a href="index.html">Back to Index</a></p>
<h1>{cat_name} ({len(emails):,})</h1>
<table><tr><th>Date</th><th>From</th><th>Subject</th><th></th></tr>{''.join(rows)}</table></body></html>"""
                with open(hdir / f"{sanitize_filename(cat_name,80)}.html", 'w', encoding='utf-8') as f: f.write(cat_html)
                self.progress.emit(done, total)
                self.log.emit(f"  {cat_name}: {len(emails):,} emails")

            self.finished_signal.emit(str(hdir))
        except Exception as e: self.error.emit(str(e))


class ConnectionTester(QObject):
    success = pyqtSignal(int); error = pyqtSignal(str)
    def __init__(self, host, addr, pw, port=993, use_ssl=True, auth_mode='password', access_token='', backend='imap'):
        super().__init__()
        self.host, self.addr, self.pw = host, addr, pw
        self.port, self.use_ssl = port, use_ssl
        self.auth_mode, self.access_token = auth_mode, access_token
        self.backend = backend
    def run(self):
        try:
            if self.backend == 'gmail_api':
                labels = GmailApiSource(self.access_token, self.addr).list_labels()
                self.success.emit(len(labels))
                return
            if self.backend == 'graph':
                count = sum(1 for _ in GraphMailSource(self.access_token).iter_message_refs())
                self.success.emit(count)
                return
            imap = open_imap_connection(
                self.host, self.port, self.use_ssl, self.addr, self.pw,
                self.auth_mode, self.access_token
            )
            _, d = imap.select('INBOX', readonly=True); c = int(d[0])
            imap.close(); imap.logout(); self.success.emit(c)
        except Exception as e: self.error.emit(str(e))


# ─── Chart Widgets ────────────────────────────────────────────────────────

class HBarChart(QWidget):
    """Horizontal bar chart widget."""
    def __init__(self, data=None, title="", parent=None):
        super().__init__(parent)
        self.data = data or []  # [(label, value)]
        self.title = title
        self.setMinimumHeight(max(60, len(self.data) * 28 + 30))

    def set_data(self, data, title=""):
        self.data = data
        if title: self.title = title
        self.setMinimumHeight(max(60, len(data) * 28 + 30))
        self.update()

    def paintEvent(self, e):
        if not self.data: return
        p = QPainter(self); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        w, h = self.width(), self.height()
        max_val = max(v for _, v in self.data) if self.data else 1
        y_off = 24 if self.title else 4
        if self.title:
            p.setPen(QColor(C.TEXT)); p.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
            p.drawText(4, 16, self.title)
        label_w = min(180, w // 3)
        bar_area = w - label_w - 80
        p.setFont(QFont("Segoe UI", 10))
        for i, (label, val) in enumerate(self.data):
            y = y_off + i * 28
            p.setPen(QColor(C.SUBTEXT0))
            p.drawText(QRect(4, y, label_w - 8, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                       label[:25] + ('...' if len(label) > 25 else ''))
            bw = int(bar_area * val / max_val) if max_val else 0
            color = QColor(CHART_COLORS[i % len(CHART_COLORS)])
            p.setBrush(color); p.setPen(Qt.PenStyle.NoPen)
            p.drawRoundedRect(label_w, y + 3, bw, 18, 4, 4)
            p.setPen(QColor(C.TEXT))
            p.drawText(label_w + bw + 6, y, 70, 24, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                       f"{val:,}")
        p.end()


class ActivityHeatmap(QWidget):
    """GitHub-style activity heatmap: 7 days x 24 hours."""
    def __init__(self, data=None, parent=None):
        super().__init__(parent)
        self.data = data or {}  # {(dow, hour): count}
        self.setMinimumHeight(210); self.setMinimumWidth(500)

    def set_data(self, data):
        self.data = data; self.update()

    def paintEvent(self, e):
        p = QPainter(self); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        max_val = max(self.data.values()) if self.data else 1
        lbl_w, top = 40, 30
        cw = min(20, (self.width() - lbl_w - 10) // 24)
        ch = min(22, (self.height() - top - 10) // 7)

        p.setPen(QColor(C.SUBTEXT0)); p.setFont(QFont("Segoe UI", 9))
        for h in range(24):
            if h % 3 == 0:
                p.drawText(lbl_w + h * cw, top - 4, f"{h:02d}")
        for d, name in enumerate(days):
            p.drawText(2, top + d * ch + ch - 4, name)

        for d in range(7):
            for h in range(24):
                val = self.data.get((d, h), 0)
                intensity = val / max_val if max_val else 0
                if intensity == 0:
                    color = QColor(C.SURFACE0)
                elif intensity < 0.25:
                    color = QColor(C.SURFACE1)
                elif intensity < 0.5:
                    color = QColor("#45a86c")
                elif intensity < 0.75:
                    color = QColor("#3dbd6a")
                else:
                    color = QColor(C.GREEN)
                p.setBrush(color); p.setPen(Qt.PenStyle.NoPen)
                p.drawRoundedRect(lbl_w + h * cw + 1, top + d * ch + 1, cw - 2, ch - 2, 3, 3)
        p.end()


# ─── Stats Dialog ─────────────────────────────────────────────────────────

class StatsDialog(QDialog):
    def __init__(self, engine, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Mailbox Statistics")
        self.setMinimumSize(900, 700); self.resize(1000, 750)
        stats = engine.get_stats()
        summary = engine.get_summary()

        layout = QVBoxLayout(self)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        content = QWidget(); cl = QVBoxLayout(content)

        # Summary cards
        cards = QHBoxLayout()
        for label, val, color in [
            ("Total Emails", f"{summary['total']:,}", C.BLUE),
            ("Categories", f"{len(summary['categories'])}", C.GREEN),
            ("Newsletters", f"{summary['newsletter_count']}", C.PEACH),
            ("Threads", f"{summary['thread_count']:,}", C.MAUVE),
            ("Storage", format_size(summary['total_size']), C.TEAL),
            ("Sensitive", f"{summary['sensitive_count']}", C.RED),
        ]:
            card = QLabel(f"<div style='text-align:center'><span style='font-size:24px;color:{color};font-weight:bold'>{val}</span><br><span style='color:{C.SUBTEXT0}'>{label}</span></div>")
            card.setStyleSheet(f"background:{C.SURFACE0}; border-radius:8px; padding:12px;")
            card.setAlignment(Qt.AlignmentFlag.AlignCenter)
            cards.addWidget(card)
        cl.addLayout(cards)

        # Emails per month
        monthly = stats['monthly']
        if monthly:
            chart = HBarChart([(k, v) for k, v in list(monthly.items())[-24:]], "Emails Per Month (last 24)")
            cl.addWidget(chart)

        # Activity heatmap
        hm_data = {}
        for k, v in stats['heatmap'].items():
            d, h = k.split(',')
            hm_data[(int(d), int(h))] = v
        if hm_data:
            lbl = QLabel("Email Activity (Day x Hour)")
            lbl.setStyleSheet(f"font-size:14px; font-weight:bold; color:{C.TEXT}; margin-top:8px;")
            cl.addWidget(lbl)
            heatmap = ActivityHeatmap(hm_data)
            cl.addWidget(heatmap)

        # Top senders + domains side by side
        row = QHBoxLayout()
        if stats['top_senders']:
            row.addWidget(HBarChart(stats['top_senders'][:15], "Top Senders"))
        if stats['top_domains']:
            row.addWidget(HBarChart(stats['top_domains'][:15], "Top Domains"))
        cl.addLayout(row)

        # Category distribution
        if stats['category_counts']:
            cl.addWidget(HBarChart(
                [(k, v) for k, v in sorted(stats['category_counts'].items(), key=lambda x: -x[1])[:15]],
                "Category Distribution"))

        # Storage by category
        if stats['category_sizes']:
            sized = [(k, v) for k, v in sorted(stats['category_sizes'].items(), key=lambda x: -x[1]) if v > 0][:15]
            if sized:
                cl.addWidget(HBarChart(sized, "Storage by Category (bytes)"))

        # Large email finder
        large_emails = sorted(engine.emails, key=lambda e: -e.size_bytes)[:20]
        if large_emails and large_emails[0].size_bytes > 0:
            lbl = QLabel("Largest Emails")
            lbl.setStyleSheet(f"font-size:14px; font-weight:bold; color:{C.TEXT}; margin-top:8px;")
            cl.addWidget(lbl)
            large_table = QTableWidget()
            large_table.setColumnCount(4)
            large_table.setHorizontalHeaderLabels(["Size", "From", "Subject", "Date"])
            lh = large_table.horizontalHeader()
            lh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            lh.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
            lh.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
            lh.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
            large_table.setColumnWidth(1, 150)
            large_table.verticalHeader().setVisible(False)
            large_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            large_table.setMaximumHeight(min(400, len(large_emails) * 28 + 30))
            large_table.setRowCount(len(large_emails))
            for r, em in enumerate(large_emails):
                large_table.setItem(r, 0, QTableWidgetItem(format_size(em.size_bytes)))
                large_table.setItem(r, 1, QTableWidgetItem(em.sender_name or em.sender))
                large_table.setItem(r, 2, QTableWidgetItem(em.subject))
                large_table.setItem(r, 3, QTableWidgetItem(
                    em.date_parsed.strftime("%Y-%m-%d") if em.date_parsed else ""))
            cl.addWidget(large_table)

        # Storage quota estimate (Gmail 15GB free tier)
        if summary['total_size'] > 0:
            quota_gb = 15
            used_gb = summary['total_size'] / (1024**3)
            pct = min(100, int(used_gb / quota_gb * 100))
            color = C.GREEN if pct < 60 else C.YELLOW if pct < 85 else C.RED
            lbl = QLabel(f"Gmail Storage Estimate: {used_gb:.1f} GB / {quota_gb} GB ({pct}%)")
            lbl.setStyleSheet(f"font-size:14px; font-weight:bold; color:{color}; margin-top:8px;")
            cl.addWidget(lbl)
            quota_bar = QProgressBar()
            quota_bar.setMaximum(100); quota_bar.setValue(pct); quota_bar.setFixedHeight(12)
            quota_bar.setStyleSheet(f"QProgressBar::chunk {{ background-color: {color}; border-radius: 4px; }}")
            cl.addWidget(quota_bar)

        health = engine.sender_health()[:15]
        if health:
            cl.addWidget(HBarChart([(item['sender'], item['score']) for item in health], "Sender Health Score"))
        latency = engine.reply_latency()
        if latency:
            cl.addWidget(HBarChart(list(latency.items()), "Reply Latency"))
        forecast = engine.storage_forecast(12)
        if forecast['forecast']:
            projected = forecast['forecast'][-1]['projected_bytes']
            forecast_label = QLabel(
                f"Storage forecast: {format_size(projected)} in 12 months "
                f"({format_size(forecast['average_monthly_bytes'])}/month average)"
            )
            forecast_label.setStyleSheet(f"color:{C.SUBTEXT0};font-size:14px;margin-top:8px;")
            cl.addWidget(forecast_label)

        cl.addStretch()
        scroll.setWidget(content)
        layout.addWidget(scroll)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)


# ─── Subscription Dialog ─────────────────────────────────────────────────

class SubscriptionDialog(QDialog):
    def __init__(self, subscriptions, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Subscriptions & Newsletters ({len(subscriptions)})")
        self.setMinimumSize(800, 500)
        self.subs = subscriptions

        layout = QVBoxLayout(self)
        info = QLabel(f"Detected {len(subscriptions)} newsletter/subscription senders")
        info.setStyleSheet(f"color:{C.SUBTEXT0};")
        layout.addWidget(info)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Sender", "Domain", "Count", "Frequency", "Last Seen", "Unsubscribe"])
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        h.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setRowCount(len(subscriptions))

        for row, si in enumerate(subscriptions):
            self.table.setItem(row, 0, QTableWidgetItem(si.sender_name))
            self.table.setItem(row, 1, QTableWidgetItem(si.domain))
            self.table.setItem(row, 2, QTableWidgetItem(f"{si.count:,}"))
            freq_item = QTableWidgetItem(si.frequency)
            if si.frequency == 'daily': freq_item.setForeground(QColor(C.RED))
            elif si.frequency == 'weekly': freq_item.setForeground(QColor(C.YELLOW))
            self.table.setItem(row, 3, freq_item)
            self.table.setItem(row, 4, QTableWidgetItem(
                si.last_seen.strftime("%Y-%m-%d") if si.last_seen else ""))
            if si.unsubscribe_url:
                btn = QPushButton("Unsubscribe")
                btn.setProperty("danger", True)
                btn.setStyleSheet("padding:4px 8px; font-size:11px;")
                btn.setCursor(Qt.CursorShape.PointingHandCursor)
                btn.clicked.connect(lambda _, url=si.unsubscribe_url: webbrowser.open(url))
                self.table.setCellWidget(row, 5, btn)
            else:
                self.table.setItem(row, 5, QTableWidgetItem("N/A"))

        layout.addWidget(self.table)
        close = QPushButton("Close"); close.clicked.connect(self.close)
        layout.addWidget(close)


# ─── Rules Editor Dialog ─────────────────────────────────────────────────

class RulesEditorDialog(QDialog):
    def __init__(self, rules_engine, categories, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Auto Clean Rules")
        self.setMinimumSize(700, 500)
        self.engine = rules_engine
        self.categories = categories

        layout = QVBoxLayout(self)
        info = QLabel("Rules are applied in order during categorization. First match wins.")
        info.setStyleSheet(f"color:{C.SUBTEXT0};")
        layout.addWidget(info)

        self.list_w = QListWidget()
        self._refresh_list()
        layout.addWidget(self.list_w)

        btns = QHBoxLayout()
        add_btn = QPushButton("Add Rule"); add_btn.clicked.connect(self._add)
        edit_btn = QPushButton("Edit"); edit_btn.setProperty("secondary", True); edit_btn.clicked.connect(self._edit)
        del_btn = QPushButton("Delete"); del_btn.setProperty("danger", True); del_btn.clicked.connect(self._delete)
        import_btn = QPushButton("Import Gmail Filters"); import_btn.setProperty("secondary", True)
        import_btn.clicked.connect(self._import_gmail)
        btns.addWidget(add_btn); btns.addWidget(edit_btn); btns.addWidget(del_btn)
        btns.addStretch(); btns.addWidget(import_btn)
        layout.addLayout(btns)

        close = QPushButton("Close"); close.clicked.connect(self.close)
        layout.addWidget(close)

    def _refresh_list(self):
        self.list_w.clear()
        for r in self.engine.rules:
            conds = ', '.join(f"{k}={v}" for k, v in r.conditions.items())
            status = "" if r.enabled else " [DISABLED]"
            self.list_w.addItem(f"{r.name}: IF {conds} THEN {r.action}={r.action_value}{status}")

    def _add(self):
        d = RuleEditDialog(self.categories, parent=self)
        if d.exec() == QDialog.DialogCode.Accepted:
            self.engine.add_rule(d.get_rule())
            self._refresh_list()

    def _edit(self):
        idx = self.list_w.currentRow()
        if idx < 0: return
        d = RuleEditDialog(self.categories, self.engine.rules[idx], self)
        if d.exec() == QDialog.DialogCode.Accepted:
            self.engine.rules[idx] = d.get_rule()
            self.engine.save()
            self._refresh_list()

    def _delete(self):
        idx = self.list_w.currentRow()
        if idx >= 0:
            self.engine.remove_rule(idx)
            self._refresh_list()

    def _import_gmail(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import Gmail Filters", "", "XML (*.xml)")
        if not path: return
        rules = CleanRulesEngine.import_gmail_filters(path)
        for r in rules:
            self.engine.add_rule(r)
        self._refresh_list()
        QMessageBox.information(self, "Imported", f"Imported {len(rules)} rules from Gmail filters")


class RuleEditDialog(QDialog):
    def __init__(self, categories, rule=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Rule" if rule else "Add Rule")
        self.setMinimumWidth(450)
        layout = QVBoxLayout(self)

        fl = QFormLayout()
        self.name_input = QLineEdit(rule.name if rule else "")
        fl.addRow("Name:", self.name_input)

        self.domain_input = QLineEdit(rule.conditions.get('domain', '') if rule else "")
        self.domain_input.setPlaceholderText("e.g. example.com")
        fl.addRow("Domain:", self.domain_input)

        self.sender_input = QLineEdit(rule.conditions.get('sender', '') if rule else "")
        self.sender_input.setPlaceholderText("e.g. newsletter@")
        fl.addRow("Sender contains:", self.sender_input)

        self.subject_input = QLineEdit(rule.conditions.get('subject_contains', '') if rule else "")
        fl.addRow("Subject contains:", self.subject_input)

        self.age_input = QSpinBox(); self.age_input.setRange(0, 9999)
        self.age_input.setValue(int(rule.conditions.get('older_than_days', 0)) if rule else 0)
        self.age_input.setSpecialValueText("Any age")
        fl.addRow("Older than (days):", self.age_input)

        self.newsletter_check = QCheckBox("Is newsletter")
        self.newsletter_check.setChecked(rule.conditions.get('is_newsletter', False) if rule else False)
        fl.addRow("", self.newsletter_check)

        self.action_combo = QComboBox()
        self.action_combo.addItems(["categorize", "flag", "skip"])
        if rule: self.action_combo.setCurrentText(rule.action)
        fl.addRow("Action:", self.action_combo)

        self.value_combo = QComboBox(); self.value_combo.setEditable(True)
        self.value_combo.addItems(sorted(categories))
        if rule: self.value_combo.setCurrentText(rule.action_value)
        fl.addRow("Category/Value:", self.value_combo)

        self.enabled_check = QCheckBox("Enabled")
        self.enabled_check.setChecked(rule.enabled if rule else True)
        fl.addRow("", self.enabled_check)

        layout.addLayout(fl)
        bb = QHBoxLayout()
        ok = QPushButton("Save"); ok.clicked.connect(self.accept)
        cancel = QPushButton("Cancel"); cancel.setProperty("secondary", True); cancel.clicked.connect(self.reject)
        bb.addStretch(); bb.addWidget(cancel); bb.addWidget(ok)
        layout.addLayout(bb)

    def get_rule(self):
        conds = {}
        if self.domain_input.text().strip(): conds['domain'] = self.domain_input.text().strip()
        if self.sender_input.text().strip(): conds['sender'] = self.sender_input.text().strip()
        if self.subject_input.text().strip(): conds['subject_contains'] = self.subject_input.text().strip()
        if self.age_input.value() > 0: conds['older_than_days'] = self.age_input.value()
        if self.newsletter_check.isChecked(): conds['is_newsletter'] = True
        return CleanRule(name=self.name_input.text().strip() or "Unnamed",
            conditions=conds, action=self.action_combo.currentText(),
            action_value=self.value_combo.currentText(), enabled=self.enabled_check.isChecked())


# ─── Contact Analysis Dialog ─────────────────────────────────────────────

class ContactDialog(QDialog):
    def __init__(self, engine, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Contact Analysis")
        self.setMinimumSize(850, 550)

        layout = QVBoxLayout(self)
        # Build contact data
        contacts = defaultdict(lambda: {'count': 0, 'sent': 0, 'received': 0,
                                        'first': None, 'last': None, 'domains': set()})
        user_domain = engine.user_domain
        for em in engine.emails:
            key = em.sender_name or em.sender
            c = contacts[key]
            c['count'] += 1
            c['domains'].add(em.sender_domain)
            if em.source_folder and 'sent' in em.source_folder.lower():
                c['sent'] += 1
            else:
                c['received'] += 1
            if em.date_parsed:
                if not c['first'] or em.date_parsed < c['first']: c['first'] = em.date_parsed
                if not c['last'] or em.date_parsed > c['last']: c['last'] = em.date_parsed

        sorted_contacts = sorted(contacts.items(), key=lambda x: -x[1]['count'])

        info = QLabel(f"{len(contacts):,} unique contacts")
        info.setStyleSheet(f"color:{C.SUBTEXT0};"); layout.addWidget(info)

        # Filter
        frow = QHBoxLayout(); frow.addWidget(QLabel("Search:"))
        self.filter_input = QLineEdit(); self.filter_input.setPlaceholderText("Filter contacts...")
        self.filter_input.textChanged.connect(lambda: self._filter(sorted_contacts))
        frow.addWidget(self.filter_input, 1); layout.addLayout(frow)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Contact", "Total", "Received", "Sent", "First Seen", "Last Seen"])
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for c in (1,2,3,4,5): h.setSectionResizeMode(c, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)
        layout.addWidget(self.table)

        self._populate(sorted_contacts)

        close = QPushButton("Close"); close.clicked.connect(self.close)
        layout.addWidget(close)
        self._all_contacts = sorted_contacts

    def _populate(self, contacts):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(min(len(contacts), 500))
        for row, (name, c) in enumerate(contacts[:500]):
            self.table.setItem(row, 0, QTableWidgetItem(name))
            ti = QTableWidgetItem(); ti.setData(Qt.ItemDataRole.DisplayRole, c['count'])
            self.table.setItem(row, 1, ti)
            ri = QTableWidgetItem(); ri.setData(Qt.ItemDataRole.DisplayRole, c['received'])
            self.table.setItem(row, 2, ri)
            si = QTableWidgetItem(); si.setData(Qt.ItemDataRole.DisplayRole, c['sent'])
            self.table.setItem(row, 3, si)
            self.table.setItem(row, 4, QTableWidgetItem(c['first'].strftime("%Y-%m-%d") if c['first'] else ""))
            self.table.setItem(row, 5, QTableWidgetItem(c['last'].strftime("%Y-%m-%d") if c['last'] else ""))
            # Color dormant contacts
            if c['last'] and (datetime.now() - c['last']).days > 365:
                for col in range(6):
                    item = self.table.item(row, col)
                    if item: item.setForeground(QColor(C.OVERLAY0))
        self.table.setSortingEnabled(True)

    def _filter(self, all_contacts):
        q = self.filter_input.text().lower()
        filtered = [(n, c) for n, c in all_contacts if q in n.lower()] if q else all_contacts
        self._populate(filtered)


# ─── UI: Connect Page ────────────────────────────────────────────────────

class ConnectPage(QWidget):
    connected = pyqtSignal(str)
    def __init__(self):
        super().__init__()
        self.imap_host = "imap.gmail.com"
        self.imap_port = 993
        self.use_ssl = True
        self.auth_mode = "oauth2"
        self.access_token = ""
        self.backend = "imap"
        self.email_addr = self.password = self.download_dir = ""
        self.loaded_engine = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.setSpacing(16)
        title = QLabel("GmailDownloader")
        title.setStyleSheet(f"font-size:32px; color:{C.BLUE}; font-weight:bold;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(title)
        sub = QLabel(f"v{VERSION} — Full Gmail Mailbox Downloader, AI Organizer & Analytics")
        sub.setStyleSheet(f"color:{C.SUBTEXT0}; font-size:12px;")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(sub)
        layout.addSpacing(20)

        form = QWidget(); form.setMaximumWidth(520); fl = QVBoxLayout(form); fl.setSpacing(12)
        backend_row = QHBoxLayout()
        backend_row.addWidget(QLabel("Backend"))
        self.backend_combo = QComboBox()
        self.backend_combo.addItem("IMAP (generic/Gmail)", "imap")
        self.backend_combo.addItem("Gmail API", "gmail_api")
        self.backend_combo.addItem("Microsoft Graph", "graph")
        self.backend_combo.currentIndexChanged.connect(self._backend_changed)
        backend_row.addWidget(self.backend_combo, 1)
        fl.addLayout(backend_row)
        server_row = QHBoxLayout()
        server_row.addWidget(QLabel("Mail server"))
        self.host_input = QLineEdit(self.imap_host); self.host_input.setPlaceholderText("imap.gmail.com")
        server_row.addWidget(self.host_input, 1)
        self.port_input = QSpinBox(); self.port_input.setRange(1, 65535); self.port_input.setValue(993)
        self.port_input.setMaximumWidth(90); server_row.addWidget(self.port_input)
        self.ssl_check = QCheckBox("SSL"); self.ssl_check.setChecked(True); server_row.addWidget(self.ssl_check)
        fl.addLayout(server_row)
        fl.addWidget(QLabel("Gmail Address"))
        self.email_input = QLineEdit(); self.email_input.setPlaceholderText("you@gmail.com")
        fl.addWidget(self.email_input)
        auth_row = QHBoxLayout()
        auth_row.addWidget(QLabel("Authentication"))
        self.auth_combo = QComboBox()
        self.auth_combo.addItems(["OAuth2 access token (recommended)", "App Password"])
        self.auth_combo.currentIndexChanged.connect(self._auth_changed)
        auth_row.addWidget(self.auth_combo, 1)
        fl.addLayout(auth_row)
        self.credential_label = QLabel("OAuth2 access token")
        fl.addWidget(self.credential_label)
        self.pass_input = QLineEdit(); self.pass_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.pass_input.setPlaceholderText("Paste an access token"); fl.addWidget(self.pass_input)
        hint = QLabel("OAuth2 is preferred. Use an App Password only when OAuth2 is unavailable.")
        hint.setStyleSheet(f"color:{C.SUBTEXT0}; font-size:11px;"); hint.setWordWrap(True)
        fl.addWidget(hint); fl.addSpacing(12)

        sync_group = QGroupBox("Sync options")
        sync_form = QFormLayout(sync_group)
        self.incremental_check = QCheckBox("Resume incrementally (new UIDs only)")
        self.incremental_check.setChecked(True)
        sync_form.addRow(self.incremental_check)
        since_row = QHBoxLayout()
        self.since_check = QCheckBox("Only messages since")
        self.since_date_edit = QDateEdit(QDate.currentDate().addYears(-1))
        self.since_date_edit.setCalendarPopup(True)
        self.since_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.since_date_edit.setEnabled(False)
        self.since_check.toggled.connect(self.since_date_edit.setEnabled)
        since_row.addWidget(self.since_check); since_row.addWidget(self.since_date_edit)
        sync_form.addRow(since_row)
        self.verify_check = QCheckBox("Verify downloaded files before re-use")
        self.verify_check.setChecked(True)
        sync_form.addRow(self.verify_check)
        self.attachments_only_check = QCheckBox("Attachments only (do not retain .eml bodies)")
        sync_form.addRow(self.attachments_only_check)
        fl.addWidget(sync_group)

        r1 = QHBoxLayout()
        self.dl_btn = QPushButton("Download Full Mailbox")
        self.dl_btn.setStyleSheet(f"background-color:{C.GREEN};")
        self.dl_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.dl_btn.clicked.connect(lambda: self._on_action("download")); r1.addWidget(self.dl_btn)
        self.scan_btn = QPushButton("Scan Inbox Headers Only")
        self.scan_btn.setProperty("secondary", True)
        self.scan_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.scan_btn.clicked.connect(lambda: self._on_action("scan")); r1.addWidget(self.scan_btn)
        fl.addLayout(r1)

        r2 = QHBoxLayout()
        self.load_scan_btn = QPushButton("Load Previous Scan")
        self.load_scan_btn.setProperty("secondary", True)
        self.load_scan_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.load_scan_btn.clicked.connect(self._on_load_scan); r2.addWidget(self.load_scan_btn)
        self.load_local_btn = QPushButton("Load Downloaded Mailbox")
        self.load_local_btn.setProperty("secondary", True)
        self.load_local_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.load_local_btn.clicked.connect(self._on_load_local); r2.addWidget(self.load_local_btn)
        fl.addLayout(r2)

        r3 = QHBoxLayout()
        self.load_mbox_btn = QPushButton("Import mbox")
        self.load_mbox_btn.setProperty("secondary", True)
        self.load_mbox_btn.clicked.connect(self._on_load_mbox); r3.addWidget(self.load_mbox_btn)
        self.load_tb_btn = QPushButton("Import Thunderbird")
        self.load_tb_btn.setProperty("secondary", True)
        self.load_tb_btn.clicked.connect(self._on_load_thunderbird); r3.addWidget(self.load_tb_btn)
        fl.addLayout(r3)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet(f"color:{C.SUBTEXT0}; font-size:12px;")
        self.status_label.setWordWrap(True); fl.addWidget(self.status_label)
        layout.addWidget(form, alignment=Qt.AlignmentFlag.AlignCenter); layout.addStretch()

    def _set_btns(self, en):
        for b in (self.scan_btn, self.dl_btn, self.load_scan_btn, self.load_local_btn,
                  self.load_mbox_btn, self.load_tb_btn): b.setEnabled(en)

    def _auth_changed(self, index):
        oauth = index == 0
        self.credential_label.setText("OAuth2 access token" if oauth else "App Password")
        self.pass_input.setPlaceholderText("Paste an access token" if oauth else "16-character app password")

    def _backend_changed(self, index):
        self.backend = self.backend_combo.itemData(index)
        is_imap = self.backend == 'imap'
        self.host_input.setEnabled(is_imap)
        self.port_input.setEnabled(is_imap)
        self.ssl_check.setEnabled(is_imap)
        if not is_imap:
            self.auth_combo.setCurrentIndex(0)
            self.auth_combo.setEnabled(False)
            self.credential_label.setText("OAuth2 access token")
            self.pass_input.setPlaceholderText("Paste an access token")
        else:
            self.auth_combo.setEnabled(True)
            self._auth_changed(self.auth_combo.currentIndex())

    def _on_action(self, mode):
        self.imap_host = self.host_input.text().strip() or "imap.gmail.com"
        self.imap_port = self.port_input.value()
        self.use_ssl = self.ssl_check.isChecked()
        self.auth_mode = "oauth2" if self.auth_combo.currentIndex() == 0 else "password"
        self.backend = self.backend_combo.currentData()
        self.email_addr = self.email_input.text().strip()
        self.password = self.pass_input.text().strip()
        self.access_token = self.password if self.auth_mode == "oauth2" else ""
        if not self.email_addr or not self.password:
            self.status_label.setText("Enter an email address and credential")
            self.status_label.setStyleSheet(f"color:{C.RED};"); return
        if mode == "download":
            f = QFileDialog.getExistingDirectory(self, "Download Folder",
                str(Path.home() / "Desktop" / "GmailDownloader"))
            if not f: return
            self.download_dir = f
        self.status_label.setText("Testing..."); self.status_label.setStyleSheet(f"color:{C.YELLOW};")
        self._set_btns(False); self._mode = mode
        self._tt = QThread(); self._tw = ConnectionTester(
            self.imap_host, self.email_addr, self.password, self.imap_port,
            self.use_ssl, self.auth_mode, self.access_token, self.backend
        )
        self._tw.moveToThread(self._tt); self._tt.started.connect(self._tw.run)
        self._tw.success.connect(self._ok); self._tw.error.connect(self._fail); self._tt.start()

    def sync_options(self):
        since = None
        if self.since_check.isChecked():
            qdate = self.since_date_edit.date()
            since = datetime(qdate.year(), qdate.month(), qdate.day())
        return SyncOptions(
            since=since,
            incremental=self.incremental_check.isChecked(),
            verify_integrity=self.verify_check.isChecked(),
            attachments_only=self.attachments_only_check.isChecked(),
        )

    def _ok(self, c):
        self._tt.quit(); self.status_label.setText(f"Connected! {c:,} in Inbox")
        self.status_label.setStyleSheet(f"color:{C.GREEN};")
        self.loaded_engine = None; self.connected.emit(self._mode)

    def _fail(self, e):
        self._tt.quit(); self.status_label.setText(e)
        self.status_label.setStyleSheet(f"color:{C.RED};"); self._set_btns(True)

    def _on_load_scan(self):
        p, _ = QFileDialog.getOpenFileName(self, "Load Scan", "", "JSON (*.json)")
        if not p: return
        eng = CategoryEngine()
        if eng.load_state(p):
            self.loaded_engine = eng; self.email_addr = self.email_input.text().strip()
            self.password = self.pass_input.text().strip()
            self.status_label.setText(f"Loaded {len(eng.emails):,} emails")
            self.status_label.setStyleSheet(f"color:{C.GREEN};"); self.connected.emit("load")
        else:
            self.status_label.setText("Failed to load"); self.status_label.setStyleSheet(f"color:{C.RED};")

    def _on_load_local(self):
        f = QFileDialog.getExistingDirectory(self, "Select Download Folder")
        if not f: return
        self.download_dir = f; mp = Path(f) / "manifest.json"
        if not mp.exists():
            self.status_label.setText("No manifest.json found")
            self.status_label.setStyleSheet(f"color:{C.RED};"); return
        self.status_label.setText("Loading..."); self.status_label.setStyleSheet(f"color:{C.YELLOW};")
        QApplication.processEvents()
        try:
            with open(mp, 'r', encoding='utf-8') as fh: manifest = json.load(fh)
            ud = self.email_input.text().strip()
            ud = ud.split('@')[1] if '@' in ud else ""
            eng = CategoryEngine(ud)
            # Load learned rules if available
            lr_path = Path(f) / "learned_rules.json"
            eng.learned = LearnedRules(str(lr_path))
            cr_path = Path(f) / "clean_rules.json"
            eng.clean_rules = CleanRulesEngine(str(cr_path))
            emails, seen = [], set()
            for fn, fd in manifest.get('folders', {}).items():
                for uid, info in fd.items():
                    mid = info.get('message_id', '')
                    if mid and mid in seen: continue
                    if mid: seen.add(mid)
                    emails.append(EmailInfo(uid=f"{fn}:{uid}", sender=info.get('sender',''),
                        sender_name=info.get('sender_name',''), subject=info.get('subject',''),
                        date=info.get('date',''), date_parsed=parse_date(info.get('date','')),
                        has_list_unsubscribe=info.get('has_list_unsubscribe',False),
                        list_unsubscribe_url=info.get('list_unsubscribe_url',''),
                        local_path=info.get('local_path',''), source_folder=fn,
                        message_id=mid, in_reply_to=info.get('in_reply_to',''),
                        references=info.get('references',''), size_bytes=info.get('size_bytes',0)))
            eng.process_all(emails); self.loaded_engine = eng
            self.status_label.setText(f"Loaded {len(emails):,} unique emails from {len(manifest.get('folders',{}))} folders")
            self.status_label.setStyleSheet(f"color:{C.GREEN};"); self.connected.emit("load")
        except Exception as e:
            self.status_label.setText(str(e)); self.status_label.setStyleSheet(f"color:{C.RED};")

    def _load_imported(self, emails, output_dir):
        if not emails:
            self.status_label.setText("No messages found")
            self.status_label.setStyleSheet(f"color:{C.RED};")
            return
        address = self.email_input.text().strip()
        domain = address.split('@', 1)[1] if '@' in address else ""
        engine = CategoryEngine(domain)
        engine.process_all(emails)
        self.download_dir = str(output_dir) if output_dir else ""
        self.loaded_engine = engine
        self.status_label.setText(f"Imported {len(emails):,} messages")
        self.status_label.setStyleSheet(f"color:{C.GREEN};")
        self.connected.emit("load")

    def _on_load_mbox(self):
        source, _ = QFileDialog.getOpenFileName(self, "Import mbox", "", "Mailbox files (*)")
        if not source:
            return
        output_dir = QFileDialog.getExistingDirectory(self, "Archive output folder")
        if not output_dir:
            return
        self.status_label.setText("Importing mbox...")
        QApplication.processEvents()
        try:
            self._load_imported(import_mbox(source, output_dir), output_dir)
        except Exception as exc:
            self.status_label.setText(f"Import failed: {exc}")
            self.status_label.setStyleSheet(f"color:{C.RED};")

    def _on_load_thunderbird(self):
        profile = QFileDialog.getExistingDirectory(self, "Select Thunderbird profile")
        if not profile:
            return
        output_dir = QFileDialog.getExistingDirectory(self, "Archive output folder")
        if not output_dir:
            return
        self.status_label.setText("Importing Thunderbird profile...")
        QApplication.processEvents()
        try:
            self._load_imported(import_thunderbird_profile(profile, output_dir), output_dir)
        except Exception as exc:
            self.status_label.setText(f"Import failed: {exc}")
            self.status_label.setStyleSheet(f"color:{C.RED};")


# ─── UI: Download Page ───────────────────────────────────────────────────

class DownloadPage(QWidget):
    download_complete = pyqtSignal()
    def __init__(self):
        super().__init__(); self.engine = self.worker = None; self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self); layout.setSpacing(12)
        title = QLabel("Downloading Full Mailbox")
        title.setStyleSheet("font-size:20px;font-weight:bold;")
        layout.addWidget(title)
        self.status_label = QLabel("Preparing..."); self.status_label.setStyleSheet(f"color:{C.SUBTEXT0};")
        layout.addWidget(self.status_label)
        self.progress = QProgressBar(); self.progress.setTextVisible(False); self.progress.setFixedHeight(8)
        layout.addWidget(self.progress)
        self.pct = QLabel("0%"); self.pct.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(self.pct)
        self.log = QPlainTextEdit(); self.log.setReadOnly(True)
        self.log.setStyleSheet("font-family:'Cascadia Code','Consolas',monospace;font-size:12px;")
        layout.addWidget(self.log, 1)
        bot = QHBoxLayout()
        self.stop_btn = QPushButton("Stop (Resumable)"); self.stop_btn.setProperty("danger", True)
        self.stop_btn.setCursor(Qt.CursorShape.PointingHandCursor); bot.addWidget(self.stop_btn)
        bot.addStretch()
        self.size_lbl = QLabel(""); self.size_lbl.setStyleSheet(f"color:{C.SUBTEXT0};"); bot.addWidget(self.size_lbl)
        self.cont_btn = QPushButton("Continue to Analysis"); self.cont_btn.setVisible(False)
        self.cont_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.cont_btn.clicked.connect(self.download_complete.emit); bot.addWidget(self.cont_btn)
        layout.addLayout(bot)

    def start_download(self, host, addr, pw, out_dir, options=None, port=993, use_ssl=True,
                       auth_mode='password', access_token=''):
        ud = addr.split('@')[1] if '@' in addr else ""
        self.engine = CategoryEngine(ud); self._out = out_dir; self._n = 0
        self.worker = ImapDownloadWorker(
            host, addr, pw, out_dir, options=options, port=port, use_ssl=use_ssl,
            auth_mode=auth_mode, access_token=access_token
        )
        self._wire_worker()

    def start_remote_download(self, source, out_dir, options=None, query=''):
        self.engine = CategoryEngine(''); self._out = str(out_dir); self._n = 0
        worker_cls = GmailApiDownloadWorker if isinstance(source, GmailApiSource) else RemoteMimeDownloadWorker
        self.worker = worker_cls(source, out_dir, options, query) if worker_cls is GmailApiDownloadWorker \
            else worker_cls(source, out_dir, type(source).__name__, options, query)
        self._wire_worker()

    def _wire_worker(self):
        self.worker.progress.connect(lambda c,t: (self.progress.setMaximum(t), self.progress.setValue(c),
            self.pct.setText(f"{int(c/t*100) if t else 0}% ({c:,}/{t:,})")))
        self.worker.status.connect(self.status_label.setText)
        self.worker.log.connect(self.log.appendPlainText)
        self.worker.email_saved.connect(self._on_saved)
        self.worker.finished_signal.connect(self._done)
        self.worker.error.connect(lambda e: (self.status_label.setText("Error"),
            self.log.appendPlainText(f"ERROR: {e}"), self.cont_btn.setVisible(True)))
        self.stop_btn.clicked.connect(lambda: (self.worker.stop(),
            self.status_label.setText("Stopping..."), self.cont_btn.setVisible(True)))
        self.worker.start()

    def _on_saved(self, em):
        self._n += 1
        if self._n % 200 == 0:
            try:
                total = sum(f.stat().st_size for f in (Path(self._out)/"folders").rglob('*.eml'))
                self.size_lbl.setText(f"{total/1024**2:,.0f} MB")
            except: pass

    def _done(self, all_emails):
        seen, unique = set(), []
        for em in all_emails:
            if em.message_id and em.message_id in seen: continue
            if em.message_id: seen.add(em.message_id)
            unique.append(em)
        # Setup learned rules and clean rules paths
        self.engine.learned = LearnedRules(str(Path(self._out) / "learned_rules.json"))
        self.engine.clean_rules = CleanRulesEngine(str(Path(self._out) / "clean_rules.json"))
        self.engine.process_all(unique)
        self.status_label.setText("Download complete!"); self.status_label.setStyleSheet(f"color:{C.GREEN};")
        self.cont_btn.setVisible(True); self.stop_btn.setEnabled(False)


# ─── UI: Analyze Page ────────────────────────────────────────────────────

class AnalyzePage(QWidget):
    analysis_complete = pyqtSignal()
    def __init__(self):
        super().__init__(); self.engine = self.worker = None; self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self); layout.setSpacing(12)
        t = QLabel("Analyzing Mailbox"); t.setStyleSheet("font-size:20px;font-weight:bold;")
        layout.addWidget(t)
        self.status_label = QLabel("Preparing..."); self.status_label.setStyleSheet(f"color:{C.SUBTEXT0};")
        layout.addWidget(self.status_label)
        self.progress = QProgressBar(); self.progress.setTextVisible(False); self.progress.setFixedHeight(8)
        layout.addWidget(self.progress)
        self.pct = QLabel("0%"); self.pct.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(self.pct)
        g = QGroupBox("Analysis Results"); gl = QVBoxLayout(g)
        self.stats_text = QPlainTextEdit(); self.stats_text.setReadOnly(True)
        self.stats_text.setStyleSheet("font-family:'Cascadia Code','Consolas',monospace;font-size:12px;")
        gl.addWidget(self.stats_text); layout.addWidget(g)
        layout.addStretch()
        bot = QHBoxLayout()
        self.save_btn = QPushButton("Save Analysis"); self.save_btn.setProperty("secondary", True)
        self.save_btn.setVisible(False); self.save_btn.clicked.connect(self._save); bot.addWidget(self.save_btn)
        self.stats_btn = QPushButton("View Statistics"); self.stats_btn.setProperty("secondary", True)
        self.stats_btn.setVisible(False); self.stats_btn.clicked.connect(self._show_stats); bot.addWidget(self.stats_btn)
        bot.addStretch()
        self.cont_btn = QPushButton("Review Categories"); self.cont_btn.setVisible(False)
        self.cont_btn.clicked.connect(self.analysis_complete.emit); bot.addWidget(self.cont_btn)
        layout.addLayout(bot)

    def start_scan(self, host, addr, pw, port=993, use_ssl=True, auth_mode='password',
                   access_token='', since=None):
        ud = addr.split('@')[1] if '@' in addr else ""
        self.engine = CategoryEngine(ud); self._dc = Counter()
        self.worker = ImapScanWorker(
            host, addr, pw, port, use_ssl, auth_mode, access_token, since
        )
        self.worker.progress.connect(lambda c,t: (self.progress.setMaximum(t), self.progress.setValue(c),
            self.pct.setText(f"{int(c/t*100) if t else 0}% ({c:,}/{t:,})")))
        self.worker.status.connect(self.status_label.setText)
        self.worker.email_batch.connect(self._batch)
        self.worker.finished_signal.connect(self._finished)
        self.worker.error.connect(lambda e: (self.status_label.setText(e),
            self.status_label.setStyleSheet(f"color:{C.RED};")))
        self.worker.start()

    def start_remote_scan(self, source, query=''):
        self.engine = CategoryEngine(''); self._dc = Counter()
        self.worker = RemoteMimeScanWorker(source, query)
        self.worker.progress.connect(lambda c,t: (self.progress.setMaximum(t), self.progress.setValue(c),
            self.pct.setText(f"{int(c/t*100) if t else 0}% ({c:,}/{t:,})")))
        self.worker.status.connect(self.status_label.setText)
        self.worker.email_batch.connect(self._batch)
        self.worker.finished_signal.connect(self._finished)
        self.worker.error.connect(lambda e: (self.status_label.setText(e),
            self.status_label.setStyleSheet(f"color:{C.RED};")))
        self.worker.start()

    def set_preloaded(self, eng): self.engine = eng; self._show_summary()

    def _batch(self, batch):
        for em in batch:
            self._dc[self.engine.extract_domain(em.sender)] += 1
        top = self._dc.most_common(15)
        lines = [f"Scanned: {sum(self._dc.values()):,}", ""]
        for d, c in top: lines.append(f"  {d:40s} {c:>6,}")
        self.stats_text.setPlainText('\n'.join(lines))

    def _finished(self, emails):
        self.status_label.setText("Categorizing..."); QApplication.processEvents()
        self.engine.process_all(emails); self._show_summary()

    def _show_summary(self):
        s = self.engine.get_summary()
        lines = [f"Total: {s['total']:,}  |  Categorized: {s['categorized']:,}  |  "
                 f"Uncategorized: {s['uncategorized']:,}",
                 f"Date range: {s['date_range'][0]} to {s['date_range'][1]}",
                 f"Threads: {s['thread_count']:,}  |  Newsletters: {s['newsletter_count']}  |  "
                 f"Storage: {format_size(s['total_size'])}"]
        if s.get('folder_counts'):
            lines += ["", "Folders:"]
            for fn, c in s['folder_counts'].items(): lines.append(f"  {fn:40s} {c:>6,}")
        lines += ["", "Categories:"]
        for cat, c in s['categories'].items(): lines.append(f"  {cat:40s} {c:>6,}")
        self.stats_text.setPlainText('\n'.join(lines))
        self.progress.setMaximum(1); self.progress.setValue(1); self.pct.setText("100%")
        self.status_label.setText("Analysis complete!"); self.status_label.setStyleSheet(f"color:{C.GREEN};")
        self.cont_btn.setVisible(True); self.save_btn.setVisible(True); self.stats_btn.setVisible(True)

    def _save(self):
        p, _ = QFileDialog.getSaveFileName(self, "Save", "gmaildownloader_scan.json", "JSON (*.json)")
        if p and self.engine: self.engine.save_state(p)

    def _show_stats(self):
        if self.engine: StatsDialog(self.engine, self).exec()


# ─── UI: Review Page ─────────────────────────────────────────────────────

class ReviewPage(QWidget):
    execute_requested = pyqtSignal()
    def __init__(self):
        super().__init__(); self.engine = None; self.has_local = False; self._dl_dir = ""
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self); layout.setSpacing(8)
        self._current_emails = []  # Currently displayed emails (for filtering)
        self._all_current_emails = []  # Before filter

        # Toolbar row 1
        tb = QHBoxLayout()
        t = QLabel("Review Categories"); t.setStyleSheet("font-size:20px;font-weight:bold;"); tb.addWidget(t)
        tb.addStretch()
        self.summary_lbl = QLabel(""); self.summary_lbl.setStyleSheet(f"color:{C.SUBTEXT0};"); tb.addWidget(self.summary_lbl)
        layout.addLayout(tb)

        # Toolbar row 2: action buttons
        tb2 = QHBoxLayout()
        for name, slot, tip in [
            ("Stats", self._show_stats, "Mailbox statistics dashboard"),
            ("Contacts", self._show_contacts, "Contact frequency analysis"),
            ("Subscriptions", self._show_subs, "Newsletter management"),
            ("Rules", self._show_rules, "Auto clean rules editor"),
            ("Inbox Zero", self._show_inbox_zero, "Preview archive and unsubscribe suggestions"),
            ("Location Timeline", self._show_location_timeline, "Audit public IP hops from Received headers"),
        ]:
            b = QPushButton(name); b.setProperty("secondary", True)
            b.setToolTip(tip); b.clicked.connect(slot); tb2.addWidget(b)
            if name == "Contacts": self.contacts_btn = b
            if name == "Location Timeline": self.location_btn = b

        # Export menu
        self.export_btn = QPushButton("Export"); self.export_btn.setProperty("secondary", True)
        export_menu = QMenu(self)
        export_menu.addAction("CSV", self._export_csv)
        export_menu.addAction("JSON", self._export_json)
        export_menu.addAction("MBOX", self._export_mbox)
        export_menu.addAction("Markdown Vault", self._export_markdown)
        export_menu.addAction("PDF", self._export_pdf)
        export_menu.addAction("Relationship Graph", self._export_graph)
        export_menu.addAction("Contact Graph", self._export_contact_graph)
        export_menu.addAction("HTML Archive", self._export_html)
        export_menu.addAction("Encrypted Archive", self._export_encrypted)
        self.export_btn.setMenu(export_menu); tb2.addWidget(self.export_btn)

        tb2.addWidget(QLabel("|")); tb2.addWidget(QLabel(""))  # spacer
        self.attach_btn = QPushButton("Attachments"); self.attach_btn.setProperty("secondary", True)
        self.attach_btn.clicked.connect(self._extract_attachments); tb2.addWidget(self.attach_btn)
        self.sensitive_btn = QPushButton("Sensitive"); self.sensitive_btn.setProperty("secondary", True)
        self.sensitive_btn.clicked.connect(self._scan_sensitive); tb2.addWidget(self.sensitive_btn)
        self.redact_btn = QPushButton("Redact Copies"); self.redact_btn.setProperty("secondary", True)
        self.redact_btn.clicked.connect(self._redact_sensitive); tb2.addWidget(self.redact_btn)
        self.ai_btn = QPushButton("AI Classify"); self.ai_btn.setProperty("secondary", True)
        self.ai_btn.clicked.connect(self._ai_classify); tb2.addWidget(self.ai_btn)
        self.local_ai_btn = QPushButton("Local AI"); self.local_ai_btn.setProperty("secondary", True)
        self.local_ai_btn.clicked.connect(self._ollama_classify); tb2.addWidget(self.local_ai_btn)
        self.receipts_btn = QPushButton("Receipts"); self.receipts_btn.setProperty("secondary", True)
        self.receipts_btn.clicked.connect(self._show_receipts); tb2.addWidget(self.receipts_btn)
        self.receipt_vision_btn = QPushButton("Receipt Vision"); self.receipt_vision_btn.setProperty("secondary", True)
        self.receipt_vision_btn.clicked.connect(self._run_receipt_vision); tb2.addWidget(self.receipt_vision_btn)
        self.thread_btn = QPushButton("Threads"); self.thread_btn.setProperty("secondary", True)
        self.thread_btn.clicked.connect(self._summarize_threads); tb2.addWidget(self.thread_btn)
        layout.addLayout(tb2)

        # Main content
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left panel with group-by
        left = QWidget(); ll = QVBoxLayout(left); ll.setContentsMargins(0,0,0,0)
        gb_row = QHBoxLayout()
        gb_row.addWidget(QLabel("Group by:"))
        self.group_combo = QComboBox()
        self.group_combo.addItems(["Category", "Sender Domain", "Sender", "Source Folder"])
        self.group_combo.currentTextChanged.connect(self._refresh_tree)
        gb_row.addWidget(self.group_combo, 1)
        ll.addLayout(gb_row)

        self.cat_tree = QTreeWidget(); self.cat_tree.setHeaderLabels(["Name", "Count"])
        self.cat_tree.setColumnWidth(0, 260); self.cat_tree.itemClicked.connect(self._on_tree_click)
        self.cat_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.cat_tree.customContextMenuRequested.connect(self._ctx_menu)
        ll.addWidget(self.cat_tree)

        cat_btns = QHBoxLayout()
        for lbl, slot, prop in [("Rename",self._rename,"secondary"),("Merge",self._merge,"secondary"),("Delete",self._delete,"danger")]:
            b = QPushButton(lbl); b.setProperty(prop, True); b.clicked.connect(slot); cat_btns.addWidget(b)
        ll.addLayout(cat_btns)
        splitter.addWidget(left)

        # Right: search + table + preview
        right = QWidget(); rl = QVBoxLayout(right); rl.setContentsMargins(0,0,0,0)

        # Search & filter bar
        filter_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search emails (subject, sender)...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._apply_filter)
        filter_row.addWidget(self.search_input, 1)
        filter_row.addWidget(QLabel("From:"))
        self.date_from = QDateEdit(); self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate(2018, 1, 1))
        self.date_from.setDisplayFormat("yyyy-MM-dd")
        self.date_from.dateChanged.connect(self._apply_filter)
        filter_row.addWidget(self.date_from)
        filter_row.addWidget(QLabel("To:"))
        self.date_to = QDateEdit(); self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setDisplayFormat("yyyy-MM-dd")
        self.date_to.dateChanged.connect(self._apply_filter)
        filter_row.addWidget(self.date_to)
        rl.addLayout(filter_row)

        self.email_count_lbl = QLabel(""); self.email_count_lbl.setStyleSheet(f"color:{C.SUBTEXT0};font-size:12px;")
        rl.addWidget(self.email_count_lbl)

        # Vertical splitter: table on top, preview below
        vsplitter = QSplitter(Qt.Orientation.Vertical)

        self.table = QTableWidget(); self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["From", "Subject", "Date", "Folder", "Conf", "Flags"])
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for c in (2,3,4,5): h.setSectionResizeMode(c, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setColumnWidth(0, 170)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.currentCellChanged.connect(self._on_email_selected)
        vsplitter.addWidget(self.table)

        # Email preview pane
        self.preview = QTextBrowser()
        self.preview.setOpenExternalLinks(True)
        self.preview.setStyleSheet(
            f"background:{C.MANTLE}; color:{C.TEXT}; border:1px solid {C.SURFACE0}; "
            f"border-radius:6px; padding:8px; font-size:13px;")
        self.preview.setPlaceholderText("Select an email to preview its content")
        vsplitter.addWidget(self.preview)
        vsplitter.setSizes([400, 200])
        rl.addWidget(vsplitter, 1)

        move_row = QHBoxLayout()
        sel_all_btn = QPushButton("Select All"); sel_all_btn.setProperty("secondary", True)
        sel_all_btn.clicked.connect(self.table.selectAll); move_row.addWidget(sel_all_btn)
        move_row.addWidget(QLabel("Move to:"))
        self.move_combo = QComboBox(); self.move_combo.setMinimumWidth(200)
        move_row.addWidget(self.move_combo, 1)
        mv_btn = QPushButton("Move"); mv_btn.setProperty("secondary", True)
        mv_btn.clicked.connect(self._move_emails); move_row.addWidget(mv_btn)
        rl.addLayout(move_row)
        splitter.addWidget(right); splitter.setSizes([320, 680])
        layout.addWidget(splitter, 1)

        # Bottom execution
        bot = QVBoxLayout(); bot.setSpacing(8)
        mr = QHBoxLayout()
        mr.addWidget(QLabel("<b>Execute:</b>"))
        self.mode_local = QRadioButton("Organize Local Files"); self.mode_local.setChecked(True)
        self.mode_gmail = QRadioButton("Apply Gmail Labels")
        mr.addWidget(self.mode_local); mr.addWidget(self.mode_gmail); mr.addStretch()
        bot.addLayout(mr)

        opts = QHBoxLayout()
        self.gmail_opts = QWidget(); go = QHBoxLayout(self.gmail_opts); go.setContentsMargins(0,0,0,0)
        go.addWidget(QLabel("Prefix:")); self.prefix_input = QLineEdit(); self.prefix_input.setMaximumWidth(150)
        go.addWidget(self.prefix_input)
        self.archive_chk = QCheckBox("Archive from Inbox"); go.addWidget(self.archive_chk)
        self.dry_run_chk = QCheckBox("Preview only (no mailbox changes)")
        self.dry_run_chk.setChecked(True); go.addWidget(self.dry_run_chk)
        opts.addWidget(self.gmail_opts)
        self.local_opts = QWidget(); lo = QHBoxLayout(self.local_opts); lo.setContentsMargins(0,0,0,0)
        self.copy_radio = QRadioButton("Copy"); self.copy_radio.setChecked(True)
        self.move_radio = QRadioButton("Move"); lo.addWidget(self.copy_radio); lo.addWidget(self.move_radio)
        opts.addWidget(self.local_opts)
        opts.addStretch()
        self.exec_btn = QPushButton("Organize Local Files")
        self.exec_btn.setStyleSheet(f"background-color:{C.GREEN};font-size:14px;padding:10px 30px;")
        self.exec_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.exec_btn.clicked.connect(self.execute_requested.emit); opts.addWidget(self.exec_btn)
        bot.addLayout(opts); layout.addLayout(bot)

        self.mode_gmail.toggled.connect(self._mode_changed); self.mode_local.toggled.connect(self._mode_changed)
        self.dry_run_chk.toggled.connect(self._mode_changed)
        self._mode_changed()

    def _mode_changed(self):
        g = self.mode_gmail.isChecked()
        self.gmail_opts.setVisible(g); self.local_opts.setVisible(not g)
        self.exec_btn.setText(
            "Preview Gmail Labels" if g and self.dry_run_chk.isChecked()
            else "Apply Gmail Labels" if g else "Organize Local Files"
        )
        self.exec_btn.setStyleSheet(f"background-color:{C.BLUE if g else C.GREEN};font-size:14px;padding:10px 30px;")

    def load_categories(self, eng, has_local=False, dl_dir=""):
        self.engine = eng; self.has_local = has_local; self._dl_dir = dl_dir
        self._refresh_tree(); self._refresh_combo()
        s = eng.get_summary()
        self.summary_lbl.setText(f"{s['total']:,} emails | {s['categorized']:,} categorized | {s['uncategorized']:,} uncategorized")
        if has_local:
            self.mode_local.setChecked(True)
        else:
            self.mode_gmail.setChecked(True)
            self.mode_local.setEnabled(False)
        self.attach_btn.setEnabled(has_local)
        self.sensitive_btn.setEnabled(has_local)
        self.redact_btn.setEnabled(has_local)
        self.ai_btn.setEnabled(has_local and HAS_ANTHROPIC)
        self.local_ai_btn.setEnabled(has_local)
        self.receipts_btn.setEnabled(has_local)
        self.receipt_vision_btn.setEnabled(has_local)
        self.location_btn.setEnabled(has_local)
        self.thread_btn.setEnabled(has_local and HAS_ANTHROPIC)

    def _refresh_tree(self, *_):
        self.cat_tree.clear()
        if not self.engine: return
        mode = self.group_combo.currentText()
        groups = {}

        if mode == "Category":
            for cat, emails in sorted(self.engine.categories.items(), key=lambda x: -len(x[1])):
                if '/' in cat:
                    parent, child = cat.split('/', 1)
                    if parent not in groups: groups[parent] = {}
                    groups[parent][child] = (cat, len(emails))
                else:
                    groups[cat] = len(emails)
        elif mode == "Sender Domain":
            domain_groups = defaultdict(list)
            for em in self.engine.emails: domain_groups[em.sender_domain].append(em)
            for d, emails in sorted(domain_groups.items(), key=lambda x: -len(x[1])):
                groups[d] = len(emails)
        elif mode == "Sender":
            sender_groups = defaultdict(list)
            for em in self.engine.emails: sender_groups[em.sender_name or em.sender].append(em)
            for s, emails in sorted(sender_groups.items(), key=lambda x: -len(x[1]))[:200]:
                groups[s] = len(emails)
        elif mode == "Source Folder":
            folder_groups = defaultdict(list)
            for em in self.engine.emails: folder_groups[em.source_folder or 'Unknown'].append(em)
            for f, emails in sorted(folder_groups.items(), key=lambda x: -len(x[1])):
                groups[sanitize_folder_name(f)] = len(emails)

        for name, val in groups.items():
            if isinstance(val, int):
                item = QTreeWidgetItem([name, f"{val:,}"])
                item.setData(0, Qt.ItemDataRole.UserRole, name)
                item.setForeground(0, QColor(C.BLUE))
                self.cat_tree.addTopLevelItem(item)
            elif isinstance(val, dict):
                total = sum(c[1] for c in val.values())
                pi = QTreeWidgetItem([name, f"{total:,}"])
                pi.setData(0, Qt.ItemDataRole.UserRole, None)
                pi.setForeground(0, QColor(C.MAUVE))
                self.cat_tree.addTopLevelItem(pi)
                for child, (full, c) in sorted(val.items(), key=lambda x: -x[1][1]):
                    ci = QTreeWidgetItem([child, f"{c:,}"])
                    ci.setData(0, Qt.ItemDataRole.UserRole, full)
                    pi.addChild(ci)
        self.cat_tree.expandAll()

    def _refresh_combo(self):
        self.move_combo.clear()
        if self.engine:
            for c in sorted(self.engine.categories.keys()): self.move_combo.addItem(c)

    def _get_emails_for_item(self, item):
        if not self.engine: return []
        name = item.data(0, Qt.ItemDataRole.UserRole)
        mode = self.group_combo.currentText()

        if mode == "Category":
            if name and name in self.engine.categories:
                return self.engine.categories[name]
            emails = []
            for i in range(item.childCount()):
                cn = item.child(i).data(0, Qt.ItemDataRole.UserRole)
                if cn and cn in self.engine.categories:
                    emails.extend(self.engine.categories[cn])
            return emails
        elif mode == "Sender Domain":
            return [em for em in self.engine.emails if em.sender_domain == name]
        elif mode == "Sender":
            return [em for em in self.engine.emails if (em.sender_name or em.sender) == name]
        elif mode == "Source Folder":
            return [em for em in self.engine.emails if sanitize_folder_name(em.source_folder or 'Unknown') == name]
        return []

    def _on_tree_click(self, item):
        emails = self._get_emails_for_item(item)
        self._all_current_emails = emails
        self._apply_filter()

    def _apply_filter(self, *_):
        """Filter current email list by search text and date range."""
        emails = self._all_current_emails
        q = self.search_input.text().strip()
        if q:
            emails = search_emails(emails, q)
        d_from = self.date_from.date().toPyDate()
        d_to = self.date_to.date().toPyDate()
        from datetime import date as dt_date
        d_from_dt = datetime(d_from.year, d_from.month, d_from.day)
        d_to_dt = datetime(d_to.year, d_to.month, d_to.day, 23, 59, 59)
        emails = [em for em in emails if not em.date_parsed or (d_from_dt <= em.date_parsed <= d_to_dt)]
        self._current_emails = emails
        self._show_emails(emails)

    def _show_emails(self, emails):
        self.email_count_lbl.setText(f"{len(emails):,} emails")
        display = sorted(emails, key=lambda e: e.date_parsed or datetime.min, reverse=True)[:2000]
        self._display_list = display  # Store for preview lookup
        self.table.setRowCount(len(display))
        for row, em in enumerate(display):
            self.table.setItem(row, 0, QTableWidgetItem(em.sender_name or em.sender))
            self.table.setItem(row, 1, QTableWidgetItem(em.subject))
            self.table.setItem(row, 2, QTableWidgetItem(em.date_parsed.strftime("%Y-%m-%d") if em.date_parsed else ""))
            self.table.setItem(row, 3, QTableWidgetItem(sanitize_folder_name(em.source_folder) if em.source_folder else ""))
            ci = QTableWidgetItem(f"{em.confidence:.0%}")
            ci.setForeground(QColor(C.GREEN if em.confidence >= 0.8 else C.YELLOW if em.confidence >= 0.5 else C.RED))
            self.table.setItem(row, 4, ci)
            flags = []
            if em.is_newsletter: flags.append("NL")
            if em.sensitive_flags: flags.append("SENS")
            self.table.setItem(row, 5, QTableWidgetItem(' '.join(flags)))
            self.table.item(row, 0).setData(Qt.ItemDataRole.UserRole, em.uid)
        if len(emails) > 2000:
            self.email_count_lbl.setText(f"{len(emails):,} emails (showing 2,000)")

    def _on_email_selected(self, row, *_):
        """Show email preview when a row is selected."""
        if row < 0 or not hasattr(self, '_display_list') or row >= len(self._display_list):
            return
        em = self._display_list[row]
        if em.local_path and Path(em.local_path).exists():
            try:
                with open(em.local_path, 'rb') as f:
                    msg = email.message_from_bytes(f.read())
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        ct = part.get_content_type()
                        if ct == 'text/html':
                            body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                            break
                        elif ct == 'text/plain' and not body:
                            text = part.get_payload(decode=True).decode('utf-8', errors='replace')
                            body = f"<pre style='white-space:pre-wrap;word-break:break-word;color:{C.TEXT};font-family:Segoe UI,sans-serif'>{text}</pre>"
                else:
                    payload = msg.get_payload(decode=True)
                    if payload:
                        if msg.get_content_type() == 'text/html':
                            body = payload.decode('utf-8', errors='replace')
                        else:
                            text = payload.decode('utf-8', errors='replace')
                            body = f"<pre style='white-space:pre-wrap;word-break:break-word;color:{C.TEXT};font-family:Segoe UI,sans-serif'>{text}</pre>"
                header = (f"<div style='background:{C.SURFACE0};padding:10px;border-radius:6px;margin-bottom:10px'>"
                          f"<b style='color:{C.BLUE}'>{em.subject}</b><br>"
                          f"<span style='color:{C.SUBTEXT0}'>From:</span> {em.sender_name} &lt;{em.sender}&gt;<br>"
                          f"<span style='color:{C.SUBTEXT0}'>Date:</span> {em.date}<br>"
                          f"<span style='color:{C.SUBTEXT0}'>Folder:</span> {em.source_folder}</div>")
                self.preview.setHtml(header + body)
            except Exception as e:
                self.preview.setPlainText(f"Error loading email: {e}")
        else:
            self.preview.setHtml(
                f"<div style='background:{C.SURFACE0};padding:10px;border-radius:6px'>"
                f"<b style='color:{C.BLUE}'>{em.subject}</b><br>"
                f"<span style='color:{C.SUBTEXT0}'>From:</span> {em.sender_name} &lt;{em.sender}&gt;<br>"
                f"<span style='color:{C.SUBTEXT0}'>Date:</span> {em.date}<br>"
                f"<span style='color:{C.SUBTEXT0}'>Folder:</span> {em.source_folder}<br><br>"
                f"<i style='color:{C.OVERLAY0}'>Email body not available (headers-only scan mode)</i></div>")

    def _ctx_menu(self, pos):
        item = self.cat_tree.itemAt(pos)
        if not item: return
        m = QMenu(self); m.addAction("Rename", self._rename); m.addAction("Merge into...", self._merge)
        m.addSeparator(); m.addAction("Delete", self._delete)
        m.exec(self.cat_tree.viewport().mapToGlobal(pos))

    def _rename(self):
        cat = self._get_sel_cat()
        if not cat: return
        new, ok = QInputDialog.getText(self, "Rename", "New name:", text=cat)
        if ok and new and new != cat:
            self.engine.rename_category(cat, new); self._refresh_tree(); self._refresh_combo()

    def _merge(self):
        src = self._get_sel_cat()
        if not src or not self.engine: return
        tgt, ok = QInputDialog.getItem(self, "Merge", f"Merge '{src}' into:",
            sorted(self.engine.categories.keys()), 0, False)
        if ok and tgt and tgt != src:
            self.engine.merge_categories([src], tgt); self._refresh_tree(); self._refresh_combo()

    def _delete(self):
        cat = self._get_sel_cat()
        if cat: self.engine.delete_category(cat); self._refresh_tree(); self._refresh_combo()

    def _get_sel_cat(self):
        item = self.cat_tree.currentItem()
        return item.data(0, Qt.ItemDataRole.UserRole) if item else None

    def _move_emails(self):
        if not self.engine: return
        tgt = self.move_combo.currentText()
        if not tgt: return
        uids = []
        for r in set(i.row() for i in self.table.selectedIndexes()):
            it = self.table.item(r, 0)
            if it:
                uid = it.data(Qt.ItemDataRole.UserRole)
                if uid: uids.append(uid)
        if uids:
            self.engine.move_emails(uids, tgt); self._refresh_tree()
            item = self.cat_tree.currentItem()
            if item: self._on_tree_click(item)

    def _show_stats(self):
        if self.engine: StatsDialog(self.engine, self).exec()

    def _show_contacts(self):
        if self.engine: ContactDialog(self.engine, self).exec()

    def _show_subs(self):
        if self.engine: SubscriptionDialog(self.engine.subscriptions, self).exec()

    def _show_rules(self):
        if self.engine:
            RulesEditorDialog(self.engine.clean_rules, list(self.engine.categories.keys()), self).exec()

    def _show_inbox_zero(self):
        if not self.engine:
            return
        suggestions = self.engine.inbox_zero_suggestions()
        dialog = QDialog(self)
        dialog.setWindowTitle("Inbox Zero Suggestions (preview only)")
        dialog.setMinimumSize(650, 450)
        layout = QVBoxLayout(dialog)
        text = QPlainTextEdit(); text.setReadOnly(True)
        lines = [f"{len(suggestions):,} suggestions — nothing will be changed"]
        for suggestion in suggestions[:1000]:
            lines.append(f"{suggestion['uid']}: {suggestion['action']} ({suggestion['reason']})")
        text.setPlainText('\n'.join(lines)); layout.addWidget(text)
        close = QPushButton("Close"); close.clicked.connect(dialog.close); layout.addWidget(close)
        dialog.exec()

    def _show_location_timeline(self):
        if not self.engine:
            return
        timeline = self.engine.location_timeline()
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Location Timeline ({len(timeline):,} public hops)")
        dialog.setMinimumSize(780, 480)
        layout = QVBoxLayout(dialog)
        table = QTableWidget(len(timeline), 5)
        table.setHorizontalHeaderLabels(["Received", "IP", "Country", "UID", "Hop"])
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        for row, item in enumerate(timeline):
            table.setItem(row, 0, QTableWidgetItem(item.get('received_at') or item.get('date', '')))
            table.setItem(row, 1, QTableWidgetItem(item.get('ip', '')))
            table.setItem(row, 2, QTableWidgetItem(item.get('country', '')))
            table.setItem(row, 3, QTableWidgetItem(item.get('uid', '')))
            table.setItem(row, 4, QTableWidgetItem(str(item.get('hop', ''))))
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(table)
        buttons = QHBoxLayout()
        export = QPushButton("Export CSV")
        export.clicked.connect(lambda: self._export_location_timeline(timeline))
        buttons.addWidget(export)
        close = QPushButton("Close"); close.clicked.connect(dialog.close); buttons.addWidget(close)
        layout.addLayout(buttons)
        dialog.exec()

    def _export_location_timeline(self, timeline):
        path, _ = QFileDialog.getSaveFileName(self, "Export location timeline", "location_timeline.csv", "CSV (*.csv)")
        if not path:
            return
        export_location_timeline_csv(timeline, path)
        QMessageBox.information(self, "Exported", f"Wrote {len(timeline):,} location hops to:\n{path}")

    def _export_csv(self):
        if not self.engine: return
        p, _ = QFileDialog.getSaveFileName(self, "Export CSV", "gmaildownloader_export.csv", "CSV (*.csv)")
        if p:
            self.engine.export_csv(p)
            QMessageBox.information(self, "Exported", f"Exported {len(self.engine.emails):,} emails to CSV")

    def _export_json(self):
        if not self.engine: return
        p, _ = QFileDialog.getSaveFileName(self, "Export JSON", "gmaildownloader_export.json", "JSON (*.json)")
        if p:
            self.engine.export_json(p)
            QMessageBox.information(self, "Exported", f"Exported {len(self.engine.emails):,} emails to JSON")

    def _export_mbox(self):
        if not self.engine: return
        p, _ = QFileDialog.getSaveFileName(self, "Export MBOX", "gmaildownloader.mbox", "MBOX (*.mbox);;All files (*)")
        if p:
            try:
                export_mbox(self.engine.emails, p)
                QMessageBox.information(self, "Exported", f"Exported {len(self.engine.emails):,} emails to MBOX")
            except Exception as exc:
                QMessageBox.warning(self, "Export failed", str(exc))

    def _export_encrypted(self):
        if not self.engine or not self._dl_dir:
            QMessageBox.warning(self, "Unavailable", "Encrypted archive export requires a downloaded mailbox.")
            return
        p, _ = QFileDialog.getSaveFileName(self, "Encrypted archive", "gmaildownloader.gd", "Encrypted archive (*.gd)")
        if not p:
            return
        password, ok = QInputDialog.getText(self, "Archive passphrase", "Passphrase:", QLineEdit.EchoMode.Password)
        if not ok or not password:
            return
        try:
            encrypt_archive(self._dl_dir, p, password)
            QMessageBox.information(self, "Encrypted", f"Encrypted archive written to:\n{p}")
        except Exception as exc:
            QMessageBox.warning(self, "Encryption failed", str(exc))

    def _export_markdown(self):
        if not self.engine: return
        destination = QFileDialog.getExistingDirectory(self, "Markdown vault folder")
        if destination:
            try:
                files = self.engine.export_markdown(destination)
                QMessageBox.information(self, "Exported", f"Wrote {len(files):,} Markdown notes")
            except Exception as exc:
                QMessageBox.warning(self, "Export failed", str(exc))

    def _export_pdf(self):
        if not self.engine: return
        p, _ = QFileDialog.getSaveFileName(self, "Export PDF", "gmaildownloader.pdf", "PDF (*.pdf)")
        if p:
            try:
                self.engine.export_pdf(p)
                QMessageBox.information(self, "Exported", f"PDF written to:\n{p}")
            except Exception as exc:
                QMessageBox.warning(self, "Export failed", str(exc))

    def _export_graph(self):
        if not self.engine: return
        p, _ = QFileDialog.getSaveFileName(
            self, "Export relationship graph", "relationships.json",
            "JSON (*.json);;GraphML (*.graphml)"
        )
        if p:
            try:
                self.engine.export_relationship_graph(p)
                QMessageBox.information(self, "Exported", f"Graph written to:\n{p}")
            except Exception as exc:
                QMessageBox.warning(self, "Export failed", str(exc))

    def _export_contact_graph(self):
        if not self.engine: return
        p, _ = QFileDialog.getSaveFileName(
            self, "Export contact graph", "contacts.json",
            "JSON (*.json);;GraphML (*.graphml)"
        )
        if p:
            try:
                self.engine.export_contact_graph(p)
                QMessageBox.information(self, "Exported", f"Graph written to:\n{p}")
            except Exception as exc:
                QMessageBox.warning(self, "Export failed", str(exc))

    def _export_html(self):
        if not self.engine or not self._dl_dir: return
        self.export_btn.setEnabled(False)
        self._html_w = HtmlArchiveWorker(self.engine, self._dl_dir)
        self._html_w.log.connect(lambda s: None)  # Silent
        self._html_w.finished_signal.connect(self._html_done)
        self._html_w.error.connect(lambda e: (self.export_btn.setEnabled(True),
            QMessageBox.warning(self, "Error", e)))
        self._html_w.start()

    def _html_done(self, path):
        self.export_btn.setEnabled(True)
        QMessageBox.information(self, "HTML Archive",
            f"Generated browseable HTML archive at:\n{path}\n\nOpen index.html to browse.")
        os.startfile(path)

    def _extract_attachments(self):
        if not self.engine or not self._dl_dir: return
        self.attach_btn.setEnabled(False); self.attach_btn.setText("Extracting...")
        self._att_worker = AttachmentExtractWorker(self.engine.emails, self._dl_dir)
        self._att_worker.status.connect(lambda s: self.attach_btn.setText(s[:30]))
        self._att_worker.finished_signal.connect(self._att_done)
        self._att_worker.error.connect(lambda e: (self.attach_btn.setEnabled(True),
            self.attach_btn.setText("Extract Attachments"), QMessageBox.warning(self, "Error", e)))
        self._att_worker.start()

    def _att_done(self, count, size, path):
        self.attach_btn.setEnabled(True); self.attach_btn.setText("Extract Attachments")
        QMessageBox.information(self, "Attachments Extracted",
            f"Extracted {count:,} unique attachments ({format_size(size)})\nSaved to: {path}")

    def _scan_sensitive(self):
        if not self.engine: return
        self.sensitive_btn.setEnabled(False); self.sensitive_btn.setText("Scanning...")
        self._sens_worker = SensitiveScanWorker(self.engine.emails)
        self._sens_worker.status.connect(lambda s: self.sensitive_btn.setText(s[:30]))
        self._sens_worker.finished_signal.connect(self._sens_done)
        self._sens_worker.error.connect(lambda e: (self.sensitive_btn.setEnabled(True),
            self.sensitive_btn.setText("Scan Sensitive")))
        self._sens_worker.start()

    def _sens_done(self, count):
        self.sensitive_btn.setEnabled(True); self.sensitive_btn.setText("Scan Sensitive")
        if count > 0:
            # Create "Sensitive" category for flagged emails
            for em in self.engine.emails:
                if em.sensitive_flags and "Sensitive" not in em.category:
                    self.engine.categories["Sensitive"].append(em)
            self._refresh_tree()
        QMessageBox.information(self, "Sensitive Scan",
            f"Found {count} emails with sensitive content (SSN, CC#, passwords, API keys)")

    def _redact_sensitive(self):
        if not self.engine or not self._dl_dir:
            return
        destination = QFileDialog.getExistingDirectory(
            self, "Redacted copies folder", str(Path(self._dl_dir) / "redacted")
        )
        if not destination:
            return
        written = 0
        for em in self.engine.emails:
            if not em.local_path or not Path(em.local_path).exists():
                continue
            target_dir = Path(destination) / sanitize_filename(em.category or "Uncategorized", 60)
            target = target_dir / f"{sanitize_filename(em.uid.replace(':', '_'), 80)}.eml"
            try:
                flags = redact_eml(em.local_path, target)
                if flags:
                    em.sensitive_flags = sorted(set(em.sensitive_flags).union(flags))
                    written += 1
            except Exception:
                continue
        QMessageBox.information(
            self, "Redacted copies", f"Wrote {written:,} redacted EML copies to:\n{destination}"
        )

    def _ollama_classify(self):
        if not self.engine:
            return
        candidates = self.engine.confidence_candidates(0.75)
        if not candidates:
            QMessageBox.information(self, "Done", "No low-confidence messages to classify.")
            return
        self.local_ai_btn.setEnabled(False); self.local_ai_btn.setText("Classifying...")
        existing = [category for category in self.engine.categories if category != "Uncategorized"]
        self._ai_candidates = candidates
        self._ollama_w = OllamaClassifyWorker(candidates, existing)
        self._ollama_w.classified.connect(self._ai_result)
        self._ollama_w.finished_signal.connect(
            lambda: (self.local_ai_btn.setEnabled(True), self.local_ai_btn.setText("Local AI"))
        )
        self._ollama_w.error.connect(
            lambda error: (self.local_ai_btn.setEnabled(True), self.local_ai_btn.setText("Local AI"),
                           QMessageBox.warning(self, "Local AI error", error))
        )
        self._ollama_w.start()

    def _show_receipts(self):
        if not self.engine:
            return
        receipts = extract_receipts(self.engine.emails)
        self._present_receipts(receipts, f"Receipt Extraction ({len(receipts):,})")

    def _present_receipts(self, receipts, title):
        dialog = QDialog(self); dialog.setWindowTitle(title)
        dialog.setMinimumSize(700, 450)
        layout = QVBoxLayout(dialog)
        table = QTableWidget(len(receipts), 5)
        table.setHorizontalHeaderLabels(["Merchant", "Amount", "Date", "Sender", "UID"])
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for row, receipt in enumerate(receipts):
            table.setItem(row, 0, QTableWidgetItem(receipt.get('merchant', '')))
            amount = receipt.get('amount')
            table.setItem(row, 1, QTableWidgetItem(f"{amount:.2f}" if amount is not None else ""))
            table.setItem(row, 2, QTableWidgetItem(receipt.get('date', '')))
            table.setItem(row, 3, QTableWidgetItem(receipt.get('sender', '')))
            table.setItem(row, 4, QTableWidgetItem(receipt.get('uid', '')))
        layout.addWidget(table)
        buttons = QHBoxLayout()
        export = QPushButton("Export OFX")
        export.clicked.connect(lambda: self._export_receipts_ofx(receipts))
        buttons.addWidget(export)
        close = QPushButton("Close"); close.clicked.connect(dialog.close); buttons.addWidget(close)
        layout.addLayout(buttons)
        dialog.exec()

    def _export_receipts_ofx(self, receipts):
        path, _ = QFileDialog.getSaveFileName(self, "Export receipts as OFX", "receipts.ofx", "OFX (*.ofx)")
        if not path:
            return
        try:
            export_receipts_ofx(receipts, path)
            QMessageBox.information(self, "Exported", f"Wrote {len(receipts):,} receipts to:\n{path}")
        except Exception as exc:
            QMessageBox.warning(self, "OFX export failed", str(exc))

    def _run_receipt_vision(self):
        if not self.engine:
            return
        choices = ["Anthropic (cloud)", "Ollama (local)"]
        choice, ok = QInputDialog.getItem(self, "Receipt Vision Backend", "Backend", choices, 0, False)
        if not ok:
            return
        backend = 'anthropic' if choice.startswith('Anthropic') else 'ollama'
        api_key = ''
        if backend == 'anthropic':
            if not HAS_ANTHROPIC:
                QMessageBox.warning(self, "Missing dependency", "Install the anthropic package first.")
                return
            api_key, ok = QInputDialog.getText(
                self, "Anthropic API Key", "Analyze receipt/invoice attachments with Claude",
                QLineEdit.EchoMode.Password
            )
            if not ok or not api_key:
                return
        self.receipt_vision_btn.setEnabled(False); self.receipt_vision_btn.setText("Vision...")
        classifier = ReceiptVisionClassifier(backend=backend, api_key=api_key)
        self._receipt_vision_worker = ReceiptVisionWorker(self.engine.emails, classifier)
        self._receipt_vision_worker.finished_signal.connect(self._receipt_vision_done)
        self._receipt_vision_worker.error.connect(self._receipt_vision_error)
        self._receipt_vision_worker.start()

    def _receipt_vision_done(self, receipts):
        self.receipt_vision_btn.setEnabled(True); self.receipt_vision_btn.setText("Receipt Vision")
        if receipts:
            self._present_receipts(receipts, f"Receipt Vision ({len(receipts):,})")
        else:
            QMessageBox.information(self, "Receipt Vision", "No receipt or invoice attachments were classified.")

    def _receipt_vision_error(self, error):
        self.receipt_vision_btn.setEnabled(True); self.receipt_vision_btn.setText("Receipt Vision")
        QMessageBox.warning(self, "Receipt Vision failed", error)

    def _ai_classify(self):
        if not HAS_ANTHROPIC:
            QMessageBox.warning(self, "Missing", "anthropic package not installed."); return
        uncat = self.engine.categories.get("Uncategorized", [])
        if not uncat:
            QMessageBox.information(self, "Done", "No uncategorized emails."); return
        key, ok = QInputDialog.getText(self, "API Key",
            f"Classify {len(uncat):,} emails via Claude Haiku", QLineEdit.EchoMode.Password)
        if not ok or not key: return
        self.ai_btn.setEnabled(False); self.ai_btn.setText("Classifying...")
        existing = [k for k in self.engine.categories if k != "Uncategorized"]
        self._ai_candidates = uncat
        self._ai_w = AiClassifyWorker(key, uncat, existing)
        self._ai_w.classified.connect(self._ai_result)
        self._ai_w.finished_signal.connect(self._ai_done)
        self._ai_w.error.connect(lambda e: (self.ai_btn.setEnabled(True), self.ai_btn.setText("AI Classify")))
        self._ai_w.start()

    def _ai_result(self, dmap):
        candidates = getattr(self, '_ai_candidates', self.engine.categories.get("Uncategorized", []))
        for domain, cat in dmap.items():
            moves = [e for e in candidates if e.sender_domain == domain]
            for em in moves:
                old = em.category
                if old in self.engine.categories:
                    self.engine.categories[old] = [item for item in self.engine.categories[old] if item.uid != em.uid]
                    if not self.engine.categories[old]:
                        del self.engine.categories[old]
                em.category = cat; em.confidence = 0.75
                self.engine.categories[cat].append(em)
                self.engine.learned.learn(em, cat)
        if not self.engine.categories.get("Uncategorized"):
            if "Uncategorized" in self.engine.categories: del self.engine.categories["Uncategorized"]
        self.engine.learned.save()
        self._ai_candidates = []
        self._refresh_tree(); self._refresh_combo()

    def _ai_done(self):
        self.ai_btn.setEnabled(True); self.ai_btn.setText("AI Classify")
        s = self.engine.get_summary()
        self.summary_lbl.setText(f"{s['total']:,} emails | {s['categorized']:,} categorized | {s['uncategorized']:,} uncategorized")

    def _summarize_threads(self):
        if not self.engine or not HAS_ANTHROPIC: return
        if not self.engine.threads:
            QMessageBox.information(self, "No Threads", "No multi-message threads found."); return
        key, ok = QInputDialog.getText(self, "API Key",
            f"Summarize {len(self.engine.threads)} threads via Claude Haiku",
            QLineEdit.EchoMode.Password)
        if not ok or not key: return
        # Summarize top 50 longest threads
        threads = sorted(self.engine.threads.items(), key=lambda x: -len(x[1]))[:50]
        self.thread_btn.setEnabled(False); self.thread_btn.setText("Summarizing...")
        self._thread_summaries = {}
        self._ts_w = ThreadSummaryWorker(key, threads)
        self._ts_w.result.connect(lambda tid, s: self._thread_summaries.update({tid: s}))
        self._ts_w.finished_signal.connect(self._threads_done)
        self._ts_w.error.connect(lambda e: (self.thread_btn.setEnabled(True),
            self.thread_btn.setText("Summarize Threads")))
        self._ts_w.start()

    def _threads_done(self):
        self.thread_btn.setEnabled(True); self.thread_btn.setText("Summarize Threads")
        if self._thread_summaries:
            text = "\n\n".join(f"Thread ({len(self.engine.threads.get(tid,[]))} msgs):\n{s}"
                              for tid, s in self._thread_summaries.items())
            dlg = QDialog(self); dlg.setWindowTitle("Thread Summaries"); dlg.setMinimumSize(700, 500)
            dl = QVBoxLayout(dlg)
            te = QPlainTextEdit(text); te.setReadOnly(True)
            te.setStyleSheet("font-family:'Cascadia Code','Consolas',monospace;font-size:12px;")
            dl.addWidget(te)
            cb = QPushButton("Close"); cb.clicked.connect(dlg.close); dl.addWidget(cb)
            dlg.exec()


# ─── UI: Execute Page ────────────────────────────────────────────────────

class ExecutePage(QWidget):
    def __init__(self):
        super().__init__(); self.worker = None; self._out = ""; self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self); layout.setSpacing(12)
        self.title_lbl = QLabel("Executing..."); self.title_lbl.setStyleSheet("font-size:20px;font-weight:bold;")
        layout.addWidget(self.title_lbl)
        self.status_label = QLabel("Starting..."); self.status_label.setStyleSheet(f"color:{C.SUBTEXT0};")
        layout.addWidget(self.status_label)
        self.progress = QProgressBar(); self.progress.setTextVisible(False); self.progress.setFixedHeight(8)
        layout.addWidget(self.progress)
        self.pct = QLabel("0%"); self.pct.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(self.pct)
        self.log = QPlainTextEdit(); self.log.setReadOnly(True)
        self.log.setStyleSheet("font-family:'Cascadia Code','Consolas',monospace;font-size:12px;")
        layout.addWidget(self.log, 1)
        bot = QHBoxLayout()
        self.stop_btn = QPushButton("Stop"); self.stop_btn.setProperty("danger", True); bot.addWidget(self.stop_btn)
        bot.addStretch()
        self.open_btn = QPushButton("Open Output"); self.open_btn.setProperty("secondary", True)
        self.open_btn.setVisible(False); bot.addWidget(self.open_btn)
        self.done_lbl = QLabel(""); bot.addWidget(self.done_lbl)
        layout.addLayout(bot)

    def start_gmail(self, host, addr, pw, cats, prefix, archive, dry_run=True,
                    port=993, use_ssl=True, auth_mode='password', access_token=''):
        self.title_lbl.setText("Previewing Gmail Labels" if dry_run else "Applying Gmail Labels")
        self.worker = ImapLabelWorker(
            host, addr, pw, cats, prefix, archive, dry_run=dry_run,
            port=port, use_ssl=use_ssl, auth_mode=auth_mode, access_token=access_token
        )
        self._wire()

    def start_local(self, cats, out_dir, copy):
        self.title_lbl.setText("Organizing Local Files")
        self._out = str(Path(out_dir) / "organized")
        self.worker = LocalOrganizeWorker(cats, out_dir, copy)
        self._wire()
        self.open_btn.clicked.connect(lambda: os.startfile(self._out))

    def _wire(self):
        self.worker.progress.connect(lambda c,t: (self.progress.setMaximum(t), self.progress.setValue(c),
            self.pct.setText(f"{int(c/t*100) if t else 0}% ({c:,}/{t:,})")))
        self.worker.status.connect(self.status_label.setText)
        self.worker.log.connect(self.log.appendPlainText)
        self.worker.finished_signal.connect(self._done)
        self.worker.error.connect(lambda e: (self.status_label.setText("Error"),
            self.log.appendPlainText(f"ERROR: {e}")))
        self.stop_btn.clicked.connect(self.worker.stop)
        self.worker.start()

    def _done(self):
        self.status_label.setText("Complete!"); self.status_label.setStyleSheet(f"color:{C.GREEN};")
        self.done_lbl.setText("Done!"); self.done_lbl.setStyleSheet(f"color:{C.GREEN};font-size:16px;font-weight:bold;")
        self.stop_btn.setEnabled(False)
        if self._out: self.open_btn.setVisible(True)


# ─── Main Window ──────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"GmailDownloader v{VERSION}")
        self.setMinimumSize(1000, 650); self.resize(1200, 800)

        # Restore window state
        self._settings = QSettings("GmailDownloader", "GmailDownloader")
        geo = self._settings.value("geometry")
        if geo: self.restoreGeometry(geo)
        state = self._settings.value("windowState")
        if state: self.restoreState(state)

        self.stack = QStackedWidget(); self.setCentralWidget(self.stack)
        self.cp = ConnectPage(); self.dp = DownloadPage()
        self.ap = AnalyzePage(); self.rp = ReviewPage(); self.ep = ExecutePage()
        for p in (self.cp, self.dp, self.ap, self.rp, self.ep): self.stack.addWidget(p)
        self._dl_dir = ""
        self.cp.connected.connect(self._connected)
        self.dp.download_complete.connect(self._dl_done)
        self.ap.analysis_complete.connect(self._review)
        self.rp.execute_requested.connect(self._execute)

        # Restore last email
        last_email = self._settings.value("last_email", "")
        if last_email: self.cp.email_input.setText(last_email)

    def closeEvent(self, event):
        self._settings.setValue("geometry", self.saveGeometry())
        self._settings.setValue("windowState", self.saveState())
        self._settings.setValue("last_email", self.cp.email_input.text())
        super().closeEvent(event)

    def _connected(self, mode):
        if mode == "load":
            self._dl_dir = self.cp.download_dir
            self.ap.set_preloaded(self.cp.loaded_engine); self.stack.setCurrentWidget(self.ap)
        elif mode == "download":
            self._dl_dir = self.cp.download_dir; self.stack.setCurrentWidget(self.dp)
            options = self.cp.sync_options()
            if self.cp.backend == 'gmail_api':
                source = GmailApiSource(self.cp.access_token, 'me')
                self.dp.start_remote_download(source, self._dl_dir, options)
            elif self.cp.backend == 'graph':
                source = GraphMailSource(self.cp.access_token)
                self.dp.start_remote_download(source, self._dl_dir, options)
            else:
                self.dp.start_download(
                    self.cp.imap_host, self.cp.email_addr, self.cp.password, self._dl_dir,
                    options, self.cp.imap_port, self.cp.use_ssl,
                    self.cp.auth_mode, self.cp.access_token
                )
        else:
            self.stack.setCurrentWidget(self.ap)
            since = self.cp.sync_options().since
            if self.cp.backend == 'gmail_api':
                self.ap.start_remote_scan(GmailApiSource(self.cp.access_token, 'me'),
                                          f"after:{since.strftime('%Y/%m/%d')}" if since else '')
            elif self.cp.backend == 'graph':
                self.ap.start_remote_scan(GraphMailSource(self.cp.access_token))
            else:
                self.ap.start_scan(
                    self.cp.imap_host, self.cp.email_addr, self.cp.password,
                    self.cp.imap_port, self.cp.use_ssl, self.cp.auth_mode,
                    self.cp.access_token, since
                )

    def _dl_done(self):
        self.ap.set_preloaded(self.dp.engine); self.stack.setCurrentWidget(self.ap)

    def _review(self):
        eng = self.ap.engine
        has_local = any(em.local_path for em in eng.emails)
        self.rp.load_categories(eng, has_local, self._dl_dir); self.stack.setCurrentWidget(self.rp)

    def _execute(self):
        eng = self.rp.engine
        if not eng: return
        cats = {k: v for k, v in eng.categories.items() if v and k != "Uncategorized"}
        if not cats: QMessageBox.warning(self, "Nothing", "No categories."); return
        if self.rp.mode_local.isChecked():
            if not any(em.local_path and Path(em.local_path).exists() for es in cats.values() for em in es):
                QMessageBox.warning(self, "No Files", "Download mailbox first."); return
            self.stack.setCurrentWidget(self.ep)
            self.ep.start_local(cats, self._dl_dir or str(Path.home()/"Desktop"/"GmailDownloader"),
                self.rp.copy_radio.isChecked())
        else:
            if self.cp.backend != 'imap':
                QMessageBox.warning(self, "Unavailable", "Gmail label application is currently available through IMAP only.")
                return
            if not self.cp.email_addr or not self.cp.password:
                QMessageBox.warning(self, "Credentials", "Enter Gmail credentials."); return
            dry_run = self.rp.dry_run_chk.isChecked()
            if not dry_run:
                answer = QMessageBox.question(
                    self, "Confirm mailbox changes",
                    "This will create Gmail labels and may archive messages from Inbox. Continue?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
            self.stack.setCurrentWidget(self.ep)
            self.ep.start_gmail(self.cp.imap_host, self.cp.email_addr, self.cp.password,
                cats, self.rp.prefix_input.text().strip(), self.rp.archive_chk.isChecked(), dry_run,
                self.cp.imap_port, self.cp.use_ssl, self.cp.auth_mode, self.cp.access_token)


def build_cli_parser():
    parser = argparse.ArgumentParser(description='GmailDownloader backup and archive tools')
    parser.add_argument('--headless', action='store_true', help='run without starting the Qt window')
    parser.add_argument('--sync', action='store_true', help='run an incremental source sync')
    parser.add_argument('--source', choices=('imap', 'gmail-api', 'graph'), default='imap')
    parser.add_argument('--output-dir', default=os.environ.get('GMAIL_OUTPUT_DIR', 'GmailDownloader'))
    parser.add_argument('--address', default=os.environ.get('GMAIL_ADDRESS', ''))
    parser.add_argument('--secret', default=os.environ.get('GMAIL_APP_PASSWORD', ''))
    parser.add_argument('--access-token', default=os.environ.get('GMAIL_ACCESS_TOKEN', ''))
    parser.add_argument('--host', default=os.environ.get('GMAIL_IMAP_HOST', 'imap.gmail.com'))
    parser.add_argument('--port', type=int, default=int(os.environ.get('GMAIL_IMAP_PORT', '993')))
    parser.add_argument('--no-ssl', action='store_true')
    parser.add_argument('--since', help='only sync messages on or after YYYY-MM-DD')
    parser.add_argument('--full-sync', action='store_true', help='ignore the last incremental UID')
    parser.add_argument('--no-verify', action='store_true', help='skip local manifest hash checks')
    parser.add_argument('--attachments-only', action='store_true')
    parser.add_argument('--query', default='')
    parser.add_argument('--import-mbox')
    parser.add_argument('--import-thunderbird')
    parser.add_argument('--load-dir', help='load an existing downloaded archive')
    parser.add_argument('--search', help='search imported messages using Gmail-like operators')
    parser.add_argument('--export-json')
    parser.add_argument('--export-markdown')
    parser.add_argument('--export-pdf')
    parser.add_argument('--export-mbox')
    parser.add_argument('--export-graph')
    parser.add_argument('--export-receipts-ofx')
    parser.add_argument('--export-location-timeline')
    return parser


def _unique_emails(emails):
    seen, result = set(), []
    for em in emails:
        if em.message_id and em.message_id in seen:
            continue
        if em.message_id:
            seen.add(em.message_id)
        result.append(em)
    return result


def run_headless(args):
    output_dir = Path(args.output_dir)
    emails = []
    errors = []
    if args.sync:
        options = SyncOptions(
            since=parse_since_date(args.since),
            incremental=not args.full_sync,
            verify_integrity=not args.no_verify,
            attachments_only=args.attachments_only,
        )
        if args.source == 'gmail-api':
            token = args.access_token or args.secret
            worker = GmailApiDownloadWorker(GmailApiSource(token, 'me'), output_dir, options, args.query)
        elif args.source == 'graph':
            token = args.access_token or args.secret
            worker = RemoteMimeDownloadWorker(GraphMailSource(token), output_dir, 'Microsoft Graph', options, args.query)
        else:
            secret = args.access_token if args.access_token and args.secret == '' else args.secret
            worker = ImapDownloadWorker(
                args.host, args.address, secret, output_dir, options=options,
                port=args.port, use_ssl=not args.no_ssl,
                auth_mode='oauth2' if args.access_token else 'password',
                access_token=args.access_token,
            )
        worker.status.connect(print)
        if hasattr(worker, 'log'):
            worker.log.connect(print)
        worker.error.connect(errors.append)
        worker.finished_signal.connect(emails.extend)
        worker.run()
    elif args.import_mbox:
        emails = import_mbox(args.import_mbox, output_dir, attachments_only=args.attachments_only)
    elif args.import_thunderbird:
        emails = import_thunderbird_profile(args.import_thunderbird, output_dir)
    elif args.load_dir:
        emails, issues = load_archive_emails(args.load_dir, not args.no_verify)
        for issue in issues:
            print(f"Manifest warning: {issue}", file=sys.stderr)
    else:
        raise ValueError('headless mode requires --sync, --import-mbox, --import-thunderbird, or --load-dir')
    if errors:
        for error in errors:
            print(error, file=sys.stderr)
        return 1
    emails = _unique_emails(emails)
    domain = args.address.split('@', 1)[1] if '@' in args.address else ''
    engine = CategoryEngine(domain)
    engine.process_all(emails)
    summary = engine.get_summary()
    print(f"{summary['total']:,} emails | {summary['categorized']:,} categorized | {summary['uncategorized']:,} uncategorized")
    if args.search:
        print(json.dumps([email_info_to_record(em) for em in engine.search(args.search)], ensure_ascii=False, indent=2))
    if args.export_json:
        engine.export_json(args.export_json)
    if args.export_markdown:
        engine.export_markdown(args.export_markdown)
    if args.export_pdf:
        engine.export_pdf(args.export_pdf)
    if args.export_mbox:
        engine.export_mbox(args.export_mbox)
    if args.export_graph:
        engine.export_relationship_graph(args.export_graph)
    if args.export_receipts_ofx:
        engine.export_receipts_ofx(args.export_receipts_ofx)
    if args.export_location_timeline:
        engine.export_location_timeline(args.export_location_timeline)
    return 0


def main(argv=None):
    args = build_cli_parser().parse_args(argv)
    if args.headless:
        try:
            return run_headless(args)
        except Exception as exc:
            print(f'Headless operation failed: {exc}', file=sys.stderr)
            return 1
    app = QApplication(sys.argv if argv is None else [sys.argv[0], *argv])
    app.setStyleSheet(STYLESHEET); app.setStyle("Fusion")
    branding_icon = QIcon(str(_branding_icon_path()))
    app.setWindowIcon(branding_icon)
    window = MainWindow(); window.show()
    return app.exec()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(main())
