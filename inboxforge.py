#!/usr/bin/env python3
"""InboxForge v1.0.0 — Full Gmail Mailbox Downloader, AI Organizer & Analytics Suite"""

VERSION = "1.0.0"

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
import csv
import io
import json
import shutil
import hashlib
import webbrowser
import traceback
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from dataclasses import dataclass, field
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

    def save_state(self, path: str):
        data = {'version': VERSION, 'user_domain': self.user_domain,
                'emails': [{'uid': em.uid, 'sender': em.sender, 'sender_name': em.sender_name,
                    'sender_domain': em.sender_domain, 'subject': em.subject,
                    'date': em.date, 'has_list_unsubscribe': em.has_list_unsubscribe,
                    'list_unsubscribe_url': em.list_unsubscribe_url,
                    'category': em.category, 'confidence': em.confidence,
                    'local_path': em.local_path, 'source_folder': em.source_folder,
                    'message_id': em.message_id, 'in_reply_to': em.in_reply_to,
                    'references': em.references, 'size_bytes': em.size_bytes,
                    'sensitive_flags': em.sensitive_flags, 'is_newsletter': em.is_newsletter}
                   for em in self.emails]}
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
                em = EmailInfo(uid=ed['uid'], sender=ed.get('sender',''),
                    sender_name=ed.get('sender_name',''), sender_domain=ed.get('sender_domain',''),
                    subject=ed.get('subject',''), date=ed.get('date',''),
                    date_parsed=parse_date(ed.get('date','')),
                    has_list_unsubscribe=ed.get('has_list_unsubscribe',False),
                    list_unsubscribe_url=ed.get('list_unsubscribe_url',''),
                    category=ed.get('category',''), confidence=ed.get('confidence',0),
                    local_path=ed.get('local_path',''), source_folder=ed.get('source_folder',''),
                    message_id=ed.get('message_id',''), in_reply_to=ed.get('in_reply_to',''),
                    references=ed.get('references',''), size_bytes=ed.get('size_bytes',0),
                    sensitive_flags=ed.get('sensitive_flags',[]),
                    is_newsletter=ed.get('is_newsletter',False))
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


# ─── IMAP Workers ─────────────────────────────────────────────────────────

class ImapScanWorker(QThread):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    email_batch = pyqtSignal(list)
    finished_signal = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, host, addr, pw):
        super().__init__()
        self.host, self.addr, self.pw = host, addr, pw
        self._stop = False

    def stop(self): self._stop = True

    def run(self):
        try:
            self.status.emit("Connecting...")
            imap = imaplib.IMAP4_SSL(self.host, 993)
            imap.login(self.addr, self.pw)
            imap.select('INBOX', readonly=True)
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
                            msg = email.message_from_bytes(item[1])
                            fd = decode_header(msg.get('From',''))
                            name, addr = email.utils.parseaddr(fd)
                            unsub = msg.get('List-Unsubscribe','')
                            em = EmailInfo(uid=uid, sender=addr or fd, sender_name=name or addr or fd,
                                subject=decode_header(msg.get('Subject','(no subject)')),
                                date=msg.get('Date',''), date_parsed=parse_date(msg.get('Date','')),
                                has_list_unsubscribe=bool(unsub), list_unsubscribe_url=unsub,
                                source_folder='INBOX', message_id=msg.get('Message-ID',''),
                                in_reply_to=msg.get('In-Reply-To',''),
                                references=msg.get('References',''), size_bytes=size)
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

    def __init__(self, host, addr, pw, output_dir, skip=None):
        super().__init__()
        self.host, self.addr, self.pw = host, addr, pw
        self.output_dir = Path(output_dir)
        self.skip = skip or GMAIL_SKIP_FOLDERS
        self._stop = False

    def stop(self): self._stop = True

    def run(self):
        try:
            folders_dir = self.output_dir / "folders"
            folders_dir.mkdir(parents=True, exist_ok=True)
            manifest_path = self.output_dir / "manifest.json"
            manifest = {'folders': {}, 'message_ids': {}}
            if manifest_path.exists():
                try:
                    with open(manifest_path, 'r', encoding='utf-8') as f:
                        manifest = json.load(f)
                    if 'folders' not in manifest:
                        manifest = {'folders': {}, 'message_ids': {}}
                except: manifest = {'folders': {}, 'message_ids': {}}
            seen_ids = manifest.get('message_ids', {})

            self.status.emit("Connecting...")
            imap = imaplib.IMAP4_SSL(self.host, 993)
            imap.login(self.addr, self.pw)
            _, folder_data = imap.list()
            all_folders = []
            for line in folder_data:
                if isinstance(line, bytes):
                    n = parse_imap_folder_list(line)
                    if n: all_folders.append(n)
            active = [f for f in all_folders if f not in self.skip]
            self.log.emit(f"Downloading {len(active)} folders, skipping {len(self.skip)}")

            folder_uids = {}
            total_count = 0
            for fn in active:
                if self._stop: break
                try:
                    if imap.select(f'"{fn}"', readonly=True)[0] != 'OK': continue
                    _, d = imap.uid('SEARCH', None, 'ALL')
                    u = d[0].split() if d[0] else []
                    folder_uids[fn] = u; total_count += len(u)
                    self.log.emit(f"  {fn}: {len(u):,}")
                except Exception as e:
                    self.log.emit(f"  {fn}: error - {e}")

            all_emails = []
            gp = 0
            for fn, fm in manifest.get('folders', {}).items():
                for uid_str, info in fm.items():
                    em = EmailInfo(uid=f"{fn}:{uid_str}", sender=info.get('sender',''),
                        sender_name=info.get('sender_name',''), subject=info.get('subject',''),
                        date=info.get('date',''), date_parsed=parse_date(info.get('date','')),
                        has_list_unsubscribe=info.get('has_list_unsubscribe',False),
                        list_unsubscribe_url=info.get('list_unsubscribe_url',''),
                        local_path=info.get('local_path',''), source_folder=fn,
                        message_id=info.get('message_id',''),
                        in_reply_to=info.get('in_reply_to',''),
                        references=info.get('references',''),
                        size_bytes=info.get('size_bytes',0))
                    all_emails.append(em)

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
                    except: gp += len(bu); continue
                    for item in md:
                        if self._stop: break
                        if not isinstance(item, tuple) or len(item) != 2: continue
                        um = re.search(rb'UID (\d+)', item[0])
                        if not um: continue
                        uid = um.group(1).decode()
                        raw = item[1]
                        if not isinstance(raw, bytes): continue
                        try:
                            msg = email.message_from_bytes(raw)
                            mid = msg.get('Message-ID','')
                            fd = decode_header(msg.get('From',''))
                            nm, ad = email.utils.parseaddr(fd)
                            unsub = msg.get('List-Unsubscribe','')
                            irt = msg.get('In-Reply-To','')
                            refs = msg.get('References','')
                        except:
                            mid=ad=nm=fd=unsub=irt=refs=''; raw=item[1]
                        if mid and mid in seen_ids:
                            eml_path = seen_ids[mid]
                        else:
                            eml_path = str(fdir / f"{uid}.eml")
                            try:
                                with open(eml_path, 'wb') as f: f.write(raw)
                            except: continue
                            if mid: seen_ids[mid] = eml_path
                        em = EmailInfo(uid=f"{fn}:{uid}", sender=ad or fd,
                            sender_name=nm or ad or fd,
                            subject=decode_header(msg.get('Subject','') if 'msg' in dir() else ''),
                            date=msg.get('Date','') if 'msg' in dir() else '',
                            date_parsed=parse_date(msg.get('Date','') if 'msg' in dir() else ''),
                            has_list_unsubscribe=bool(unsub), list_unsubscribe_url=unsub,
                            local_path=eml_path, source_folder=fn, message_id=mid,
                            in_reply_to=irt, references=refs,
                            size_bytes=len(raw))
                        all_emails.append(em)
                        self.email_saved.emit(em)
                        if fn not in manifest['folders']: manifest['folders'][fn] = {}
                        manifest['folders'][fn][uid] = {
                            'sender': em.sender, 'sender_name': em.sender_name,
                            'subject': em.subject, 'date': em.date,
                            'has_list_unsubscribe': em.has_list_unsubscribe,
                            'list_unsubscribe_url': em.list_unsubscribe_url,
                            'local_path': eml_path, 'message_id': mid,
                            'in_reply_to': irt, 'references': refs,
                            'size_bytes': em.size_bytes}
                    gp += len(bu)
                    self.progress.emit(gp, total_count)
                manifest['message_ids'] = seen_ids
                try:
                    with open(manifest_path, 'w', encoding='utf-8') as f:
                        json.dump(manifest, f, ensure_ascii=False)
                except: pass
            try: imap.logout()
            except: pass
            self.finished_signal.emit(all_emails)
        except Exception as e:
            self.error.emit(f"Error: {e}\n{traceback.format_exc()}")


class ImapLabelWorker(QThread):
    progress = pyqtSignal(int, int); status = pyqtSignal(str)
    log = pyqtSignal(str); finished_signal = pyqtSignal(); error = pyqtSignal(str)
    def __init__(self, host, addr, pw, cats, prefix="", archive=False):
        super().__init__()
        self.host, self.addr, self.pw = host, addr, pw
        self.cats, self.prefix, self.archive = cats, prefix, archive
        self._stop = False
    def stop(self): self._stop = True
    def run(self):
        try:
            imap = imaplib.IMAP4_SSL(self.host, 993); imap.login(self.addr, self.pw)
            total = sum(len(v) for v in self.cats.values()); done = 0
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


class ConnectionTester(QObject):
    success = pyqtSignal(int); error = pyqtSignal(str)
    def __init__(self, host, addr, pw):
        super().__init__(); self.host, self.addr, self.pw = host, addr, pw
    def run(self):
        try:
            imap = imaplib.IMAP4_SSL(self.host, 993); imap.login(self.addr, self.pw)
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


# ─── UI: Connect Page ────────────────────────────────────────────────────

class ConnectPage(QWidget):
    connected = pyqtSignal(str)
    def __init__(self):
        super().__init__()
        self.imap_host = "imap.gmail.com"
        self.email_addr = self.password = self.download_dir = ""
        self.loaded_engine = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.setSpacing(16)
        title = QLabel("InboxForge")
        title.setStyleSheet(f"font-size:32px; color:{C.BLUE}; font-weight:bold;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(title)
        sub = QLabel(f"v{VERSION} — Full Gmail Mailbox Downloader, AI Organizer & Analytics")
        sub.setStyleSheet(f"color:{C.SUBTEXT0}; font-size:12px;")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(sub)
        layout.addSpacing(20)

        form = QWidget(); form.setMaximumWidth(520); fl = QVBoxLayout(form); fl.setSpacing(12)
        fl.addWidget(QLabel("Gmail Address"))
        self.email_input = QLineEdit(); self.email_input.setPlaceholderText("you@gmail.com")
        fl.addWidget(self.email_input)
        fl.addWidget(QLabel("App Password"))
        self.pass_input = QLineEdit(); self.pass_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.pass_input.setPlaceholderText("16-character app password"); fl.addWidget(self.pass_input)
        hint = QLabel("Google Account > Security > 2-Step Verification > App Passwords")
        hint.setStyleSheet(f"color:{C.SUBTEXT0}; font-size:11px;"); hint.setWordWrap(True)
        fl.addWidget(hint); fl.addSpacing(12)

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

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet(f"color:{C.SUBTEXT0}; font-size:12px;")
        self.status_label.setWordWrap(True); fl.addWidget(self.status_label)
        layout.addWidget(form, alignment=Qt.AlignmentFlag.AlignCenter); layout.addStretch()

    def _set_btns(self, en):
        for b in (self.scan_btn, self.dl_btn, self.load_scan_btn, self.load_local_btn): b.setEnabled(en)

    def _on_action(self, mode):
        self.email_addr = self.email_input.text().strip()
        self.password = self.pass_input.text().strip()
        if not self.email_addr or not self.password:
            self.status_label.setText("Enter email and app password")
            self.status_label.setStyleSheet(f"color:{C.RED};"); return
        if mode == "download":
            f = QFileDialog.getExistingDirectory(self, "Download Folder",
                str(Path.home() / "Desktop" / "InboxForge"))
            if not f: return
            self.download_dir = f
        self.status_label.setText("Testing..."); self.status_label.setStyleSheet(f"color:{C.YELLOW};")
        self._set_btns(False); self._mode = mode
        self._tt = QThread(); self._tw = ConnectionTester(self.imap_host, self.email_addr, self.password)
        self._tw.moveToThread(self._tt); self._tt.started.connect(self._tw.run)
        self._tw.success.connect(self._ok); self._tw.error.connect(self._fail); self._tt.start()

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


# ─── UI: Download Page ───────────────────────────────────────────────────

class DownloadPage(QWidget):
    download_complete = pyqtSignal()
    def __init__(self):
        super().__init__(); self.engine = self.worker = None; self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self); layout.setSpacing(12)
        QLabel("Downloading Full Mailbox", self).setStyleSheet("font-size:20px;font-weight:bold;")
        layout.addWidget(self.findChild(QLabel))
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

    def start_download(self, host, addr, pw, out_dir):
        ud = addr.split('@')[1] if '@' in addr else ""
        self.engine = CategoryEngine(ud); self._out = out_dir; self._n = 0
        self.worker = ImapDownloadWorker(host, addr, pw, out_dir)
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

    def start_scan(self, host, addr, pw):
        ud = addr.split('@')[1] if '@' in addr else ""
        self.engine = CategoryEngine(ud); self._dc = Counter()
        self.worker = ImapScanWorker(host, addr, pw)
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
        p, _ = QFileDialog.getSaveFileName(self, "Save", "inboxforge_scan.json", "JSON (*.json)")
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

        # Toolbar
        tb = QHBoxLayout()
        t = QLabel("Review Categories"); t.setStyleSheet("font-size:20px;font-weight:bold;"); tb.addWidget(t)
        tb.addStretch()
        self.summary_lbl = QLabel(""); self.summary_lbl.setStyleSheet(f"color:{C.SUBTEXT0};"); tb.addWidget(self.summary_lbl)

        # Toolbar buttons
        self.stats_btn = QPushButton("Stats"); self.stats_btn.setProperty("secondary", True)
        self.stats_btn.clicked.connect(self._show_stats); tb.addWidget(self.stats_btn)
        self.subs_btn = QPushButton("Subscriptions"); self.subs_btn.setProperty("secondary", True)
        self.subs_btn.clicked.connect(self._show_subs); tb.addWidget(self.subs_btn)
        self.rules_btn = QPushButton("Rules"); self.rules_btn.setProperty("secondary", True)
        self.rules_btn.clicked.connect(self._show_rules); tb.addWidget(self.rules_btn)

        # Export menu
        self.export_btn = QPushButton("Export"); self.export_btn.setProperty("secondary", True)
        export_menu = QMenu(self)
        export_menu.addAction("CSV", self._export_csv)
        export_menu.addAction("JSON", self._export_json)
        self.export_btn.setMenu(export_menu); tb.addWidget(self.export_btn)

        # Scan buttons
        self.attach_btn = QPushButton("Extract Attachments"); self.attach_btn.setProperty("secondary", True)
        self.attach_btn.clicked.connect(self._extract_attachments); tb.addWidget(self.attach_btn)
        self.sensitive_btn = QPushButton("Scan Sensitive"); self.sensitive_btn.setProperty("secondary", True)
        self.sensitive_btn.clicked.connect(self._scan_sensitive); tb.addWidget(self.sensitive_btn)
        self.ai_btn = QPushButton("AI Classify"); self.ai_btn.setProperty("secondary", True)
        self.ai_btn.clicked.connect(self._ai_classify); tb.addWidget(self.ai_btn)
        self.thread_btn = QPushButton("Summarize Threads"); self.thread_btn.setProperty("secondary", True)
        self.thread_btn.clicked.connect(self._summarize_threads); tb.addWidget(self.thread_btn)
        layout.addLayout(tb)

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

        # Right: email table
        right = QWidget(); rl = QVBoxLayout(right); rl.setContentsMargins(0,0,0,0)
        self.email_count_lbl = QLabel(""); self.email_count_lbl.setStyleSheet(f"color:{C.SUBTEXT0};font-size:12px;")
        rl.addWidget(self.email_count_lbl)
        self.table = QTableWidget(); self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["From", "Subject", "Date", "Folder", "Conf", "Flags"])
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for c in (2,3,4,5): h.setSectionResizeMode(c, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setColumnWidth(0, 170)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False); rl.addWidget(self.table)

        move_row = QHBoxLayout(); move_row.addWidget(QLabel("Move to:"))
        self.move_combo = QComboBox(); self.move_combo.setMinimumWidth(200)
        move_row.addWidget(self.move_combo, 1)
        mv_btn = QPushButton("Move"); mv_btn.setProperty("secondary", True)
        mv_btn.clicked.connect(self._move_emails); move_row.addWidget(mv_btn)
        rl.addLayout(move_row)
        splitter.addWidget(right); splitter.setSizes([350, 650])
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
        self._mode_changed()

    def _mode_changed(self):
        g = self.mode_gmail.isChecked()
        self.gmail_opts.setVisible(g); self.local_opts.setVisible(not g)
        self.exec_btn.setText("Apply Gmail Labels" if g else "Organize Local Files")
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
        self._show_emails(emails)

    def _show_emails(self, emails):
        self.email_count_lbl.setText(f"{len(emails):,} emails")
        display = sorted(emails, key=lambda e: e.date_parsed or datetime.min, reverse=True)[:2000]
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

    def _show_subs(self):
        if self.engine: SubscriptionDialog(self.engine.subscriptions, self).exec()

    def _show_rules(self):
        if self.engine:
            RulesEditorDialog(self.engine.clean_rules, list(self.engine.categories.keys()), self).exec()

    def _export_csv(self):
        if not self.engine: return
        p, _ = QFileDialog.getSaveFileName(self, "Export CSV", "inboxforge_export.csv", "CSV (*.csv)")
        if p: self.engine.export_csv(p)

    def _export_json(self):
        if not self.engine: return
        p, _ = QFileDialog.getSaveFileName(self, "Export JSON", "inboxforge_export.json", "JSON (*.json)")
        if p: self.engine.export_json(p)

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
        self._ai_w = AiClassifyWorker(key, uncat, existing)
        self._ai_w.classified.connect(self._ai_result)
        self._ai_w.finished_signal.connect(self._ai_done)
        self._ai_w.error.connect(lambda e: (self.ai_btn.setEnabled(True), self.ai_btn.setText("AI Classify")))
        self._ai_w.start()

    def _ai_result(self, dmap):
        for domain, cat in dmap.items():
            moves = [e for e in self.engine.categories.get("Uncategorized", []) if e.sender_domain == domain]
            for em in moves:
                self.engine.categories["Uncategorized"].remove(em)
                em.category = cat; em.confidence = 0.75
                self.engine.categories[cat].append(em)
                self.engine.learned.learn(em, cat)
        if not self.engine.categories.get("Uncategorized"):
            if "Uncategorized" in self.engine.categories: del self.engine.categories["Uncategorized"]
        self.engine.learned.save()
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

    def start_gmail(self, host, addr, pw, cats, prefix, archive):
        self.title_lbl.setText("Applying Gmail Labels")
        self.worker = ImapLabelWorker(host, addr, pw, cats, prefix, archive)
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
        self.setWindowTitle(f"InboxForge v{VERSION}")
        self.setMinimumSize(1000, 650); self.resize(1200, 800)
        self.stack = QStackedWidget(); self.setCentralWidget(self.stack)
        self.cp = ConnectPage(); self.dp = DownloadPage()
        self.ap = AnalyzePage(); self.rp = ReviewPage(); self.ep = ExecutePage()
        for p in (self.cp, self.dp, self.ap, self.rp, self.ep): self.stack.addWidget(p)
        self._dl_dir = ""
        self.cp.connected.connect(self._connected)
        self.dp.download_complete.connect(self._dl_done)
        self.ap.analysis_complete.connect(self._review)
        self.rp.execute_requested.connect(self._execute)

    def _connected(self, mode):
        if mode == "load":
            self._dl_dir = self.cp.download_dir
            self.ap.set_preloaded(self.cp.loaded_engine); self.stack.setCurrentWidget(self.ap)
        elif mode == "download":
            self._dl_dir = self.cp.download_dir; self.stack.setCurrentWidget(self.dp)
            self.dp.start_download(self.cp.imap_host, self.cp.email_addr, self.cp.password, self._dl_dir)
        else:
            self.stack.setCurrentWidget(self.ap)
            self.ap.start_scan(self.cp.imap_host, self.cp.email_addr, self.cp.password)

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
            self.ep.start_local(cats, self._dl_dir or str(Path.home()/"Desktop"/"InboxForge"),
                self.rp.copy_radio.isChecked())
        else:
            if not self.cp.email_addr or not self.cp.password:
                QMessageBox.warning(self, "Credentials", "Enter Gmail credentials."); return
            self.stack.setCurrentWidget(self.ep)
            self.ep.start_gmail(self.cp.imap_host, self.cp.email_addr, self.cp.password,
                cats, self.rp.prefix_input.text().strip(), self.rp.archive_chk.isChecked())


def main():
    app = QApplication(sys.argv); app.setStyleSheet(STYLESHEET); app.setStyle("Fusion")
    w = MainWindow(); w.show(); sys.exit(app.exec())

if __name__ == "__main__":
    main()
