import email
import base64
import json
import tempfile
from datetime import datetime
from pathlib import Path
from email.message import EmailMessage

import pytest

import gmaildownloader as app


def test_manifest_hashes_detect_tampering():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        message = root / "message.eml"
        message.write_bytes(b"Subject: test\n\nbody\n")
        manifest = app.new_manifest()
        manifest["folders"] = {"INBOX": {"1": {"local_path": str(message)}}}
        manifest_path = root / "manifest.json"
        app.save_manifest(manifest_path, manifest)

        loaded = app.load_manifest(manifest_path)
        assert app.validate_manifest(loaded, root, update_missing=True) == []
        assert loaded["folders"]["INBOX"]["1"]["sha256"]

        message.write_bytes(b"tampered")
        issues = app.validate_manifest(loaded, root)
        assert issues[0]["reason"] == "sha256 mismatch"


def test_parse_and_redact_email_without_touching_original():
    raw = (
        b"From: Alice <alice@example.com>\n"
        b"Subject: Credentials\n"
        b"Date: Mon, 01 Jan 2024 00:00:00 +0000\n"
        b"Content-Type: text/plain; charset=utf-8\n\n"
        b"password: super-secret-value\n"
    )
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source = root / "source.eml"
        redacted = root / "redacted.eml"
        source.write_bytes(raw)
        info = app.parse_email_message(raw, "1", "INBOX")
        assert info.sender == "alice@example.com"
        assert info.subject == "Credentials"
        flags = app.redact_eml(source, redacted)
        assert "Password" in flags
        assert b"super-secret-value" in source.read_bytes()
        assert b"super-secret-value" not in redacted.read_bytes()


def test_mbox_round_trip():
    raw = b"From: sender@example.com\nSubject: Hello\n\nMessage\n"
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source = root / "takeout.mbox"
        box = app.mailbox.mbox(str(source))
        box.add(email.message_from_bytes(raw))
        box.flush()
        box.close()
        imported = app.import_mbox(source, root / "archive")
        assert len(imported) == 1
        assert imported[0].local_path.endswith(".eml")

        exported = root / "roundtrip.mbox"
        app.export_mbox(imported, exported)
        roundtrip = app.import_mbox(exported)
        assert len(roundtrip) == 1
        assert roundtrip[0].subject == "Hello"


def test_xoauth2_payload():
    assert app.build_xoauth2_string("me@example.com", "token") == (
        b"user=me@example.com\x01auth=Bearer token\x01\x01"
    )


class FakeImap:
    def __init__(self):
        self.searches = []
        self.fetches = 0

    def list(self):
        return "OK", [b'(\\HasNoChildren) "/" "INBOX"']

    def select(self, *_args, **_kwargs):
        return "OK", [b"2"]

    def response(self, name):
        assert name == "UIDVALIDITY"
        return "OK", [b"42"]

    def uid(self, command, *_args):
        if command == "SEARCH":
            self.searches.append(_args)
            return "OK", [b"1 2"]
        if command == "FETCH":
            self.fetches += 1
            requested = _args[0].decode().split(",")
            messages = []
            for uid in requested:
                raw = (
                    f"From: sender{uid}@example.com\n"
                    f"Subject: Message {uid}\n"
                    "Date: Mon, 01 Jan 2024 00:00:00 +0000\n\n"
                    f"Body {uid}\n"
                ).encode()
                messages.append((f"* 1 FETCH (UID {uid} RFC822 {{{len(raw)}}})".encode(), raw))
            return "OK", messages
        raise AssertionError(command)

    def logout(self):
        return "OK", []


def test_incremental_sync_and_since_query(monkeypatch):
    fake = FakeImap()
    monkeypatch.setattr(app, "open_imap_connection", lambda *args, **kwargs: fake)
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        completed = []
        worker = app.ImapDownloadWorker("imap.test", "me@example.com", "secret", root)
        worker.finished_signal.connect(completed.append)
        worker.run()
        assert len(completed[0]) == 2
        assert fake.fetches == 1

        worker = app.ImapDownloadWorker("imap.test", "me@example.com", "secret", root)
        worker.run()
        assert fake.fetches == 1
        assert app.load_manifest(root / "manifest.json")["folder_metadata"]["INBOX"]["last_uid"] == 2

    fake = FakeImap()
    monkeypatch.setattr(app, "open_imap_connection", lambda *args, **kwargs: fake)
    with tempfile.TemporaryDirectory() as directory:
        worker = app.ImapDownloadWorker(
            "imap.test", "me@example.com", "secret", directory,
            options=app.SyncOptions(since=datetime(2024, 1, 2), incremental=False),
        )
        worker.run()
        assert fake.searches[0][1:] == ("SINCE", "02-Jan-2024")


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self):
        return self.payload


def test_gmail_api_pagination_and_raw_decode():
    raw = b"From: api@example.com\nSubject: API\n\nBody\n"
    encoded = base64.urlsafe_b64encode(raw).decode().rstrip("=")
    calls = []
    auth_headers = []

    def opener(request, timeout=0):
        calls.append((request.full_url, timeout))
        auth_headers.append(request.headers["Authorization"])
        if request.full_url.endswith("/messages?maxResults=100"):
            return FakeResponse(json.dumps({"messages": [{"id": "one"}], "nextPageToken": "next"}).encode())
        if "pageToken=next" in request.full_url:
            return FakeResponse(json.dumps({"messages": [{"id": "two"}]}).encode())
        if "/messages/one?format=raw" in request.full_url:
            return FakeResponse(json.dumps({"raw": encoded}).encode())
        if "/messages/two?format=raw" in request.full_url:
            return FakeResponse(json.dumps({"raw": encoded}).encode())
        raise AssertionError(request.full_url)

    source = app.GmailApiSource("access", opener=opener)
    messages = list(source.iter_messages())
    assert [message_id for message_id, _, _ in messages] == ["one", "two"]
    assert messages[0][1] == raw
    assert all(value == "Bearer access" for value in auth_headers)
    assert len(calls) == 4


def test_oauth_url_and_multi_account_isolation():
    client = app.GoogleOAuthClient("client-id")
    url, state = client.authorization_url("state-value")
    assert "client_id=client-id" in url
    assert "state=state-value" in url
    assert state == "state-value"

    with tempfile.TemporaryDirectory() as directory:
        manager = app.MultiAccountManager(directory)
        manager.add(app.AccountConfig("personal", "me@example.com"))
        manager.add(app.AccountConfig("work", "me@work.example"))
        assert manager.output_dir("personal") != manager.output_dir("work")
        config = Path(directory) / "accounts.json"
        manager.save_config(config)
        restored = app.MultiAccountManager(directory)
        restored.load_config(config)
        assert sorted(restored.accounts) == ["personal", "work"]


def test_remote_mime_worker_writes_manifest_and_eml():
    raw = b"From: remote@example.com\nSubject: Remote\n\nBody\n"

    class Source:
        def iter_messages(self, query=''):
            assert query == ''
            yield "remote-id", raw, ["INBOX"]

    with tempfile.TemporaryDirectory() as directory:
        completed = []
        worker = app.GmailApiDownloadWorker(Source(), directory)
        worker.finished_signal.connect(completed.append)
        worker.run()
        assert len(completed[0]) == 1
        assert Path(completed[0][0].local_path).exists()
        manifest = app.load_manifest(Path(directory) / "manifest.json")
        assert manifest["folders"]["Gmail API"]["remote-id"]["sha256"]


def test_search_analytics_and_exports():
    first = app.EmailInfo(
        uid="1", sender="alice@example.com", sender_name="Alice", subject="Invoice paid",
        date="Mon, 01 Jan 2024 00:00:00 +0000", date_parsed=datetime(2024, 1, 1),
        message_id="<one>", source_folder="INBOX", size_bytes=100,
    )
    second = app.EmailInfo(
        uid="2", sender="bob@example.com", sender_name="Bob", subject="Re: Invoice paid",
        date="Mon, 01 Jan 2024 02:00:00 +0000", date_parsed=datetime(2024, 1, 1, 2),
        message_id="<two>", in_reply_to="<one>", source_folder="Sent Mail", size_bytes=200,
    )
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        raw = b"From: alice@example.com\nSubject: Invoice paid\n\nTotal: $12.50\n"
        first.local_path = str(root / "one.eml")
        Path(first.local_path).write_bytes(raw)
        first.has_attachments = True
        engine = app.CategoryEngine("example.net")
        engine.process_all([first, second])
        assert engine.search("from:alice subject:invoice") == [first]
        assert engine.search("has:attachment") == [first]
        assert app.extract_receipts([first])[0]["amount"] == 12.5
        assert engine.reply_latency()["1-4h"] == 1
        assert engine.storage_forecast()["average_monthly_bytes"] == 300
        graph = engine.relationship_graph()
        assert graph["nodes"]
        assert engine.thread_clusters()
        notes = engine.export_markdown(root / "vault")
        assert len(notes) == 2
        graph_path = engine.export_relationship_graph(root / "graph.json")
        assert Path(graph_path).exists()
        pdf_path = engine.export_pdf(root / "emails.pdf")
        assert Path(pdf_path).stat().st_size > 0


def test_location_timeline_extracts_public_hops_and_resolver_data():
    raw = (
        b"From: travel@example.com\n"
        b"Subject: Travel\n"
        b"Received: from mx.example (8.8.8.8) by local.example; Tue, 02 Jan 2024 12:00:00 +0000\n"
        b"Received: from [2001:4860:4860::8888] by mx.example; Tue, 02 Jan 2024 11:59:00 +0000\n"
        b"Received: from [192.168.1.4] by mx.example; Tue, 02 Jan 2024 11:58:00 +0000\n"
        b"\nBody\n"
    )
    with tempfile.TemporaryDirectory() as directory:
        source = Path(directory) / "travel.eml"
        source.write_bytes(raw)
        message = app.parse_email_message(raw, "travel", "INBOX", str(source))
        timeline = app.location_timeline(
            [message], resolver=lambda ip: {"country": "US", "city": ip}
        )
        assert [item["ip"] for item in timeline] == ["2001:4860:4860::8888", "8.8.8.8"]
        assert all(item["country"] == "US" for item in timeline)
        assert timeline[0]["received_at"].startswith("2024-01-02T11:59:00")


def test_receipt_vision_renders_pdf_and_exports_ofx():
    pytest.importorskip("pypdfium2")
    canvas = pytest.importorskip("reportlab.pdfgen.canvas")
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        pdf = root / "invoice.pdf"
        document = canvas.Canvas(str(pdf))
        document.drawString(72, 720, "ACME total $19.95")
        document.save()

        pages = app.render_pdf_pages(pdf, root / "pages")
        assert len(pages) == 1
        assert pages[0].suffix == ".png"

        def fake_classifier(image_path, ocr_text):
            assert image_path.suffix == ".png"
            return {
                "merchant": "ACME",
                "date": "2024-01-02",
                "total": "19.95",
                "currency": "USD",
                "line_items": [{"name": "Widget", "amount": 19.95}],
            }

        classifier = app.ReceiptVisionClassifier(image_classifier=fake_classifier)
        receipt = classifier.classify_attachment(pdf, uid="INBOX:1", sender="billing@acme.test")
        assert receipt["merchant"] == "ACME"
        assert receipt["amount"] == 19.95
        assert receipt["pages"] == 1
        ofx = root / "receipts.ofx"
        app.export_receipts_ofx([receipt], ofx)
        parsed = app.ET.parse(ofx)
        assert parsed.findtext(".//TRNAMT") == "-19.95"
        assert parsed.findtext(".//NAME") == "ACME"

        message = EmailMessage()
        message["From"] = "billing@acme.test"
        message["Subject"] = "Invoice"
        message.set_content("Please see attached invoice.")
        message.add_attachment(pdf.read_bytes(), maintype="application", subtype="pdf", filename="invoice.pdf")
        eml = root / "message.eml"
        eml.write_bytes(message.as_bytes())
        email_info = app.parse_email_message(eml.read_bytes(), "1", "INBOX", str(eml))
        extracted = app.extract_receipt_attachments([email_info], classifier)
        assert extracted[0]["attachment"] == "invoice.pdf"
        assert extracted[0]["uid"] == "1"


def test_headless_import_and_scheduler_helpers(capsys):
    raw = b"From: cli@example.com\nSubject: CLI\n\nBody\n"
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source = root / "source.mbox"
        box = app.mailbox.mbox(str(source)); box.add(email.message_from_bytes(raw)); box.flush(); box.close()
        output = root / "archive"
        exported = root / "export.json"
        args = app.build_cli_parser().parse_args([
            "--headless", "--import-mbox", str(source), "--output-dir", str(output),
            "--export-json", str(exported),
        ])
        assert app.run_headless(args) == 0
        assert exported.exists()
        assert "--headless --sync" in app.build_cron_entry("C:/app.py", "C:/archive")
        task = app.install_windows_scheduled_backup("Gmail", "C:/app.py", "C:/archive")
        assert task[0] == "schtasks"
    assert "1 emails" in capsys.readouterr().out


@pytest.mark.skipif(not hasattr(app, "encrypt_archive"), reason="archive helpers unavailable")
def test_encrypted_archive_round_trip():
    pytest.importorskip("cryptography")
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source = root / "source"
        source.mkdir()
        (source / "file.txt").write_text("private archive", encoding="utf-8")
        encrypted = root / "archive.gd"
        restored = root / "restored"
        app.encrypt_archive(source, encrypted, "correct horse battery staple")
        app.decrypt_archive(encrypted, restored, "correct horse battery staple")
        assert (restored / "file.txt").read_text(encoding="utf-8") == "private archive"
