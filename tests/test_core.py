import email
import tempfile
from datetime import datetime
from pathlib import Path

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
