#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

python3 -m PyInstaller --noconfirm --clean --windowed \
  --name GmailDownloader \
  --osx-bundle-identifier com.gmaildownloader.app \
  --add-data "icon.png:." \
  gmaildownloader.py

test -d "$ROOT_DIR/dist/GmailDownloader.app"
printf 'Built %s\n' "$ROOT_DIR/dist/GmailDownloader.app"
