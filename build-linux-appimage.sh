#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="${APP_DIR:-$ROOT_DIR/build/GmailDownloader.AppDir}"
OUTPUT="${APPIMAGE_OUT:-$ROOT_DIR/dist/GmailDownloader-x86_64.AppImage}"
cd "$ROOT_DIR"

python3 -m PyInstaller --noconfirm --clean GmailDownloader.spec
rm -rf -- "$APP_DIR"
mkdir -p "$APP_DIR/usr/bin" "$APP_DIR/usr/share/icons/hicolor/256x256/apps"
cp -a "$ROOT_DIR/dist/GmailDownloader/." "$APP_DIR/usr/bin/"
cp "$ROOT_DIR/icon.png" "$APP_DIR/usr/share/icons/hicolor/256x256/apps/gmaildownloader.png"
cp "$ROOT_DIR/icon.png" "$APP_DIR/gmaildownloader.png"

cat > "$APP_DIR/AppRun" <<'APPRUN'
#!/bin/sh
set -eu
HERE="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
exec "$HERE/usr/bin/GmailDownloader" "$@"
APPRUN
chmod +x "$APP_DIR/AppRun"

cat > "$APP_DIR/gmaildownloader.desktop" <<'DESKTOP'
[Desktop Entry]
Type=Application
Name=GmailDownloader
Comment=Local-first Gmail archive and analytics
Exec=GmailDownloader
Icon=gmaildownloader
Categories=Office;Email;
Terminal=false
DESKTOP

if ! command -v appimagetool >/dev/null 2>&1; then
  printf 'appimagetool is required to finish the AppImage build.\n' >&2
  exit 1
fi
mkdir -p "$(dirname -- "$OUTPUT")"
appimagetool "$APP_DIR" "$OUTPUT"
printf 'Built %s\n' "$OUTPUT"
