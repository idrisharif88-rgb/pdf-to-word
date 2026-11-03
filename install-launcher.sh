#!/bin/sh
# Install the "PDF to Word" launcher for this checkout (app menu + Desktop).
set -e
APP_DIR=$(cd "$(dirname "$0")" && pwd)
OUT="$HOME/.local/share/applications/pdf-to-word.desktop"
mkdir -p "$(dirname "$OUT")"
sed "s|@APP_DIR@|$APP_DIR|g" "$APP_DIR/pdf-to-word.desktop.in" > "$OUT"
chmod +x "$OUT"
if [ -d "$HOME/Desktop" ]; then
    cp "$OUT" "$HOME/Desktop/pdf-to-word.desktop"
    gio set "$HOME/Desktop/pdf-to-word.desktop" metadata::trusted true 2>/dev/null || true
fi
echo "Installed launcher for $APP_DIR"
