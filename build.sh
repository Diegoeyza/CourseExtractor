#!/usr/bin/env bash
# ── CourseExtractor — PyInstaller build script ─────────────────────────────────
# Produces a single self-contained executable in dist/
# Usage: bash build.sh
# Requirements: pip install -r requirements.txt  (must be done first)

set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "==> Installing/updating dependencies…"
pip install -r requirements.txt

echo "==> Building executable…"
pyinstaller \
  --onefile \
  --windowed \
  --name "CourseExtractor" \
  --hidden-import tkinter \
  --add-data "extractor_service.py:." \
  app.py

echo ""
echo "==> Done! Executable is at: dist/CourseExtractor"
echo "    On Windows it will be:  dist/CourseExtractor.exe"
