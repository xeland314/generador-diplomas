#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
DIST_DIR="$PROJECT_DIR/dist"

echo "=== Building generador-diplomas for Linux ==="

cd "$PROJECT_DIR"

echo "--- Installing dependencies ---"
uv sync --group dev

echo "--- Running PyInstaller ---"
uv run --group dev pyinstaller \
    --noconfirm \
    --clean \
    --windowed \
    --name generador-diplomas \
    --distpath "$DIST_DIR" \
    --workpath "$PROJECT_DIR/build/build-linux" \
    --specpath "$SCRIPT_DIR" \
    main.py

echo ""
echo "=== Build complete ==="
echo "Binary: $DIST_DIR/generador-diplomas/generador-diplomas"
