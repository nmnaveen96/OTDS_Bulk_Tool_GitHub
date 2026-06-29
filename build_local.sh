#!/usr/bin/env bash
# Local build – Linux / macOS
set -e
echo "======================================================"
echo "  OTDS Bulk Tool - Local Build (Linux / macOS)"
echo "======================================================"
python -m pip install --upgrade pip
pip install -r requirements.txt
rm -rf build dist __pycache__
echo ""
echo ">> Running PyInstaller..."
pyinstaller otds_bulk_tool.spec --clean --noconfirm
chmod +x dist/otds_bulk_tool
echo ""
echo "======================================================"
echo "  SUCCESS -> dist/otds_bulk_tool"
echo "======================================================"
echo "  Test: ./dist/otds_bulk_tool --init"
echo "        ./dist/otds_bulk_tool --dry-run"
