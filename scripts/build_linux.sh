#!/usr/bin/env bash
set -euo pipefail

# cd to project root (this script lives in scripts/)
cd "$(dirname "$0")/.."

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python -m unittest discover -s tests

python -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name ExcelSplitter-Linux \
  src/excel_splitter.py

echo
echo "Build complete. Output: dist/ExcelSplitter-Linux"
