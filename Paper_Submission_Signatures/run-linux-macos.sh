#!/bin/sh
# Turns the raw CSV of signatures into Word files and then PDFs.
#
#   ./run.sh                 do the whole job: CSV -> .docx files -> PDFs
#   ./run.sh --libreoffice   ... converting with LibreOffice (faster)
#
# The first run sets everything up, which takes a minute; later runs skip
# straight to the work.

set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
    echo "Python is not installed. On Ubuntu/Mint/Debian, run:"
    echo "    sudo apt install python3 python3-venv"
    exit 1
fi

if [ ! -d .venv ]; then
    echo "Setting up (first run only)..."
    python3 -m venv .venv
fi

. .venv/bin/activate
pip install --quiet --requirement requirements.txt
python split_signatures_into_folders.py
python convert_docx_to_pdf.py "$@"
