#!/bin/sh
# Builds the book of initials.
#
#   ./run.sh              build the book
#   ./run.sh --dry-run    only report the layout and the page count
#
# The first run sets everything up, which takes a minute; later runs skip
# straight to building.

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
python generate_initials_book.py "$@"
