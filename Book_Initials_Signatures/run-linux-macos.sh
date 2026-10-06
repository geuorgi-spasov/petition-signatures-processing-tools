#!/bin/sh
# Builds the book of initials.
#
#   ./run-linux-macos.sh              build the book
#   ./run-linux-macos.sh --dry-run    only report the layout and the page count
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

# The font is not a Python dependency: the script only writes its *name*
# into the document, and whatever opens the document has to have it. So
# it is installed here, for this user only - no root, nothing touched
# outside the home folder.
FONT_FILE="oswald.ttf"
FONT_NAME="Oswald"

font_is_installed() {
    if command -v fc-match >/dev/null 2>&1; then
        [ "$(fc-match -f '%{family}' "$FONT_NAME" 2>/dev/null)" = "$FONT_NAME" ]
    else
        # macOS has no fontconfig: look where the installer would put it.
        [ -f "$HOME/Library/Fonts/$FONT_FILE" ] || [ -f "/Library/Fonts/$FONT_FILE" ]
    fi
}

if [ -f "$FONT_FILE" ] && ! font_is_installed; then
    case "$(uname -s)" in
        Darwin) font_dir="$HOME/Library/Fonts" ;;
        *)      font_dir="$HOME/.local/share/fonts" ;;
    esac
    echo
    echo "The book is set in Oswald, which is not installed yet."
    echo "The font is bundled here and free to install (SIL Open Font"
    echo "License, see OFL.txt). It goes in your own home folder:"
    echo "    $font_dir/$FONT_FILE"
    echo
    if [ -t 0 ]; then
        printf "Install it now? [Y/n] "
        read -r answer || answer=""
    else
        # The question exists to ask before writing to the home folder,
        # so with no one there to answer, the answer is no.
        answer="n"
        echo "This run has no terminal, so there is no one to ask. Run this"
        echo "file from a terminal to be asked, or copy the font to the"
        echo "folder above yourself."
    fi
    case "$answer" in
        [Nn]*)
            echo
            echo "Skipped. The book will still be built, but whatever opens it"
            echo "will substitute another font and the spacing will be wrong."
            ;;
        *)
            mkdir -p "$font_dir"
            cp "$FONT_FILE" "$font_dir/"
            if command -v fc-cache >/dev/null 2>&1; then
                fc-cache -f "$font_dir" >/dev/null 2>&1 || true
            fi
            if font_is_installed; then
                echo "Installed."
            else
                echo "Copied to $font_dir, but the system does not report it"
                echo "yet. Log out and in again if the book looks wrong."
            fi
            ;;
    esac
    echo
fi

python generate_initials_book.py "$@"
