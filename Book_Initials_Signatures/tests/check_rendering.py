"""
Builds the book, has a real program turn it into a PDF, and checks that
the result matches what the layout predicted.

This is the check no unit test can do. The layout arithmetic can be
perfectly self-consistent and still be wrong, because the program that
paginates the document gets the final say — Word fits fewer lines on a
page than the arithmetic predicts, LibreOffice fits exactly as many. That
disagreement once turned a 378-page book into a 755-page one - back
when the book ran 7 columns to a line - and every
unit test passed throughout.

What it reports:

    lines on a page  what the layout predicted, against what the PDF
                     actually contains (worked out from the page count)
    pages            predicted, against the PDF's own page count

A few percent of difference between programs is normal and fine. A large
one in either direction means the document will not come out as intended,
and the check fails - too few pages is as much a surprise as too many.

Usage:
    python tests/check_rendering.py
    python tests/check_rendering.py --limit 20000
    python tests/check_rendering.py --columns 5 --font-size 12
    python tests/check_rendering.py --word

The last one matters on Windows. LibreOffice is the default because it
runs everywhere, but Word is the renderer that fits fewer lines on a page
and so the one that produced the 755-page book - measure it where it is
the program that will open the file.
"""

from __future__ import annotations

import argparse
import contextlib
import glob
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from generate_initials_book import (  # noqa: E402
    FONT_NAME,
    INPUT_CSV,
    Layout,
    build_initials_document,
    group_into_lines,
    names_to_initials,
    read_two_column_csv,
)

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUNDLED_FONT = os.path.join(HERE, "oswald.ttf")

DEFAULT_LIMIT = 20_000
TOLERANCE = 0.10  # a tenth more pages than predicted is still acceptable


def say(message: str) -> None:
    """Print a progress line that appears immediately, not at the end."""
    print(message, flush=True)


def find_libreoffice() -> str | None:
    """Return the path to the LibreOffice CLI, or None."""
    for name in ("libreoffice", "soffice"):
        path = shutil.which(name)
        if path:
            return path

    candidates = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    for directory in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                      os.environ.get("ProgramFiles(x86)", "")):
        if directory:
            candidates.append(
                os.path.join(directory, "LibreOffice", "program", "soffice.exe")
            )
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return None


def find_word() -> str | None:
    """Return ``"docx2pdf"`` if the package that drives Word is installed."""
    try:
        import docx2pdf  # noqa: F401
    except ImportError:
        return None
    return "docx2pdf"


def find_converter(prefer: str | None = None) -> tuple[str, str] | None:
    """Return ``(kind, path)`` for a program that can make PDFs, or None.

    Which program matters, and not only for speed. The two disagree about
    how many lines fit on a page - Word fits fewer - and that disagreement
    is the whole reason this check exists. So whichever one is about to
    print the book is the one worth measuring, and on a machine with both
    it has to be possible to say which.

    Unasked, LibreOffice wins: it is the one that can be installed
    anywhere, so it is the one a check that must run everywhere defaults
    to. That is a default, not a verdict - pass --word on a Windows
    machine where Word is what will open the book.
    """
    if prefer == "word":
        path = find_word()
        return ("word", path) if path else None
    if prefer == "libreoffice":
        path = find_libreoffice()
        return ("libreoffice", path) if path else None

    path = find_libreoffice()
    if path:
        return ("libreoffice", path)
    path = find_word()
    return ("word", path) if path else None


def font_environment(work_dir: str) -> dict[str, str]:
    """An environment in which the bundled font is findable.

    On Linux the check should not depend on whether anyone remembered to
    install the font: a layout measured with a substitute is not the
    layout that will be printed. fontconfig can be pointed at the copy in
    this folder for the length of one conversion, without installing
    anything. Elsewhere this does nothing and the system fonts are used.
    """
    environment = dict(os.environ)
    if not sys.platform.startswith("linux") or not os.path.exists(BUNDLED_FONT):
        return environment

    config = os.path.join(work_dir, "fonts.conf")
    with open(config, "w", encoding="utf-8") as handle:
        handle.write(
            "<?xml version='1.0'?>\n<fontconfig>\n"
            "  <include ignore_missing='yes'>/etc/fonts/fonts.conf</include>\n"
            f"  <dir>{HERE}</dir>\n"
            f"  <cachedir>{os.path.join(work_dir, 'fontcache')}</cachedir>\n"
            "</fontconfig>\n"
        )
    environment["FONTCONFIG_FILE"] = config
    return environment


def fonts_in_pdf(path: str) -> set[str]:
    """The font names a PDF actually embeds, without their subset prefixes."""
    names = re.findall(rb"/BaseFont\s*/([A-Za-z0-9+\-,._]+)", pathlib_read(path))
    return {name.decode("latin-1").split("+")[-1] for name in names}


def convert_to_pdf(kind: str, path: str, docx: str, out_dir: str) -> str:
    """Convert ``docx`` to a PDF in ``out_dir`` and return the PDF's path."""
    if kind == "libreoffice":
        subprocess.run(
            [path, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            env=font_environment(out_dir),
        )
    else:
        from docx2pdf import convert
        convert(docx, os.path.join(out_dir, "book.pdf"))

    produced = glob.glob(os.path.join(out_dir, "*.pdf"))
    if not produced:
        raise RuntimeError("the converter produced no PDF")
    return produced[0]


def count_pdf_pages(path: str) -> int:
    """Count the pages of a PDF without needing any extra library.

    /Type /Page marks a page object and nothing else, so it is counted
    first. /Count is the fallback for the one case that defeats it - a
    writer that packs its page objects into compressed object streams,
    where the bytes are not there to be found. /Count is not the first
    choice because it is not page-specific: an outline (bookmark) tree
    carries one too, and on a file with bookmarks the largest /Count in
    the file can be the number of bookmarks rather than of pages.
    """
    data = pathlib_read(path)
    page_objects = len(re.findall(rb"/Type\s*/Page[^s]", data))
    if page_objects:
        return page_objects
    counts = [int(n) for n in re.findall(rb"/Count\s+(\d+)", data)]
    return max(counts) if counts else 0


def pathlib_read(path: str) -> bytes:
    with open(path, "rb") as handle:
        return handle.read()


def load_initials(path: str, limit: int) -> list[str]:
    """Read initials from the CSV, or invent some if it is not there."""
    if os.path.isfile(path):
        with contextlib.redirect_stdout(io.StringIO()):
            initials = names_to_initials(read_two_column_csv(path))
    else:
        letters = [chr(code) for code in range(0x410, 0x430)]
        initials = [f"{letters[i % 32]}. {letters[(i * 7) % 32]}."
                    for i in range(limit)]
    return initials[:limit]


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Check the real PDF against what the layout predicted."
    )
    parser.add_argument("--input", default=INPUT_CSV, help="CSV of names")
    parser.add_argument("--limit", type=int, default=DEFAULT_LIMIT,
                        help="How many initials to render")
    parser.add_argument("--columns", type=int, default=Layout().columns)
    parser.add_argument("--font-size", type=float, default=Layout().font_size_pt)
    parser.add_argument("--margin", type=float, default=Layout().margin_cm)
    parser.add_argument("--font-name", default=FONT_NAME,
                        help="Font to set the book in")
    choice = parser.add_mutually_exclusive_group()
    choice.add_argument("--word", dest="prefer", action="store_const",
                        const="word", help="Measure Microsoft Word, which "
                                           "fits fewer lines on a page")
    choice.add_argument("--libreoffice", dest="prefer", action="store_const",
                        const="libreoffice", help="Measure LibreOffice")
    args = parser.parse_args()

    # Every step below is slow enough to look like a hang, and the two
    # slowest say nothing while they work. Keep the output unbuffered so
    # each line appears when it happens rather than all at once at the
    # end - which is what Python does by default when its output is not
    # going straight to a console.
    say("Looking for a program that can make a PDF...")

    converter = find_converter(args.prefer)
    if converter is None:
        if args.prefer == "word":
            say("--word was asked for, but the 'docx2pdf' package that "
                "drives Word is not installed:")
            say("    pip install -r tests/test_requirements.txt")
            return 1
        if args.prefer == "libreoffice":
            say("--libreoffice was asked for, but LibreOffice was not found.")
            say("Install it from https://www.libreoffice.org/ and try again.")
            return 1
        say("No program found that can make a PDF.")
        say("Install LibreOffice (free) or, on Windows/macOS, Microsoft Word.")
        say("On Windows with Word already installed, the package that "
            "drives it may be missing: pip install -r "
            "tests/test_requirements.txt")
        return 1
    kind, path = converter
    program = "LibreOffice" if kind == "libreoffice" else "Microsoft Word"
    say(f"  found {program} ({path}).")

    source = args.input if os.path.isfile(args.input) else "invented initials"
    say(f"Reading {source}...")
    initials = load_initials(args.input, args.limit)
    say(f"  {len(initials):,} initials.")

    layout = Layout(columns=args.columns, font_size_pt=args.font_size,
                    margin_cm=args.margin)
    layout.validate()
    lines = len(group_into_lines(initials, layout.columns))

    with tempfile.TemporaryDirectory() as folder:
        docx = os.path.join(folder, "book.docx")
        say(f"Building the document ({lines:,} lines)...")
        build_initials_document(initials, layout,
                                font_name=args.font_name).save(docx)
        say(f"Converting it to PDF with {program} — this is the slow part, "
            f"and it prints nothing while it runs...")
        pdf = convert_to_pdf(kind, path, docx, folder)
        say("  converted.")
        actual_pages = count_pdf_pages(pdf)
        embedded = fonts_in_pdf(pdf)

    predicted_pages = layout.pages_for(len(initials))

    if not actual_pages:
        print("\nFAILED — could not read a page count out of the PDF.")
        return 1

    # The page count does not say outright how many lines fit on a page,
    # but it bounds it: pages == ceil(lines / per_page) holds only for a
    # narrow range of per_page, and the range tightens the more we render.
    lowest = -(-lines // actual_pages)
    highest = (lines - 1) // (actual_pages - 1) if actual_pages > 1 else lines
    if lowest > highest:
        # The pages are not simply as full as they can be - something is
        # forcing breaks - so only an average can be given.
        implied = f"~{lines / actual_pages:.0f}"
    elif lowest == highest:
        implied = str(lowest)
    else:
        implied = f"{lowest}-{highest}"

    print()
    print(f"{'':16}{'predicted':>12}{'actual':>12}")
    print(f"{'lines on a page':16}{layout.lines_per_page:>12}{implied:>12}")
    print(f"{'pages':16}{predicted_pages:>12,}{actual_pages:>12,}")

    # A page measured with a stand-in font is not the page that gets
    # printed, so say which font actually went into the PDF.
    # A PDF names the face, not the family: Oswald goes in as
    # "Oswald-Regular". Match the family and let a style suffix follow,
    # or this reports the right font as missing.
    wanted = args.font_name.replace(" ", "").lower()
    if any(name.lower() == wanted or name.lower().startswith(wanted + "-")
           for name in embedded):
        print(f"\nFont: {args.font_name} — the real one.")
    else:
        print(f"\nFont: {args.font_name} was NOT used. The PDF contains "
              f"{', '.join(sorted(embedded)) or 'nothing recognisable'}.")
        print("The spacing you see measured here is a substitute's, not "
              "the font the book is set in.")

    overshoot = (actual_pages - predicted_pages) / predicted_pages
    print()
    if abs(overshoot) > TOLERANCE:
        more_or_fewer = "more" if overshoot > 0 else "fewer"
        print(f"FAILED — {abs(overshoot):.0%} {more_or_fewer} pages than "
              f"predicted. This program fits {implied} lines on a page, not "
              f"{layout.lines_per_page}.")
        return 1
    print(f"OK — {overshoot:+.0%} against the prediction, within the "
          f"{TOLERANCE:.0%} allowed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
