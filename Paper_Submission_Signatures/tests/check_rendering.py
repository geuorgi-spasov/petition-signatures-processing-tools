"""
Builds a submission folder, converts it to PDF, and checks that the
document really has the number of pages it says it has.

This matters more here than in the book of initials. The book lets the
program fill pages as it likes, so its page count is an estimate. These
documents cannot: they break the page after every tenth signature and
print the page number in the footer, continuing across all 111 files. If
the renderer disagrees about what fits, the breaks land in the wrong
place and every footer is wrong — which is exactly what happened when the
page size moved from US Letter to A4 and a 100-page folder came out at
199 pages while the script still reported 100.

So the rule here is strict: declared and actual must match exactly.

Usage:
    python tests/check_rendering.py
    python tests/check_rendering.py --rows 200
    python tests/check_rendering.py --rows 200 --rows-per-page 12
    python tests/check_rendering.py --libreoffice

--rows-per-page is the one worth running after any change to the row
height, the footer or the offsets: it is what moves the page breaks, and
the generator caps it for exactly that reason. The cap is arithmetic; this
checks the arithmetic against a renderer.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from split_signatures_into_folders import (  # noqa: E402
    BODY_FONT,
    FOOTER_FONT,
    INPUT_CSV,
    ROWS_PER_FILE,
    ROWS_PER_PAGE,
    build_folder_document,
    fits_on_the_page,
    max_rows_per_page,
    read_signatures_csv,
    scale_column_widths,
)
from convert_docx_to_pdf import get_converter  # noqa: E402


def count_pdf_pages(path: str) -> int:
    """Count a PDF's pages without needing any extra library."""
    with open(path, "rb") as handle:
        data = handle.read()
    counts = [int(n) for n in re.findall(rb"/Count\s+(\d+)", data)]
    if counts:
        return max(counts)
    return len(re.findall(rb"/Type\s*/Page[^s]", data))


def fonts_in_pdf(path: str) -> set[str]:
    """The font names a PDF embeds, without their subset prefixes."""
    with open(path, "rb") as handle:
        data = handle.read()
    names = re.findall(rb"/BaseFont\s*/([A-Za-z0-9+\-,._]+)", data)
    return {name.decode("latin-1").split("+")[-1] for name in names}


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Check a submission document's real page count."
    )
    parser.add_argument("--input", default=INPUT_CSV, help="CSV of signatures")
    parser.add_argument("--rows", type=int, default=ROWS_PER_FILE,
                        help="Signatures to put in the folder being checked")
    parser.add_argument("--rows-per-page", type=int, default=ROWS_PER_PAGE,
                        help="Signatures on a page — this is what moves the "
                             "page breaks, so it is the setting most worth "
                             "checking against a real renderer")
    choice = parser.add_mutually_exclusive_group()
    choice.add_argument("--word", dest="prefer", action="store_const",
                        const="word", help="Convert with Microsoft Word")
    choice.add_argument("--libreoffice", dest="prefer", action="store_const",
                        const="libreoffice", help="Convert with LibreOffice")
    args = parser.parse_args()

    if not fits_on_the_page(args.rows_per_page):
        print(f"ERROR: {args.rows_per_page} rows a page do not fit; the "
              f"generator allows at most {max_rows_per_page()}.")
        return 1

    if not os.path.isfile(args.input):
        print(f"ERROR: '{args.input}' not found. Run this from the folder "
              f"that holds the CSV.")
        return 1

    convert = get_converter(args.prefer)

    signatures = read_signatures_csv(args.input).head(args.rows)
    header = list(signatures.columns)
    widths = scale_column_widths(len(header))

    print(f"\nBuilding one folder of {len(signatures):,} signatures "
          f"({args.rows_per_page} to a page)...")

    with tempfile.TemporaryDirectory() as folder:
        docx = os.path.join(folder, "folder.docx")
        pdf = os.path.join(folder, "folder.pdf")
        document, declared_pages = build_folder_document(
            1, signatures, header, widths, args.rows_per_page
        )
        document.save(docx)
        convert(docx, pdf)
        if not os.path.exists(pdf):
            print("FAILED — the converter produced no PDF.")
            return 1
        actual_pages = count_pdf_pages(pdf)
        embedded = fonts_in_pdf(pdf)

    rendered_rows_per_page = (
        len(signatures) / actual_pages if actual_pages else 0
    )

    print()
    print(f"{'':20}{'declared':>10}{'actual':>10}")
    print(f"{'pages':20}{declared_pages:>10,}{actual_pages:>10,}")
    print(f"{'signatures a page':20}{args.rows_per_page:>10}"
          f"{rendered_rows_per_page:>10.1f}")

    missing = [font for font in (BODY_FONT, FOOTER_FONT)
               if not any(font.replace(" ", "").lower() == name.lower()
                          for name in embedded)]
    print()
    if missing:
        print(f"Fonts: {' and '.join(missing)} not available here — the PDF "
              f"contains {', '.join(sorted(embedded))}.")
        print("Those stand-ins have the same character widths, so the layout "
              "holds, but the letters are drawn differently.")
    else:
        print(f"Fonts: {BODY_FONT} and {FOOTER_FONT} — the real ones.")

    print()
    if actual_pages != declared_pages:
        print(f"FAILED — the document says {declared_pages:,} pages and "
              f"renders as {actual_pages:,}.")
        print("The page breaks are in the wrong place, so the page numbers "
              "printed in the footers are wrong too. The table and footer no "
              "longer fit the page: reduce ROW_HEIGHT_CM, FOOTER_FONT_SIZE_PT "
              "or the top and bottom offsets.")
        return 1

    print(f"OK — {actual_pages:,} pages, exactly as declared, with "
          f"{args.rows_per_page} signatures on each.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
