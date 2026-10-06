"""
Splits a raw CSV export of signatures into multiple Word documents for
paper submission. Each document contains ROWS_PER_FILE signatures
(default: 1000), laid out ROWS_PER_PAGE per A4 landscape page (default: 10),
with a footer that page-numbers continuously across all files.

All generated .docx files are written into OUTPUT_DOCX_FOLDER, which is
created automatically if it doesn't exist.

Usage:
    python split_signatures_into_folders.py

See README.md for the full workflow.
"""

from __future__ import annotations

import argparse
import math
import os
import time
from collections.abc import Iterator

import pandas as pd
from docx import Document
from docx.document import Document as DocumentType
from docx.enum.section import WD_ORIENT
from docx.enum.table import (
    WD_ALIGN_VERTICAL,
    WD_ROW_HEIGHT_RULE,
    WD_TABLE_ALIGNMENT,
)
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.table import Table

# ---------------------------------------------------------------------------
# Configuration — change these if your file names or layout differ
# ---------------------------------------------------------------------------

INPUT_CSV = "Signatures_from_the_database_raw.csv"
OUTPUT_DOCX_FOLDER = "signatures_docx"

# Grouping
ROWS_PER_FILE = 1000   # signatures per submission folder/document
ROWS_PER_PAGE = 10     # data rows per printed page

# Page layout, in centimeters (A4 landscape — the paper this is printed on)
PAGE_WIDTH_CM = 29.7
PAGE_HEIGHT_CM = 21.0
LEFT_MARGIN_CM = 2.0
RIGHT_MARGIN_CM = 2.0
TOP_MARGIN_CM = 1.5
BOTTOM_MARGIN_CM = 2.0

# How much height the footer takes off every page. Measured, not derived:
# the footer is a blank line and two 12 pt italic lines, and what actually
# fits is settled by the renderer, not by adding those up. Ten rows a page
# is the most that renders as declared; eleven does not. Re-measure with
# `python tests/check_rendering.py --rows-per-page N` if the row height,
# the footer or the offsets change.
FOOTER_BLOCK_CM = 2.0

# Row layout
ROW_HEIGHT_CM = 1.3
BODY_FONT = "Arial"
BODY_FONT_SIZE_PT = 12
FOOTER_FONT = "Cambria"
FOOTER_FONT_SIZE_PT = 12   # lower this to give the table more height

# Footer text (Bulgarian)
ORGANIZATION_NAME = 'Сдружение „Невидими животни"'

# Column widths in centimeters, tuned for a 6-column CSV. Only their
# proportions matter: they are scaled so their total exactly fills the
# available page width.
DEFAULT_COLUMN_WIDTHS_CM = [1.8, 3.45, 3.7, 8.75, 3.35, 3.0]

# Separators and encodings tried when auto-detecting the CSV format.
#
# Separators are ordered least-likely-to-appear-inside-a-field first.
# A separator cannot fail the way an encoding can: if a field holds the
# character being tried, that split also yields at least two columns and
# wins, silently, and the column count cannot tell the two splits apart.
#
# The order is a judgement, not a measurement. None of the four appears
# anywhere in the current export, so 110,942 rows rule them out equally
# and cannot rank them. What that sample does show is that fields are
# not clean - '?', '.', '-', digits, an apostrophe, an ellipsis, '%',
# ':' and '@' all turn up inside names - so none of these is a "never".
# A tab is the least plausible of the four: unlike the others it is not
# a character anyone types in the middle of a word. Then a pipe, then a
# semicolon, and a comma by far the most likely, so it is tried last.
#
# The first encoding that decodes the file wins, so the order is what makes
# this correct rather than merely successful. An encoding earns a place here
# only if it can be reached - that is, only if the ones before it fail on
# the files it is meant to catch:
#
#   utf-8    what a modern export should be. Fails loudly on anything else,
#            and pandas strips the byte-order mark Excel writes, so a
#            separate utf-8-sig entry would never be reached.
#   cp1251   Windows Cyrillic, the usual encoding for Bulgarian text
#            out of Excel. utf-8 rejects those bytes, so this is
#            reachable - that part is tested; which tool produced any
#            given file is not something this script can know.
#   latin-1  LAST RESORT, and it must stay last: it maps every one of the
#            256 byte values to a character, so it can never fail. Put it
#            earlier and it swallows the file, turning "Иван" into "Èâàí"
#            without raising anything at all.
#
# Deliberately absent: cp866 (DOS Cyrillic) decodes the same bytes as
# cp1251 without error, so after cp1251 it could never be reached - listing
# it would only suggest a coverage that is not there. The same goes for
# cp1252 and iso-8859-1, which latin-1 shadows completely (iso-8859-1 is
# not even a different codec - Python resolves both names to the same one).
CSV_SEPARATORS = ["\t", "|", ";", ","]
CSV_ENCODINGS = ["utf-8", "cp1251", "latin-1"]
LAST_RESORT_ENCODING = "latin-1"


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def in_blocks_of(rows: pd.DataFrame, size: int) -> Iterator[pd.DataFrame]:
    """Cut ``rows`` into successive blocks of at most ``size`` rows.

    The export is cut twice, by the same rule at two scales: into the
    ~1000-row folders that each become a file, and then into the 10-row
    pages that each become a table. The last block of either is whatever
    is left over - ``iloc`` stops at the end rather than running past it.
    """
    for start in range(0, len(rows), size):
        yield rows.iloc[start:start + size]


def max_rows_per_page() -> int:
    """How many data rows fit between the top and bottom offsets.

    One row is ROW_HEIGHT_CM high and the header row takes one more, so
    this is the usable height - less what the footer takes - divided by
    the row height, less that header. It is the same arithmetic the page
    breaks rely on: if more rows are asked for than this, the table runs
    off the page and the declared page count stops matching the rendered
    one.

    The footer term was missing until tests/check_rendering.py was given
    --rows-per-page and could finally render the thing. Without it this
    returned 12, and at 11 and 12 rows a page a document declaring 6
    pages came out as 11. The shipped default of 10 was never affected.
    """
    usable_cm = (PAGE_HEIGHT_CM - TOP_MARGIN_CM - BOTTOM_MARGIN_CM
                 - FOOTER_BLOCK_CM)
    return int(usable_cm // ROW_HEIGHT_CM) - 1


def fits_on_the_page(rows_per_page: int) -> bool:
    """Whether that many data rows, plus the header row, fit on a page."""
    return 1 <= rows_per_page <= max_rows_per_page()


def _format_duration(seconds: float) -> str:
    """Format seconds as ``'Xs'`` or ``'Xm Ys'``."""
    if seconds < 60:
        return f"{seconds:.0f}s"
    return f"{int(seconds) // 60}m {int(seconds) % 60}s"


# ---------------------------------------------------------------------------
# CSV reading
# ---------------------------------------------------------------------------

def _try_read_csv(path: str, sep: str, encoding: str = "utf-8") -> pd.DataFrame | None:
    """Return the DataFrame only if it parses into at least two columns."""
    try:
        df = pd.read_csv(path, sep=sep, encoding=encoding)
    except FileNotFoundError:
        raise  # a missing file is not a parsing problem - report it
    except Exception:
        return None
    if len(df.columns) > 1:
        return df
    return None


def read_signatures_csv(path: str) -> pd.DataFrame:
    """Read the raw CSV, trying each encoding and separator in turn.

    The first combination giving at least two columns wins, so the
    order of CSV_ENCODINGS is what makes this correct rather than merely
    successful - see the note beside it.
    """
    for encoding in CSV_ENCODINGS:
        for sep in CSV_SEPARATORS:
            df = _try_read_csv(path, sep, encoding)
            if df is None:
                continue
            print(f"Read '{path}' using separator {sep!r} and "
                  f"encoding '{encoding}'.")
            if encoding == LAST_RESORT_ENCODING:
                print(f"  NOTE: '{encoding}' accepts any file at all, so this "
                      f"is a guess.\n  Check the names below look right - if "
                      f"they read like 'Èâàí' instead of 'Иван', re-save the "
                      f"CSV as UTF-8.")
            return df

    raise ValueError(
        f"Could not parse '{path}' as a multi-column CSV.\n"
        f"Tried separators {CSV_SEPARATORS} with encodings {CSV_ENCODINGS}."
    )


# ---------------------------------------------------------------------------
# Document building
# ---------------------------------------------------------------------------

def scale_column_widths(num_columns: int) -> list[float]:
    """Return column widths in centimeters that fit the available width."""
    available_cm = PAGE_WIDTH_CM - LEFT_MARGIN_CM - RIGHT_MARGIN_CM

    if num_columns == len(DEFAULT_COLUMN_WIDTHS_CM):
        widths = DEFAULT_COLUMN_WIDTHS_CM
    elif num_columns < len(DEFAULT_COLUMN_WIDTHS_CM):
        widths = DEFAULT_COLUMN_WIDTHS_CM[:num_columns]
    else:
        # More columns than we have presets for: distribute evenly
        return [available_cm / num_columns] * num_columns

    scale = available_cm / sum(widths)
    return [w * scale for w in widths]


# A thin solid black line, as the four attributes OOXML wants on every
# border edge. python-docx has no API for table borders - it models
# fonts, widths and alignment, but not these - so the only way to draw a
# gridline is to build the XML elements by hand.
BORDER_STYLE = {
    qn("w:val"): "single",   # a solid line, not dashed or doubled
    qn("w:sz"): "4",         # thickness in eighths of a point: 0.5 pt
    qn("w:space"): "0",      # no padding between the line and the text
    qn("w:color"): "000000",
}

# A table styles its outer frame and its inner gridlines in one element;
# a cell has only its own four sides.
TABLE_EDGES = ("top", "left", "bottom", "right", "insideH", "insideV")
CELL_EDGES = ("top", "left", "bottom", "right")


def _borders_element(tag: str, edges: tuple[str, ...]):
    """Build a ``<w:tblBorders>`` or ``<w:tcBorders>`` with every edge drawn."""
    container = OxmlElement(tag)
    for edge_name in edges:
        edge = OxmlElement(f"w:{edge_name}")
        for attribute, value in BORDER_STYLE.items():
            edge.set(attribute, value)
        container.append(edge)
    return container


def apply_full_table_borders(table: Table) -> None:
    """Draw visible single-line borders on every cell of the table."""
    table._tbl.tblPr.append(_borders_element("w:tblBorders", TABLE_EDGES))

    # And again on each cell: belt and braces, because some viewers
    # honour only the cell-level borders.
    for row in table.rows:
        for cell in row.cells:
            cell._tc.get_or_add_tcPr().append(
                _borders_element("w:tcBorders", CELL_EDGES)
            )


def create_landscape_document() -> DocumentType:
    """Create a new Word document set up for landscape printing."""
    doc = Document()
    doc.styles["Normal"].font.name = BODY_FONT

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(PAGE_WIDTH_CM)
    section.page_height = Cm(PAGE_HEIGHT_CM)
    section.left_margin = Cm(LEFT_MARGIN_CM)
    section.right_margin = Cm(RIGHT_MARGIN_CM)
    section.top_margin = Cm(TOP_MARGIN_CM)
    section.bottom_margin = Cm(BOTTOM_MARGIN_CM)

    return doc


def add_page_table(
    doc: DocumentType,
    header: list[str],
    page_rows: pd.DataFrame,
    column_widths: list[float],
) -> None:
    """Add one page's worth of data as a bordered table."""
    table = doc.add_table(rows=len(page_rows) + 1, cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    for i, width in enumerate(column_widths):
        if i < len(table.columns):
            table.columns[i].width = Cm(width)

    # Header row: bold, centered, body font size
    for i, col_name in enumerate(header):
        cell = table.rows[0].cells[i]
        cell.text = str(col_name)
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(BODY_FONT_SIZE_PT)

    # Data rows: left-aligned, body font size
    for row_idx, data_row in enumerate(page_rows.itertuples(index=False)):
        table_row = table.rows[row_idx + 1]
        for col_idx, value in enumerate(data_row):
            if col_idx >= len(table_row.cells):
                continue
            cell = table_row.cells[col_idx]
            cell.text = str(value) if pd.notna(value) else ""
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in paragraph.runs:
                    run.font.size = Pt(BODY_FONT_SIZE_PT)

    # Fixed row height with vertically centered cells
    for row in table.rows:
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row.height = Cm(ROW_HEIGHT_CM)
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    apply_full_table_borders(table)


def add_page_footer(
    doc: DocumentType, global_page_number: int, folder_number: int
) -> None:
    """Write the per-page footer: blank line, page/folder info, org name."""
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.LEFT
    footer.add_run().add_break()  # blank line before the footer

    info_run = footer.add_run(
        f"Стр. {global_page_number}, папка {folder_number}"
    )
    info_run.font.name = FOOTER_FONT
    info_run.font.size = Pt(FOOTER_FONT_SIZE_PT)
    info_run.italic = True
    info_run.add_break()

    org_run = footer.add_run(ORGANIZATION_NAME)
    org_run.font.name = FOOTER_FONT
    org_run.font.size = Pt(FOOTER_FONT_SIZE_PT)
    org_run.italic = True


def build_folder_document(
    folder_number: int,
    folder_rows: pd.DataFrame,
    header: list[str],
    column_widths: list[float],
    rows_per_page: int = ROWS_PER_PAGE,
    rows_per_file: int = ROWS_PER_FILE,
) -> tuple[DocumentType, int]:
    """Build the Word document for a single submission folder (~1000 rows)."""
    doc = create_landscape_document()

    pages_in_file = math.ceil(len(folder_rows) / rows_per_page)
    # Pages in a *full* folder — used so page numbering continues across files
    pages_per_full_folder = math.ceil(rows_per_file / rows_per_page)

    for page_num, page_rows in enumerate(
            in_blocks_of(folder_rows, rows_per_page), start=1):
        add_page_table(doc, header, page_rows, column_widths)

        global_page = page_num + (folder_number - 1) * pages_per_full_folder
        add_page_footer(doc, global_page, folder_number)

        if page_num < pages_in_file:
            doc.add_page_break()

    return doc, pages_in_file


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def build_arg_parser() -> argparse.ArgumentParser:
    """The command line. Every default is the constant of the same name."""
    parser = argparse.ArgumentParser(
        description="Split a signatures CSV into Word submission documents."
    )
    parser.add_argument("--input", default=INPUT_CSV,
                        help="CSV file to read")
    parser.add_argument("--output", default=OUTPUT_DOCX_FOLDER,
                        help="Folder to write the documents into")
    parser.add_argument("--rows-per-file", type=int, default=ROWS_PER_FILE,
                        help="Signatures in one submission folder")
    parser.add_argument("--rows-per-page", type=int, default=ROWS_PER_PAGE,
                        help="Signatures on one printed page. Changing this "
                             "changes where the page breaks fall, so re-run "
                             "tests/check_rendering.py afterwards")
    parser.add_argument("--limit", type=int, default=None,
                        help="Only process the first N signatures")
    parser.add_argument("--dry-run", action="store_true",
                        help="Report what would be written, write nothing")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_arg_parser().parse_args(argv)

    if args.rows_per_file < 1 or args.rows_per_page < 1:
        print("ERROR: --rows-per-file and --rows-per-page must be at least 1.")
        return 1
    if not fits_on_the_page(args.rows_per_page):
        print(f"ERROR: {args.rows_per_page} rows of {ROW_HEIGHT_CM} cm plus a "
              f"header do not fit between the {TOP_MARGIN_CM} cm and "
              f"{BOTTOM_MARGIN_CM} cm offsets of a "
              f"{PAGE_WIDTH_CM} x {PAGE_HEIGHT_CM} cm page.")
        print(f"Use at most {max_rows_per_page()} row(s) a page, or a smaller "
              f"ROW_HEIGHT_CM.")
        return 1

    print(f"Reading '{args.input}'...")
    try:
        df = read_signatures_csv(args.input)
    except FileNotFoundError:
        print(f"ERROR: file '{args.input}' not found in this folder.")
        return 1
    except ValueError as exc:
        print(f"ERROR: {exc}")
        return 1

    if args.limit is not None:
        df = df.head(args.limit)

    print(f"Read {len(df)} rows with {len(df.columns)} columns.")
    print(f"Columns: {list(df.columns)}")

    header = df.columns.tolist()
    column_widths = scale_column_widths(len(header))
    print(f"Column widths (cm): {[round(w, 2) for w in column_widths]}")

    total_files = math.ceil(len(df) / args.rows_per_file)
    total_pages = math.ceil(len(df) / args.rows_per_page)

    if args.dry_run:
        print(f"\n{len(df)} signatures -> {total_files} file(s), "
              f"{total_pages} page(s) at {args.rows_per_page} a page.")
        print("Dry run — nothing written.")
        return 0

    # exist_ok: a second run writes into the same folder rather than failing
    os.makedirs(args.output, exist_ok=True)
    print(f"Writing output into '{args.output}/'.")

    print(f"\nCreating {total_files} submission file(s)...\n")

    overall_start = time.perf_counter()

    # One file per folder of signatures, named for the range of database
    # IDs it holds: "Папка 7 с подписи от 6001 до 7000". Column Номер is
    # column 0 and runs 1 to 110,942 in order, so the range identifies a
    # folder exactly - which is why the first and last row are read before
    # the document is built.
    for folder_number, folder_rows in enumerate(
            in_blocks_of(df, args.rows_per_file), start=1):
        first_id = str(folder_rows.iloc[0, 0])
        last_id = str(folder_rows.iloc[-1, 0])

        doc, pages = build_folder_document(
            folder_number, folder_rows, header, column_widths,
            args.rows_per_page, args.rows_per_file
        )
        filename = (
            f"Папка {folder_number} с подписи от {first_id} до {last_id}.docx"
        )
        output_path = os.path.join(args.output, filename)
        doc.save(output_path)

        # Every folder takes about as long as the last, so the average so
        # far is a good enough estimate of what is left. This run takes
        # minutes, and a line that only said "[7/111]" would leave the
        # person watching it with no idea whether to wait.
        elapsed = time.perf_counter() - overall_start
        avg_per_folder = elapsed / folder_number
        remaining = avg_per_folder * (total_files - folder_number)
        print(
            f"  [{folder_number}/{total_files}] '{filename}' — "
            f"{len(folder_rows)} signatures, {pages} pages"
            + (
                f"  (~{_format_duration(remaining)} remaining)"
                if folder_number < total_files
                else ""
            )
        )

    total_time = time.perf_counter() - overall_start
    print(
        f"\nDone in {_format_duration(total_time)}. "
        f"Created {total_files} file(s) in '{args.output}/'."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
