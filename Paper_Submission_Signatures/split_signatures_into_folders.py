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
# Separators are ordered least-likely-to-appear-inside-a-field first, and
# that order is load-bearing for the same reason the encoding order is.
# Unlike an encoding, a separator cannot fail: several can each split the
# same file into more than one column, and the first one tried wins. The
# column count cannot break that tie - both splits look equally valid -
# so only the characters themselves can:
#
#   \t and |   never occur inside this export's fields, so a file that
#              splits on one of them really is separated by it.
#   ;          occurs in prose but almost never in a name or an address.
#   ,          LAST, and it must stay last: it is the one separator that
#              turns up inside the data itself. Today's export is clean -
#              no field in it contains a comma - but a later one carrying
#              a free-text column (an address, say: "гр. София, ул.
#              Витоша 5") would break the moment the comma were tried
#              first, shifting every field one column to the left.
#
# The first encoding that decodes the file wins, so the order is what makes
# this correct rather than merely successful. An encoding earns a place here
# only if it can be reached - that is, only if the ones before it fail on
# the files it is meant to catch:
#
#   utf-8    what a modern export should be. Fails loudly on anything else,
#            and pandas strips the byte-order mark Excel writes, so a
#            separate utf-8-sig entry would never be reached.
#   cp1251   Windows Cyrillic - what Excel saves Bulgarian text as. utf-8
#            rejects those bytes, so this is reachable.
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


def _format_duration(seconds: float) -> str:
    """Format seconds as ``'Xs'`` or ``'Xm Ys'``."""
    if seconds < 60:
        return f"{seconds:.0f}s"
    return f"{int(seconds) // 60}m {int(seconds) % 60}s"


# ---------------------------------------------------------------------------
# CSV reading
# ---------------------------------------------------------------------------

def _try_read_csv(path: str, sep: str, encoding: str = "utf-8") -> pd.DataFrame | None:
    """Return the DataFrame only if it parses into more than one column."""
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

    The first combination that yields more than one column wins, so the
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
) -> tuple[DocumentType, int]:
    """Build the Word document for a single submission folder (~1000 rows)."""
    doc = create_landscape_document()

    pages_in_file = math.ceil(len(folder_rows) / ROWS_PER_PAGE)
    # Pages in a *full* folder — used so page numbering continues across files
    pages_per_full_folder = math.ceil(ROWS_PER_FILE / ROWS_PER_PAGE)

    for page_num, page_rows in enumerate(
            in_blocks_of(folder_rows, ROWS_PER_PAGE), start=1):
        add_page_table(doc, header, page_rows, column_widths)

        global_page = page_num + (folder_number - 1) * pages_per_full_folder
        add_page_footer(doc, global_page, folder_number)

        if page_num < pages_in_file:
            doc.add_page_break()

    return doc, pages_in_file


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main() -> None:
    print(f"Reading '{INPUT_CSV}'...")
    try:
        df = read_signatures_csv(INPUT_CSV)
    except FileNotFoundError:
        print(f"ERROR: file '{INPUT_CSV}' not found in this folder.")
        return
    except ValueError as exc:
        print(f"ERROR: {exc}")
        return

    print(f"Read {len(df)} rows with {len(df.columns)} columns.")
    print(f"Columns: {list(df.columns)}")

    header = df.columns.tolist()
    column_widths = scale_column_widths(len(header))
    print(f"Column widths (cm): {[round(w, 2) for w in column_widths]}")

    os.makedirs(OUTPUT_DOCX_FOLDER, exist_ok=True)
    print(f"Writing output into '{OUTPUT_DOCX_FOLDER}/'.")

    total_files = math.ceil(len(df) / ROWS_PER_FILE)
    print(f"\nCreating {total_files} submission file(s)...\n")

    overall_start = time.perf_counter()

    for folder_number, folder_rows in enumerate(
            in_blocks_of(df, ROWS_PER_FILE), start=1):
        first_id = str(folder_rows.iloc[0, 0])
        last_id = str(folder_rows.iloc[-1, 0])

        doc, pages = build_folder_document(
            folder_number, folder_rows, header, column_widths
        )
        filename = (
            f"Папка {folder_number} с подписи от {first_id} до {last_id}.docx"
        )
        output_path = os.path.join(OUTPUT_DOCX_FOLDER, filename)
        doc.save(output_path)

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
        f"Created {total_files} file(s) in '{OUTPUT_DOCX_FOLDER}/'."
    )


if __name__ == "__main__":
    main()
