"""
Generates a printable book of signatory initials from a CSV of names.

The workflow is three steps:

    1. read the CSV of names          -> read_two_column_csv()
    2. turn the names into initials    -> names_to_initials()
    3. lay the initials out on A4      -> Layout + build_initials_document()

A Layout is the five numbers that decide how the book looks: columns per
line, font size, offset from the page edges, distance between the initials
and line spacing. It works out how many pages the book will have, and
refuses — with an explanation — anything that cannot be printed on A4.

Usage:
    python generate_initials_book.py
    python generate_initials_book.py --columns 10 --column-gap 1.0
    python generate_initials_book.py --dry-run     # only report the plan
    python generate_initials_book.py --help        # all parameters

See README.md for the full workflow.
"""

from __future__ import annotations

import argparse
import math
import time
from dataclasses import dataclass

import numpy as np
import pandas as pd
from docx import Document
from docx.document import Document as DocumentType
from docx.shared import Cm, Pt

# ---------------------------------------------------------------------------
# Configuration — the default for every parameter.
# Each one can also be set on the command line, see --help.
# ---------------------------------------------------------------------------

INPUT_CSV = "book_signatures_only_names.csv"
OUTPUT_DOCX = "book_signatures_initials.docx"

# The font must be installed on your system. The exact name has to match
# what your OS reports for the installed font. The bundled file
# `bebasneuecyrillic.ttf` registers itself as "Bebas Neue Cyrillic".
FONT_NAME = "Bebas Neue Cyrillic"
FONT_SIZE_PT = 10.0

COLUMNS = 7           # sets of initials next to each other on one line
MARGIN_CM = 1.5       # offset from all four page edges
COLUMN_GAP_CM = 1.5   # distance between the initials on a line
LINE_SPACING = 1.2    # line height as a multiple of the font size

# --- Fixed facts the layout is calculated from -----------------------------

# A4 portrait, the paper this book is printed on.
PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7

POINTS_PER_CM = 72 / 2.54  # 1 cm = 28.35 pt

# A set of initials is at most "И. И." — five characters. In Bebas Neue
# Cyrillic one character is about 0.42 of the font size wide, so one set is
# 5 × 0.42 × font size wide. Raise 0.42 if you switch to a wider font.
LONGEST_INITIALS = 5
CHARACTER_WIDTH_EM = 0.42

# Separators tried when auto-detecting the CSV format
CSV_SEPARATORS = [",", ";", "\t", "|"]


class LayoutError(ValueError):
    """Raised when a layout cannot be printed on an A4 page."""


# ---------------------------------------------------------------------------
# The layout
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Layout:
    """The five numbers that decide how the book looks.

    Everything else — how wide a set of initials is, how many lines fit on
    a page, how many pages the book takes — follows from these and from the
    size of an A4 sheet, and is a property below. Call :meth:`validate`
    first: it is what guarantees the numbers can actually be printed.
    """

    columns: int = COLUMNS
    font_size_pt: float = FONT_SIZE_PT
    margin_cm: float = MARGIN_CM
    column_gap_cm: float = COLUMN_GAP_CM
    line_spacing: float = LINE_SPACING

    # -- the space the initials have to fit into ----------------------------

    @property
    def text_width_cm(self) -> float:
        """The page width minus the left and right offsets."""
        return PAGE_WIDTH_CM - 2 * self.margin_cm

    @property
    def text_height_cm(self) -> float:
        """The page height minus the top and bottom offsets."""
        return PAGE_HEIGHT_CM - 2 * self.margin_cm

    # -- one set of initials, one line, one page ----------------------------

    @property
    def initials_width_cm(self) -> float:
        """How wide one set of initials is at this font size."""
        return (
            LONGEST_INITIALS * CHARACTER_WIDTH_EM * self.font_size_pt / POINTS_PER_CM
        )

    @property
    def width_needed_cm(self) -> float:
        """How wide a full line is: n sets of initials with n-1 gaps."""
        return (
            self.columns * self.initials_width_cm
            + (self.columns - 1) * self.column_gap_cm
        )

    @property
    def line_height_pt(self) -> float:
        return self.font_size_pt * self.line_spacing

    @property
    def lines_per_page(self) -> int:
        """How many lines fit between the top and bottom offsets."""
        return int(self.text_height_cm * POINTS_PER_CM // self.line_height_pt)

    @property
    def initials_per_page(self) -> int:
        return self.columns * self.lines_per_page

    def pages_for(self, total_initials: int) -> int:
        """How many pages that many sets of initials take in this layout."""
        return math.ceil(total_initials / self.initials_per_page)

    def column_offsets_cm(self) -> list[float]:
        """Where each column starts, measured from the left offset.

        The block is centred, so width the columns do not use is split
        evenly between the left and the right.
        """
        pitch = self.initials_width_cm + self.column_gap_cm
        padding = (self.text_width_cm - self.width_needed_cm) / 2
        return [padding + column * pitch for column in range(self.columns)]

    def describe(self, total_initials: int) -> str:
        """A short summary, printed before the document is built."""
        return (
            f"Layout: A4, {self.columns} column(s) x {self.lines_per_page} "
            f"line(s) = {self.initials_per_page:,} initials per page\n"
            f"        {self.font_size_pt:g} pt font, {self.margin_cm:g} cm "
            f"offsets, {self.column_gap_cm:g} cm between the initials, "
            f"line spacing {self.line_spacing:g}\n"
            f"        {total_initials:,} initials -> "
            f"{self.pages_for(total_initials):,} page(s)"
        )

    # -- the rules ----------------------------------------------------------

    def validate(self) -> None:
        """Check that this layout can be printed on A4.

        Raises:
            LayoutError: saying what does not fit and what to change.
        """
        self._check_parameter_values()
        self._check_columns_fit_the_width()
        self._check_lines_fit_the_height()

    def _check_parameter_values(self) -> None:
        """Every number has to be in a range that makes sense at all."""
        if self.columns < 1:
            raise LayoutError(
                f"The number of columns must be at least 1, got {self.columns}."
            )
        if self.font_size_pt <= 0 or self.line_spacing <= 0:
            raise LayoutError(
                f"The font size and the line spacing must be greater than 0, "
                f"got {self.font_size_pt:g} pt and {self.line_spacing:g}."
            )
        if self.margin_cm < 0 or self.column_gap_cm < 0:
            raise LayoutError(
                f"The offset and the gap cannot be negative, got "
                f"{self.margin_cm:g} cm and {self.column_gap_cm:g} cm."
            )
        if self.text_width_cm <= 0 or self.text_height_cm <= 0:
            raise LayoutError(
                f"An offset of {self.margin_cm:g} cm on every side leaves no "
                f"room on an A4 page ({PAGE_WIDTH_CM:g} x {PAGE_HEIGHT_CM:g} "
                f"cm). Use less than {PAGE_WIDTH_CM / 2:g} cm."
            )

    def _check_columns_fit_the_width(self) -> None:
        """A full line of initials has to fit across the page."""
        if self.width_needed_cm <= self.text_width_cm:
            return
        fitting = int(
            (self.text_width_cm + self.column_gap_cm)
            // (self.initials_width_cm + self.column_gap_cm)
        )
        hint = (
            f"Use at most {fitting} column(s), a smaller gap, a smaller font "
            f"size, or smaller offsets."
            if fitting >= 1
            else "Not even one column fits — use a smaller font size."
        )
        raise LayoutError(
            f"{self.columns} column(s) of {self.font_size_pt:g} pt initials "
            f"with {self.column_gap_cm:g} cm between them need "
            f"{self.width_needed_cm:.1f} cm, but only {self.text_width_cm:.1f} "
            f"cm are left between the {self.margin_cm:g} cm offsets on A4.\n"
            f"{hint}"
        )

    def _check_lines_fit_the_height(self) -> None:
        """At least one line has to fit down the page."""
        if self.lines_per_page >= 1:
            return
        raise LayoutError(
            f"A line of {self.font_size_pt:g} pt text at line spacing "
            f"{self.line_spacing:g} is {self.line_height_pt / POINTS_PER_CM:.1f}"
            f" cm high, but only {self.text_height_cm:.1f} cm are left between "
            f"the {self.margin_cm:g} cm offsets on A4.\n"
            f"Use a smaller font size, a smaller line spacing, or smaller "
            f"offsets."
        )


# ---------------------------------------------------------------------------
# Step 1 & 2 — from a CSV of names to a list of initials
# ---------------------------------------------------------------------------

def read_two_column_csv(path: str) -> pd.DataFrame:
    """Read a CSV, trying common separators until one yields >1 column."""
    for separator in CSV_SEPARATORS:
        try:
            names = pd.read_csv(path, header=None, sep=separator)
        except FileNotFoundError:
            raise  # a missing file is not a parsing problem — report it
        except Exception:
            continue
        if len(names.columns) > 1:
            print(f"  Read '{path}' using separator {separator!r}.")
            return names
    raise ValueError(
        f"Could not parse '{path}' into at least two columns. "
        f"Tried separators: {CSV_SEPARATORS}"
    )


def names_to_initials(names: pd.DataFrame) -> list[str]:
    """Turn each (first name, last name) row into initials like ``И. И.``.

    Rows with no name at all are dropped; a row with only one name gives a
    single initial. Uses vectorized pandas operations (~20× faster than
    row-by-row iteration on large datasets).
    """
    first_names = names.iloc[:, 0].fillna("").astype(str).str.strip()
    last_names = names.iloc[:, 1].fillna("").astype(str).str.strip()

    first_initial = first_names.str[:1].str.upper()
    last_initial = last_names.str[:1].str.upper()
    has_first = first_initial != ""
    has_last = last_initial != ""

    # np.select takes the first matching condition, so "has_first" below
    # already means "first name but no last name".
    initials = np.select(
        [has_first & has_last, has_first, has_last],
        [first_initial + ". " + last_initial + ".", first_initial + ".",
         last_initial + "."],
        default="",
    )
    return [text for text in initials if text]


# ---------------------------------------------------------------------------
# Step 3 — from a list of initials to an A4 document
# ---------------------------------------------------------------------------

def group_into_lines(initials: list[str], columns: int) -> list[list[str]]:
    """Split the initials into lines of at most ``columns`` entries."""
    return [initials[i:i + columns] for i in range(0, len(initials), columns)]


def apply_page_style(
    doc: DocumentType, layout: Layout, font_name: str = FONT_NAME
) -> None:
    """Set up the A4 page, the font and the column positions.

    The line height is exact and every column gets a tab stop, so a line
    always lands where the layout counted on.
    """
    section = doc.sections[0]
    section.page_width = Cm(PAGE_WIDTH_CM)
    section.page_height = Cm(PAGE_HEIGHT_CM)
    section.left_margin = section.right_margin = Cm(layout.margin_cm)
    section.top_margin = section.bottom_margin = Cm(layout.margin_cm)

    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(layout.font_size_pt)

    paragraph_style = style.paragraph_format
    paragraph_style.space_before = Pt(0)
    paragraph_style.space_after = Pt(0)
    paragraph_style.line_spacing = Pt(layout.line_height_pt)

    column_offsets = layout.column_offsets_cm()
    paragraph_style.left_indent = Cm(column_offsets[0])  # centres the columns
    for offset_cm in column_offsets[1:]:
        paragraph_style.tab_stops.add_tab_stop(Cm(offset_cm))


def build_initials_document(
    initials: list[str], layout: Layout, font_name: str = FONT_NAME
) -> DocumentType:
    """Build the A4 document: one paragraph per line, columns split by tabs.

    Assumes ``layout`` has been validated.
    """
    doc = Document()
    apply_page_style(doc, layout, font_name)

    for line_number, line in enumerate(group_into_lines(initials, layout.columns)):
        paragraph = doc.add_paragraph("\t".join(line))
        # Break exactly where the layout says a page ends, so the printed
        # page count is the reported one.
        if line_number and line_number % layout.lines_per_page == 0:
            paragraph.paragraph_format.page_break_before = True

    return doc


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

def build_arg_parser() -> argparse.ArgumentParser:
    """The command line: one option per configuration value."""
    parser = argparse.ArgumentParser(
        description="Generate an A4 book of initials from a CSV of names.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--input", default=INPUT_CSV, help="CSV file to read")
    parser.add_argument("--output", default=OUTPUT_DOCX, help="Document to write")
    parser.add_argument("--columns", type=int, default=COLUMNS,
                        help="Sets of initials on one line")
    parser.add_argument("--font-size", type=float, default=FONT_SIZE_PT,
                        help="Font size in points")
    parser.add_argument("--font-name", default=FONT_NAME, help="Font family")
    parser.add_argument("--margin", type=float, default=MARGIN_CM,
                        help="Offset from all four page edges, in cm")
    parser.add_argument("--column-gap", type=float, default=COLUMN_GAP_CM,
                        help="Distance between the initials on a line, in cm")
    parser.add_argument("--line-spacing", type=float, default=LINE_SPACING,
                        help="Line height as a multiple of the font size")
    parser.add_argument("--dry-run", action="store_true",
                        help="Only report the layout and the page count")
    return parser


def layout_from_args(args: argparse.Namespace) -> Layout:
    """Turn the parsed command line into a :class:`Layout`."""
    return Layout(
        columns=args.columns,
        font_size_pt=args.font_size,
        margin_cm=args.margin,
        column_gap_cm=args.column_gap,
        line_spacing=args.line_spacing,
    )


def main(argv: list[str] | None = None) -> int:
    """Run the whole workflow. Returns a process exit code."""
    args = build_arg_parser().parse_args(argv)
    started = time.perf_counter()

    print(f"Reading '{args.input}'...")
    try:
        names = read_two_column_csv(args.input)
        layout = layout_from_args(args)
        layout.validate()
    except FileNotFoundError:
        print(f"ERROR: file '{args.input}' not found in this folder.")
        return 1
    except ValueError as error:  # an unreadable CSV, or a layout that cannot fit
        print(f"ERROR: {error}")
        return 1

    initials = names_to_initials(names)
    print(f"  {len(names):,} rows, {len(initials):,} initials extracted.\n")
    print(layout.describe(len(initials)))

    if args.dry_run:
        print("\nDry run — nothing written.")
        return 0

    print(f"\nBuilding '{args.output}'...")
    document = build_initials_document(initials, layout, font_name=args.font_name)
    document.save(args.output)

    print(
        f"Done in {time.perf_counter() - started:.0f}s — "
        f"{layout.pages_for(len(initials)):,} page(s) when printed or "
        f"exported to PDF."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
