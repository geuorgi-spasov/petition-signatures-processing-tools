"""
Compares how long the document takes to build with different numbers of
columns, and how many pages each one produces.

One line of the document is one paragraph, so 6 columns means 6× fewer
paragraphs than 1 column — and the number of paragraphs is what the build
time depends on.

6 is the most that fits on A5 in Oswald, so it is the last of the defaults
below. Asking for more is not an error: the layout refuses it and says how
much width it would have needed, which is the useful answer.

Usage:
    python tests/benchmark_initials_book.py
    python tests/benchmark_initials_book.py --limit 20000
    python tests/benchmark_initials_book.py --columns 1 6 10 --column-gap 1.0
"""

from __future__ import annotations

import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from generate_initials_book import (  # noqa: E402
    COLUMN_GAP_CM,
    INPUT_CSV,
    Layout,
    LayoutError,
    build_initials_document,
    group_into_lines,
    names_to_initials,
    read_names_csv,
)


def load_initials(path: str, limit: int | None) -> list[str]:
    """Read the initials from ``path``, or make some up if it is missing."""
    if os.path.isfile(path):
        initials = names_to_initials(read_names_csv(path))
    else:
        print(f"  '{path}' not found — using generated names.")
        letters = [chr(code) for code in range(0x410, 0x430)]
        initials = [
            f"{letters[i % 32]}. {letters[(i * 7) % 32]}."
            for i in range(limit or 20_000)
        ]
    return initials[:limit] if limit else initials


def main() -> int:
    parser = argparse.ArgumentParser(description="Benchmark the column layouts.")
    parser.add_argument("--input", default=INPUT_CSV, help="CSV of names")
    parser.add_argument("--limit", type=int, help="Use only the first N initials")
    parser.add_argument("--columns", type=int, nargs="+", default=[1, 3, 5, 6],
                        help="Column counts to compare")
    parser.add_argument("--column-gap", type=float, default=COLUMN_GAP_CM,
                        help="Distance between the initials, in cm")
    args = parser.parse_args()

    initials = load_initials(args.input, args.limit)
    print(f"\nBuilding {len(initials):,} initials in each layout.\n")
    print(f"{'columns':>8} {'lines':>9} {'pages':>7} {'time':>8}")

    for columns in args.columns:
        layout = Layout(columns=columns, column_gap_cm=args.column_gap)
        try:
            layout.validate()
        except LayoutError as error:
            print(f"{columns:>8} skipped — {str(error).splitlines()[0]}")
            continue

        start = time.perf_counter()
        build_initials_document(initials, layout)
        seconds = time.perf_counter() - start
        lines = len(group_into_lines(initials, columns))
        pages = layout.pages_for(len(initials))
        print(f"{columns:>8} {lines:>9,} {pages:>7,} {seconds:>7.1f}s")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
