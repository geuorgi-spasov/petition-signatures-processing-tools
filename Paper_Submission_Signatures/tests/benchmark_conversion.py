"""
Benchmarks two .docx → .pdf conversion strategies and reports which is
faster on your machine.

Strategy 1 (per-file):  call ``convert(file, pdf)`` once per file
Strategy 2 (batch):     call ``convert(input_dir, output_dir)`` once

Both strategies use the same conversion backend as the main workflow,
chosen automatically for your platform: Microsoft Word (via docx2pdf)
on Windows and macOS, and LibreOffice in headless mode on Linux.

Both strategies operate on identical inputs (a staged copy of the same
.docx files) and write into fresh temporary folders, so the comparison
is fair. The actual speedup depends on your platform — the converter's
startup overhead is the dominant cost, and that's what batch mode
amortises.

Usage:
    python tests/benchmark_conversion.py
    python tests/benchmark_conversion.py --input-folder some_other_folder
    python tests/benchmark_conversion.py --limit 10
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from convert_docx_to_pdf import INPUT_FOLDER as DEFAULT_INPUT_FOLDER  # noqa: E402
from convert_docx_to_pdf import get_converter  # noqa: E402


def time_per_file_conversion(docx_dir: str, pdf_dir: str) -> float:
    """Convert each file with its own ``convert()`` call. Returns seconds."""
    convert = get_converter()

    files = sorted(f for f in os.listdir(docx_dir) if f.lower().endswith(".docx"))
    start = time.perf_counter()
    for filename in files:
        convert(
            os.path.join(docx_dir, filename),
            os.path.join(pdf_dir, filename[:-5] + ".pdf"),
        )
    return time.perf_counter() - start


def time_batch_conversion(docx_dir: str, pdf_dir: str) -> float:
    """Convert the whole directory in one ``convert()`` call. Returns seconds."""
    convert = get_converter()

    start = time.perf_counter()
    convert(docx_dir, pdf_dir)
    return time.perf_counter() - start


def stage_inputs(source_dir: str, dest_dir: str, files: list[str]) -> None:
    """Copy ``files`` from ``source_dir`` into ``dest_dir``."""
    for filename in files:
        shutil.copy(
            os.path.join(source_dir, filename), os.path.join(dest_dir, filename)
        )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Benchmark per-file vs batch .docx → .pdf conversion."
    )
    parser.add_argument(
        "--input-folder",
        default=DEFAULT_INPUT_FOLDER,
        help=f"Folder containing .docx files (default: {DEFAULT_INPUT_FOLDER!r})",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Only benchmark with the first N files (useful for quick checks)",
    )
    args = parser.parse_args()

    # Validate that a conversion backend is available for this platform
    # (docx2pdf/Word on Windows/macOS, LibreOffice on Linux). This exits
    # with a clear message if neither is present.
    get_converter()

    if not os.path.isdir(args.input_folder):
        print(f"ERROR: folder '{args.input_folder}' not found.")
        sys.exit(1)

    all_files = sorted(
        f for f in os.listdir(args.input_folder) if f.lower().endswith(".docx")
    )
    if not all_files:
        print(f"No .docx files in '{args.input_folder}'.")
        sys.exit(1)

    files = all_files[: args.limit] if args.limit else all_files

    print(
        f"Benchmarking with {len(files)} of {len(all_files)} .docx file(s) "
        f"from '{args.input_folder}'.\n"
    )

    with tempfile.TemporaryDirectory() as staging:
        stage_inputs(args.input_folder, staging, files)

        # Strategy 1
        with tempfile.TemporaryDirectory() as out1:
            print("Strategy 1: per-file convert() calls...")
            t_per_file = time_per_file_conversion(staging, out1)
            print(
                f"  Total: {t_per_file:.1f}s — "
                f"{t_per_file / len(files):.2f}s per file"
            )

        # Strategy 2
        with tempfile.TemporaryDirectory() as out2:
            print("\nStrategy 2: single batch convert() call...")
            t_batch = time_batch_conversion(staging, out2)
            print(
                f"  Total: {t_batch:.1f}s — "
                f"{t_batch / len(files):.2f}s per file"
            )

    speedup = t_per_file / max(t_batch, 1e-9)
    print(
        f"\nBatch is {speedup:.1f}× faster — "
        f"saved {t_per_file - t_batch:.1f}s on {len(files)} file(s)."
    )


if __name__ == "__main__":
    main()
