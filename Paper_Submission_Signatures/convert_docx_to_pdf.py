"""
Converts every .docx file inside INPUT_FOLDER to a PDF in OUTPUT_FOLDER,
skipping already-converted files and Word lock files. Safe to re-run.

Performance: pending files are converted in a single batch so that
Microsoft Word (or LibreOffice) is started just once instead of once per
file. If the batch fails (e.g. a corrupt file), the script automatically
falls back to per-file conversion so that one bad file doesn't block the
rest.

Usage:
    python convert_docx_to_pdf.py

See README.md for the full workflow.
"""

from __future__ import annotations

import argparse
import glob
import os
import shutil
import sys
import tempfile
import time

# Note: ``docx2pdf`` is imported lazily inside ``main()`` so that this module
# can be imported (and its pure functions unit-tested) on machines without
# Microsoft Word or LibreOffice installed.

# ---------------------------------------------------------------------------
# Configuration — change these if your folder names differ
# ---------------------------------------------------------------------------

INPUT_FOLDER = "signatures_docx"
OUTPUT_FOLDER = "signatures_pdf"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def is_word_lock_file(filename: str) -> bool:
    """Return True for Word's ``~$<name>.docx`` lock files.

    Word creates a small hidden file whose name starts with ``~$`` whenever
    a document is open. These are not valid .docx files and ``docx2pdf``
    will throw a "file appears to be corrupted" error if it tries to
    convert one.
    """
    return filename.startswith("~$")


def discover_conversion_jobs(
    input_folder: str, output_folder: str
) -> tuple[list[str], list[str], list[str]]:
    """Inspect ``input_folder`` and classify its .docx files.

    Returns:
        ``(pending, already_converted, skipped_lock_files)`` — three lists
        of filenames (not full paths).
    """
    pending: list[str] = []
    already_converted: list[str] = []
    skipped_lock_files: list[str] = []

    for filename in sorted(os.listdir(input_folder)):
        if not filename.lower().endswith(".docx"):
            continue
        if is_word_lock_file(filename):
            skipped_lock_files.append(filename)
            continue
        pdf_path = os.path.join(output_folder, filename[:-5] + ".pdf")
        if os.path.exists(pdf_path):
            already_converted.append(filename)
        else:
            pending.append(filename)

    return pending, already_converted, skipped_lock_files


def verify_conversion_results(
    pending: list[str], output_folder: str
) -> tuple[list[str], list[str]]:
    """Split ``pending`` into (successful, failed) based on output presence.

    After a batch conversion, the only way to know which files actually
    succeeded is to check which .pdf files now exist in ``output_folder``.
    """
    successful: list[str] = []
    failed: list[str] = []
    for filename in pending:
        pdf_name = filename[:-5] + ".pdf"
        if os.path.exists(os.path.join(output_folder, pdf_name)):
            successful.append(filename)
        else:
            failed.append(filename)
    return successful, failed


# ---------------------------------------------------------------------------
# Conversion backends (cross-platform)
# ---------------------------------------------------------------------------
#
# Microsoft Word gives the best fidelity, so it is preferred wherever it is
# installed; the ``docx2pdf`` package drives it on Windows and macOS. When
# Word is not there - always on Linux, and on any Windows or Mac without the
# desktop Word application - LibreOffice is driven in headless mode instead.
# ``get_converter()`` picks a backend and returns a ``convert(src, dst)``
# style callable with the same folder-in / folder-out contract docx2pdf
# uses, so the rest of ``main()`` doesn't need to care which backend runs.


def _program_files_dirs() -> list[str]:
    """The Program Files folders on Windows, whatever drive they are on."""
    found = [
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
    ]
    return [directory for directory in found if directory]


def _find_word() -> str | None:
    """Return the path to the Microsoft Word application, or None.

    Only the installed desktop application counts - ``docx2pdf`` automates
    it, so a Microsoft 365 subscription used only through office.com in a
    browser cannot be used.
    """
    if sys.platform == "darwin":
        app = "/Applications/Microsoft Word.app"
        return app if os.path.exists(app) else None
    if not sys.platform.startswith("win"):
        return None

    for directory in _program_files_dirs():
        for pattern in ("Microsoft Office/root/Office*/WINWORD.EXE",
                        "Microsoft Office/Office*/WINWORD.EXE"):
            matches = glob.glob(os.path.join(directory, pattern))
            if matches:
                return matches[0]
    return shutil.which("winword")


def _find_libreoffice() -> str | None:
    """Return the path to the LibreOffice CLI, or None if not found."""
    for name in ("libreoffice", "soffice"):
        path = shutil.which(name)
        if path:
            return path

    # On Windows and macOS LibreOffice is not normally on the PATH, so look
    # where its installer puts it.
    candidates = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    candidates += [
        os.path.join(directory, "LibreOffice", "program", "soffice.exe")
        for directory in _program_files_dirs()
    ]
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return None


def _libreoffice_convert(src: str, dst: str) -> None:
    """Convert a .docx (file) or folder of .docx to PDF using LibreOffice.

    Mirrors docx2pdf's dual behaviour: if ``src`` is a directory, every
    .docx inside it is converted into ``dst`` (also a directory); if
    ``src`` is a single file, it is converted to the single file ``dst``.
    """
    import subprocess

    soffice = _find_libreoffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice not found. Install it (e.g. `sudo apt install "
            "libreoffice`) so .docx files can be converted to PDF on Linux."
        )

    if os.path.isdir(src):
        out_dir = dst
        os.makedirs(out_dir, exist_ok=True)
        docx_files = [
            os.path.join(src, f)
            for f in sorted(os.listdir(src))
            if f.lower().endswith(".docx") and not is_word_lock_file(f)
        ]
        if not docx_files:
            return
        # One invocation converts the whole batch — LibreOffice starts once.
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir",
             out_dir, *docx_files],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    else:
        out_dir = os.path.dirname(dst) or "."
        os.makedirs(out_dir, exist_ok=True)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir",
             out_dir, src],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        # LibreOffice names the output after the input stem; rename if the
        # caller asked for a specific destination name.
        produced = os.path.join(
            out_dir, os.path.basename(src)[:-5] + ".pdf"
        )
        if produced != dst and os.path.exists(produced):
            os.replace(produced, dst)


def _word_converter():
    """Return docx2pdf's ``convert``, which drives the installed Word."""
    try:
        from docx2pdf import convert
    except ImportError:
        print(
            "ERROR: the 'docx2pdf' package isn't installed. Run "
            "`pip install -r requirements.txt` and try again."
        )
        sys.exit(1)
    print("  Converting with Microsoft Word.")
    return convert


def get_converter(prefer: str | None = None):
    """Return a ``convert(src, dst)`` callable for this computer.

    Args:
        prefer: ``"word"`` or ``"libreoffice"`` to insist on one of them,
            or None to choose automatically — Word where it is installed,
            LibreOffice otherwise.

    Word is the default because it is the program the .docx format is
    defined by, so its rendering is the reference. LibreOffice is several
    times faster; ``--libreoffice`` is there for when that matters more.

    Exits with an explanation if the chosen program is not available.
    """
    if prefer == "word":
        if _find_word() is None:
            print(
                "ERROR: --word was asked for, but Microsoft Word is not "
                "installed on this computer.\nLeave the option off to use "
                "whatever is available."
            )
            sys.exit(1)
        return _word_converter()

    if prefer == "libreoffice":
        if _find_libreoffice() is None:
            print(
                "ERROR: --libreoffice was asked for, but LibreOffice is not "
                "installed.\nGet it free from "
                "https://www.libreoffice.org/download/, or leave the option "
                "off to use whatever is available."
            )
            sys.exit(1)
        print("  Converting with LibreOffice.")
        return _libreoffice_convert

    if _find_word() is not None:
        return _word_converter()

    if _find_libreoffice() is not None:
        print("  Converting with LibreOffice.")
        return _libreoffice_convert

    if sys.platform.startswith(("win", "darwin")):
        print(
            "ERROR: no program found that can turn .docx files into PDFs.\n"
            "Either install Microsoft Word (the desktop application - the\n"
            "browser version at office.com cannot be used), or install\n"
            "LibreOffice, which is free: https://www.libreoffice.org/download/"
        )
    else:
        print(
            "ERROR: LibreOffice not found. On Linux the conversion uses "
            "LibreOffice in headless mode.\nInstall it with e.g. "
            "`sudo apt install libreoffice` and try again."
        )
    sys.exit(1)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def build_arg_parser() -> argparse.ArgumentParser:
    """The command line: which program should do the converting."""
    parser = argparse.ArgumentParser(
        description=(
            f"Convert every .docx in '{INPUT_FOLDER}' to a PDF in "
            f"'{OUTPUT_FOLDER}'. Already-converted files are skipped."
        )
    )
    choice = parser.add_mutually_exclusive_group()
    choice.add_argument(
        "--word", dest="prefer", action="store_const", const="word",
        help="Convert with Microsoft Word (the default where it is installed)",
    )
    choice.add_argument(
        "--libreoffice", dest="prefer", action="store_const",
        const="libreoffice",
        help="Convert with LibreOffice instead — several times faster",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    args = build_arg_parser().parse_args(argv)
    convert = get_converter(args.prefer)

    if not os.path.isdir(INPUT_FOLDER):
        print(f"ERROR: folder '{INPUT_FOLDER}' not found.")
        print(
            "Run 'split_signatures_into_folders.py' first, or create the "
            "folder and put .docx files inside it."
        )
        sys.exit(1)

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    pending, already_converted, lock_files = discover_conversion_jobs(
        INPUT_FOLDER, OUTPUT_FOLDER
    )
    real_files = len(pending) + len(already_converted)

    if lock_files:
        print(
            f"Skipping {len(lock_files)} Word lock file(s) — "
            f"close Microsoft Word if you still have documents open:"
        )
        for name in lock_files:
            print(f"  - {name}")
        print()

    if real_files == 0:
        print(f"No .docx files found in '{INPUT_FOLDER}'.")
        return

    if not pending:
        print(
            f"All {real_files} file(s) in '{INPUT_FOLDER}' are already "
            f"converted. Nothing to do."
        )
        return

    print(
        f"Found {real_files} .docx file(s) in '{INPUT_FOLDER}' "
        f"({len(already_converted)} already converted, "
        f"{len(pending)} pending).\n"
    )

    start = time.perf_counter()

    # Stage only the pending files in a temp folder, then convert that
    # folder in one shot. This is dramatically faster than calling
    # convert() per file because Word/LibreOffice starts up only once.
    with tempfile.TemporaryDirectory() as staging:
        for filename in pending:
            shutil.copy(
                os.path.join(INPUT_FOLDER, filename),
                os.path.join(staging, filename),
            )

        batch_ok = True
        print(
            f"Converting {len(pending)} file(s) in a single batch "
            f"(Word / LibreOffice opens once)..."
        )
        try:
            convert(staging, OUTPUT_FOLDER)
        except Exception as exc:
            batch_ok = False
            print(f"\nBatch conversion failed: {exc}")

        # If batch mode crashed, fall back to converting each file
        # individually so one bad file doesn't block the rest.
        if not batch_ok:
            print("Retrying each file individually...\n")
            for filename in pending:
                src = os.path.join(staging, filename)
                dst = os.path.join(OUTPUT_FOLDER, filename[:-5] + ".pdf")
                try:
                    convert(src, dst)
                    print(f"  OK: {filename}")
                except Exception as file_exc:
                    print(f"  FAIL: {filename} — {file_exc}")

    elapsed = time.perf_counter() - start
    successful, failed = verify_conversion_results(pending, OUTPUT_FOLDER)

    per_file = elapsed / max(1, len(successful))
    print(
        f"\nDone in {elapsed:.1f}s "
        f"(~{per_file:.2f}s per file). "
        f"Converted: {len(successful)}, skipped: {len(already_converted)}, "
        f"failed: {len(failed)}."
    )
    if failed:
        print("Failed files:")
        for filename in failed:
            print(f"  - {filename}")
    print(f"PDFs are in '{OUTPUT_FOLDER}/'.")


if __name__ == "__main__":
    main()
