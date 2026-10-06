"""
Converts every .docx file inside INPUT_FOLDER to a PDF in OUTPUT_FOLDER,
or wherever --input and --output point instead,
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
from collections.abc import Callable

# What every conversion backend looks like from the outside: it takes a
# source and a destination, and it accepts either two file paths or two
# folder paths. docx2pdf works both ways and _libreoffice_convert below
# was written to match, which is what makes one batch call possible.
Converter = Callable[[str, str], None]

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


def pdf_name_for(docx_filename: str) -> str:
    """The PDF a .docx becomes: the same name, with a .pdf extension."""
    return os.path.splitext(docx_filename)[0] + ".pdf"


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
        pdf_path = os.path.join(output_folder, pdf_name_for(filename))
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
        if os.path.exists(os.path.join(output_folder,
                                        pdf_name_for(filename))):
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


# Where the Windows installer puts Word, as path components rather than
# one "a/b/c" string - see the note inside _find_word.
WORD_PATTERNS = (
    ("Microsoft Office", "root", "Office*", "WINWORD.EXE"),
    ("Microsoft Office", "Office*", "WINWORD.EXE"),
)


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

    # Path components, not one "a/b/c" string. os.path.join leaves the
    # inside of a string alone, so the old single-string patterns came
    # back from Windows as "...\Microsoft Office/root\Office16\
    # WINWORD.EXE" - glob found Word, but the path did not equal the same
    # path spelled natively, which is what failed the test for it.
    for directory in _program_files_dirs():
        for parts in WORD_PATTERNS:
            matches = glob.glob(os.path.join(directory, *parts))
            if matches:
                return os.path.normpath(matches[0])
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
    """The command line: which folders, and which program converts them."""
    parser = argparse.ArgumentParser(
        description=(
            f"Convert every .docx in '{INPUT_FOLDER}' to a PDF in "
            f"'{OUTPUT_FOLDER}'. Already-converted files are skipped."
        )
    )
    # The splitting step takes --output, so this step has to take --input:
    # send the first one somewhere else and the second has to be able to
    # follow it. The defaults are the two folders the pair agrees on.
    parser.add_argument("--input", default=INPUT_FOLDER,
                        help=f"Folder of .docx files (default: "
                             f"{INPUT_FOLDER!r})")
    parser.add_argument("--output", default=OUTPUT_FOLDER,
                        help=f"Folder to write the PDFs into (default: "
                             f"{OUTPUT_FOLDER!r})")
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


def report_lock_files(lock_files: list[str]) -> None:
    """Say which files were skipped because Word still has them open."""
    if not lock_files:
        return
    print(f"Skipping {len(lock_files)} Word lock file(s) — "
          f"close Microsoft Word if you still have documents open:")
    for name in lock_files:
        print(f"  - {name}")
    print()


def convert_pending(
    convert: Converter, pending: list[str],
    input_folder: str, output_folder: str,
) -> None:
    """Convert ``pending`` into ``output_folder``, in one batch if it can.

    The files are copied into a single staging folder first so the
    converter can be handed one directory and started once. That is the
    difference between minutes and an hour: the startup cost of Word or
    LibreOffice dominates, and a batch pays it once.

    Both strategies read from that same staging copy, which is why the
    fallback runs here, inside the staging folder's lifetime, rather than
    being called separately by main().
    """
    with tempfile.TemporaryDirectory() as staging:
        for filename in pending:
            shutil.copy(os.path.join(input_folder, filename),
                        os.path.join(staging, filename))

        if not _convert_as_one_batch(convert, len(pending), staging,
                                     output_folder):
            _convert_one_at_a_time(convert, pending, staging, output_folder)


def _convert_as_one_batch(
    convert: Converter, count: int, staging: str, output_folder: str,
) -> bool:
    """Convert the whole staging folder in one call. False if it failed."""
    print(f"Converting {count} file(s) in a single batch "
          f"(Word / LibreOffice opens once)...")
    try:
        convert(staging, output_folder)
        return True
    except Exception as exc:
        print(f"\nBatch conversion failed: {exc}")
        return False


def _convert_one_at_a_time(
    convert: Converter, pending: list[str], staging: str, output_folder: str,
) -> None:
    """Convert each file alone, so one bad file blocks only itself."""
    print("Retrying each file individually...\n")
    for filename in pending:
        source = os.path.join(staging, filename)
        destination = os.path.join(output_folder, pdf_name_for(filename))
        try:
            convert(source, destination)
            print(f"  OK: {filename}")
        except Exception as exc:
            print(f"  FAIL: {filename} — {exc}")


def report_results(
    pending: list[str], already_converted: list[str],
    output_folder: str, elapsed: float,
) -> None:
    """Say how it went: what converted, what was skipped, what failed."""
    successful, failed = verify_conversion_results(pending, output_folder)

    # No per-file figure when nothing converted: dividing the whole run by
    # one file would report the entire elapsed time as the cost of a file
    # that was never produced.
    rate = f" (~{elapsed / len(successful):.2f}s per file)" if successful else ""

    print(f"\nDone in {elapsed:.1f}s{rate}. "
          f"Converted: {len(successful)}, skipped: {len(already_converted)}, "
          f"failed: {len(failed)}.")
    if failed:
        print("Failed files:")
        for filename in failed:
            print(f"  - {filename}")
    print(f"PDFs are in '{output_folder}/'.")


def main(argv: list[str] | None = None) -> None:
    args = build_arg_parser().parse_args(argv)
    convert = get_converter(args.prefer)

    if not os.path.isdir(args.input):
        print(f"ERROR: folder '{args.input}' not found.")
        print("Run 'split_signatures_into_folders.py' first, or create the "
              "folder and put .docx files inside it.")
        sys.exit(1)

    os.makedirs(args.output, exist_ok=True)

    pending, already_converted, lock_files = discover_conversion_jobs(
        args.input, args.output
    )
    report_lock_files(lock_files)

    found = len(pending) + len(already_converted)
    if found == 0:
        print(f"No .docx files found in '{args.input}'.")
        return
    if not pending:
        print(f"All {found} file(s) in '{args.input}' are already converted. "
              f"Nothing to do.")
        return

    print(f"Found {found} .docx file(s) in '{args.input}' "
          f"({len(already_converted)} already converted, "
          f"{len(pending)} pending).\n")

    start = time.perf_counter()
    convert_pending(convert, pending, args.input, args.output)
    report_results(pending, already_converted, args.output,
                   time.perf_counter() - start)


if __name__ == "__main__":
    main()
