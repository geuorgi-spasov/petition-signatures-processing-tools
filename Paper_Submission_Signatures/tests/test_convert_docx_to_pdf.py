"""Tests for convert_docx_to_pdf.py.

Driving Microsoft Word or LibreOffice for real is not something a test
can do, so the backends themselves are not tested here. Everything
around them is: file discovery, lock-file filtering, post-conversion
verification, which backend gets chosen — and the batch-then-fallback
strategy, which takes the converter as an argument and so can be handed
one that fails on purpose.
"""
from __future__ import annotations

import os
import sys
import types

import pytest

import convert_docx_to_pdf as converter
from convert_docx_to_pdf import (
    convert_pending,
    discover_conversion_jobs,
    is_word_lock_file,
    pdf_name_for,
    verify_conversion_results,
)


class Folders:
    """An input folder of .docx files and an output folder of PDFs.

    Every discovery test needs the same two folders and spent four lines
    making them, which buried the one line that was the point of the
    test. Here the folders exist as soon as the fixture runs, and a test
    says only what it puts in them.
    """

    def __init__(self, root):
        self.input = root / "in"
        self.output = root / "out"
        self.input.mkdir()
        self.output.mkdir()

    def add_docx(self, *names):
        """Create empty .docx files in the input folder."""
        for name in names:
            (self.input / name).touch()

    def add_pdf(self, *names):
        """Create empty PDFs in the output folder, as if already converted."""
        for name in names:
            (self.output / name).touch()

    def discover(self):
        """``(pending, already_converted, lock_files)`` for these folders."""
        return discover_conversion_jobs(str(self.input), str(self.output))


@pytest.fixture
def folders(tmp_path):
    return Folders(tmp_path)


# ---------------------------------------------------------------------------
# is_word_lock_file
# ---------------------------------------------------------------------------

class TestIsWordLockFile:
    def test_normal_docx_is_not_lock(self):
        assert is_word_lock_file("Папка 1.docx") is False

    def test_tilde_dollar_prefix_is_lock(self):
        assert is_word_lock_file("~$пка 1.docx") is True

    def test_tilde_without_dollar_is_not_lock(self):
        assert is_word_lock_file("~something.docx") is False

    def test_empty_string(self):
        assert is_word_lock_file("") is False


# ---------------------------------------------------------------------------
# discover_conversion_jobs
# ---------------------------------------------------------------------------

class TestDiscoverConversionJobs:
    def test_empty_input_folder(self, folders):
        assert folders.discover() == ([], [], [])

    def test_all_pending_when_output_is_empty(self, folders):
        folders.add_docx("a.docx", "b.docx", "c.docx")
        assert folders.discover() == (["a.docx", "b.docx", "c.docx"], [], [])

    def test_skips_files_with_existing_pdf(self, folders):
        folders.add_docx("a.docx", "b.docx", "c.docx")
        folders.add_pdf("b.pdf")
        assert folders.discover() == (["a.docx", "c.docx"], ["b.docx"], [])

    def test_ignores_non_docx_files(self, folders):
        folders.add_docx("real.docx", "readme.txt", "image.png", ".hidden")
        assert folders.discover() == (["real.docx"], [], [])

    def test_handles_uppercase_docx_extension(self, folders):
        folders.add_docx("MIXED.DOCX")
        pending, _, _ = folders.discover()
        assert pending == ["MIXED.DOCX"]

    def test_results_are_sorted(self, folders):
        folders.add_docx("zebra.docx", "alpha.docx", "mango.docx")
        pending, _, _ = folders.discover()
        assert pending == ["alpha.docx", "mango.docx", "zebra.docx"]

    def test_handles_bulgarian_filenames(self, folders):
        folders.add_docx("Папка 1 с подписи от 1 до 1000.docx",
                         "Папка 2 с подписи от 1001 до 2000.docx")
        folders.add_pdf("Папка 1 с подписи от 1 до 1000.pdf")
        assert folders.discover() == (
            ["Папка 2 с подписи от 1001 до 2000.docx"],
            ["Папка 1 с подписи от 1 до 1000.docx"],
            [],
        )

    def test_filters_out_word_lock_files(self, folders):
        """Lock files like ``~$пка 11.docx`` must be excluded."""
        folders.add_docx("Папка 11.docx", "~$пка 11.docx",
                         "~$some_other_doc.docx")
        pending, done, locks = folders.discover()
        assert pending == ["Папка 11.docx"]
        assert done == []
        assert sorted(locks) == ["~$some_other_doc.docx", "~$пка 11.docx"]


# ---------------------------------------------------------------------------
# verify_conversion_results
# ---------------------------------------------------------------------------

class TestVerifyConversionResults:
    def test_all_successful(self, folders):
        folders.add_pdf("a.pdf", "b.pdf")
        successful, failed = verify_conversion_results(
            ["a.docx", "b.docx"], str(folders.output)
        )
        assert successful == ["a.docx", "b.docx"]
        assert failed == []

    def test_all_failed(self, folders):
        successful, failed = verify_conversion_results(
            ["a.docx", "b.docx"], str(folders.output)
        )
        assert successful == []
        assert failed == ["a.docx", "b.docx"]

    def test_mixed(self, folders):
        folders.add_pdf("a.pdf", "c.pdf")
        successful, failed = verify_conversion_results(
            ["a.docx", "b.docx", "c.docx"], str(folders.output)
        )
        assert successful == ["a.docx", "c.docx"]
        assert failed == ["b.docx"]

    def test_handles_bulgarian_filenames(self, folders):
        folders.add_pdf("Папка 1 с подписи от 1 до 1000.pdf")
        successful, failed = verify_conversion_results(
            [
                "Папка 1 с подписи от 1 до 1000.docx",
                "Папка 2 с подписи от 1001 до 2000.docx",
            ],
            str(folders.output),
        )
        assert successful == ["Папка 1 с подписи от 1 до 1000.docx"]
        assert failed == ["Папка 2 с подписи от 1001 до 2000.docx"]


# ---------------------------------------------------------------------------
# pdf_name_for
# ---------------------------------------------------------------------------

class TestPdfNameFor:
    @pytest.mark.parametrize("docx,pdf", [
        ("Папка 1.docx", "Папка 1.pdf"),
        ("MIXED.DOCX", "MIXED.pdf"),
        ("two.dots.docx", "two.dots.pdf"),
    ])
    def test_the_extension_is_replaced_and_the_name_kept(self, docx, pdf):
        assert pdf_name_for(docx) == pdf


# ---------------------------------------------------------------------------
# convert_pending — the batch, and the fallback when the batch fails
# ---------------------------------------------------------------------------

class FakeConverter:
    """A stand-in for Word or LibreOffice that records how it was called.

    The real backends need an installed office suite, which is why this
    path had no test until it was lifted out of main(). This one writes
    empty PDFs where the real one would write real ones, and can be told
    to fail on the batch call, on one named file, or on both.
    """

    def __init__(self, fail_batch=False, corrupt=()):
        self.fail_batch = fail_batch
        self.corrupt = set(corrupt)
        self.calls: list[str] = []

    def __call__(self, source: str, destination: str) -> None:
        self.calls.append(source)
        if os.path.isdir(source):
            if self.fail_batch:
                raise RuntimeError("Word stopped responding")
            for name in os.listdir(source):
                open(os.path.join(destination, pdf_name_for(name)), "w").close()
            return
        if os.path.basename(source) in self.corrupt:
            raise RuntimeError("file is corrupt")
        open(destination, "w").close()


class TestConvertPending:
    def _run(self, folders, convert, names):
        convert_pending(convert, list(names),
                        str(folders.input), str(folders.output))

    def _pdfs(self, folders):
        return sorted(path.name for path in folders.output.iterdir())

    def test_one_batch_call_converts_everything(self, folders):
        folders.add_docx("a.docx", "b.docx", "c.docx")
        convert = FakeConverter()

        self._run(folders, convert, ["a.docx", "b.docx", "c.docx"])

        assert len(convert.calls) == 1, "the batch should be a single call"
        assert self._pdfs(folders) == ["a.pdf", "b.pdf", "c.pdf"]

    def test_a_failed_batch_retries_each_file(self, folders):
        folders.add_docx("a.docx", "b.docx", "c.docx")
        convert = FakeConverter(fail_batch=True)

        self._run(folders, convert, ["a.docx", "b.docx", "c.docx"])

        # one batch attempt, then one call per file
        assert len(convert.calls) == 4
        assert self._pdfs(folders) == ["a.pdf", "b.pdf", "c.pdf"]

    def test_one_bad_file_does_not_block_the_others(self, folders):
        # The whole reason the fallback exists: a batch is all-or-nothing,
        # so a single unreadable document would otherwise cost the lot.
        folders.add_docx("a.docx", "b.docx", "c.docx")
        convert = FakeConverter(fail_batch=True, corrupt=["b.docx"])

        self._run(folders, convert, ["a.docx", "b.docx", "c.docx"])

        assert self._pdfs(folders) == ["a.pdf", "c.pdf"]

    def test_the_originals_are_only_read(self, folders):
        # Staging exists so the converter never opens the input folder;
        # on Windows it would leave lock files in it if it did.
        folders.add_docx("a.docx", "b.docx")
        convert = FakeConverter()

        self._run(folders, convert, ["a.docx", "b.docx"])

        assert sorted(path.name for path in folders.input.iterdir()) == [
            "a.docx", "b.docx"
        ]

    def test_the_converter_is_handed_a_staging_copy_not_the_input(self, folders):
        folders.add_docx("a.docx")
        convert = FakeConverter()

        self._run(folders, convert, ["a.docx"])

        assert convert.calls[0] != str(folders.input)

    def test_the_staging_folder_is_cleaned_up(self, folders):
        folders.add_docx("a.docx")
        convert = FakeConverter()

        self._run(folders, convert, ["a.docx"])

        assert not os.path.exists(convert.calls[0])

    def test_nothing_pending_converts_nothing(self, folders):
        convert = FakeConverter()

        self._run(folders, convert, [])

        # The batch call still happens, on an empty folder, and the real
        # backends treat that as a no-op rather than an error.
        assert self._pdfs(folders) == []


# ---------------------------------------------------------------------------
# Choosing a conversion backend
# ---------------------------------------------------------------------------

class TestFindWord:
    def test_finds_word_where_the_windows_installer_puts_it(
        self, monkeypatch, tmp_path
    ):
        word = tmp_path / "Microsoft Office" / "root" / "Office16" / "WINWORD.EXE"
        word.parent.mkdir(parents=True)
        word.touch()
        monkeypatch.setattr(sys, "platform", "win32")
        monkeypatch.setattr(converter, "_program_files_dirs", lambda: [str(tmp_path)])
        assert converter._find_word() == str(word)

    def test_the_path_it_returns_uses_one_kind_of_separator(self):
        # Regression: the search patterns used to be single strings with
        # "/" inside them. os.path.join keeps those as-is, so on Windows
        # the result came back as "...\\Microsoft Office/root\\Office16
        # \\WINWORD.EXE" - found, but not equal to the same path spelled
        # natively. This checks the joining rule itself under Windows
        # semantics, so it fails on Linux too if the patterns regress.
        import ntpath
        for parts in converter.WORD_PATTERNS:
            joined = ntpath.join(r"C:\Program Files", *parts)
            assert "/" not in joined, joined

    def test_finds_word_where_the_macos_installer_puts_it(self, monkeypatch):
        # The only branch that is not a Program Files search: on macOS
        # Word is one fixed bundle path, so there is nothing to glob.
        app = "/Applications/Microsoft Word.app"
        monkeypatch.setattr(sys, "platform", "darwin")
        monkeypatch.setattr(converter.os.path, "exists",
                            lambda path: path == app)
        assert converter._find_word() == app

    def test_no_word_on_macos_without_it(self, monkeypatch):
        monkeypatch.setattr(sys, "platform", "darwin")
        monkeypatch.setattr(converter.os.path, "exists", lambda path: False)
        assert converter._find_word() is None

    def test_no_word_on_linux(self, monkeypatch):
        monkeypatch.setattr(sys, "platform", "linux")
        assert converter._find_word() is None


class TestFindLibreOffice:
    def test_finds_libreoffice_where_the_windows_installer_puts_it(
        self, monkeypatch, tmp_path
    ):
        soffice = tmp_path / "LibreOffice" / "program" / "soffice.exe"
        soffice.parent.mkdir(parents=True)
        soffice.touch()
        monkeypatch.setattr(converter.shutil, "which", lambda name: None)
        monkeypatch.setattr(converter, "_program_files_dirs", lambda: [str(tmp_path)])
        assert converter._find_libreoffice() == str(soffice)

    def test_prefers_whatever_is_on_the_path(self, monkeypatch):
        monkeypatch.setattr(converter.shutil, "which",
                            lambda name: "/usr/bin/soffice")
        assert converter._find_libreoffice() == "/usr/bin/soffice"


class TestGetConverter:
    def test_prefers_word_when_it_is_installed(self, monkeypatch):
        fake_docx2pdf = types.ModuleType("docx2pdf")
        fake_docx2pdf.convert = lambda src, dst=None: None
        monkeypatch.setitem(sys.modules, "docx2pdf", fake_docx2pdf)
        monkeypatch.setattr(converter, "_find_word", lambda: r"C:\WINWORD.EXE")
        assert converter.get_converter() is fake_docx2pdf.convert

    def test_falls_back_to_libreoffice_when_word_is_missing(self, monkeypatch):
        monkeypatch.setattr(converter, "_find_word", lambda: None)
        monkeypatch.setattr(converter, "_find_libreoffice", lambda: "/usr/bin/soffice")
        assert converter.get_converter() is converter._libreoffice_convert

    def test_windows_without_either_names_both_options(self, monkeypatch, capsys):
        monkeypatch.setattr(sys, "platform", "win32")
        monkeypatch.setattr(converter, "_find_word", lambda: None)
        monkeypatch.setattr(converter, "_find_libreoffice", lambda: None)
        with pytest.raises(SystemExit):
            converter.get_converter()
        message = capsys.readouterr().out
        assert "Microsoft Word" in message
        assert "libreoffice.org" in message
        assert "office.com" in message

    def test_linux_without_libreoffice_says_how_to_install_it(
        self, monkeypatch, capsys
    ):
        monkeypatch.setattr(sys, "platform", "linux")
        monkeypatch.setattr(converter, "_find_word", lambda: None)
        monkeypatch.setattr(converter, "_find_libreoffice", lambda: None)
        with pytest.raises(SystemExit):
            converter.get_converter()
        assert "apt install libreoffice" in capsys.readouterr().out


class TestChoosingTheConverterByHand:
    def test_libreoffice_can_be_insisted_on_even_when_word_is_there(
        self, monkeypatch
    ):
        monkeypatch.setattr(converter, "_find_word", lambda: r"C:\WINWORD.EXE")
        monkeypatch.setattr(converter, "_find_libreoffice", lambda: "/usr/bin/soffice")
        assert converter.get_converter("libreoffice") is converter._libreoffice_convert

    def test_word_can_be_insisted_on_even_when_libreoffice_is_there(
        self, monkeypatch
    ):
        fake_docx2pdf = types.ModuleType("docx2pdf")
        fake_docx2pdf.convert = lambda src, dst=None: None
        monkeypatch.setitem(sys.modules, "docx2pdf", fake_docx2pdf)
        monkeypatch.setattr(converter, "_find_word", lambda: r"C:\WINWORD.EXE")
        monkeypatch.setattr(converter, "_find_libreoffice", lambda: "/usr/bin/soffice")
        assert converter.get_converter("word") is fake_docx2pdf.convert

    def test_asking_for_word_without_word_says_so(self, monkeypatch, capsys):
        monkeypatch.setattr(converter, "_find_word", lambda: None)
        monkeypatch.setattr(converter, "_find_libreoffice", lambda: "/usr/bin/soffice")
        with pytest.raises(SystemExit):
            converter.get_converter("word")
        assert "--word was asked for" in capsys.readouterr().out

    def test_asking_for_libreoffice_without_it_says_so(self, monkeypatch, capsys):
        monkeypatch.setattr(converter, "_find_word", lambda: r"C:\WINWORD.EXE")
        monkeypatch.setattr(converter, "_find_libreoffice", lambda: None)
        with pytest.raises(SystemExit):
            converter.get_converter("libreoffice")
        assert "libreoffice.org" in capsys.readouterr().out

    @pytest.mark.parametrize(
        "argv,expected",
        [([], None), (["--word"], "word"), (["--libreoffice"], "libreoffice")],
    )
    def test_the_command_line_maps_to_a_preference(self, argv, expected):
        assert converter.build_arg_parser().parse_args(argv).prefer == expected

    def test_both_options_at_once_is_refused(self):
        with pytest.raises(SystemExit):
            converter.build_arg_parser().parse_args(["--word", "--libreoffice"])
