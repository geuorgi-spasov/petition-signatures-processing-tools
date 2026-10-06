"""Tests for tests/check_rendering.py — which program it decides to measure.

The check exists because Word and LibreOffice disagree about how many
lines fit on a page, so measuring the wrong one answers the wrong
question. Which one it picks therefore has to be right, and it was not:
find_word() asked whether the docx2pdf package was importable, which it
is on any Windows machine, and reported Word as found on machines that
have no Word at all.
"""
from __future__ import annotations

import os
import sys
import types

import pytest

import check_rendering as checker


@pytest.fixture
def no_libreoffice(monkeypatch):
    """A machine where LibreOffice cannot be found."""
    monkeypatch.setattr(checker.shutil, "which", lambda name: None)
    monkeypatch.setattr(checker.os.path, "exists", lambda path: False)
    monkeypatch.setattr(checker, "program_files_dirs", list)


def _with_docx2pdf(monkeypatch, installed: bool):
    monkeypatch.setattr(checker, "docx2pdf_installed", lambda: installed)


class TestFindWord:
    def test_no_word_on_linux(self, monkeypatch):
        monkeypatch.setattr(sys, "platform", "linux")
        assert checker.find_word() is None

    def test_finds_word_where_the_windows_installer_puts_it(
        self, monkeypatch, tmp_path
    ):
        word = tmp_path / "Microsoft Office" / "root" / "Office16" / "WINWORD.EXE"
        word.parent.mkdir(parents=True)
        word.touch()
        monkeypatch.setattr(sys, "platform", "win32")
        monkeypatch.setattr(checker, "program_files_dirs", lambda: [str(tmp_path)])
        assert checker.find_word() == str(word)

    def test_no_word_on_windows_without_it(self, monkeypatch, tmp_path):
        # The case that was broken: a Windows machine with the package
        # installed and no Word. Word is what is being measured, so the
        # absence of Word has to be the answer.
        monkeypatch.setattr(sys, "platform", "win32")
        monkeypatch.setattr(checker, "program_files_dirs", lambda: [str(tmp_path)])
        monkeypatch.setattr(checker.shutil, "which", lambda name: None)
        assert checker.find_word() is None

    def test_finds_word_where_the_macos_installer_puts_it(self, monkeypatch):
        app = "/Applications/Microsoft Word.app"
        monkeypatch.setattr(sys, "platform", "darwin")
        monkeypatch.setattr(checker.os.path, "exists", lambda path: path == app)
        assert checker.find_word() == app

    def test_the_path_it_returns_uses_one_kind_of_separator(self):
        # Same guard the paper toolkit carries: a pattern written as one
        # string keeps its "/" through os.path.join, so on Windows the
        # result comes back with both separators mixed into it.
        import ntpath
        for parts in checker.WORD_PATTERNS:
            joined = ntpath.join(r"C:\Program Files", *parts)
            assert "/" not in joined, joined


class TestFindConverter:
    def test_asking_for_word_without_word_is_not_satisfied(
        self, monkeypatch, no_libreoffice
    ):
        # The package being importable is not Word being installed.
        _with_docx2pdf(monkeypatch, True)
        monkeypatch.setattr(checker, "find_word", lambda: None)
        assert checker.find_converter("word") is None

    def test_asking_for_word_without_the_package_is_not_satisfied(
        self, monkeypatch, no_libreoffice
    ):
        _with_docx2pdf(monkeypatch, False)
        monkeypatch.setattr(checker, "find_word", lambda: r"C:\WINWORD.EXE")
        assert checker.find_converter("word") is None

    def test_word_is_used_when_both_are_there(self, monkeypatch):
        _with_docx2pdf(monkeypatch, True)
        monkeypatch.setattr(checker, "find_word", lambda: r"C:\WINWORD.EXE")
        assert checker.find_converter("word") == ("word", r"C:\WINWORD.EXE")

    def test_asking_for_libreoffice_without_it_is_not_satisfied(self, monkeypatch):
        monkeypatch.setattr(checker, "find_libreoffice", lambda: None)
        assert checker.find_converter("libreoffice") is None

    def test_unasked_libreoffice_wins_because_it_installs_anywhere(
        self, monkeypatch
    ):
        _with_docx2pdf(monkeypatch, True)
        monkeypatch.setattr(checker, "find_libreoffice", lambda: "/usr/bin/soffice")
        monkeypatch.setattr(checker, "find_word", lambda: r"C:\WINWORD.EXE")
        assert checker.find_converter() == ("libreoffice", "/usr/bin/soffice")

    def test_unasked_word_is_the_fallback(self, monkeypatch):
        _with_docx2pdf(monkeypatch, True)
        monkeypatch.setattr(checker, "find_libreoffice", lambda: None)
        monkeypatch.setattr(checker, "find_word", lambda: r"C:\WINWORD.EXE")
        assert checker.find_converter() == ("word", r"C:\WINWORD.EXE")

    def test_nothing_installed_is_nothing_found(self, monkeypatch):
        _with_docx2pdf(monkeypatch, False)
        monkeypatch.setattr(checker, "find_libreoffice", lambda: None)
        monkeypatch.setattr(checker, "find_word", lambda: None)
        assert checker.find_converter() is None


class TestDocx2pdfInstalled:
    def test_true_when_the_module_imports(self, monkeypatch):
        monkeypatch.setitem(sys.modules, "docx2pdf", types.ModuleType("docx2pdf"))
        assert checker.docx2pdf_installed() is True


class TestCountPdfPages:
    def test_page_objects_are_counted_first(self, tmp_path):
        # /Count is not page-specific - an outline tree carries one too -
        # so a file whose largest /Count is a bookmark count must still
        # report its real page count.
        pdf = tmp_path / "x.pdf"
        pdf.write_bytes(
            b"/Type /Outlines /Count 99\n"
            b"/Type /Pages /Count 3\n"
            b"/Type /Page /Contents\n" * 3
        )
        assert checker.count_pdf_pages(str(pdf)) == 3

    def test_count_is_the_fallback_when_no_page_object_is_visible(self, tmp_path):
        pdf = tmp_path / "x.pdf"
        pdf.write_bytes(b"/Type /Pages /Count 7\n")
        assert checker.count_pdf_pages(str(pdf)) == 7

    def test_an_unreadable_pdf_counts_nothing(self, tmp_path):
        pdf = tmp_path / "x.pdf"
        pdf.write_bytes(b"not a pdf at all")
        assert checker.count_pdf_pages(str(pdf)) == 0
