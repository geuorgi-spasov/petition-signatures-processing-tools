"""Tests for split_signatures_into_folders.py."""
from __future__ import annotations

import pandas as pd
import pytest

from split_signatures_into_folders import (
    DEFAULT_COLUMN_WIDTHS_CM,
    LEFT_MARGIN_CM,
    PAGE_WIDTH_CM,
    RIGHT_MARGIN_CM,
    build_folder_document,
    read_signatures_csv,
    scale_column_widths,
)

AVAILABLE_WIDTH_CM = PAGE_WIDTH_CM - LEFT_MARGIN_CM - RIGHT_MARGIN_CM


# ---------------------------------------------------------------------------
# scale_column_widths
# ---------------------------------------------------------------------------

class TestScaleColumnWidths:
    def test_six_columns_fills_available_width(self):
        widths = scale_column_widths(6)
        assert len(widths) == 6
        assert sum(widths) == pytest.approx(AVAILABLE_WIDTH_CM)

    def test_six_columns_preserves_default_proportions(self):
        widths = scale_column_widths(6)
        ratios = [w / d for w, d in zip(widths, DEFAULT_COLUMN_WIDTHS_CM)]
        for r in ratios[1:]:
            assert r == pytest.approx(ratios[0])

    def test_fewer_columns_uses_first_n_defaults(self):
        widths = scale_column_widths(3)
        assert len(widths) == 3
        assert sum(widths) == pytest.approx(AVAILABLE_WIDTH_CM)

    def test_more_columns_than_defaults_distributes_evenly(self):
        widths = scale_column_widths(8)
        assert len(widths) == 8
        assert all(w == pytest.approx(widths[0]) for w in widths)
        assert sum(widths) == pytest.approx(AVAILABLE_WIDTH_CM)

    def test_one_column_takes_full_width(self):
        widths = scale_column_widths(1)
        assert widths == pytest.approx([AVAILABLE_WIDTH_CM])


# ---------------------------------------------------------------------------
# read_signatures_csv
# ---------------------------------------------------------------------------

class TestReadSignaturesCsv:
    def _write(self, path, content, encoding="utf-8"):
        path.write_text(content, encoding=encoding)

    def test_reads_comma_separated_with_headers(self, tmp_path):
        csv = tmp_path / "sigs.csv"
        self._write(csv, "id,name,city\n1,Иван,София\n2,Петър,Пловдив\n")
        df = read_signatures_csv(str(csv))
        assert list(df.columns) == ["id", "name", "city"]
        assert len(df) == 2

    def test_reads_semicolon_separated(self, tmp_path):
        csv = tmp_path / "sigs.csv"
        self._write(csv, "id;name;city\n1;Иван;София\n2;Петър;Пловдив\n")
        df = read_signatures_csv(str(csv))
        assert len(df) == 2
        assert df.iloc[0]["city"] == "София"

    def test_raises_on_single_column_file(self, tmp_path):
        csv = tmp_path / "broken.csv"
        self._write(csv, "only_one_column\nrow1\nrow2\n")
        with pytest.raises(ValueError):
            read_signatures_csv(str(csv))


# ---------------------------------------------------------------------------
# build_folder_document — smoke test that exercises the whole pipeline
# ---------------------------------------------------------------------------

class TestBuildFolderDocument:
    def _make_chunk(self, n_rows):
        return pd.DataFrame(
            {
                "id": range(1, n_rows + 1),
                "name": [f"Name {i}" for i in range(1, n_rows + 1)],
                "city": ["София"] * n_rows,
                "address": ["Address line " * 3] * n_rows,
                "date": ["2024-01-01"] * n_rows,
                "signature": ["OK"] * n_rows,
            }
        )

    def test_returns_document_and_correct_page_count(self):
        chunk = self._make_chunk(25)  # 25 rows = 3 pages at 10 per page
        header = chunk.columns.tolist()
        widths = scale_column_widths(len(header))

        doc, pages = build_folder_document(1, chunk, header, widths)

        assert pages == 3
        # One table per page
        assert len(doc.tables) == 3

    def test_single_page_for_small_chunk(self):
        chunk = self._make_chunk(5)
        header = chunk.columns.tolist()
        widths = scale_column_widths(len(header))

        doc, pages = build_folder_document(1, chunk, header, widths)
        assert pages == 1
        assert len(doc.tables) == 1

    def test_full_folder_size(self):
        chunk = self._make_chunk(1000)
        header = chunk.columns.tolist()
        widths = scale_column_widths(len(header))

        doc, pages = build_folder_document(1, chunk, header, widths)
        assert pages == 100  # 1000 / 10
        assert len(doc.tables) == 100
