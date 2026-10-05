"""Tests for split_signatures_into_folders.py."""
from __future__ import annotations

import pandas as pd
import pytest

from split_signatures_into_folders import (
    CSV_ENCODINGS,
    CSV_SEPARATORS,
    DEFAULT_COLUMN_WIDTHS_CM,
    LEFT_MARGIN_CM,
    PAGE_WIDTH_CM,
    RIGHT_MARGIN_CM,
    apply_full_table_borders,
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
    def _make_folder_rows(self, n_rows):
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
        folder_rows = self._make_folder_rows(25)  # 25 rows = 3 pages at 10 per page
        header = folder_rows.columns.tolist()
        widths = scale_column_widths(len(header))

        doc, pages = build_folder_document(1, folder_rows, header, widths)

        assert pages == 3
        # One table per page
        assert len(doc.tables) == 3

    def test_single_page_for_small_chunk(self):
        folder_rows = self._make_folder_rows(5)
        header = folder_rows.columns.tolist()
        widths = scale_column_widths(len(header))

        doc, pages = build_folder_document(1, folder_rows, header, widths)
        assert pages == 1
        assert len(doc.tables) == 1

    def test_full_folder_size(self):
        folder_rows = self._make_folder_rows(1000)
        header = folder_rows.columns.tolist()
        widths = scale_column_widths(len(header))

        doc, pages = build_folder_document(1, folder_rows, header, widths)
        assert pages == 100  # 1000 / 10
        assert len(doc.tables) == 100


# ---------------------------------------------------------------------------
# Encodings
# ---------------------------------------------------------------------------

BULGARIAN_ROWS = "Номер\tИме\tФамилия\n1\tИван\tИванов\n2\tМария\tМаринова\n"


def _write(path, text, encoding):
    path.write_bytes(text.encode(encoding))
    return str(path)


class TestEncodings:
    @pytest.mark.parametrize("encoding", ["utf-8", "utf-8-sig", "cp1251"])
    def test_bulgarian_names_survive_the_round_trip(self, tmp_path, encoding):
        path = _write(tmp_path / "signatures.csv", BULGARIAN_ROWS, encoding)
        df = read_signatures_csv(path)
        assert list(df.columns) == ["Номер", "Име", "Фамилия"]
        assert list(df.iloc[0])[1:] == ["Иван", "Иванов"]

    def test_the_last_resort_is_last(self):
        # latin-1 decodes any byte at all, so anything after it could never
        # be reached - and anything before it must be able to fail.
        assert CSV_ENCODINGS[-1] == "latin-1"
        assert bytes(range(256)).decode("latin-1")
        for earlier in CSV_ENCODINGS[:-1]:
            with pytest.raises(UnicodeDecodeError):
                bytes(range(256)).decode(earlier)

    def test_every_encoding_in_the_list_is_reachable(self):
        # An encoding is only worth listing if the ones before it reject
        # the bytes it is there to catch. The last resort is exempt - its
        # whole job is to catch what nothing else did.
        middle = CSV_ENCODINGS[1:-1]
        assert middle, "nothing between utf-8 and the last resort to check"
        for position, encoding in enumerate(middle, start=1):
            sample = BULGARIAN_ROWS.encode(encoding)
            assert any(
                not _decodes(sample, earlier)
                for earlier in CSV_ENCODINGS[:position]
            ), f"{encoding} is shadowed by the encodings before it"

    @pytest.mark.parametrize("shadowed", ["cp866", "cp1252", "iso-8859-1"])
    def test_the_encodings_left_out_really_are_unreachable(self, shadowed):
        # These were in the list once. Each is decoded without error by an
        # encoding that comes before it, so it could never have been used.
        sample = BULGARIAN_ROWS.encode("cp1251")
        assert _decodes(sample, "cp1251")
        assert _decodes(bytes(range(256)), "latin-1")
        assert shadowed not in CSV_ENCODINGS

    def test_a_missing_file_is_reported_as_missing(self, tmp_path):
        # Not as a parse failure: there is nothing to parse, and listing
        # every separator and encoding "tried" would be a lie. main()
        # has always had a handler for this - it just could not fire.
        with pytest.raises(FileNotFoundError):
            read_signatures_csv(str(tmp_path / "does_not_exist.csv"))

    def test_an_unreadable_file_says_what_was_tried(self, tmp_path):
        path = tmp_path / "one_column.csv"
        path.write_text("just_one_column\nstill_one\n", encoding="utf-8")
        with pytest.raises(ValueError) as error:
            read_signatures_csv(str(path))
        assert "Tried separators" in str(error.value)


def _decodes(data: bytes, encoding: str) -> bool:
    try:
        data.decode(encoding)
        return True
    except UnicodeDecodeError:
        return False


# ---------------------------------------------------------------------------
# Separators
# ---------------------------------------------------------------------------

class TestSeparators:
    def test_a_comma_in_the_address_does_not_beat_the_real_separator(
        self, tmp_path
    ):
        # Bulgarian addresses are full of commas. The file is tab-separated,
        # so the tab must win or every field shifts one column left.
        path = tmp_path / "signatures.csv"
        path.write_text(
            "Номер\tИме\tАдрес\n"
            "1\tИван Иванов\tгр. София, ул. Витоша 5\n"
            "2\tМария Маринова\tгр. Пловдив, бул. Руски 12\n",
            encoding="utf-8",
        )
        df = read_signatures_csv(str(path))
        assert list(df.columns) == ["Номер", "Име", "Адрес"]
        assert df.iloc[0]["Адрес"] == "гр. София, ул. Витоша 5"

    def test_a_genuine_comma_file_still_works(self, tmp_path):
        # Ordering the comma last must not stop it being found when it
        # really is the separator.
        path = tmp_path / "signatures.csv"
        path.write_text("Номер,Име\n1,Иван\n2,Мария\n", encoding="utf-8")
        df = read_signatures_csv(str(path))
        assert list(df.columns) == ["Номер", "Име"]

    def test_the_comma_is_tried_last(self):
        # A separator cannot fail the way an encoding can, so the one most
        # likely to appear inside the data has to be the last resort.
        assert CSV_SEPARATORS[-1] == ","
        assert set(CSV_SEPARATORS) == {"\t", "|", ";", ","}


# ---------------------------------------------------------------------------
# Table borders
# ---------------------------------------------------------------------------

class TestTableBorders:
    def _table(self):
        from docx import Document
        doc = Document()
        table = doc.add_table(rows=3, cols=4)
        apply_full_table_borders(table)
        return table

    def test_the_table_gets_an_outer_frame_and_inner_gridlines(self):
        xml = self._table()._tbl.xml
        assert "<w:tblBorders>" in xml
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            assert f"<w:{edge} " in xml, f"table edge {edge} missing"

    def test_every_cell_gets_its_own_four_sides(self):
        # Belt and braces: some viewers honour only the cell-level borders.
        table = self._table()
        for row in table.rows:
            for cell in row.cells:
                xml = cell._tc.xml
                assert "<w:tcBorders>" in xml
                assert xml.count("w:val=\"single\"") >= 4

    def test_the_line_is_thin_solid_and_black(self):
        xml = self._table()._tbl.xml
        assert 'w:val="single"' in xml   # solid, not dashed
        assert 'w:sz="4"' in xml         # eighths of a point -> 0.5 pt
        assert 'w:color="000000"' in xml
