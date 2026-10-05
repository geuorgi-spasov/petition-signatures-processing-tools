"""Tests for generate_initials_book.py.

They follow the three steps of the script: reading the CSV, turning names
into initials, and laying the initials out on the page — plus the rules
that say when a layout cannot be printed.
"""
from __future__ import annotations

import dataclasses

import pandas as pd
import pytest

from generate_initials_book import (
    Layout,
    LayoutError,
    build_initials_document,
    group_into_lines,
    main,
    names_to_initials,
    read_two_column_csv,
)

# The defaults on A5: 1.5 cm offsets leave 11.8 × 18.0 cm. At 10 pt and
# line spacing 1.2 a line is 12 pt high, so 42 lines fit, 7 to a line.
DEFAULT = Layout()
LINES_PER_PAGE = 42
INITIALS_PER_PAGE = 294


def layout(**changes) -> Layout:
    """The default layout with a few numbers changed."""
    return dataclasses.replace(DEFAULT, **changes)


# ---------------------------------------------------------------------------
# Step 1 — reading the CSV
# ---------------------------------------------------------------------------

class TestReadTwoColumnCsv:
    @pytest.mark.parametrize("separator", [",", ";", "\t", "|"])
    def test_detects_the_common_separators(self, tmp_path, separator):
        csv = tmp_path / "names.csv"
        csv.write_text(f"Иван{separator}Иванов\n", encoding="utf-8")
        names = read_two_column_csv(str(csv))
        assert (names.iloc[0, 0], names.iloc[0, 1]) == ("Иван", "Иванов")

    def test_raises_on_a_single_column_file(self, tmp_path):
        csv = tmp_path / "broken.csv"
        csv.write_text("just_one_column\n", encoding="utf-8")
        with pytest.raises(ValueError):
            read_two_column_csv(str(csv))

    def test_raises_on_a_missing_file(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            read_two_column_csv(str(tmp_path / "does_not_exist.csv"))


# ---------------------------------------------------------------------------
# Step 2 — names to initials
# ---------------------------------------------------------------------------

class TestNamesToInitials:
    def test_takes_the_first_letter_of_each_name(self):
        names = pd.DataFrame([["Иван", "Иванов"], ["Мария", "Маринова"]])
        assert names_to_initials(names) == ["И. И.", "М. М."]

    def test_lowercase_and_whitespace_are_cleaned_up(self):
        assert names_to_initials(pd.DataFrame([["  иван  ", " иванов "]])) == ["И. И."]

    def test_a_missing_name_leaves_a_single_initial(self):
        names = pd.DataFrame([["Иван", None], [None, "Петров"]])
        assert names_to_initials(names) == ["И.", "П."]

    def test_rows_without_any_name_are_skipped(self):
        names = pd.DataFrame([[None, None], ["", ""], ["Иван", "Иванов"]])
        assert names_to_initials(names) == ["И. И."]

    # The export really does contain names like these - 60 rows of it.
    # The initial is the first *letter*, whatever junk precedes it.
    @pytest.mark.parametrize(
        "name,initial",
        [
            ("2milyanov", "M"),      # digit
            ("0milyanov", "M"),
            ("?milyanov", "M"),      # question mark
            ("^milyanov", "M"),      # caret
            ("'milyanov", "M"),      # quote
            (".milyanov", "M"),      # dot
            ("-milyanov", "M"),      # dash
            ("…milyanov", "M"),      # ellipsis
            ("#@!milyanov", "M"),    # several at once
            ("  2milyanov  ", "M"),  # and with spaces around it
            ("Иван", "И"),           # the ordinary case still works
            ("иван", "И"),
        ],
    )
    def test_the_initial_is_the_first_letter_whatever_precedes_it(
        self, name, initial
    ):
        names = pd.DataFrame([[name, "Иванов"]])
        assert names_to_initials(names) == [f"{initial}. И."]

    @pytest.mark.parametrize("name", ["2", ".", "?", "^", "...", "123", "   "])
    def test_a_name_with_no_letters_at_all_counts_as_missing(self, name):
        names = pd.DataFrame([[name, "Иванов"], [name, name]])
        assert names_to_initials(names) == ["И."]

    def test_empty_input(self):
        assert names_to_initials(pd.DataFrame(columns=[0, 1])) == []


# ---------------------------------------------------------------------------
# Step 3a — what the layout works out
# ---------------------------------------------------------------------------

class TestLayoutMaths:
    def test_the_defaults_are_an_a5_book_page(self):
        assert (DEFAULT.page_width_cm, DEFAULT.page_height_cm) == (14.8, 21.0)
        assert DEFAULT.page_name == "A5"
        assert DEFAULT.columns == 7
        assert DEFAULT.lines_per_page == LINES_PER_PAGE
        assert DEFAULT.initials_per_page == INITIALS_PER_PAGE

    def test_the_text_area_is_the_page_minus_the_offsets(self):
        assert DEFAULT.text_width_cm == pytest.approx(11.8)
        assert DEFAULT.text_height_cm == pytest.approx(18.0)

    def test_the_page_size_can_be_changed(self):
        a4 = layout(page_width_cm=21.0, page_height_cm=29.7, column_gap_cm=1.5)
        a4.validate()
        assert a4.page_name == "A4"
        assert a4.lines_per_page == 63
        assert a4.pages_for(110_942) == 252

    def test_an_unusual_page_size_is_named_by_its_measurements(self):
        assert layout(page_width_cm=13.0, page_height_cm=20.0).page_name == (
            "13 x 20 cm"
        )

    def test_a_line_is_the_font_size_times_the_line_spacing(self):
        assert layout(font_size_pt=10, line_spacing=1.2).line_height_pt == 12.0

    def test_a_full_line_has_one_gap_fewer_than_it_has_columns(self):
        three = layout(columns=3, column_gap_cm=2.0)
        assert three.width_needed_cm == pytest.approx(
            3 * three.initials_width_cm + 2 * 2.0
        )

    def test_the_page_count_is_rounded_up(self):
        assert DEFAULT.pages_for(INITIALS_PER_PAGE) == 1
        assert DEFAULT.pages_for(INITIALS_PER_PAGE + 1) == 2
        assert DEFAULT.pages_for(110_942) == 378

    def test_more_columns_means_fewer_pages(self):
        assert layout(columns=7).pages_for(10_000) < (
            layout(columns=4).pages_for(10_000)
        )

    def test_a_bigger_font_means_fewer_lines_and_more_pages(self):
        small, big = layout(font_size_pt=8), layout(font_size_pt=16, columns=4)
        assert big.lines_per_page < small.lines_per_page
        assert big.pages_for(10_000) > small.pages_for(10_000)

    def test_wider_line_spacing_means_more_pages(self):
        assert layout(line_spacing=2.0).pages_for(10_000) > (
            layout(line_spacing=1.0).pages_for(10_000)
        )

    def test_bigger_offsets_mean_more_pages(self):
        assert layout(margin_cm=3.0, columns=3).pages_for(10_000) > (
            DEFAULT.pages_for(10_000)
        )

    def test_columns_are_spaced_by_their_width_plus_the_gap(self):
        offsets = layout(columns=3, column_gap_cm=2.0).column_offsets_cm()
        pitch = DEFAULT.initials_width_cm + 2.0
        assert [b - a for a, b in zip(offsets, offsets[1:])] == pytest.approx(
            [pitch, pitch]
        )

    def test_the_columns_are_centred_between_the_offsets(self):
        three = layout(columns=3, column_gap_cm=2.0)
        offsets = three.column_offsets_cm()
        space_on_the_right = (
            three.text_width_cm - offsets[-1] - three.initials_width_cm
        )
        assert offsets[0] == pytest.approx(space_on_the_right)

    def test_the_summary_mentions_the_columns_and_the_pages(self):
        summary = DEFAULT.describe(1000)
        assert "A5" in summary
        assert "7 column(s)" in summary
        assert "about 4 pages" in summary


# ---------------------------------------------------------------------------
# Step 3b — the rules a layout has to obey
# ---------------------------------------------------------------------------

class TestLayoutValidation:
    def test_the_defaults_are_valid(self):
        DEFAULT.validate()  # must not raise

    def test_too_many_columns_for_the_page_is_rejected(self):
        with pytest.raises(LayoutError) as error:
            layout(columns=20).validate()
        assert "Use at most 7 column(s)" in str(error.value)
        assert "on A5" in str(error.value)

    def test_the_suggested_number_of_columns_really_fits(self):
        layout(columns=7).validate()  # must not raise

    def test_too_big_a_font_for_the_columns_is_rejected(self):
        with pytest.raises(LayoutError) as error:
            layout(font_size_pt=40).validate()
        assert "40 pt initials" in str(error.value)

    def test_too_wide_a_gap_is_rejected(self):
        with pytest.raises(LayoutError):
            layout(column_gap_cm=5.0).validate()

    def test_offsets_that_leave_no_room_are_rejected(self):
        with pytest.raises(LayoutError) as error:
            layout(margin_cm=11.0).validate()
        assert "leaves no room" in str(error.value)

    def test_lines_too_tall_for_the_page_are_rejected(self):
        with pytest.raises(LayoutError) as error:
            layout(line_spacing=100).validate()
        assert "only 18.0 cm are left" in str(error.value)

    @pytest.mark.parametrize(
        "impossible", [{"columns": 0}, {"columns": -1}, {"font_size_pt": 0},
                       {"line_spacing": 0}, {"margin_cm": -1},
                       {"column_gap_cm": -1}]
    )
    def test_numbers_that_make_no_sense_are_rejected(self, impossible):
        with pytest.raises(LayoutError):
            layout(**impossible).validate()


# ---------------------------------------------------------------------------
# Step 3c — the document that comes out
# ---------------------------------------------------------------------------

class TestGroupIntoLines:
    def test_fills_one_line_at_a_time(self):
        assert group_into_lines(["a", "b", "c", "d"], 2) == [["a", "b"], ["c", "d"]]

    def test_the_last_line_may_be_short(self):
        assert group_into_lines(["a", "b", "c"], 2) == [["a", "b"], ["c"]]

    def test_no_initials_means_no_lines(self):
        assert group_into_lines([], 7) == []


class TestBuildInitialsDocument:
    def test_one_paragraph_per_line_with_tabs_between_the_columns(self):
        initials = ["a", "b", "c", "d", "e", "f", "g"]
        doc = build_initials_document(initials, layout(columns=3))
        assert [p.text for p in doc.paragraphs] == ["a\tb\tc", "d\te\tf", "g"]

    def test_no_initial_is_lost(self):
        initials = [f"{i}." for i in range(100)]
        doc = build_initials_document(initials, DEFAULT)
        placed = [cell for p in doc.paragraphs for cell in p.text.split("\t")]
        assert placed == initials

    def test_empty_input_produces_an_empty_document(self):
        assert build_initials_document([], DEFAULT).paragraphs == []

    def test_the_page_matches_the_layout(self):
        doc = build_initials_document(["И. И."], layout(margin_cm=2.0))
        section = doc.sections[0]
        assert section.page_width.cm == pytest.approx(14.8, abs=0.01)
        assert section.page_height.cm == pytest.approx(21.0, abs=0.01)
        assert section.left_margin.cm == pytest.approx(2.0, abs=0.01)
        assert section.top_margin.cm == pytest.approx(2.0, abs=0.01)

    def test_the_font_and_the_line_height_are_applied(self):
        doc = build_initials_document(
            ["И. И."], layout(font_size_pt=12), font_name="Arial"
        )
        style = doc.styles["Normal"]
        assert style.font.name == "Arial"
        assert style.font.size.pt == pytest.approx(12.0)
        assert style.paragraph_format.line_spacing.pt == pytest.approx(14.4)

    def test_a_tab_stop_sits_at_every_column_after_the_first(self):
        three = layout(columns=3, column_gap_cm=1.0)
        doc = build_initials_document(["a", "b", "c"], three)
        stops = doc.styles["Normal"].paragraph_format.tab_stops
        assert [stop.position.cm for stop in stops] == pytest.approx(
            three.column_offsets_cm()[1:], abs=0.01
        )

    def test_no_page_breaks_are_forced(self):
        # Word fits fewer lines on a page than the arithmetic predicts, so a
        # forced break left the last lines of each page orphaned on the next
        # one and doubled the length of the book. Pages are the program's job.
        initials = [f"{i}." for i in range(1000)]
        doc = build_initials_document(initials, DEFAULT)
        assert not any(
            p.paragraph_format.page_break_before for p in doc.paragraphs
        )

    def test_initials_are_not_spell_checked(self):
        doc = build_initials_document(["И. И."], DEFAULT)
        assert doc.styles["Normal"].font.no_proof is True


# ---------------------------------------------------------------------------
# The whole script, end to end
# ---------------------------------------------------------------------------

@pytest.fixture
def names_csv(tmp_path, monkeypatch):
    """A small CSV in a temporary working directory."""
    monkeypatch.chdir(tmp_path)
    (tmp_path / "names.csv").write_text(
        "Иван,Иванов\nПетър,Петров\n", encoding="utf-8"
    )
    return tmp_path


class TestMain:
    def test_writes_the_document(self, names_csv):
        assert main(["--input", "names.csv", "--output", "out.docx"]) == 0
        assert (names_csv / "out.docx").exists()

    def test_dry_run_reports_the_pages_without_writing(self, names_csv, capsys):
        assert main(["--input", "names.csv", "--dry-run"]) == 0
        assert not (names_csv / "book_signatures_initials.docx").exists()
        assert "about 1 page" in capsys.readouterr().out

    def test_a_missing_input_file_is_reported(self, names_csv, capsys):
        assert main(["--input", "nope.csv"]) == 1
        assert "not found" in capsys.readouterr().out

    def test_an_impossible_layout_is_reported_and_nothing_is_written(
        self, names_csv, capsys
    ):
        code = main(["--input", "names.csv", "--output", "out.docx",
                     "--columns", "20"])
        assert code == 1
        assert not (names_csv / "out.docx").exists()
        assert "Use at most 7 column(s)" in capsys.readouterr().out
