# Book of Initials

A single Python script that turns a CSV of signatory names into a
printable A4 document listing just their initials (e.g. `И. И.`), in the
**Bebas Neue Cyrillic** font, laid out in several columns per line.

The layout is yours to choose: how many sets of initials go on one line
(**7** by default), the font size, the offset from the page edges, the
distance between the initials on a line and the line spacing. The script
reports **how many pages the book will have** before it builds anything,
and refuses — with an explanation — any combination that cannot fit on A4.

This workflow is independent of the paper-submission one — you only
need what's listed below.

---

## About this code

This toolkit was written with help from an AI assistant (Claude by
Anthropic) and reviewed by a human. The code is intentionally written
to be readable by people who aren't programmers.

If anything here is unclear, if you want to verify how a script works,
or if you have a question this README doesn't answer — paste the
relevant code or text into an AI assistant (Claude, ChatGPT, etc.) and
ask. The AI can explain how each part works or help you adapt it.

---

## Quick start

1. Install Python 3.9+ from <https://www.python.org/downloads/>
   (on Windows, tick *"Add Python to PATH"* during installation).
2. Install the bundled **Bebas Neue Cyrillic** font
   (`bebasneuecyrillic.ttf`) — see [Installing the font](#installing-the-font).
3. Open PowerShell (Windows) or Terminal (macOS/Linux) **in this
   folder** and install the dependencies. Using a virtual environment
   is recommended on every OS, and is **required on most current Linux
   distributions** (Debian/Ubuntu/Mint 24+), which block installing
   into the system Python:

   **Linux / macOS:**
   ```
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

   **Windows (PowerShell):**
   ```
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```

   The virtual environment lives in a `.venv/` folder next to the
   script. Activate it once per terminal session (run the `activate`
   line again in a new terminal) — you'll see `(.venv)` at the start
   of your prompt when it's active.

   > If you prefer not to use a virtual environment on Linux and
   > understand the risk, you can instead run
   > `pip install -r requirements.txt --break-system-packages`, but the
   > virtual environment above is the clean, recommended approach.

4. Put your CSV file `book_signatures_only_names.csv` next to the
   script and run (with the virtual environment active):

   ```
   python generate_initials_book.py
   ```

   With the default 7 columns a list of ~111,000 names takes about
   15 seconds and produces a 252-page book.

   To see the layout and the page count *without* writing anything:

   ```
   python generate_initials_book.py --dry-run
   ```

The output `book_signatures_initials.docx` appears in the same folder.
Print it, or export it to PDF from Word (*File → Save as → PDF*) — the
PDF has exactly the number of pages the script reported.

---

## Installing the font

The font file `bebasneuecyrillic.ttf` is bundled in this folder. After
installing, it registers itself on your system as **`Bebas Neue
Cyrillic`** — that is the exact name the script looks for.

### Windows

The simplest way is to double-click `bebasneuecyrillic.ttf` and press
**Install**. Or, from PowerShell in this folder:

```
Start-Process .\bebasneuecyrillic.ttf
```

…then press **Install** in the window that opens.

### macOS

Double-click the file and press **Install Font**, or from Terminal in
this folder:

```
cp bebasneuecyrillic.ttf ~/Library/Fonts/
```

### Linux

From Terminal in this folder:

```
mkdir -p ~/.local/share/fonts && cp bebasneuecyrillic.ttf ~/.local/share/fonts/ && fc-cache -f
```

After installing, confirm the exact family name your system reports —
on Linux it should match the `FONT_NAME` in the script
(`Bebas Neue Cyrillic`):

```
fc-list | grep -i bebas
```

If the name shown differs, set `FONT_NAME` in
`generate_initials_book.py` to exactly what `fc-list` reports.

---

## What the input should look like

A CSV with **two columns**: first name and last name, one signatory
per row. Any common separator works (comma, semicolon, tab, pipe).
Example:

```
Иван,Иванов
Петър,Петров
Мария,Маринова
```

Headers are *not* expected — the script reads from the very first row.

## What the output looks like

An A4 document with 7 sets of initials per line:

```
И. И.    П. П.    М. М.    Г. Г.    А. А.    Н. Н.    С. С.
Д. Д.    В. В.    Р. Р.    К. К.    Т. Т.    Б. Б.    Ж. Ж.
```

Font: **Bebas Neue Cyrillic** at 10 pt, 1.5 cm offsets on all sides. That
leaves an 18.0 × 26.7 cm text area, which holds 63 lines, so one page
carries 7 × 63 = **441 sets of initials**.

Before building, the script prints the plan:

```
Layout: A4, 7 column(s) × 63 line(s) = 441 initials per page
        10 pt font, 1.5 cm offsets, 1.5 cm between the initials, line spacing 1.2
        252 page(s) in total
```

---

## Making the book look the way you want

Every setting can be changed in two ways — whichever you find easier:

- **On the command line**, e.g.
  `python generate_initials_book.py --columns 10 --column-gap 1.0`.
  Run `python generate_initials_book.py --help` to see them all.
- **In the file**: open `generate_initials_book.py` in any text editor.
  The first thing in it is a `# Configuration` block of `UPPER_CASE`
  variables with the same meanings. Save and re-run.

| Command line | In the file | Default | What it does |
| --- | --- | --- | --- |
| `--columns` | `COLUMNS` | `7` | Sets of initials next to each other on one line |
| `--font-size` | `FONT_SIZE_PT` | `10` | Font size in points. Also decides how many lines fit on a page |
| `--margin` | `MARGIN_CM` | `1.5` | Offset from all four page edges, in cm |
| `--column-gap` | `COLUMN_GAP_CM` | `1.5` | Distance between the initials on a line, in cm |
| `--line-spacing` | `LINE_SPACING` | `1.2` | Height of a line, as a multiple of the font size |
| `--font-name` | `FONT_NAME` | `Bebas Neue Cyrillic` | Font family, exactly as your system reports it |
| `--input`, `--output` | `INPUT_CSV`, `OUTPUT_DOCX` | see the script | The files to read and write |
| `--dry-run` | — | off | Print the layout and the page count, write nothing |

The page is always A4 (21 × 29.7 cm) — that is what the book is printed
on, so it is fixed rather than a parameter. The columns are centred
between the offsets, so whatever width they don't use is split evenly on
both sides.

### How the page count is worked out

1. The text area is the page minus the offsets — 18.0 × 26.7 cm by
   default.
2. One line is `font size × line spacing` high: 10 × 1.2 = 12 pt. 26.7 cm
   is 756.9 pt, so **63 lines** fit.
3. One line holds `--columns` sets of initials, so a page holds
   7 × 63 = **441**, and the page count is the number of initials divided
   by that, rounded up.

So **a bigger font gives more pages** and **more columns give fewer
pages**. For the ~111,000 names in the sample CSV:

| Settings | Initials per page | Pages |
| --- | --- | --- |
| defaults (7 columns, 10 pt) | 441 | 252 |
| `--columns 10 --column-gap 1.0` | 630 | 177 |
| `--columns 5 --font-size 14` | 225 | 494 |
| `--columns 1` (one per line) | 63 | 1,761 |

### When the settings don't fit

Nothing is written and the script says what is wrong and what to change:

```
ERROR: 20 column(s) of 10 pt initials with 1.5 cm between them need 43.3 cm,
but only 18.0 cm are left between the 1.5 cm offsets on A4.
Use at most 8 column(s), a smaller gap, a smaller font size, or smaller offsets.
```

The same happens when the offsets leave no room on the page, or when the
lines are too tall to fit even one between the top and bottom offsets.

---

## How the code is organised

`generate_initials_book.py` is one file with three steps and one small
class:

| | |
| --- | --- |
| `read_two_column_csv()` | reads the CSV, trying the usual separators |
| `names_to_initials()` | `Иван, Иванов` → `И. И.`, in one vectorised pandas pass |
| `Layout` | the five numbers above. Every other figure — the width of one set of initials, the lines per page, the page count — is a property derived from them and from the size of an A4 sheet |
| `Layout.validate()` | the three rules: the values make sense, the columns fit across the page, a line fits down it. Raises `LayoutError` with the message you saw above |
| `build_initials_document()` | one paragraph per line, the columns separated by tabs, a page break where the layout says a page ends |
| `main()` | reads, validates, prints the plan, writes the document |

`LayoutError` extends `ValueError`, so `main()` catches an unreadable CSV
and an impossible layout in the same place.

---

## Running the tests (for developers)

The project ships with a small `pytest` suite covering the initials, the
CSV reading, the layout maths and its limits, and the document that comes
out. To run it (with the virtual environment from
[Quick start](#quick-start) active):

```
pip install -r tests/test_requirements.txt
python -m pytest
```

A successful run looks like:

```
======================== test session starts ========================
collected 50 items

tests/test_generate_initials_book.py ....................    [100%]

======================== 50 passed in 0.8s ==========================
```

`tests/benchmark_initials_book.py` compares the build time and the page
count of different column counts — one line is one paragraph, so more
columns means fewer paragraphs to write:

```
python tests/benchmark_initials_book.py --limit 20000
```

```
 columns     lines   pages     time
       1    20,000     318    10.1s
       3     6,667     106     1.6s
       5     4,000      64     1.1s
       7     2,858      46     0.9s
```

---

## Troubleshooting

- **`file '…' not found in this folder`** — make sure the CSV is in
  the same folder as the script and the name matches exactly,
  including capitalisation.
- **`Could not parse '…' into at least two columns`** — open the file
  in a text editor and check that columns are separated by commas,
  semicolons, tabs, or `|`. If possible, re-save as UTF-8.
- **The initials look like rectangles or the font looks wrong** — the
  Bebas Neue Cyrillic font isn't installed on your system. Install it
  using one of the commands above and re-run. If you're on Windows,
  close Word completely before re-running so it picks up the newly
  installed font.
- **`… need 43.3 cm, but only 18.0 cm are left`** — the columns, the
  gap and the font size together need more width than A4 has. The
  message says how many columns do fit; use that number, or a smaller
  `--column-gap`, `--font-size` or `--margin`.
- **`A line of … is 35.3 cm high`** — the font size times the line
  spacing is taller than what is left between the offsets. Use a
  smaller `--font-size` or `--line-spacing`.
- **`pip` is not recognised** — Python wasn't added to PATH. Re-install
  Python and tick *"Add Python to PATH"*, or use `py -m pip …`
  instead.
- **`error: externally-managed-environment`** (Linux) — your
  distribution blocks installing packages into the system Python. Use
  the virtual environment shown in [Quick start](#quick-start)
  (`python3 -m venv .venv && source .venv/bin/activate`) and run `pip`
  inside it. This is the recommended fix.
- **The initials look wrong only on Linux** — the installed font may
  register under a slightly different family name. Run
  `fc-list | grep -i bebas` and set `FONT_NAME` in the script to match
  exactly what it reports.

If your problem isn't here, paste the error and the relevant script
into an AI assistant (Claude, ChatGPT, etc.) — it can usually
diagnose it from the script and the message alone.

---

## Folder layout

```
Book_Initials_Signatures/
├── README.md
├── requirements.txt
├── bebasneuecyrillic.ttf                   (bundled font, install once)
├── generate_initials_book.py
├── tests/
│   ├── conftest.py
│   ├── test_requirements.txt
│   ├── benchmark_initials_book.py          (layout speed comparison)
│   └── test_generate_initials_book.py
├── book_signatures_only_names.csv          (your input)
└── book_signatures_initials.docx           (the generated output)
```

> The `.venv/` folder (created by the Quick start) and the generated
> `book_signatures_initials.docx` are intentionally excluded from
> version control by the `.gitignore` at the root of the repository.
