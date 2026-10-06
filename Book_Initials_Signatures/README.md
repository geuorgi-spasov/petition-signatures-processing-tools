# Book of Initials

A single Python script that turns a CSV of signatory names into a
printable A5 book listing just their initials (e.g. `И. И.`), in the
**Oswald** font, laid out in several columns per line.

The layout is yours to choose: the page size (**A5** by default — half an
A4 sheet, the usual book format), how many sets of initials go on one line
(**7**), the font size, the offset from the page edges, the distance
between the initials on a line and the line spacing. The script reports
**how many pages the book will have** before it builds anything, and
refuses — with an explanation — any combination that cannot fit.

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

Follow **one** of the two paths below — whichever matches your computer.

### Windows

1. **Install Python** from <https://www.python.org/downloads/>, ticking
   *"Add Python to PATH"* during the installation.
2. **Put your `book_signatures_only_names.csv` in this folder.**
3. **Double-click `run-windows.bat`.**

### Linux and macOS

1. **Install Python** — on Ubuntu/Mint/Debian:

   ```
   sudo apt install python3 python3-venv
   ```

2. **Put your `book_signatures_only_names.csv` in this folder.**
3. **Open Terminal in this folder and run:**

   ```
   ./run-linux-macos.sh
   ```

That's it. The launcher sets up everything the script needs, offers to
install the font the first time, and then builds the book. The first run
takes an extra minute to set itself up; after that it's about 15 seconds
for ~111,000 names.

The font prompt looks like this, and pressing Enter accepts:

```
The book is set in Oswald, which is not installed yet.
The font is bundled here and free to install (SIL Open Font
License, see OFL.txt). It goes in your own home folder:
    /home/you/.local/share/fonts/oswald.ttf

Install it now? [Y/n]
```

Nothing is installed outside your own user account, and no
administrator or root rights are needed. If you would rather do it
yourself, answer `n` and see [Installing the font](#installing-the-font)
below — the book is still built either way, but whatever opens it will
substitute a different font and the spacing will not be what the layout
intends.

To see the layout and the page count *without* writing anything, add
`--dry-run` — `run-windows.bat --dry-run` or `./run-linux-macos.sh --dry-run`.

### Prefer to run it yourself?

The launchers do no magic. On Windows:

```
pip install -r requirements.txt
python generate_initials_book.py
```

On Linux, Python has to be installed into a folder of its own, so it is
one line more:

```
python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt
python generate_initials_book.py
```

Run the `source .venv/bin/activate` part again in any new terminal.

The output `book_signatures_initials.docx` appears in the same folder.
Print it, or export it to PDF from Word (*File → Save as → PDF*) — the
PDF will be about as long as the script reported; the exact number is up
to whichever program prints it, and they differ a little.

---

## Installing the font

**You normally do not have to do this** — the launcher offers to install
the font on the first run, for your user only. This section is for doing
it by hand, or for checking what the launcher did.

The font file `oswald.ttf` is bundled in this folder and is licensed
under the SIL Open Font License (`OFL.txt`), which is what makes it
free to pass around with the program. After installing, it registers
itself on your system as **`Oswald`** — the exact name the script looks
for.

### Windows

The simplest way is to double-click `oswald.ttf` and press
**Install**. Or, from PowerShell in this folder:

```
Start-Process .\oswald.ttf
```

…then press **Install** in the window that opens.

### macOS

Double-click the file and press **Install Font**, or from Terminal in
this folder:

```
cp oswald.ttf ~/Library/Fonts/
```

### Linux

From Terminal in this folder:

```
mkdir -p ~/.local/share/fonts && cp oswald.ttf ~/.local/share/fonts/ && fc-cache -f
```

After installing, confirm the exact family name your system reports —
on Linux it should match the `FONT_NAME` in the script
(`Oswald`):

```
fc-list | grep -i bebas
```

If the name shown differs, set `FONT_NAME` in
`generate_initials_book.py` to exactly what `fc-list` reports.

---

## What the input should look like

A CSV with **two columns**: first name and last name, one signatory
per row. Any common separator works (comma, semicolon, tab, pipe),
and the script auto-detects the encoding — UTF-8, or the Windows
Cyrillic `cp1251` that Excel saves Bulgarian text as. Example:

```
Иван,Иванов
Петър,Петров
Мария,Маринова
```

Headers are *not* expected — the script reads from the very first row.

## What the output looks like

An A5 book with 6 sets of initials per line:

```
И. И.    П. П.    М. М.    Г. Г.    А. А.    Н. Н.    С. С.
Д. Д.    В. В.    Р. Р.    К. К.    Т. Т.    Б. Б.    Ж. Ж.
```

Font: **Oswald** at 10 pt, 1.5 cm offsets on all sides. On
A5 that leaves an 11.8 × 18.0 cm text area, which holds 42 lines, so one
page carries 6 × 42 = **252 sets of initials**.

Before building, the script prints the plan:

```
Layout: A5, 6 column(s) x 42 line(s) = 252 initials per page
        10 pt font, 1.5 cm offsets, 1 cm between the initials, line spacing 1.2
        110,942 initials -> 441 page(s)
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
| `--columns` | `COLUMNS` | `6` | Sets of initials next to each other on one line |
| `--font-size` | `FONT_SIZE_PT` | `10` | Font size in points. Also decides how many lines fit on a page |
| `--margin` | `MARGIN_CM` | `1.5` | Offset from all four page edges, in cm |
| `--column-gap` | `COLUMN_GAP_CM` | `1.0` | Distance between the initials on a line, in cm |
| `--line-spacing` | `LINE_SPACING` | `1.2` | Height of a line, as a multiple of the font size |
| `--page-width` | `PAGE_WIDTH_CM` | `14.8` | Page width in cm (A5 is 14.8, A4 is 21.0) |
| `--page-height` | `PAGE_HEIGHT_CM` | `21.0` | Page height in cm (A5 is 21.0, A4 is 29.7) |
| `--font-name` | `FONT_NAME` | `Oswald` | Font family, exactly as your system reports it |
| `--input`, `--output` | `INPUT_CSV`, `OUTPUT_DOCX` | see the script | The files to read and write |
| `--dry-run` | — | off | Print the layout and the page count, write nothing |

The default page is **A5** (14.8 × 21.0 cm). A4 is a document format, not
a book format; A5 is half an A4 sheet, so two book pages print on one
sheet with nothing wasted. Other common Bulgarian book formats, should
your printer prefer one of them:

```
python generate_initials_book.py --page-width 14.5 --page-height 20    # 60x84/16
python generate_initials_book.py --page-width 13   --page-height 20    # 84x108/32
python generate_initials_book.py --page-width 17   --page-height 24    # 70x100/16
```

The columns are centred between the offsets, so whatever width they don't
use is split evenly on both sides.

### How the page count is worked out

1. The text area is the page minus the offsets — 11.8 × 18.0 cm on A5
   with the default offsets.
2. One line is `font size × line spacing` high: 10 × 1.2 = 12 pt. 18.0 cm
   is 510.2 pt, so **42 lines** fit.
3. One line holds `--columns` sets of initials, so a page holds
   6 × 42 = **252**, and the page count is the number of initials divided
   by that, rounded up.

That number is an **estimate**. Word, Word on the web and LibreOffice each
fit a slightly different number of lines on a page — Word fits two fewer
than the arithmetic predicts on A5 — so treat it as "about 441 pages",
not a guarantee. The document itself is unaffected: no page breaks are
forced, so every page comes out as full as the program can make it.

So **a bigger font gives more pages** and **more columns give fewer
pages**. For the ~111,000 names in the sample CSV:

| Settings | Page | Initials per page | Pages |
| --- | --- | --- | --- |
| defaults (A5, 6 columns, 10 pt) | A5 | 252 | 441 |
| `--columns 5 --font-size 12` | A5 | 175 | 634 |
| `--page-width 21 --page-height 29.7` | A4 | 441 | 252 |
| `--columns 1` (one per line) | A5 | 42 | 2,642 |

### When the settings don't fit

Nothing is written and the script says what is wrong and what to change:

```
ERROR: 20 column(s) of 10 pt initials with 1 cm between them need 33.8 cm, but only 11.8 cm are left between the 1.5 cm offsets on A5.
Use at most 6 column(s), a smaller gap, a smaller font size, or smaller offsets.
```

The same happens when the offsets leave no room on the page, or when the
lines are too tall to fit even one between the top and bottom offsets.

---

## How the code is organised

`generate_initials_book.py` is one file with three steps and one small
class:

| | |
| --- | --- |
| `read_names_csv()` | reads the CSV, trying the usual separators and encodings |
| `names_to_initials()` | `Иван, Иванов` → `И. И.`, in one vectorised pandas pass |
| `Layout` | the five numbers above. Every other figure — the width of one set of initials, the lines per page, the page count — is a property derived from them and from the page size |
| `Layout.validate()` | the three rules: the values make sense, the columns fit across the page, a line fits down it. Raises `LayoutError` with the message you saw above |
| `build_initials_document()` | one paragraph per line, the columns separated by tabs, a page break where the layout says a page ends |
| `main()` | reads, validates, prints the plan, writes the document |

`LayoutError` extends `ValueError`, so `main()` catches an unreadable CSV
and an impossible layout in the same place.

---

## Running the tests (for developers)

The project ships with a small `pytest` suite covering the initials, the
CSV reading, the layout maths and its limits, and the document that comes
out. To run it:

```
pip install -r tests/test_requirements.txt
python -m pytest
```

A successful run looks like:

```
======================== test session starts ========================
collected 77 items

tests/test_generate_initials_book.py ....................    [100%]

======================== 77 passed in 0.9s ==========================
```

### Checking the real PDF

Unit tests can only prove the layout arithmetic is self-consistent. They
cannot tell you how many lines Word will actually put on a page — and when
the two disagree the book comes out wrong while every test still passes.
That is how a 378-page book once became a 755-page one, back when the
book was set in a narrower font and ran 7 columns to a line.

`tests/check_rendering.py` closes that gap: it builds the book, has
LibreOffice or Word turn it into a PDF, and compares the result with the
prediction. It needs one of those programs installed.

```
python tests/check_rendering.py
```

```
Rendering 20,000 initials with LibreOffice...

                   predicted      actual
lines on a page           42          42
pages                     69          69

Font: Oswald — the real one.

OK — +0% against the prediction, within the 10% allowed.
```

It also says which font went into the PDF, because a page measured with a
stand-in font is not the page that gets printed. On Linux it points
fontconfig at the `oswald.ttf` in this folder for the length of
the conversion, so the check works whether or not the font was ever
installed — nothing is added to your system.

A few percent either way is normal — the programs genuinely differ. A big
difference fails the check and says what the program really fits:

```
FAILED — 87% more pages than predicted. This program fits ~22 lines on a
page, not 42.
```

Worth running on any machine whose output matters, a Windows one above
all, since Word and LibreOffice do not paginate alike.

### Benchmarking the layouts

`tests/benchmark_initials_book.py` compares the build time and the page
count of different column counts — one line is one paragraph, so more
columns means fewer paragraphs to write:

```
python tests/benchmark_initials_book.py --limit 20000
```

```
 columns     lines   pages     time
       1    20,000     477    11.8s
       3     6,667     159     1.9s
       5     4,000      96     1.2s
       7     2,858      69     0.9s
```

---

## Troubleshooting

- **`file '…' not found in this folder`** — make sure the CSV is in
  the same folder as the script and the name matches exactly,
  including capitalisation.
- **`Could not parse '…' into at least two columns`** — open the file
  in a text editor and check that columns are separated by commas,
  semicolons, tabs, or `|`. If possible, re-save as UTF-8.
- **The initials read `È. È.` instead of `И. И.`** — the script will
  have warned that it fell back to `latin-1`, which accepts any file
  at all and so is only ever a guess. Re-save the CSV as UTF-8 and
  run it again.
- **The initials look like rectangles or the font looks wrong** — the
  Oswald font isn't installed on your system. Install it
  using one of the commands above and re-run. If you're on Windows,
  close Word completely before re-running so it picks up the newly
  installed font.
- **`… need 43.3 cm, but only 18.0 cm are left`** — the columns, the
  gap and the font size together need more width than the page has. The
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
  the [Linux steps](#linux-and-macos) in Quick start instead. This is the recommended fix.
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
├── oswald.ttf                              (bundled font, install once)
├── OFL.txt                                 (the font's licence)
├── generate_initials_book.py
├── run-windows.bat                         (Windows: double-click it)
├── run-linux-macos.sh                      (Linux/macOS: ./run-linux-macos.sh)
├── tests/
│   ├── conftest.py
│   ├── test_requirements.txt
│   ├── benchmark_initials_book.py          (layout speed comparison)
│   ├── check_rendering.py                  (real PDF vs the prediction)
│   └── test_generate_initials_book.py
├── book_signatures_only_names.csv          (your input)
└── book_signatures_initials.docx           (the generated output)
```

> The `.venv/` folder (if you made one on Linux) and the generated
> `book_signatures_initials.docx` are intentionally excluded from
> version control by the `.gitignore` at the root of the repository.
