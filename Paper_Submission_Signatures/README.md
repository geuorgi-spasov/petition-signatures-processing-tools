# Paper Submission

Two Python scripts that turn a raw database export of signatures into
PDFs ready for paper submission:

1. **`split_signatures_into_folders.py`** — splits the export into Word
   documents of 1000 signatures each (10 rows per A4 landscape page), and
   writes them into `signatures_docx/`. Each page footer reads
   `Стр. X, папка Y` and `Сдружение „Невидими животни"`.
2. **`convert_docx_to_pdf.py`** — converts every `.docx` in
   `signatures_docx/` into a `.pdf` in `signatures_pdf/`, skipping any
   already-converted file.

Both output folders are created automatically — you don't have to make
them by hand.

This workflow is independent of the book-of-initials one — you only
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

The sample data is synthetic and anonymised and contains no real personal information in compliance with GDPR.

---

## Quick start

Follow **one** of the two paths below — whichever matches your computer.

### Windows

1. **Install Python** from <https://www.python.org/downloads/>, ticking
   *"Add Python to PATH"* during the installation.
2. **Check you can make PDFs.** If you have Microsoft Word on this
   computer, there is nothing to do. If you don't, install
   [LibreOffice](https://www.libreoffice.org/download/) — it's free.
   (Word in a browser at office.com doesn't count; the script opens the
   Word program itself.)
3. **Put your `Signatures_from_the_database_raw.csv` in this folder.**
4. **Double-click `run.bat`.**

### Linux and macOS

1. **Install Python** — on Ubuntu/Mint/Debian:

   ```
   sudo apt install python3 python3-venv
   ```

2. **Install LibreOffice**, which is what makes the PDFs:

   ```
   sudo apt install libreoffice
   ```

3. **Put your `Signatures_from_the_database_raw.csv` in this folder.**
4. **Open Terminal in this folder and run:**

   ```
   ./run.sh
   ```

### What happens then

`run.bat` and `run.sh` do the whole job. First they fill
`signatures_docx/` with one Word file per 1000 signatures — about 13
minutes for 111 files. Then they fill `signatures_pdf/` with a PDF of
each one, printing which program they are using. That takes about 5
minutes with LibreOffice, or around 40 with Word, which is much slower
to start up.

Both halves skip anything they have already done, so if it is
interrupted you can start it again and it picks up where it left off.

### Prefer to run the steps yourself?

The launchers do no magic. On Windows:

```
pip install -r requirements.txt
python split_signatures_into_folders.py
python convert_docx_to_pdf.py
```

On Linux, Python has to be installed into a folder of its own, so it is
one line more:

```
python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt
python split_signatures_into_folders.py
python convert_docx_to_pdf.py
```

Run the `source .venv/bin/activate` part again in any new terminal.

---

## The two steps in detail

### Step 1 — split signatures into submission documents

Place `Signatures_from_the_database_raw.csv` (the unmodified database
export, headers and all) next to the scripts. The script auto-detects
the separator and encoding.

Run:

```
python split_signatures_into_folders.py
```

For every 1000 signatures the script creates one Word document in
`signatures_docx/`, named:

```
signatures_docx/Папка 1 с подписи от <first-id> до <last-id>.docx
signatures_docx/Папка 2 с подписи от <first-id> до <last-id>.docx
…
```

Each A4 landscape page (29.7 × 21.0 cm) contains 10 rows. The footer shows a page number
that runs *continuously across all folders* (`Стр. 1, папка 1` …
`Стр. 100, папка 1`, then `Стр. 101, папка 2` …) followed by the
organisation name.

### Step 2 — convert the .docx files to PDF

Once step 1 has produced the `.docx` files in `signatures_docx/`, run:

```
python convert_docx_to_pdf.py
```

The script creates `signatures_pdf/` if it doesn't exist, converts
every `.docx` from `signatures_docx/` and writes the matching `.pdf`
into `signatures_pdf/`. Any file whose `.pdf` already exists is
skipped, so this is safe to re-run.

The conversion tool is chosen automatically: Microsoft Word (via
`docx2pdf`) wherever it is installed, and LibreOffice in headless mode
otherwise — which is always the case on Linux. The script prints which
one it is using. Either way the pending files are converted in a single
batch so the tool starts only once.

---

## Customising the scripts

Both scripts begin with a `# Configuration` block of `UPPER_CASE`
variables. Edit, save, and re-run.

**`split_signatures_into_folders.py`** exposes:

- `INPUT_CSV` — the CSV file name
- `OUTPUT_DOCX_FOLDER` — where the `.docx` files are written
  (default `"signatures_docx"`)
- `ROWS_PER_FILE`, `ROWS_PER_PAGE` — grouping (defaults 1000 and 10)
- `PAGE_WIDTH_CM`, `PAGE_HEIGHT_CM` and the four margin constants —
  page layout, in centimeters
- `ROW_HEIGHT_CM` — height of each table row in centimeters
- `BODY_FONT`, `FOOTER_FONT`, `BODY_FONT_SIZE_PT` — fonts
- `DEFAULT_COLUMN_WIDTHS_CM` — preset column widths in centimeters,
  tuned for 6 columns and auto-scaled to fit the page
- `ORGANIZATION_NAME` — text printed at the bottom of every page

**`convert_docx_to_pdf.py`** exposes:

- `INPUT_FOLDER` — defaults to `"signatures_docx"`
- `OUTPUT_FOLDER` — defaults to `"signatures_pdf"`

If you change the folder names in step 1, change them to match in
step 2.

---

## Performance

The slow part of this workflow is the .docx → .pdf conversion, because
the converter (Microsoft Word where it is installed, LibreOffice
otherwise) has to render each document, and starting it has a cost.

`convert_docx_to_pdf.py` converts **every pending file in a single
batch**: it stages the pending .docx files in a temporary folder and
converts them in one shot, so the converter starts once rather than
once per file.

How much this batching helps depends heavily on the platform:

- **Windows / macOS (Microsoft Word):** Word has a high start-up cost
  (several seconds) and `docx2pdf` starts it afresh for every
  per-file call. Batching amortises that start-up over the whole run,
  so it is typically **3–10× faster** than one call per file.
- **Linux (LibreOffice):** LibreOffice's start-up cost is comparatively
  small, so per-file and batch runs come out **roughly equal** — the
  actual per-document rendering (~2–3 s/file) dominates either way.
  Batching still does no harm, and keeps the behaviour identical
  across platforms.

You can measure the difference on your own machine with
`benchmark_conversion.py` (see below).

When the script finishes you'll see a line like:

```
Done in 42.3s (~0.42s per converted file). Converted: 100, skipped: 0, failed: 0.
```

so you can see exactly how long the run took.

### Measuring it yourself

`benchmark_conversion.py` runs both strategies (per-file vs batch) on
the same files and reports the speedup. It uses the same conversion
backend as the main workflow — Word where it is installed, LibreOffice
otherwise — so it works on every platform. Run it after
step 1 has produced some .docx files:

```
python tests/benchmark_conversion.py
```

Or, to do a quick check with only the first 10 files:

```
python tests/benchmark_conversion.py --limit 10
```

You'll get output similar to:

```
Benchmarking with 10 of 100 .docx file(s) from 'signatures_docx'.

Strategy 1: per-file convert() calls...
  Total: 52.4s — 5.24s per file

Strategy 2: single batch convert() call...
  Total: 8.7s — 0.87s per file

Batch is 6.0× faster — saved 43.7s on 10 file(s).
```

---

## Running the tests (for developers)

The project ships with a small `pytest` test suite covering the pure
helpers (CSV reading, column-width scaling, document construction,
file discovery). The tests don't require Microsoft Word or LibreOffice
to be installed — the conversion call itself is not unit-tested.

To run them:

```
pip install -r tests/test_requirements.txt
python -m pytest
```

A successful run looks like:

```
======================== test session starts ========================
collected 35 items

tests/test_convert_docx_to_pdf.py ........................         [ 68%]
tests/test_split_signatures.py ...........                       [100%]

======================== 35 passed in 6.3s ==========================
```

---

## Troubleshooting

- **`file '…' not found in this folder`** — the CSV must be in the
  same folder as the script and the name must match exactly.
- **`Could not parse '…' as a multi-column CSV`** — open the file in
  a text editor and check that columns are separated by commas,
  semicolons, tabs, or `|`. Re-save as UTF-8 if possible.
- **`folder '…' not found`** when running step 2 — run step 1 first
  so that `signatures_docx/` gets created.
- **"Skipping N Word lock file(s)"** — the script found one or more
  `~$<name>.docx` files in `signatures_docx/`. These are temporary
  lock files Word creates while a document is open, not real `.docx`
  files. Close Word and they'll disappear. The script already skips
  them, so this is just informational.
- **`ERROR: the 'docx2pdf' package isn't installed`** (Windows/macOS)
  — run `pip install -r requirements.txt`.
- **`ERROR: LibreOffice not found`** (Linux) — install it with
  `sudo apt install libreoffice`.
- **`ERROR: no program found that can turn .docx files into PDFs`**
  (Windows/macOS) — neither Microsoft Word nor LibreOffice is
  installed. Install either one; LibreOffice is free, see
  [Quick start](#quick-start) step 2.
- **Step 2 hangs or errors out** — the tool it reported using (Word or
  LibreOffice) must be able to open .docx files. On Windows, close any
  open Word windows before running.
- **`error: externally-managed-environment`** (Linux) — your
  distribution blocks installing packages into the system Python. Use
  the [Linux steps](#linux-and-macos) in Quick start instead.
- **`pip` is not recognised** — Python wasn't added to PATH. Re-install
  Python and tick *"Add Python to PATH"*, or use `py -m pip …`
  instead.

If your problem isn't here, paste the error and the relevant script
into an AI assistant (Claude, ChatGPT, etc.) — it can usually
diagnose it from the script and the message alone.

---

## Folder layout

```
Paper_Submission_Signatures/
├── README.md
├── requirements.txt
├── split_signatures_into_folders.py
├── convert_docx_to_pdf.py
├── run.bat                                 (Windows: double-click to start)
├── run.sh                                  (Linux/macOS: ./run.sh)
├── tests/
│   ├── conftest.py
│   ├── test_requirements.txt               (extras for running tests)
│   ├── benchmark_conversion.py             (optional, for measuring speed)
│   ├── test_split_signatures.py
│   └── test_convert_docx_to_pdf.py
├── Signatures_from_the_database_raw.csv    (your input)
├── signatures_docx/                        (auto-created by step 1)
│   ├── Папка 1 с подписи от ... до ....docx
│   ├── Папка 2 с подписи от ... до ....docx
│   └── ...
└── signatures_pdf/                         (auto-created by step 2)
    ├── Папка 1 с подписи от ... до ....pdf
    ├── Папка 2 с подписи от ... до ....pdf
    └── ...
```

> The `.venv/` folder and the generated `signatures_docx/` and
> `signatures_pdf/` folders are excluded from version control via
> `.gitignore`.
