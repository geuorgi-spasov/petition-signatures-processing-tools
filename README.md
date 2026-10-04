# NGO Petition Signatures Processing Tools

Two independent Python toolkits built to help a Bulgarian animal-welfare
NGO (*Сдружение „Невидими животни"*) process the signatures collected
for its petitions and turn them into print-ready documents.

Both toolkits read an exported list of signatories, but the
paper-submission toolkit uses the full record for each signatory
(names, email, date, etc.), while the book-of-initials toolkit needs
only the names. Each toolkit lives in its own folder with its own
detailed README, dependencies, and test suite, and can be used
completely independently of the other.

The code is written to be readable and maintainable: small, documented
functions, configuration grouped at the top of each script, and a
`pytest` suite covering the core logic. It runs on Windows, macOS, and
Linux.

---

## The two toolkits

### 📖 [`Book_Initials_Signatures/`](Book_Initials_Signatures/)

Turns a CSV of signatory names into a printable A5 book listing just
their initials (e.g. `И. И.`), in the Bebas Neue Cyrillic font, several
per line. The page size, the number of columns, the page offsets, the
distance between the initials, the line spacing and the font size are all
parameters with sensible defaults; the script reports how many pages the
book will have before building it, and rejects settings that cannot fit.

See [`Book_Initials_Signatures/README.md`](Book_Initials_Signatures/README.md)
for setup and usage.

### 📄 [`Paper_Submission_Signatures/`](Paper_Submission_Signatures/)

Turns a raw database export of signatures into paginated Word documents
(1000 signatures each, 10 per A4 landscape page, with continuous page
numbering and a footer) and converts them to PDFs ready for paper
submission. The PDF conversion runs on Microsoft Word (Windows/macOS) or
LibreOffice (Linux), whichever is available.

See [`Paper_Submission_Signatures/README.md`](Paper_Submission_Signatures/README.md)
for setup and usage.

---

## A note on the sample data

The CSV files included for demonstration are **synthetic and
anonymised** — names have been randomised and emails replaced with
placeholders, so the repository contains **no real personal
information**, in line with GDPR.

---

## About this code

These tools were written with the help of an AI assistant and reviewed
by a human. They are intentionally written to be understandable by
people who aren't programmers — if any part is unclear, pasting the
relevant code into an AI assistant will usually explain or adapt it.

## Tech

Python 3.9+ · pandas · python-docx · pytest · docx2pdf (Windows/macOS) /
LibreOffice (Linux)
