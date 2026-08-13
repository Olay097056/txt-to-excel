*English · [ภาษาไทย](README.th.md)*

# txt-to-excel

Converts fixed-width scan logs into a colour-coded Excel workbook, flagging the
records that look wrong — so a human reviews 5 suspicious rows instead of
eyeballing 2,700 identical-looking ones.

Each input line is exactly 20 characters:

```
04  2026  08  08  0231  029405
──  ────  ──  ──  ────  ──────
unit year month day time  EN
(2)  (4)  (2) (2)  (4)   (6)
```

## The interesting part: two layers of detection

A fixed-width parser is trivial. The problem this actually solves is that
**corrupt records often look perfectly valid** — a scanner glitch turns a digit
into another digit, and the result is still six numerals in the right place.
So validation runs in two independent layers:

**Layer 1 — hard rules (marked red).** Structurally impossible values: a field
containing a non-digit, a year outside 2000–2099, month outside 1–12, a day
that doesn't exist in that month (leap years included), an invalid time, a line
whose length isn't 20. Control characters that Excel physically cannot write —
the usual souvenir of OCR and scanner output — are stripped before parsing
rather than crashing the export.

**Layer 2 — per-unit frequency (marked yellow).** For each machine
independently, the tool learns which values that machine normally emits. A unit
that reports month `08` in 99% of its records and `09` once has probably
glitched, even though `09` is a perfectly legal month and layer 1 will never
object.

Frequency analysis is applied **only** to year, month and day. Time and EN are
deliberately excluded: they have high cardinality — nearly every record has a
unique value — so "unusual for this machine" would flag everything and drown the
real findings. Knowing which fields a statistical check *shouldn't* run on is
most of what makes the check useful.

The output workbook colours each row accordingly and lists, per problem row,
exactly which field failed and why.

## Usage

```bash
pip install openpyxl
```

GUI — paste text directly, or import a file:

```bash
python txt_to_excel.py
```

CLI:

```bash
python txt_to_excel.py --file sample_data.txt --out result.xlsx
```

Try it against the included synthetic sample, which contains four hard-rule
violations and one frequency anomaly:

```
Total 60 records | clean 55 | suspicious 1 | invalid 4
```

## Notes

- `sample_data.txt` is generated, not real. Actual scan exports contain machine
  IDs, timestamps and EN numbers belonging to real people, and are gitignored.
- The Windows `.exe` is a PyInstaller build (~30 MB) and is not versioned —
  build it from source if you need one.

## Stack

Python 3, Tkinter for the GUI, openpyxl for the workbook. No other dependencies.
