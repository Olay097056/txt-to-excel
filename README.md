*English · [ภาษาไทย](README.th.md)*

# txt-to-excel

### ▶ [Try the detector in your browser](https://olay097056.github.io/txt-to-excel/)

Paste your own records or load the bundled synthetic sample. The validator is
ported line-for-line from `txt_to_excel.py` and runs entirely client-side —
nothing is uploaded.

---

Reads the raw export format of **ETWIN card-reader terminals**, converts it into
a colour-coded Excel workbook, and flags the records that look wrong — so a
human reviews 5 suspicious rows instead of eyeballing 2,700 identical-looking
ones.

The real job it does is answering **"which terminal is malfunctioning?"** When a
card reader starts dropping or mangling digits, nothing announces it. The
records keep arriving, they keep looking like records, and the problem only
surfaces weeks later when someone's attendance doesn't add up. This tool reads
the export, groups the damage by terminal, and names the units that are
misbehaving:

```
Total 60 records | clean 55 | suspicious 1 | invalid 4
Terminals with problems: 01, 02, 03, 04, 06
```

Each exported line is exactly 20 characters:

```
04  2026  08  08  0231  029405
──  ────  ──  ──  ────  ──────
unit year month day time  EN
(2)  (4)  (2) (2)  (4)   (6)
```

`unit` is the terminal number, `EN` the card/employee number, and the middle
four fields are the swipe timestamp.

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

**Layer 2 — per-terminal frequency (marked yellow).** For each terminal
independently, the tool learns which values that terminal normally emits. A unit
that reports month `08` in 99% of its records and `09` once has probably
glitched, even though `09` is a perfectly legal month and layer 1 will never
object. Doing this *per terminal* rather than across the whole file is what
turns "this file has errors" into "terminal 03 needs attention".

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

- `sample_data.txt` is generated, not real. Actual ETWIN exports contain terminal
  IDs, swipe times and EN numbers belonging to real people, and are gitignored.
- The Windows `.exe` is a PyInstaller build (~30 MB) and is not versioned —
  build it from source if you need one.

## Stack

Python 3, Tkinter for the GUI, openpyxl for the workbook. No other dependencies.
