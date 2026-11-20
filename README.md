# ECI Electoral Roll → Excel (Punjab 2027 Draft Roll)

Download, organise and extract the **Election Commission of India 2027 draft electoral
rolls** (Punjab, state code `S19`, Punjabi/Gurmukhi PDFs) into validated Excel workbooks:
one row per voter, plus part-level metadata and a full validation report.

> **Data privacy:** the PDFs and the Excel files built from them contain personal data
> (names, relatives, house addresses, ages, EPIC numbers). They are git-ignored and must
> never be committed or uploaded anywhere.

---

## What it does

| Step | Script | Input → Output |
|---|---|---|
| 1. Download | `download_rolls.py` | ECI website → `input/S19/AC<ac>/<part>.pdf` (verified PDFs) |
| 2. Organise *(only for bulk downloads with original names)* | `organize_rolls.py` | `2027-EROLLGEN-S19-86-…-PAN-18-WI.pdf` → `<dest>/S19/AC86/18.pdf` |
| 3. Extract + validate | `extract_voters_2027.py` | `input/S19/AC*/` → `output/<AC>.xlsx` + `output/RUN_SUMMARY.xlsx` |

Tested on 234 PDFs (2 random parts from every one of Punjab's 117 assembly
constituencies, 186,238 voters): **all validation checks pass; 0 errors**.

## Setup

```bash
python -m pip install -r requirements.txt     # Python 3.10+
```

## Usage

### 1. Download rolls

URL pattern (the script builds it):
`https://voters.eci.gov.in/eroll/2027/s19/ssr-draftroll/<ac>/2027-EROLLGEN-S19-<ac>-SSR-DraftRoll-Revision1-PAN-<part>-WI.pdf`

```bash
python download_rolls.py                         # 2 random parts from each AC 1-117 (sample)
python download_rolls.py --parts all             # every part of every AC (~24,700 PDFs)
python download_rolls.py --acs 1-10 --parts all  # every part of AC 1-10
python download_rolls.py --acs 2 --parts 1-10    # fixed parts
python download_rolls.py --per-ac 5 --seed 7     # different random sample
```

* Finds each AC's number of parts with HEAD requests (binary search), then downloads.
* Every file is verified (size = Content-Length, `%PDF` header, `%%EOF`, opens in PyMuPDF)
  and written atomically (`.partial` → rename). Re-running skips valid files (resume).
* Log: `input/S19/_download_log.xlsx`.

### 2. Organise an existing bulk download *(optional)*

Only needed if you already have thousands of PDFs with ECI's original file names in one
folder. Moves them into `<dest>/<STATE>/AC<ac>/<part>.pdf`.

```bash
python organize_rolls.py "D:\dwld" "D:\2027_voter_data" --dry-run   # preview, changes nothing
python organize_rolls.py "D:\dwld" "D:\2027_voter_data"             # move
python organize_rolls.py --verify "D:\2027_voter_data\_organize_logs\manifest_<time>.csv"
python organize_rolls.py --undo   "D:\2027_voter_data\_organize_logs\manifest_<time>.csv"
```

No-data-loss design: never overwrites (rename fails if target exists), byte-compares
duplicates, copy+verify+delete across drives, live crash-safe manifest, resumable, undo.
Files with unrecognised names are left untouched.

### 3. Extract to Excel

```bash
python extract_voters_2027.py input/S19 -o output -w 8      # all AC folders, 8 processes
python extract_voters_2027.py input/S19/AC2 -o output       # one AC
python extract_voters_2027.py input/S19 -o output --force   # redo ACs already done
```

* One workbook per AC: `output/AC2.xlsx` with sheets **Voters**, **Parts**, **Issues**.
* `output/RUN_SUMMARY.xlsx`: one row per AC (PASS/FAIL, voters vs cover total, errors).
* Resumable: ACs whose workbook exists are skipped; files are written atomically.
* Exit code `0` = all checks passed, `1` = at least one ERROR.
* Speed: ~0.4 s per PDF with `-w 8` on a 4-core laptop → full Punjab ≈ 2.5–3 h.
  Memory ≈ 0.55 GB per worker for a 200-part AC (use `-w 4` on 8 GB RAM).

## Output

See [docs/DATA_DICTIONARY.md](docs/DATA_DICTIONARY.md) for every column.

**Voters sheet** (one row per voter): state, assembly no/name/reservation, parliament
no/name, part no, district, tehsil, police station, main village, ward, pincode,
polling station no/name, section no/name, page, serial no, EPIC, voter name, relation
type (English + Punjabi), relative name, house no, age, gender (English + Punjabi),
photo status, warnings.

**Parts sheet** (one row per PDF): all cover-page metadata (dates, roll identification,
sections, polling station address/type, counts), summary-page counts and the result of
every validation check.

**Issues sheet**: every ERROR / WARN with PDF, serial number, check name and detail.

## Validation

Every PDF is validated; see [docs/VALIDATION.md](docs/VALIDATION.md).

* Voters parsed = cover total = summary-page total; serial numbers contiguous;
  male / female / third-gender counts match the cover.
* **Two independent extractions** of every page (text stream vs. drawn card boxes) must
  agree on every field of every voter.
* Page footers (AC, part, dates, page x of N) agree with the cover on every page.
* Field checks: EPIC present, name, relation, age 18–125, gender; no unparsed text.

Values are kept **exactly as printed** (including odd EPICs, Hindi-script names and
legacy-font glyphs); unusual values are flagged in `warnings`, never "corrected".

## Repository layout

```
download_rolls.py         step 1 - download + verify PDFs
organize_rolls.py         step 2 - safe bulk re-organiser (optional)
extract_voters_2027.py    step 3 - PDF -> Excel extractor + validator
requirements.txt
AGENTS.md                 instructions/context for AI coding agents (and humans)
CLAUDE.md                 Claude Code entry point (imports AGENTS.md)
docs/
  DATA_DICTIONARY.md      every output column
  PDF_LAYOUT.md           structure of the ECI PDFs + variants handled
  VALIDATION.md           checks, test procedure, latest results
input/   (git-ignored)    PDFs:  input/S19/AC<ac>/<part>.pdf
output/  (git-ignored)    Excel output
```

## Limitations

* Built for the **Punjabi (S19) 2027 draft-roll layout**. Other states/languages
  (e.g. `S24` Uttar Pradesh, Hindi) need their labels added — see AGENTS.md.
* Text is taken from the PDF text layer (no OCR). Scanned/image-only PDFs are not supported.
