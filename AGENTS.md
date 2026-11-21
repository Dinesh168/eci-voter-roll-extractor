# AGENTS.md

Instructions and context for AI coding agents (Claude, Codex, Copilot, …) and new
developers working on this repository. Read this file fully before changing code.
Human-facing usage is in [README.md](README.md).

---

## 1. Project goal

Turn the Election Commission of India (ECI) **2027 draft electoral-roll PDFs** into
**Excel workbooks with exact, validated voter data**. Supported:

* **Punjab** – state code `S19`, 117 assembly constituencies, ~24,700 parts,
  Punjabi/Gurmukhi text → `extract_voters_2027.py`.
* **Uttar Pradesh** – state code `S24`, 403 assembly constituencies, Hindi/Devanagari
  text → `extract_voters_up_2027.py` (standalone, same design; see section 4b).

The owner's requirements (non-negotiable):

1. **Exact data.** Every value must be exactly what the PDF contains. Never "correct",
   transliterate, guess or drop data. Unusual values are kept and *flagged*.
2. **Validated.** Every PDF must be proven complete and consistent (counts, serial
   numbers, cross-checks). A run with any ERROR is a failed run.
3. **Excel only** for output (no CSV).
4. **No data loss** when moving/organising files (never overwrite, never delete before
   a verified copy exists).
5. Must scale to ~25k PDFs on a Windows laptop and be resumable.

## 2. Repository map

| File | Purpose |
|---|---|
| `download_rolls.py` | Download PDFs from ECI into `input/<STATE>/AC<ac>/<part>.pdf`. Finds part count per AC via HEAD + binary search; verifies each PDF; resumable. |
| `organize_rolls.py` | Safely move/copy bulk downloads with original ECI names into `<dest>/<STATE>/AC<ac>/<part>.pdf`. Manifest, verify, undo. |
| `extract_voters_2027.py` | Core: PDF → Excel extractor + validator for Punjab (multiprocess, resumable). |
| `extract_voters_up_2027.py` | Same for Uttar Pradesh (Hindi), plus glyph check and proven hidden-text repairs. |
| `docs/PDF_LAYOUT.md` | How the PDFs are structured and every layout variant found so far (Punjab). |
| `docs/PDF_LAYOUT_UP.md` | Uttar Pradesh differences, Hindi text-layer quirks, glyph fixes. |
| `docs/DATA_DICTIONARY.md` | Every output column. |
| `docs/VALIDATION.md` | Validation rules, test procedure, latest results. |
| `input/`, `output/` | **Git-ignored data.** Never commit. |

## 3. Environment

* Owner works on **Windows 11**, Python **3.10**, PowerShell / cmd. Paths like `D:\...`.
* Dependencies: `pip install -r requirements.txt` (PyMuPDF, xlsxwriter, openpyxl).
* Run Python with `-I` when reading downloaded files (isolated mode).
* Console encoding: scripts call `sys.stdout.reconfigure(encoding="utf-8")`; Punjabi
  output breaks on the default cp1252 console otherwise. `python -I` ignores
  `PYTHONIOENCODING`, so keep the in-script reconfigure.
* Agent tooling gotcha: inline `python - <<'PY'` heredocs run through Git-Bash may
  mangle backslashes (`"\\n"`, `'\\'`). For multi-line code edits or patch scripts,
  write a `.py` file and run it instead.
* Multiprocessing on Windows needs the `if __name__ == "__main__":` guard and top-level
  (picklable) worker functions; keep it that way.

## 4. How `extract_voters_2027.py` works

Pipeline per PDF (`process_pdf`):

1. `page_items(page)`: PyMuPDF `get_text("dict")` lines → text + bbox. Re-joins line
   pieces PyMuPDF splits mid-word on the same baseline (`ਹਰਪ੍ਰੀ` + `ਤ ਕੌਰ`), then
   `norm()` fixes PDF glyph glitches only (doubled ੰ/ੱ marks, sihari stored before its
   consonant `ਿਪਤਾ` → `ਪਿਤਾ`). **norm must never change real content.**
2. `parse_cover(page 1)`: assembly/parliament/part, dates, roll identification, sections,
   village, ward, police station, tehsil, district, pincode, polling station, counts.
   `_after()` reads "label → value" pairs; a blank value (next line is `:` or another
   label in `COVER_LABELS`) returns `""`.
3. Summary page = page whose heading is `ਵੋਟਰਾਂ ਦੀ ਸਾਰਣੀ` (not always the last page).
   A later page starting with `ਭਾਗ ਵਿੱਚ ਸੈਕਸ਼ਨਾਂ…` holds the full section list when the
   cover list overflows; it is merged into `cover["sections"]`.
4. Voter pages (contain `ਉਮਰ :`):
   * **Method A** `parse_voter_page`: one text stream, cut at each card start
     (serial-number line followed by EPIC/photo/name line) → `parse_card`.
   * **Method B** `box_cards`: card rectangles from the drawn grid (`card_boxes`, vertical
     border lines), each text line assigned to the box containing its centre; inside a box
     the top row (SL, EPIC) comes first, the right 30% is the photo placeholder column,
     the rest is the text column → `parse_card`.
   * Every field of every card must be identical in A and B, else `crosscheck` ERROR.
5. Per-card checks, per-PDF checks, metadata checks (see `docs/VALIDATION.md`).
6. `process_ac` writes `<AC>.xlsx` (Voters / Parts / Issues) via `write_xlsx`
   (xlsxwriter constant-memory, temp file + atomic rename, retries if file is open in
   Excel). `main` runs ACs in a `ProcessPoolExecutor`, skips finished ACs, and rewrites
   `RUN_SUMMARY.xlsx` after each AC.

## 4b. How `extract_voters_up_2027.py` (Uttar Pradesh, Hindi) differs

Same pipeline and output layout as section 4, with Hindi labels. The Hindi PDFs need
extra care (details and examples in `docs/PDF_LAYOUT_UP.md`):

1. **Text source = ActualText** (logical order). Glyph text (visual order) is used only
   for the **glyph check**: per page, extracted chars == drawn glyph chars
   (order-independent) — ERROR otherwise.
2. `patch_cmaps(doc)` corrects wrong ToUnicode entries in memory (`GLYPH_FIXES`, keyed by
   font base name + glyph id). **Only add an entry after drawing the glyph outline from
   the embedded font and looking at it** (fontTools `SVGPathPen` on the `FontFile2`
   stream). Fonts are subsets: a glyph with no outline in one PDF may have one in another.
3. `page_items` repairs hidden-text artefacts (zero-width duplicate after a joiner,
   doubled signs, repeated pieces — `dedupe_chars`, `undouble`, `collapse_repair`) **only
   when the drawn glyphs of that line prove the repair** (repaired chars == drawn chars,
   unrepaired ≠). Drawn glyphs are assigned to the nearest line (glyph left edge, then
   vertical centre). Never repair without that proof.
4. `explain_render_quirks`: two producer rendering bugs (consonant drawn twice around a
   reph; `रृ` drawn as reph + `ऋ`) are proven by rendering; the page gets WARN
   `render_quirk` only when the difference is *exactly* the predicted occurrences.
5. EPIC: national `AAA9999999` and old UP `UP/12/345/0123456` are valid; card-start
   detection also accepts backslash/lowercase variants (WARN `epic_format`).
6. `PyMuPDF get_texttrace()` crashes (`none_dealloc`) after many calls in one process —
   don't use it in the extractor; for ad-hoc tools run one PDF per process
   (`multiprocessing.Pool(maxtasksperchild=1)`; Python 3.10 has no `max_tasks_per_child`).

ERROR vs WARN policy:
* **ERROR** = extraction may be wrong or incomplete (count mismatch, SL gap, A≠B,
  unparsed text, missing mandatory field, footer mismatch). Makes the PDF/AC `FAIL`.
* **WARN** = the PDF itself contains something unusual, extracted faithfully
  (odd EPIC format, blank relative/house no., Hindi script, legacy glyphs `Ã Ç`,
  blank cover field, section list truncated on cover).
* Do not downgrade an ERROR to WARN unless you have *proven* (render the page, look at
  it) that the PDF really contains that and the extraction is exact.

## 5. Rules for agents

1. **Never commit or upload data** (`input/`, `output/`, `*.pdf`, `*.xlsx`). They contain
   personal data. Don't paste voter rows into issues/PRs/external services.
2. **Never weaken validation** to make a run pass. Fix the parser, or prove the PDF is
   the source of the oddity and flag it as WARN.
3. **Never alter extracted values** beyond `norm()` glyph-glitch fixes.
4. After any change to an extractor, run its regression test (section 6)
   and confirm **0 errors and the same voter total**. Report the numbers.
5. When you meet a new layout variant: render the page/card
   (`page.get_pixmap(clip=rect, dpi=200).save(...)`) and look at it before deciding.
   Then add it to `docs/PDF_LAYOUT.md` (Punjab) or `docs/PDF_LAYOUT_UP.md` (UP).
9. Examples in code/docs must be synthetic or word fragments — never real voter names,
   EPICs or addresses copied from the PDFs (the repo is public).
6. `organize_rolls.py` must keep its no-data-loss guarantees: no `os.replace` onto
   existing files, verify before delete, live manifest.
7. Keep output Excel-only unless the owner asks otherwise.
8. Keep Windows compatibility (paths, file locking, multiprocessing guard).

## 6. Regression test (do this after every change)

No PDFs are in git. Re-create the sample set (2 random parts × 117 ACs, fixed seed):

```bash
python download_rolls.py                               # ~3 min, 234 PDFs -> input/S19/
python extract_voters_2027.py input/S19 -o output/S19_test -w 8 --force
```

Expected (seed `2027`, verified 2026-11-20):

```
Done in ~2 min.  117 ACs, 186,238 voters, 0 AC(s) with errors.
```

`output/S19_test/RUN_SUMMARY.xlsx` → all 117 rows `PASS`; Issues sheets contain only
WARN rows (22 in total: 17 `epic_format`, 4 `metadata_blank`, 1 `relative_missing`).
If ECI republishes the rolls, totals may change; then re-establish the baseline and
update this section and `docs/VALIDATION.md`.

**Uttar Pradesh** (1 random part × 403 ACs, seed `2027`):

```bash
python download_rolls.py --state S24 --lang HIN --acs 1-403 --per-ac 1   # 403 PDFs -> input/S24/
python extract_voters_up_2027.py input/S24 -o output/S24_test -w 6 --force
```

Expected (verified 2026-11-21):

```
Done in ~4.3 min.  403 ACs, 309,733 voters, 0 AC(s) with errors.
```

All 403 rows `PASS`; WARN rows only (136: 71 `metadata_blank`, 53 `epic_format`,
12 `render_quirk`). On 8 GB RAM use `-w 4`.

## 7. What the owner (human) needs to do

1. On the production PC: `pip install -r requirements.txt`.
2. Get all PDFs, either:
   * `python download_rolls.py --parts all` (writes `input/S19/AC*/<part>.pdf`), or
   * organise an existing bulk download:
     `python organize_rolls.py "D:\dwld" "D:\2027_voter_data" --dry-run`, then without
     `--dry-run`. (The owner has ~199k PDFs in `D:\dwld`, several states; 588 had
     unrecognised names and are left in place.)
3. Run: `python extract_voters_2027.py D:\2027_voter_data\S19 -o D:\voter_excel -w 8`
   (≈ 2.5–3 h for all of Punjab; resumable — just re-run after any stop).
4. Open `RUN_SUMMARY.xlsx`. Every AC must be `PASS`. For any `FAIL`, open that AC's
   **Issues** sheet and hand the PDF + issue to an agent.
5. Optionally review rows with a non-empty `warnings` column (~0.3% of rows: values
   printed oddly in the source PDF).
6. **Uttar Pradesh:** same steps with `--state S24 --lang HIN` for downloading, and
   `python extract_voters_up_2027.py D:\2027_voter_data\S24 -o D:\voter_excel_up -w 6`
   (≈ 0.6 s per PDF). ~1.6% of UP rows carry a `warnings` flag (mostly `मो0`-style
   digits or Latin letters in names).

## 8. Open tasks / roadmap

* **UP full run:** only a 403-PDF sample (1 part per AC) has been validated. On the
  full UP set, new glyph variants may appear → `glyph_check` ERROR; verify the glyph by
  drawing it, add to `GLYPH_FIXES`, re-run both regressions.
* **Other states / languages:** copy the closer extractor (Hindi states:
  `extract_voters_up_2027.py`), replace the label strings (name, relation types, house
  no., age/gender, section header, footer, cover labels, summary heading), download a
  random sample per AC and iterate until clean. Do not break S19 or S24.
* Optional: a final "ALL_ACS" workbook is impossible for all of Punjab (~21M rows >
  Excel's 1,048,576 limit); per-AC workbooks are the design.
* Optional: unit tests with small synthetic PDFs (no real voter data in the repo).

## 9. Glossary

| Term | Meaning |
|---|---|
| AC | Assembly Constituency (ਵਿਧਾਨ ਸਭਾ ਚੋਣ ਹਲਕਾ / विधानसभा निर्वाचन क्षेत्र), 1–117 Punjab, 1–403 UP |
| Part / ਭਾਗ | One polling-station roll = one PDF (usually 400–1,500 voters) |
| Section / ਸੈਕਸ਼ਨ | Locality inside a part |
| EPIC | Elector's Photo Identity Card number (normally `AAA9999999`) |
| SL / ਲੜੀ ਨੰ | Serial number of a voter inside the part |
| Cover | Page 1: metadata + counts; Summary: `ਵੋਟਰਾਂ ਦੀ ਸਾਰਣੀ` page with totals |
| ਪਿਤਾ / ਪਤੀ / ਮਾਤਾ / ਪਤਨੀ / ਹੋਰ | Father / Husband / Mother / Wife / Other |
| ਪੁਰਸ਼ / ਇਸਤਰੀ / ਤੀਜਾ ਲਿੰਗ | Male / Female / Third gender |
| पिता / पति / माता / पत्नी / अन्य | Father / Husband / Mother / Wife / Other (UP) |
| पुरुष / महिला / तृतीय लिंग | Male / Female / Third gender (UP) |
| ActualText | Hidden logical-order text attached to drawn glyphs in the PDF (UP text source) |
