# Validation

`extract_voters_2027.py` validates every PDF while extracting. Any **ERROR** marks the
PDF and its AC as `FAIL` (exit code 1). **WARN** means the PDF itself contains something
unusual that was extracted faithfully.

## Checks per PDF

| Check (Parts column / Issues `check`) | Rule | Level |
|---|---|---|
| `count_vs_cover` | voters parsed = cover total | ERROR |
| `count_vs_summary` | voters parsed = summary-page total | ERROR |
| `sl_sequence` | serial numbers exactly `start_sl … end_sl`, no gaps/duplicates | ERROR |
| `male_count` / `female_count` / `third_count` | parsed gender counts = cover counts | ERROR |
| `crosscheck` (`crosscheck_cards`, `crosscheck_field`) | **two independent extractions** of every voter page — A: text stream cut at card starts; B: text assigned to drawn card boxes by position — give the same serial numbers and identical `epic, voter_name, relation_type_pa, relative_name, house_no, age, gender_pa` | ERROR |
| `cover_totals` | male + female + third = total; end_sl − start_sl + 1 = total | ERROR |
| `summary_vs_cover` | summary main row = cover counts; summary total row = main row | ERROR |
| `page_footers` | on every page: AC no/name, publication date, qualifying date, `page k of N` agree with cover / page position | ERROR |
| `part_no_footer` | part number in every footer = cover | ERROR |
| `part_no_filename` | part in file name (`…-PAN-18-WI.pdf` or `18.pdf`) = cover | ERROR |
| `assembly_no_folder` | folder `AC<n>` = cover assembly number | ERROR |
| `epic_missing`, `name_missing`, `relation_type`, `age` (18–125), `gender` | per voter | ERROR |
| `unparsed` | any text inside a card not assigned to a field | ERROR |
| `section_missing` | voter page without section footer | ERROR |
| `metadata_missing` | assembly, part, dates or state code not found on cover | ERROR |
| `exception` | PDF could not be processed | ERROR |
| `epic_format` | EPIC not `AAA9999999` | WARN |
| `relative_missing` | relative name blank in PDF | WARN |
| `metadata_blank` | district/tehsil/police station/village/pincode/PS name/address/type blank on cover | WARN |
| `pincode_format`, `ps_no_vs_part` | pincode not 6 digits; polling-station no ≠ part no | WARN |
| `section_not_on_cover`, `sections_seen` | page section not in cover list / cover section without pages | WARN |
| `crosscheck_skipped` | no card boxes found on a voter page | WARN |

Per AC: `duplicate_epic` (WARN), `assembly_mixed` (ERROR).

## Regression test

```bash
python download_rolls.py                                     # seed 2027: 2 random parts x 117 ACs
python extract_voters_2027.py input/S19 -o output/S19_test -w 8 --force
```

## Latest results (2026-11-20, rolls published 12-11-2026)

Sample: 234 PDFs (2 random parts from each of the 117 ACs, seed `2027`).

| Metric | Result |
|---|---|
| ACs PASS | **117 / 117** |
| PDFs PASS | **234 / 234** |
| Voters extracted | **186,238** = sum of cover totals **186,238** |
| ERRORs | **0** |
| Stream vs card-box cross-check | identical on every field of every voter, 0 pages skipped |
| Page footers vs cover | consistent on every page of every PDF |
| Text on voter pages outside a card | none |
| WARNs | 22 — `epic_format` 17, `metadata_blank` 4, `relative_missing` 1 |
| Rows with a `warnings` flag | ≈0.3% (blank house no., wraps, Hindi script, legacy glyphs) |
| Gender | Male 96,113 · Female 90,120 · Third gender 5 |
| Relation | Father 110,011 · Husband 73,970 · Mother 1,358 · Other 899 |

Extra independent checks done once on this sample:

* **Different PDF engine:** pdfplumber re-extraction of serial number + EPIC for 12 random
  PDFs (8,532 voters): identical for all voters (2 apparent differences were the ad-hoc
  checker matching a house number; Excel values confirmed against rendered pages).
* **Visual:** 12 random cards (incl. wrapped values, blank fields, legacy glyphs) rendered
  and compared by eye with the Excel row: all identical.
* **Cover audit:** every cover line is either a label or captured in a field.

## What "100% accurate" means here

The Excel equals the PDF's text content, verified by counts, two independent
segmentations, a second PDF engine and visual checks. Where the PDF itself prints
something odd (Hindi script, legacy glyphs such as `Ã`, non-standard EPICs, blank
fields) the value is kept exactly and flagged — it is not "corrected".
