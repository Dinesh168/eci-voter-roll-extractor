# Data dictionary

Output of `extract_voters_2027.py`: one workbook per AC (`<AC>.xlsx`, sheets **Voters**,
**Parts**, **Issues**) plus `RUN_SUMMARY.xlsx` (sheet **Summary**).

All text is Unicode, exactly as in the PDF text layer (Gurmukhi unless the PDF itself
uses another script). Blank cells mean the PDF has no value there.

## Voters sheet: one row per voter (33 columns)

| Column | Type | Source in PDF | Example |
|---|---|---|---|
| `ac_folder` | text | input folder name | `AC2` |
| `source_pdf` | text | file name | `148.pdf` |
| `state_code` | text | cover title `ਵੋਟਰ ਸੂਚੀ 2027 S19 ਪੰਜਾਬ` | `S19` |
| `state_name` | text | cover title | `ਪੰਜਾਬ` |
| `assembly_no` | int | cover | `2` |
| `assembly_name` | text | cover | `ਭੋਆ` |
| `reservation` | text | cover `(…)` after assembly name | `ਅ.ਜ.` (SC) or `ਜਨਰਲ` |
| `parliament_no` | int | cover | `1` |
| `parliament_name` | text | cover | `ਗੁਰਦਾਸਪੁਰ` |
| `part_no` | int | cover `ਭਾਗ ਨੰਬਰ` | `148` |
| `district` | text | cover `ਜ਼ਿਲ੍ਹਾ` | `ਪਠਾਨਕੋਟ` |
| `tehsil` | text | cover `ਤਹਿਸੀਲ` | `ਪਠਾਨਕੋਟ` |
| `police_station` | text | cover `ਪੁਲਿਸ ਸਟੇਸ਼ਨ` | `ਨਰੋਟ ਜੈਮਲ ਸਿੰਘ` |
| `main_village` | text | cover `ਮੁੱਖ ਪਿੰਡ/ ਸ਼ਹਿਰ` | `ਸਕੋਲ` |
| `ward` | text | cover `ਵਾਰਡ` (blank for most villages) | `ਵਾਰਡ ਨੰ. 9` |
| `pincode` | text | cover `ਪਿੰਨ ਕੋਡ` | `145027` |
| `polling_station_no` | int | cover | `148` |
| `polling_station_name` | text | cover | `ਕੋਟਦੁੱਨਾ` |
| `section_no` | int | footer of the voter's page `ਸੈਕਸ਼ਨ ਨੰ. ਅਤੇ ਨਾਮ` | `1` |
| `section_name` | text | same | `ਸਕੋਲ` |
| `page_no` | int | 1-based page of the PDF | `3` |
| `sl_no` | int | card serial number | `1` |
| `epic` | text | card, top right; kept exactly (odd formats flagged) | `ABC1234567` |
| `voter_name` | text | `ਨਾਮ :` | `ਹਰਜੀਤ ਸਿੰਘ` |
| `relation_type` | text | English of `relation_type_pa` | `Father` / `Husband` / `Mother` / `Wife` / `Other` |
| `relation_type_pa` | text | label before `ਦਾ ਨਾਮ:` | `ਪਿਤਾ` / `ਪਤੀ` / `ਮਾਤਾ` / `ਪਤਨੀ` / `ਹੋਰ` |
| `relative_name` | text | value after the relation label | `ਜਸਵੰਤ ਸਿੰਘ` |
| `house_no` | text | `ਮਕਾਨ ਨੰ. :` (may be a full address; may be blank) | `24B` |
| `age` | int | `ਉਮਰ :` (age as on the qualifying date) | `48` |
| `gender` | text | English of `gender_pa` | `Male` / `Female` / `Third Gender` |
| `gender_pa` | text | `ਲਿੰਗ :` | `ਪੁਰਸ਼` / `ਇਸਤਰੀ` / `ਤੀਜਾ ਲਿੰਗ` |
| `photo_status` | text | photo placeholder text | `ਫੋਟੋ ਉਪਲਬਧ ਹੈ` |
| `warnings` | text | `; `-separated flags, blank = clean (see below) | `wrapped house_no` |

### `warnings` values

| Flag | Meaning |
|---|---|
| `wrapped <field>` | value continued on the next line in the PDF; joined with a space |
| `house_no empty` | no house number printed |
| `relative_name empty in PDF` | relation label printed with no name |
| `Devanagari (Hindi) script in <field>` | value printed in Hindi script |
| `latin/digit in <field>` | Latin letters or digits inside a name |
| `unusual character X in <field> (as printed)` | legacy-font glyphs such as `Ã`, `Ç` printed in the PDF |

## Parts sheet: one row per PDF

| Group | Columns |
|---|---|
| Identity | `source_pdf`, `status` (PASS/FAIL), `state_code`, `state_name`, `roll_year`, `pages` |
| Counts | `parsed_voters`, `cover_total`, `sum_total` (summary page, main-list row), `sumrow_total` (summary total row), `parsed_male`/`cover_male`, `parsed_female`/`cover_female`, `parsed_third`/`cover_third`, `start_sl`, `end_sl` |
| Constituency | `assembly_no`, `assembly_name`, `reservation`, `parliament_no`, `parliament_name`, `part_no` |
| Area | `sections` (`1-name; 2-name; …`), `main_village`, `ward`, `police_station`, `tehsil`, `district`, `pincode` |
| Polling station | `polling_station_type`, `polling_station_no`, `polling_station_name`, `polling_station_address`, `auxiliary_ps` |
| Roll | `revision_year`, `qualifying_date`, `revision_type`, `publication_date`, `roll_identification` |
| Checks (`OK`/`FAIL`) | `count_vs_cover`, `count_vs_summary`, `sl_sequence`, `male_count`, `female_count`, `third_count`, `part_no_footer`, `part_no_filename`, `assembly_no_folder`, `cover_totals`, `summary_vs_cover`, `page_footers`, `crosscheck` |
| Check info | `crosscheck_pages` (`N checked, M skipped`), `sections_from_extra_page` (page number when the full section list is on an extra page) |

## Issues sheet

| Column | Meaning |
|---|---|
| `severity` | `ERROR` (extraction may be wrong → FAIL) or `WARN` (PDF content is unusual, extracted faithfully) |
| `source_pdf` | PDF file (blank for AC-level checks) |
| `sl_no` | voter serial number, if the issue is about one voter |
| `check` | check name, e.g. `crosscheck_field`, `epic_format`, `metadata_blank` |
| `detail` | human-readable detail |

## RUN_SUMMARY.xlsx (sheet Summary): one row per AC

`ac_folder`, `status` (PASS/FAIL), `assembly_no`, `assembly_name`, `pdfs`, `pdfs_pass`,
`pdfs_fail`, `voters`, `cover_total`, `male`, `female`, `third`, `errors`, `warnings`,
`seconds`, `file` (path of the AC workbook).
