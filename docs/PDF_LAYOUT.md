# ECI 2027 draft-roll PDF layout (Punjab, S19)

> Uttar Pradesh (S24, Hindi) differences and text-layer quirks: [PDF_LAYOUT_UP.md](PDF_LAYOUT_UP.md).

File name: `2027-EROLLGEN-S19-<ac>-SSR-DraftRoll-Revision1-PAN-<part>-WI.pdf`
URL: `https://voters.eci.gov.in/eroll/2027/s19/ssr-draftroll/<ac>/<file name>`
Producer: iText Core 8 + pdfCalligraph. Text-based PDF (real Unicode text layer, font
NirmalaUI), A4 595×842 pt. No OCR needed. Missing parts return HTTP 404 (nginx).

Punjab has 117 ACs with 141–329 parts each (~24,700 PDFs, ~20–30 pages each).

## Page structure

| Page | Content |
|---|---|
| 1 – Cover | Title `ਵੋਟਰ ਸੂਚੀ 2027 S19 ਪੰਜਾਬ`; assembly no–name (reservation); part no; parliament; 1. revision details (year, qualifying date, type, publication date, roll identification); 2. part & area (section list, main village, ward, police station, tehsil, district, pincode); 3. polling station (type, no–name, address, auxiliary PS count); 4. counts (start SL, end SL, male, female, third gender, total) |
| 2 – Map/photos | 6 images (Nazri Naksha, Google map, PS building photos, CAD, key map). No data. |
| 3 … N-1 – Voters | 3 columns × up to 10 rows of voter cards (≤30 per page). Footer: section no–name, qualifying date, publication date, AC no–name, part no, `ਕੁੱਲ ਪੰਨੇ N - ਪੰਨਾ k` |
| N – Summary | `ਵੋਟਰਾਂ ਦੀ ਸਾਰਣੀ`: row I main list (male, female, third, total) + total row |
| N+1 (rare) | Full section list `ਭਾਗ ਵਿੱਚ ਸੈਕਸ਼ਨਾਂ ਦੇ ਨੰ. ਅਤੇ ਨਾਮ` when the cover box overflows |

Each section starts on a new page; the page footer names the section for every card on it.

## Voter card

Drawn as a box (≈186 × 75 pt; vertical border lines are used to find boxes):

```
┌──────────────────────────────────────────┐
│ [ SL ]                        EPIC        │
│ ਨਾਮ : <voter name>               ┌──────┐ │
│ ਪਿਤਾ/ਪਤੀ/ਮਾਤਾ/ਪਤਨੀ ਦਾ ਨਾਮ: <name>│ ਫੋਟੋ  │ │
│   (or  ਹੋਰ: <name>)               │ਉਪਲਬਧ│ │
│ ਮਕਾਨ ਨੰ. : <house no / address>  │  ਹੈ  │ │
│ ਉਮਰ : <age> ਲਿੰਗ : <gender>       └──────┘ │
└──────────────────────────────────────────┘
```

Long values wrap onto the next line inside the text column (name, relative, house no.).
The photo placeholder text sits in the right ~30% of the card, at the same heights as
the name/relative/house lines.

## Variants and quirks found (all handled)

| # | Variant | Example | Handling |
|---|---|---|---|
| 1 | PyMuPDF splits one visual line mid-word | `ਹਰਪ੍ਰੀ` + `ਤ ਕੌਰ` | re-join pieces on the same baseline that touch |
| 2 | Doubled combining marks in text layer | `ਲਿੰੰਗ`, `ਵਿੱੱਚ` | collapsed in `norm()` |
| 3 | Sihari stored before its consonant | `ਿਪਤਾ` | swapped to `ਪਿਤਾ` in `norm()` |
| 4 | Values wrapped onto 2+ lines | `ਜਸਵਿੰਦਰ ਸਿੰਘ ਉਰਫ ਬਲਵੀਰ` / `ਸਿੰਘ` | continuation lines joined with a space, flagged `wrapped` |
| 5 | EPIC in non-standard format | `aBC1234567`, `A1B2345678`, `AbC1234567`, `ABCO123456`, `ABCDO12345` | accepted as card start; kept exactly; WARN `epic_format` |
| 6 | Relation `ਹੋਰ` (Other), `ਮਾਤਾ` (Mother) | `ਹੋਰ: <name>` | supported |
| 7 | Third gender | `ਤੀਜਾ ਲਿੰਗ` | supported |
| 8 | Blank relative name / house no. | `ਪਿਤਾ ਦਾ ਨਾਮ:` | kept blank, flagged |
| 9 | Hindi (Devanagari) names/addresses | `राम कुमार` | kept, flagged |
| 10 | Legacy-font glyphs printed in addresses | `ਵÇਰ੍ਡ`, `ਗÃਸ` (font SolaimanLipi) | kept as printed, flagged |
| 11 | Ward present / blank | `ਵਾਰਡ ਨੰ. 9`, `ਨਗਰ ਨਿਗਮ ਅੰਮ੍ਰਿਤਸਰ ਵਾਰਡ ਨੰ. 14`, `:` | value or blank |
| 12 | Blank cover value followed by next label | polling-station type, pincode | `_after()` returns blank |
| 13 | Section list item wraps / has empty name | `14-…ਵਾਰਡ` + `ਨੰ.10`; `2-` | joined / empty name allowed |
| 14 | Cover section list truncated | AC106 part 12: cover lists 15, pages use 24 | full list read from extra page after summary |
| 15 | Summary not on last page | same PDF | summary located by heading |
| 16 | Card SL and EPIC baselines differ slightly | EPIC 0.3 pt higher | box reader groups top row with 3 pt tolerance |

## Not seen in the 2027 draft roll (would raise ERROR if they appear)

* Supplementary/deletion lists on the summary page (only row I "ਮੁੱਖ ਸੂਚੀ").
* Deleted/shifted stamps on cards, "photo not available" text variants.
* Image-only (scanned) pages.
