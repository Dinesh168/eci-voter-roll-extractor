# ECI 2027 draft-roll PDF layout: Uttar Pradesh (S24, Hindi)

File name: `2027-EROLLGEN-S24-<ac>-SSR-DraftRoll-Revision1-HIN-<part>-WI.pdf`
URL: `https://voters.eci.gov.in/eroll/2027/s24/ssr-draftroll/<ac>/<file name>`
Producer: iText + pdfCalligraph, fonts NotoSans-Regular / NotoSans-SemiBold (subsets,
CID Identity-H). Text-based PDF, no OCR needed. UP has 403 ACs.

Extractor: `extract_voters_up_2027.py` (standalone; same design as the Punjab extractor,
see [PDF_LAYOUT.md](PDF_LAYOUT.md) for the shared page/card structure).

All examples below are synthetic or word fragments, not real voter data.

## Differences from Punjab

| Item | Punjab (S19) | Uttar Pradesh (S24) |
|---|---|---|
| Cover title | `ਵੋਟਰ ਸੂਚੀ 2027 S19 ਪੰਜਾਬ` | `निर्वाचक नामावली 2027 S24 उत्तर प्रदेश` |
| Cover area fields | village, ward, police station, tehsil, district, pincode | `मुख्य कस्बा`, `वार्ड`, `डाकघर` (post office), `थाना`, `पंचायत`, `ब्लॉक`, `तहसील`, `जिला`, `पिन कोड` |
| Parliament | no/name | no/name **with reservation** |
| Section list header | `ਭਾਗ ਵਿੱਚ ਸੈਕਸ਼ਨਾਂ…` | `भाग में अनुभागों की संख्या और नाम` |
| Summary heading | `ਵੋਟਰਾਂ ਦੀ ਸਾਰਣੀ` | `निर्वाचकों का सार` |
| Card labels | `ਨਾਮ :`, `ਪਿਤਾ/ਪਤੀ/ਮਾਤਾ/ਪਤਨੀ ਦਾ ਨਾਮ:`, `ਹੋਰ:`, `ਮਕਾਨ ਨੰ. :`, `ਉਮਰ : … ਲਿੰਗ :` | `नाम :`, `पिता/पति/माता/पत्नी का नाम:`, `अन्य:`, `मकान संख्या :`, `आयु : … लिंग :` |
| Gender values | `ਪੁਰਸ਼ / ਇਸਤਰੀ / ਤੀਜਾ ਲਿੰਗ` | `पुरुष / महिला / तृतीय लिंग` |
| Footer | `ਕੁੱਲ ਪੰਨੇ N - ਪੰਨਾ k` | `कुल पृष्ठ N - पृष्ठ k`, `भाग संख्या : <part>` |
| EPIC formats | `AAA9999999` | `AAA9999999` and the old UP format `UP/12/345/0123456` (≈14% of voters); rarely `UP\12\345\0123456`, lowercase or 11 characters (WARN) |

## Hindi text layer: how the text is read

The PDF contains two text sources:

1. **ActualText** (marked content) – the original string in correct *logical* order
   (`ि` after its consonant, reph `र्` before its cluster). **This is the extracted text.**
2. **Glyph text** – each drawn glyph's ToUnicode mapping, in *visual* order
   (`िप` instead of `पि`). Used only for the **glyph check**: per page, the characters
   of the extracted text must equal the characters of the drawn glyphs
   (order-independent). Any difference is an ERROR.

## Quirks found and how they are handled (403-PDF sample, 2026-11-21)

| # | Quirk | Example (synthetic) | Handling |
|---|---|---|---|
| 1 | Invisible joiners in ActualText | `न्<ZWJ>द` | removed by `norm()` (do not change printed text) |
| 2 | Zero-width **duplicate** after a joiner: text repeats a cluster that is drawn once | `न्<ZWJ>द्रद्र` printed `न्द्र` | removed **only if the drawn glyphs of that line prove it** (repaired line chars == drawn chars) |
| 3 | Doubled signs in ActualText | `दिुुर`, `मेंंद्र` | same proven-repair rule |
| 4 | Repeated piece in ActualText | `प्पापा` printed `प्पा`, `ट्स्स` printed `ट्स` | same proven-repair rule; accepted only if exactly one repair fits |
| 5 | Zero-width run **without** joiner is real text (tail of a multi-letter glyph `न्न`, `प्र`) | `अन्नन`, `प्ररव` (printed so) | never touched |
| 6 | Wrong ToUnicode in the embedded font | `ृ` mapped as `र्`, `ऋ` as `र`, `ड्ड` as `ड`, `द्रृ` as `द्रर्`, `ऋ़` as `ऱ` | `GLYPH_FIXES`, each verified by drawing the glyph outline from the embedded font |
| 7 | `ि` glyph variants that also draw `ं`, `र्` (reph), `र्ं` or `ँ`, all mapped to bare `ि` | glyphs 4354–4364, 4366–4372, 4380/4382, 4485/4486 | `GLYPH_FIXES` (classified by drawing all variants found in the sample) |
| 8 | Empty glyph (no outline, 0 advance) mapped to `ं` | glyph 4389 | mapped to nothing |
| 9 | **Producer rendering quirk**: consonant drawn twice around a reph | text `उर्रह` drawn `उर्ररह` | ActualText kept (the real value); page WARN `render_quirk` |
| 10 | **Producer rendering quirk**: `रृ` (data-entry oddity) drawn as reph + `ऋ` | text `…रृ` | ActualText kept; WARN `render_quirk` |
| 11 | Wrapped name: continuation row overlaps the previous row vertically | — | drawn glyphs are assigned to the nearest line (glyph left edge, then vertical centre) |
| 12 | Devanagari digits in house numbers | `१०७/२` | kept as printed |

Glyph ids are those of the NotoSans subsets embedded by ECI's producer; they were stable
across all 403 sample PDFs. If ECI changes fonts, the glyph check will fail loudly and the
fixes must be re-verified (scratch tools: draw the glyph outline with fontTools
`SVGPathPen` from the PDF's `FontFile2` stream and render it).
