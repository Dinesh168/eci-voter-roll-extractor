#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Uttar Pradesh (S24) Draft Electoral Roll 2027 PDF -> Excel, with full validation.

Hindi (Devanagari) counterpart of extract_voters_2027.py (Punjab, Punjabi).
Built for "2027-EROLLGEN-S24-<AC>-SSR-DraftRoll-Revision1-HIN-<part>-WI.pdf"
(text-based PDF, 3 voter cards per row).

USAGE
    python extract_voters_up_2027.py input\\S24 -o output_up -w 8   # folder of AC1..AC403 folders
    python extract_voters_up_2027.py input\\S24\\AC244               # one AC folder

OUTPUT (Excel only, in <output>\\)
    <AC>.xlsx              sheets: Voters, Parts (cover metadata + checks), Issues
    RUN_SUMMARY.xlsx       one row per AC: PASS/FAIL, voters vs cover total, errors
    Resumable: ACs whose <AC>.xlsx exists are skipped (--force redoes them).

HINDI TEXT LAYER (ECI's iText/pdfCalligraph output) - handled
    * Text is read from the PDF's ActualText (correct logical order: ि and reph र्).
    * ActualText sometimes contains text that is NOT drawn: a zero-width copy of a
      cluster after a joiner ("न्‍द्रद्र" printed "न्द्र") or doubled signs ("ुु").
      Such a repair is applied to a line ONLY when the drawn glyphs of that line prove it
      (the repaired text then has exactly the drawn characters); otherwise the text is
      kept and the glyph check fails the page.
    * Invisible joiners U+200D/U+200C (ZWJ/ZWNJ) are removed (they do not change the
      printed text).
    * Some PDFs embed a wrong ToUnicode entry for a glyph (e.g. ृ read as र्, ऋ as र,
      ठ्ठ as ठ, ि-with-anusvara as ि). Entries listed in GLYPH_FIXES were verified against
      the glyph shapes in the embedded fonts and are corrected in memory. (They matter
      for the glyph check; the extracted text itself comes from ActualText.)
    * Known producer rendering quirks (consonant drawn twice around a reph; 'रृ' drawn
      as reph+ऋ): the ActualText value is kept and the page is flagged WARN render_quirk.

VALIDATION (per PDF)  - ERROR makes the PDF/AC FAIL, WARN is informational
    * voters parsed == cover total == summary total; SL contiguous; gender counts (ERROR)
    * CROSS-CHECK 1: text stream vs card-box segmentation, every field identical (ERROR)
    * CROSS-CHECK 2 (glyph check): per page, the characters extracted (after the fixes
      above) == the characters of the glyphs actually drawn (order-independent)  (ERROR)
    * page footers (AC, part, dates, page x/N) == cover; cover/summary totals      (ERROR)
    * every card: EPIC, name, relation, age 18-125, gender; no unparsed text       (ERROR)
    * odd EPIC, blank relative/house no., unusual characters, blank cover field     (WARN)

Exit code 0 = all checks passed, 1 = at least one ERROR (see Issues sheet).
Requires: pip install PyMuPDF xlsxwriter openpyxl
"""

import argparse
import glob
import os
import re
import sys
import time
import unicodedata
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed

import fitz  # PyMuPDF

# ----------------------------------------------------------------------------
# Font ToUnicode corrections (verified by drawing each glyph from the embedded font)
#   font base name -> {glyph id (= CID, Identity-H): correct text}
# ----------------------------------------------------------------------------
GLYPH_FIXES = {
    "NotoSans-Regular": {
        3811: "ृ",      # uni0943; some PDFs map it to र्
        3755: "ऋ",      # uni090B; many PDFs map it to र
        4252: "ड्ड",    # ligature; some PDFs map it to ड
        4244: "ट्ट",    # ligature; some PDFs map it to ट
        4144: "हृ",     # ligature; some PDFs map it to हर्
        4158: "दृ",     # ligature; some PDFs map it to दर्
        4249: "ठ्ठ",    # ligature; many PDFs map it to ठ
        4246: "ट्ठ",    # ligature; some PDFs map it to ठ
        # 4389: EMPTY glyph (no outline in any embedded subset, 0 advance) that the PDFs
        # map to ं; the dot is drawn inside the preceding i-matra glyph (see below).
        4389: "​",  # -> no text (zero-width space, removed by norm)
    },
    "NotoSans-SemiBold": {
        4389: "​",
    },
}
# i-matra (ि) width variants. The PDFs map every one of them to a bare ि, but some also
# draw the anusvara and/or the reph of the syllable. Classified by drawing every variant
# found in the 403 UP sample PDFs (scratch tool imatra_gallery.py, 2026-11-21):
#   plain ि  : 4341-4352            ि + ं  : 4354-4364 (composite glyphs)
#   ि + र्   : 4366-4372            ि + र् + ं : 4380, 4382
# (4393 ीं and 4396/4398 ीर् are already mapped correctly by the PDFs.)
for _g in range(4354, 4365):
    GLYPH_FIXES["NotoSans-Regular"][_g] = "िं"
for _g in range(4366, 4373):
    GLYPH_FIXES["NotoSans-Regular"][_g] = "िर्"
for _g in (4380, 4382):
    GLYPH_FIXES["NotoSans-Regular"][_g] = "िर्ं"
GLYPH_FIXES["NotoSans-SemiBold"].update({4356: "िं", 4358: "िं", 4368: "िर्"})
GLYPH_FIXES["NotoSans-Regular"].update({
    4485: "िँ", 4486: "िँ",   # ि + candrabindu (one glyph); PDFs map it to ि
    3880: "ऋ़",          # ऋ with nukta; PDFs map it to ऱ
    4161: "द्रृ",              # ligature द्र + ृ; PDFs map it to द्रर्
})
_HEX = re.compile(rb"<([0-9A-Fa-f]+)>")


def _utf16hex(s):
    return s.encode("utf-16-be").hex().encode()


def patch_cmaps(doc):
    """Correct known-wrong ToUnicode entries in memory. Returns number of entries fixed."""
    fixed = 0
    for x in range(1, doc.xref_length()):
        try:
            t, v = doc.xref_get_key(x, "ToUnicode")
        except Exception:
            continue
        if t != "xref":
            continue
        base = re.sub(r"^/?[A-Z]{6}\+", "", doc.xref_get_key(x, "BaseFont")[1].lstrip("/"))
        fixes = GLYPH_FIXES.get(base)
        if not fixes:
            continue
        sx = int(v.split()[0])
        data = doc.xref_stream(sx)
        out, in_char, changed = [], False, 0
        for line in data.split(b"\n"):
            s = line.strip()
            if s.endswith(b"beginbfchar"):
                in_char = True
            elif s.startswith(b"endbfchar"):
                in_char = False
            toks = _HEX.findall(line)
            if len(toks) == 3 and toks[0].lower() == toks[1].lower() and b"[" not in line:
                cid = int(toks[0], 16)
                if cid in fixes and toks[2].lower() != _utf16hex(fixes[cid]):
                    line = b"<%s><%s><%s>" % (toks[0], toks[1], _utf16hex(fixes[cid]))
                    changed += 1
            elif len(toks) == 2 and in_char:
                cid = int(toks[0], 16)
                if cid in fixes and toks[1].lower() != _utf16hex(fixes[cid]):
                    line = b"<%s> <%s>" % (toks[0], _utf16hex(fixes[cid]))
                    changed += 1
            out.append(line)
        if changed:
            doc.update_stream(sx, b"\n".join(out))
            fixed += changed
    return fixed


# ----------------------------------------------------------------------------
# Text normalisation
# ----------------------------------------------------------------------------
INVISIBLE = dict.fromkeys(map(ord, "‍‌​﻿"), None)


def norm(s):
    if not s:
        return ""
    s = s.translate(INVISIBLE).replace(" ", " ")
    return re.sub(r"[ \t]+", " ", s).strip()


def dedupe_chars(chars, force=False):
    """chars: [(char, width)]. ActualText duplication artefact: a zero-width run that
    contains a ZERO WIDTH JOINER, whose part after the joiner is repeated immediately by the
    visible glyphs, e.g. 'न्‍[द्र]द्र' or 'क्ष्‍[मी]मी'; the repeated part is
    dropped. Zero-width runs WITHOUT a joiner are the tail of a multi-letter glyph
    (न्न -> 'न' + zero-width '्न', प्र) and are real text: never touched
    ('अन्नन', 'प्ररव' are what the PDF draws). Verified per page by the glyph check."""
    text = "".join(c for c, _ in chars)
    keep, i, n = [], 0, len(chars)
    removed = 0
    while i < n:
        if chars[i][1] < 0.05 and not chars[i][0].isspace():
            j = i
            while j < n and chars[j][1] < 0.05 and not chars[j][0].isspace():
                j += 1
            seg = text[i:j]
            cut = 0
            z = seg.rfind("‍")
            if z >= 0 and j < n:
                tail = seg[z + 1:]
                if force and tail and not chars[j][0].isspace():
                    cut = len(tail)   # candidate only; accepted if the glyphs prove it
                for k in (range(len(tail), 0, -1) if not cut else ()):
                    suf = tail[len(tail) - k:]
                    if unicodedata.category(suf[0]) == "Lo" and text[j:j + k] == suf:
                        cut = k
                        break
            if cut:
                removed += 1
            keep.append(seg[:len(seg) - cut])
            i = j
            continue
        keep.append(chars[i][0])
        i += 1
    return "".join(keep), removed


# ----------------------------------------------------------------------------
# Patterns (applied to normalised text)
# ----------------------------------------------------------------------------
# valid EPIC formats: national AAA9999999, and UP's old state format UP/12/345/0123456
EPIC_RE = re.compile(r"^(?:[A-Z]{3}[0-9]{7}|[A-Z]{2}/\d{1,3}/\d{2,3}/\d{5,7})$")
# card-start detection is looser (kept exactly, unusual ones flagged): also backslash
# variants such as UP\12\345\0123456
EPIC_LOOSE_RE = re.compile(r"^(?=(?:[^0-9]*[0-9]){4})[A-Za-z0-9/\\\-]{6,24}$")
SL_RE = re.compile(r"^\d{1,5}$")
NAME_RE = re.compile(r"^नाम\s*:\s*(.*)$")
REL_RE = re.compile(r"^(पिता|पति|माता|पत्नी|अन्य)(?:\s*का नाम)?\s*:\s*(.*)$")
HOUSE_RE = re.compile(r"^मकान संख्या\s*:\s*(.*)$")
AGE_RE = re.compile(r"^आयु\s*:\s*(\d*)\s*लिंग\s*:\s*(.*)$")
PHOTO_RE = re.compile(r"^फोटो")
SECTION_RE = re.compile(r"^अनुभाग संख्या और नाम\s*:\s*(\d+)\s*-\s*(.*)$")
FOOTER_RE = re.compile(r"कुल पृष्ठ\s*(\d+)\s*-\s*पृष्ठ\s*(\d+)")
PART_FOOT_RE = re.compile(r"^भाग संख्या\s*:\s*:?\s*(\d+)")
AC_FOOT_RE = re.compile(r"निर्वाचक नामावली\s*-\s*\d{4}\s*,\s*विधानसभा निर्वाचन क्षेत्र\s*:\s*:?\s*(\d+)\s*-\s*(.+)")
PAGE_NOISE = ("अर्हता दिनांक", "प्रकाशन की तिथि", "निर्वाचक नामावली -", "निर्वाचक नामावली-",
              "भाग संख्या", "कुल पृष्ठ")
VOTER_PAGE_MARK = "आयु :"

REL_EN = {"पिता": "Father", "पति": "Husband", "माता": "Mother", "पत्नी": "Wife", "अन्य": "Other"}
GENDER_EN = {"पुरुष": "Male", "महिला": "Female", "तृतीय लिंग": "Third Gender"}

COLUMNS = [
    "ac_folder", "source_pdf", "state_code", "state_name", "assembly_no", "assembly_name",
    "reservation", "parliament_no", "parliament_name", "part_no", "district", "tehsil",
    "block", "panchayat", "police_station", "post_office", "main_village", "ward", "pincode",
    "polling_station_no", "polling_station_name", "section_no", "section_name", "page_no",
    "sl_no", "epic", "voter_name", "relation_type", "relation_type_hi", "relative_name",
    "house_no", "age", "gender", "gender_hi", "photo_status", "warnings",
]


def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


# ----------------------------------------------------------------------------
# Line reconstruction
# ----------------------------------------------------------------------------
# a repeated run of 1-2 identical Devanagari signs ('ुु', 'ंं', 'ेंें', '््'): never valid
DOUBLE_MARK_RE = re.compile(r"([ऀ-ःऺ-ॏॕ-ॗॢॣ]{1,2})\1")


def undouble(s):
    return DOUBLE_MARK_RE.sub(lambda m: m.group(1), s)


DOUBLE_ANY_RE = re.compile(r"([\u0900-\u097f][^\s]{0,7})\1")


def collapse_repair(raw, g, depth=3):
    """Remove immediately repeated pieces ('प्पापा' -> 'प्पा', 'ट्स्स' -> 'ट्स')
    until the characters equal the drawn glyph characters g. Returns the repaired text
    only if exactly one distinct result is found (else None: keep raw, page is flagged)."""
    s0 = raw.translate(INVISIBLE)
    have = _chars_key(s0)
    if have - g == Counter() or g - have != Counter():
        return None
    results = set()

    def walk(s, excess, d):
        if not excess:
            results.add(s)
            return
        if d == 0 or len(results) > 1:
            return
        n = len(s)
        for L in range(1, 9):
            for i in range(0, n - 2 * L + 1):
                x = s[i:i + L]
                if x.isspace() or s[i + L:i + 2 * L] != x:
                    continue
                cx = Counter(ch for ch in x if not ch.isspace())
                if cx and not (cx - excess):
                    walk(s[:i + L] + s[i + 2 * L:], excess - cx, d - 1)

    walk(s0, have - g, depth)
    return results.pop() if len(results) == 1 else None


def _chars_key(s):
    return Counter(ch for ch in norm(s) if not ch.isspace())


def _drawn_chars(page):
    """[(x, cy, char)] of the drawn glyphs (ActualText ignored). x is the glyph's LEFT
    edge (+0.1): that is where the line's text places it; a wide last glyph of a line can
    extend up to the start of the next line piece."""
    res = []
    for b in page.get_text("rawdict", flags=fitz.TEXTFLAGS_RAWDICT | fitz.TEXT_IGNORE_ACTUALTEXT)["blocks"]:
        for ln in b.get("lines", []):
            for sp in ln["spans"]:
                for c in sp["chars"]:
                    x0, y0, x1, y1 = c["bbox"]
                    res.append((x0 + 0.1, (y0 + y1) / 2, c["c"]))
    return res


def page_items(page):
    """[(text, x0, y0, x1, y1)] from ActualText (logical order), split lines on the same
    baseline re-joined. Returns (items, repairs).

    The ActualText of these PDFs has two artefacts that duplicate text which is NOT drawn:
    a zero-width duplicate after a joiner ('न्‍द्रद्र', see dedupe_chars) and doubled
    signs ('दिुुर', 'मेंंद्र'). A repair is applied to a line ONLY when the drawn glyphs
    of that line prove it: the line's characters must then equal the drawn glyph characters
    (and must not equal them without the repair). Otherwise the text is kept as is and the
    page-level glyph check reports any difference."""
    out, repairs = [], 0
    lines = [ln for b in page.get_text("rawdict")["blocks"] for ln in b.get("lines", [])]
    owner = None   # drawn glyph chars per line index, computed only when needed

    def drawn_of(idx):
        nonlocal owner
        if owner is None:
            # each drawn glyph goes to the nearest line on its baseline (a line's bbox can
            # end before its last glyph when its trailing ActualText chars are zero-width)
            # (vertically overlapping rows, e.g. a wrapped name: nearest vertical centre)
            owner = [[] for _ in lines]
            rows = {}
            for k, l in enumerate(lines):
                for yb in range(int(l["bbox"][1]), int(l["bbox"][3]) + 1):
                    rows.setdefault(yb, []).append(k)
            for cx, cy, t in _drawn_chars(page):
                best, bd = None, None
                for k in rows.get(int(cy), ()):
                    x0, y0, x1, y1 = lines[k]["bbox"]
                    if not (y0 <= cy <= y1):
                        continue
                    dx = 0.0 if x0 <= cx <= x1 else min(abs(cx - x0), abs(cx - x1))
                    if dx > 3.0:
                        continue
                    score = (dx, abs(cy - (y0 + y1) / 2))
                    if bd is None or score < bd:
                        best, bd = k, score
                if best is not None:
                    owner[best].append(t)
        return _chars_key("".join(owner[idx]))

    for idx, ln in enumerate(lines):
        chars = [(c["c"], c["bbox"][2] - c["bbox"][0]) for s in ln["spans"] for c in s["chars"]]
        raw = "".join(c for c, _ in chars)
        txt = raw
        dd, r = dedupe_chars(chars)
        cands = [dd, undouble(raw), undouble(dd)]
        if any(c != raw for c in cands) or DOUBLE_ANY_RE.search(raw):
            g = drawn_of(idx)
            if _chars_key(raw) != g:
                df = dedupe_chars(chars, force=True)[0]
                cands += [df, undouble(df)]
                found = next((c for c in cands if c != raw and _chars_key(c) == g), None)
                if found is None:
                    found = collapse_repair(raw, g)
                if found is not None:
                    txt = found
                    repairs += 1
        if not txt.strip():
            continue
        x0, y0, x1, y1 = ln["bbox"]
        if out:
            p = out[-1]
            if abs(y0 - p[2]) < 1.5 and abs(y1 - p[4]) < 1.5 and -0.5 <= x0 - p[3] < 4.0:
                p[0] += txt
                p[3] = x1
                continue
        out.append([txt, x0, y0, x1, y1])
    res = []
    for t in out:
        n = norm(t[0])
        if n:
            res.append((n, t[1], t[2], t[3], t[4]))
    return res, repairs


def glyph_chars(page):
    """Characters of the glyphs actually drawn (ActualText ignored), for the glyph check.
    Returns (Counter, drawn text in visual order)."""
    t = norm(page.get_text("text", flags=fitz.TEXTFLAGS_TEXT | fitz.TEXT_IGNORE_ACTUALTEXT))
    return Counter(ch for ch in t if not ch.isspace()), t


# Rendering quirks of the PDF producer, proven by rendering the pages (2026-11-21).
# The hidden text (ActualText) holds the real value; only the drawing differs:
#  * consonant drawn twice around a reph: text 'उर्रह' / 'र्मू' / 'र्वी' is drawn
#    'उर्ररह' / 'र्ममू' / 'र्ववी' (drawn glyph order: C, reph, C).
#  * text 'रृ' (र + vocalic-r sign, a data-entry oddity) is drawn as reph + ऋ.
REPH_DUP_RE = re.compile(r"([\u0915-\u0939])\u093c?\u0930\u094d\1")


def explain_render_quirks(a, g, drawn_text, page_text):
    """True if the whole difference between extracted chars a and drawn chars g is
    explained by the rendering quirks above (each occurrence accounted for exactly)."""
    exp_a, exp_g = Counter(), Counter()
    for m in REPH_DUP_RE.finditer(drawn_text):
        exp_g[m.group(1)] += 1
    n = page_text.count("\u0930\u0943")
    if n:
        exp_a["\u0943"] += n
        exp_g["\u090b"] += n
        exp_g["\u094d"] += n
    if not exp_a and not exp_g:
        return False
    extra, missing = a - g, g - a
    # every predicted occurrence must be present and nothing else may differ
    return extra == exp_a and missing == exp_g


# ----------------------------------------------------------------------------
# Cover page and summary page
# ----------------------------------------------------------------------------
COVER_LABELS = (
    "पुनरीक्षण का वर्ष", "अर्हता तिथि", "पुनरीक्षण का प्रकार", "प्रकाशन की तिथि",
    "निर्वाचक नामावली की पहचान", "1. ", "2. ", "3. ", "4. ", "भाग में अनुभागों",
    "मुख्य कस्बा", "वार्ड", "डाकघर", "थाना", "पंचायत", "ब्लॉक", "तहसील", "जिला", "पिन कोड",
    "मतदान स्थल", "इस भाग में सहायक", "संख्या :", "कुल मतदाता", "प्रारम्भिक क्रम", "अंतिम क्रम",
)


def _is_label(s):
    return any(s.startswith(x) for x in COVER_LABELS)


def _idx(lines, label):
    return next((k for k, l in enumerate(lines) if l.startswith(label)), None)


def _after(lines, label):
    """Value after a label: on the same line after ':' or on the next line (leading ':'
    stripped). Blank if the next line is ':' only or another label."""
    i = _idx(lines, label)
    if i is None:
        return ""
    rest = lines[i][len(label):].strip()
    if rest.startswith(":") and rest.lstrip(":").strip():
        return rest.lstrip(":").strip()
    if i + 1 < len(lines):
        v = lines[i + 1]
        if v.startswith(":"):
            return v[1:].strip()
        return "" if _is_label(v) else v.strip()
    return ""


def _block(lines, label, stop):
    """All lines after a label line up to (not incl.) the first line starting with stop."""
    i = _idx(lines, label)
    if i is None:
        return ""
    out = []
    for l in lines[i + 1:]:
        if any(l.startswith(s) for s in stop):
            break
        out.append(l)
    return " ".join(out).strip()


SECTION_LIST_HDR = "भाग में अनुभागों की संख्या और नाम"
SUMMARY_HDR = "निर्वाचकों का सार"


def parse_section_list(lines):
    sections = {}
    i = _idx(lines, SECTION_LIST_HDR)
    if i is None:
        return sections
    last = None
    for l in lines[i + 1:]:
        if l.startswith("मुख्य कस्बा") or l.startswith(PAGE_NOISE):
            break
        mm = re.match(r"^(\d+)\s*-\s*(.*)$", l)
        if mm:
            last = int(mm.group(1))
            sections[last] = mm.group(2).strip()
        elif last is not None:
            sections[last] = (sections[last] + " " + l).strip()
    return sections


def parse_cover(lines):
    c = {}
    txt = "\n".join(lines)
    m = re.search(r"^निर्वाचक नामावली\s+(\d{4})\s+(S\d+)\s+(.+)$", txt, re.M)
    c["roll_year"], c["state_code"], c["state_name"] = (
        (int(m.group(1)), m.group(2), m.group(3).strip()) if m else (None, "", ""))
    m = re.search(r"विधानसभा निर्वाचन क्षेत्र की संख्या व नाम और आरक्षण स्थिति\s*:\s*(\d+)\s*-\s*"
                  r"([^\n(]+?)\s*(\([^)]*\))?\s*$", txt, re.M)
    if m:
        c["assembly_no"], c["assembly_name"] = int(m.group(1)), m.group(2).strip()
        c["reservation"] = (m.group(3) or "").strip("() ")
    m = re.search(r"स्थित है\s*:\s*(\d+)\s*-\s*([^\n(]+?)\s*(\([^)]*\))?\s*$", txt, re.M)
    if m:
        c["parliament_no"], c["parliament_name"] = int(m.group(1)), m.group(2).strip()
        c["parliament_reservation"] = (m.group(3) or "").strip("() ")
    m = re.search(r"^भाग संख्या\s*:\s*:?\s*(\d+)", txt, re.M)
    c["part_no"] = int(m.group(1)) if m else None
    c["revision_year"] = _after(lines, "पुनरीक्षण का वर्ष")
    c["qualifying_date"] = _after(lines, "अर्हता तिथि")
    c["revision_type"] = _block(lines, "पुनरीक्षण का प्रकार", COVER_LABELS)
    c["publication_date"] = _after(lines, "प्रकाशन की तिथि")
    c["roll_identification"] = _block(lines, "निर्वाचक नामावली की पहचान", COVER_LABELS)
    c["sections"] = parse_section_list(lines)
    c["main_village"] = _after(lines, "मुख्य कस्बा")
    for key, label in (("ward", "वार्ड"), ("post_office", "डाकघर"), ("police_station", "थाना"),
                       ("panchayat", "पंचायत"), ("block", "ब्लॉक"), ("tehsil", "तहसील"),
                       ("district", "जिला"), ("pincode", "पिन कोड")):
        c[key] = _after(lines, label)
    c["polling_station_type"] = _after(lines, "मतदान स्थल के प्रकार")
    ps = _block(lines, "मतदान स्थल की संख्या और नाम", ("मतदान स्थल का पता",))
    mm = re.match(r"^(\d+)\s*-\s*(.*)$", ps)
    c["polling_station_no"] = int(mm.group(1)) if mm else None
    c["polling_station_name"] = mm.group(2).strip() if mm else ps
    c["polling_station_address"] = _block(lines, "मतदान स्थल का पता", ("इस भाग में सहायक",))
    i = _idx(lines, "इस भाग में सहायक")
    aux = next((l for l in lines[i + 1:i + 4] if re.fullmatch(r"\d+", l)), None) if i is not None else None
    c["auxiliary_ps"] = int(aux) if aux else None
    nums = []
    i = next((k for k, l in enumerate(lines) if l == "कुल"), None)
    if i is not None:
        for l in lines[i + 1:]:
            if re.fullmatch(r"\d+", l):
                nums.append(int(l))
                if len(nums) == 6:
                    break
            elif nums:
                break
    for k, v in zip(("start_sl", "end_sl", "cover_male", "cover_female", "cover_third", "cover_total"),
                    nums + [None] * 6):
        c[k] = v
    return c


def parse_summary(lines):
    nums = []
    i = next((k for k, l in enumerate(lines) if l == "कुल"), None)
    if i is not None:
        for l in lines[i + 1:]:
            if re.fullmatch(r"\d+", l):
                nums.append(int(l))
                if len(nums) == 8:
                    break
    nums += [None] * 8
    keys = ("sum_male", "sum_female", "sum_third", "sum_total",
            "sumrow_male", "sumrow_female", "sumrow_third", "sumrow_total")
    return dict(zip(keys, nums[:8]))


# ----------------------------------------------------------------------------
# Voter pages
# ----------------------------------------------------------------------------
def is_card_start(lines, i):
    if not SL_RE.match(lines[i]):
        return False
    nxt = lines[i + 1] if i + 1 < len(lines) else ""
    return bool(EPIC_LOOSE_RE.match(nxt) or PHOTO_RE.match(nxt) or NAME_RE.match(nxt))


def split_page(lines):
    info = {"section_no": None, "section_name": None, "footer_total": None,
            "footer_page": None, "part_no": None}
    body = []
    for l in lines:
        m = SECTION_RE.match(l)
        if m:
            info["section_no"], info["section_name"] = int(m.group(1)), m.group(2).strip()
            continue
        m = FOOTER_RE.search(l)
        if m:
            info["footer_total"], info["footer_page"] = int(m.group(1)), int(m.group(2))
            continue
        m = PART_FOOT_RE.match(l)
        if m:
            info["part_no"] = int(m.group(1))
            continue
        if l.startswith(PAGE_NOISE):
            continue
        body.append(l)
    return body, info


def parse_voter_page(lines, page_no):
    """Method A: one text stream cut at each card start."""
    body, info = split_page(lines)
    starts = [i for i in range(len(body)) if is_card_start(body, i)]
    cards = [parse_card(body[s:e], page_no) for s, e in zip(starts, starts[1:] + [len(body)])]
    return cards, info, (body[:starts[0]] if starts else body)


def card_boxes(page):
    bands = {}
    for g in page.get_drawings():
        r = g["rect"]
        if r.width < 2.5 and 55 < r.height < 140:
            bands.setdefault((round(r.y0), round(r.y1)), set()).add(round((r.x0 + r.x1) / 2, 1))
    boxes = []
    for (y0, y1), xs in sorted(bands.items()):
        merged = []
        for x in sorted(xs):
            if not merged or x - merged[-1] > 2:
                merged.append(x)
        for a, b in zip(merged[0::2], merged[1::2]):
            if b - a > 100:
                boxes.append(fitz.Rect(a, y0, b, y1))
    return boxes


def box_cards(page, items, page_no):
    """Method B: text lines assigned to drawn card boxes by position."""
    boxes = card_boxes(page)
    buckets = [[] for _ in boxes]
    for t, x0, y0, x1, y1 in items:
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        for k, r in enumerate(boxes):
            if r.x0 <= cx <= r.x1 and r.y0 <= cy <= r.y1:
                buckets[k].append((y0, x0, t))
                break
    out = []
    for r, b in zip(boxes, buckets):
        b = [z for z in b if not z[2].startswith(PAGE_NOISE)]
        if not b:
            continue
        b.sort(key=lambda z: (z[0], z[1]))
        top_y = b[0][0]
        top = sorted([z for z in b if z[0] - top_y < 3], key=lambda q: q[1])
        rest = [z for z in b if z[0] - top_y >= 3]
        split_x = r.x0 + 0.70 * r.width
        photo = [z for z in rest if z[1] >= split_x]
        text = [z for z in rest if z[1] < split_x]
        lines = [t for _, _, t in top] + [t for _, _, t in photo] + [t for _, _, t in text]
        if lines and SL_RE.match(lines[0]):
            out.append(parse_card(lines, page_no))
    return out


def parse_card(seg, page_no):
    cur = {"sl_no": int(seg[0]), "epic": "", "photo_status": "", "voter_name": "",
           "relation_type_hi": "", "relative_name": "", "house_no": "",
           "age": None, "gender_hi": "", "page_no": page_no, "_extra": []}
    last = None
    for l in seg[1:]:
        if not cur["epic"] and not cur["voter_name"] and EPIC_LOOSE_RE.match(l):
            cur["epic"] = l
            last = None
        elif PHOTO_RE.match(l):
            cur["photo_status"] = l
            last = "photo_status"
        elif (m := NAME_RE.match(l)):
            cur["voter_name"] = m.group(1).strip()
            last = "voter_name"
        elif (m := REL_RE.match(l)):
            cur["relation_type_hi"], cur["relative_name"] = m.group(1), m.group(2).strip()
            last = "relative_name"
        elif (m := HOUSE_RE.match(l)):
            cur["house_no"] = m.group(1).strip()
            last = "house_no"
        elif (m := AGE_RE.match(l)):
            cur["age"] = int(m.group(1)) if m.group(1) else None
            cur["gender_hi"] = m.group(2).strip()
            last = None
        elif last:
            cur[last] = (cur[last] + " " + l).strip()
            if last != "photo_status":
                cur["_extra"].append(f"wrapped {last}")
        else:
            cur["_extra"].append(f"unparsed line: {l}")
    return cur


CROSS_FIELDS = ("epic", "voter_name", "relation_type_hi", "relative_name", "house_no", "age", "gender_hi")


# ----------------------------------------------------------------------------
# One PDF
# ----------------------------------------------------------------------------
def process_pdf(path, ac_folder):
    name = os.path.basename(path)
    issues = []

    def issue(sev, check, detail, sl=None):
        issues.append({"severity": sev, "source_pdf": name, "sl_no": sl, "check": check, "detail": detail})

    doc = fitz.open(path)
    cmap_fixes = patch_cmaps(doc)
    items, dup_removed = [], 0
    for p in doc:
        it, r = page_items(p)
        items.append(it)
        dup_removed += r
    pages = [[t[0] for t in it] for it in items]
    cover = parse_cover(pages[0])
    sum_idx = next((k for k in range(len(pages) - 1, 0, -1)
                    if any(l.startswith(SUMMARY_HDR) for l in pages[k])), len(pages) - 1)
    summary = parse_summary(pages[sum_idx])
    for k in range(1, len(pages)):
        if k != sum_idx and any(l.startswith(SECTION_LIST_HDR) for l in pages[k]) \
                and not any(l.startswith(VOTER_PAGE_MARK) for l in pages[k]):
            extra = parse_section_list(pages[k])
            if extra:
                cover["sections"].update(extra)
                cover["sections_from_extra_page"] = k + 1

    # ---- glyph check on every page: extracted characters == drawn glyph characters
    glyph_bad = []
    for k, p in enumerate(doc):
        a = Counter(ch for l in pages[k] for ch in l if not ch.isspace())
        g, drawn_text = glyph_chars(p)
        if a != g:
            extra, missing = a - g, g - a
            if explain_render_quirks(a, g, drawn_text, "\n".join(pages[k])):
                issue("WARN", "render_quirk",
                      f"page {k+1}: drawing differs from the text only by known producer "
                      f"quirks (consonant doubled after reph / 'रृ' drawn as reph+ऋ): "
                      f"text-only {dict(extra)} drawn-only {dict(missing)}")
                continue
            glyph_bad.append(k + 1)
            issue("ERROR", "glyph_check",
                  f"page {k+1}: extracted-not-drawn {dict(extra.most_common(6))} "
                  f"drawn-not-extracted {dict(missing.most_common(6))}")

    voters, seen_sections, footer_parts = [], set(), set()
    cross_pages = cross_skipped = cross_mismatch = 0
    for pno in range(1, doc.page_count):
        lines = pages[pno]
        if not any(l.startswith(VOTER_PAGE_MARK) for l in lines):
            continue
        cards, info, pre = parse_voter_page(lines, pno + 1)
        for l in pre:
            issue("ERROR", "text_outside_cards", f"page {pno+1}: {l}")
        sec_no, sec_name = info["section_no"], info["section_name"]
        if info["part_no"]:
            footer_parts.add(info["part_no"])
        if sec_no is None:
            issue("ERROR", "section_missing", f"page {pno+1} has no section header")
        else:
            seen_sections.add(sec_no)
            if cover["sections"] and sec_no not in cover["sections"]:
                issue("WARN", "section_not_on_cover", f"page {pno+1} section {sec_no} not in cover list")
        bcards = box_cards(doc[pno], items[pno], pno + 1)
        if not bcards:
            cross_skipped += 1
            issue("WARN", "crosscheck_skipped", f"page {pno+1}: no card boxes found")
        else:
            cross_pages += 1
            amap = {c["sl_no"]: c for c in cards}
            bmap = {c["sl_no"]: c for c in bcards}
            if set(amap) != set(bmap):
                cross_mismatch += 1
                issue("ERROR", "crosscheck_cards", f"page {pno+1}: stream SL {sorted(set(amap) - set(bmap))} "
                      f"vs box SL {sorted(set(bmap) - set(amap))}")
            for sl in set(amap) & set(bmap):
                for fld in CROSS_FIELDS:
                    if amap[sl][fld] != bmap[sl][fld]:
                        cross_mismatch += 1
                        issue("ERROR", "crosscheck_field",
                              f"page {pno+1} {fld}: stream '{amap[sl][fld]}' vs box '{bmap[sl][fld]}'", sl)
        for c in cards:
            c["section_no"], c["section_name"] = sec_no, sec_name
            voters.append(c)
    n_pages = doc.page_count
    doc.close()

    # ---- per-card checks
    gcount = Counter()
    for v in voters:
        sl = v["sl_no"]
        w = list(v.pop("_extra"))
        v["relation_type"] = REL_EN.get(v["relation_type_hi"], "")
        v["gender"] = GENDER_EN.get(v["gender_hi"], "")
        gcount[v["gender"] or "?"] += 1
        if not v["epic"]:
            issue("ERROR", "epic_missing", "no EPIC number", sl)
        elif not EPIC_RE.match(v["epic"]):
            issue("WARN", "epic_format", f"unusual EPIC '{v['epic']}'", sl)
        if not v["voter_name"]:
            issue("ERROR", "name_missing", "voter name empty", sl)
        if not v["relative_name"]:
            issue("WARN", "relative_missing", "relative name empty in PDF", sl)
            w.append("relative_name empty in PDF")
        if not v["relation_type"]:
            issue("ERROR", "relation_type", f"unknown relation '{v['relation_type_hi']}'", sl)
        if v["age"] is None or not (18 <= v["age"] <= 125):
            issue("ERROR", "age", f"age invalid: {v['age']}", sl)
        if not v["gender"]:
            issue("ERROR", "gender", f"unknown gender '{v['gender_hi']}'", sl)
        if not v["house_no"]:
            w.append("house_no empty")
        for fld in ("voter_name", "relative_name"):
            if re.search(r"[A-Za-z]", v[fld]):
                w.append(f"latin letters in {fld}")
            if re.search(r"[0-9०-९]", v[fld]):
                w.append(f"digit in {fld}")
        for fld in ("voter_name", "relative_name", "house_no"):
            odd = sorted({ch for ch in v[fld] if not ("ऀ" <= ch <= "ॿ" or ch.isascii())})
            if odd:
                w.append(f"unusual character {''.join(odd)} in {fld} (as printed)")
        for x in w:
            if x.startswith("unparsed"):
                issue("ERROR", "unparsed", x, sl)
        v["warnings"] = "; ".join(w)

    # ---- per-PDF checks
    n = len(voters)
    sls = [v["sl_no"] for v in voters]
    checks = {}

    def chk(key, ok, detail):
        checks[key] = "OK" if ok else "FAIL"
        if not ok:
            issue("ERROR", key, detail)

    tot = cover.get("cover_total")
    chk("count_vs_cover", n == tot, f"parsed {n} voters, cover says {tot}")
    chk("count_vs_summary", n == summary["sum_total"], f"parsed {n}, summary says {summary['sum_total']}")
    s0, s1 = cover.get("start_sl"), cover.get("end_sl")
    exp_sl = list(range(s0, s1 + 1)) if s0 and s1 else []
    chk("sl_sequence", sls == exp_sl, f"SL not contiguous {s0}-{s1}; missing="
        f"{sorted(set(exp_sl) - set(sls))[:20]} dup={sorted(k for k, c in Counter(sls).items() if c > 1)[:20]}")
    chk("male_count", gcount["Male"] == cover.get("cover_male"), f"male {gcount['Male']} vs {cover.get('cover_male')}")
    chk("female_count", gcount["Female"] == cover.get("cover_female"),
        f"female {gcount['Female']} vs {cover.get('cover_female')}")
    chk("third_count", gcount["Third Gender"] == cover.get("cover_third"),
        f"third {gcount['Third Gender']} vs {cover.get('cover_third')}")
    chk("part_no_footer", footer_parts <= {cover.get("part_no")}, f"footer parts {footer_parts} vs {cover.get('part_no')}")
    chk("crosscheck", cross_mismatch == 0, f"{cross_mismatch} stream vs card-box differences")
    chk("glyph_check", not glyph_bad, f"pages {glyph_bad[:15]}: text differs from drawn glyphs")
    m = re.search(r"HIN-(\d+)", name) or re.fullmatch(r"(\d+)\.pdf", name, re.I)
    if m and cover.get("part_no") is not None:
        chk("part_no_filename", int(m.group(1)) == cover["part_no"], f"file part {m.group(1)} vs {cover['part_no']}")
    m = re.fullmatch(r"AC0*(\d+)", ac_folder, re.I)
    if m and cover.get("assembly_no") is not None:
        chk("assembly_no_folder", int(m.group(1)) == cover["assembly_no"],
            f"folder {ac_folder} vs cover assembly {cover['assembly_no']}")
    chk("cover_totals", tot is not None and
        (cover.get("cover_male") or 0) + (cover.get("cover_female") or 0) + (cover.get("cover_third") or 0) == tot
        and s0 is not None and s1 is not None and s1 - s0 + 1 == tot,
        f"cover counts inconsistent (total {tot})")
    chk("summary_vs_cover",
        (summary["sum_male"], summary["sum_female"], summary["sum_third"], summary["sum_total"]) ==
        (cover.get("cover_male"), cover.get("cover_female"), cover.get("cover_third"), tot) and
        (summary["sumrow_male"], summary["sumrow_female"], summary["sumrow_third"], summary["sumrow_total"]) ==
        (summary["sum_male"], summary["sum_female"], summary["sum_third"], summary["sum_total"]),
        f"summary {summary} vs cover")
    bad_foot = []
    for k in range(1, len(pages)):
        t = "\n".join(pages[k])
        fm = AC_FOOT_RE.search(t)
        pm = re.search(r"प्रकाशन की तिथि\s*:-\s*(\S+)", t)
        qm = re.search(r"अर्हता दिनांक\s*:\s*(\d\d-\d\d-\d{4})", t)
        gm = FOOTER_RE.search(t)
        if not (fm and int(fm.group(1)) == cover.get("assembly_no")
                and fm.group(2).strip() == cover.get("assembly_name")
                and pm and pm.group(1) == cover.get("publication_date")
                and qm and qm.group(1) == cover.get("qualifying_date")
                and gm and int(gm.group(1)) == n_pages and int(gm.group(2)) == k + 1):
            bad_foot.append(k + 1)
    chk("page_footers", not bad_foot, f"pages {bad_foot[:15]}: footer differs from cover")
    for fld in ("assembly_no", "assembly_name", "part_no", "publication_date", "qualifying_date", "state_code"):
        if cover.get(fld) in (None, ""):
            issue("ERROR", "metadata_missing", f"cover field '{fld}' not found")
    for fld in ("district", "tehsil", "police_station", "main_village", "pincode", "post_office",
                "block", "polling_station_name", "polling_station_address", "polling_station_type"):
        if cover.get(fld) in (None, ""):
            issue("WARN", "metadata_blank", f"cover field '{fld}' is blank in the PDF")
    if cover.get("pincode") and not re.fullmatch(r"\d{6}", cover["pincode"]):
        issue("WARN", "pincode_format", f"pincode '{cover['pincode']}'")
    if cover["sections"] and not set(cover["sections"]) <= seen_sections:
        issue("WARN", "sections_seen", f"cover sections {sorted(set(cover['sections']) - seen_sections)} without pages")

    meta = {k: cover.get(k) for k in ("state_code", "state_name", "assembly_no", "assembly_name", "reservation",
                                      "parliament_no", "parliament_name", "part_no", "district", "tehsil",
                                      "block", "panchayat", "police_station", "post_office", "main_village",
                                      "ward", "pincode", "polling_station_no", "polling_station_name")}
    for v in voters:
        v.update(meta)
        v["ac_folder"], v["source_pdf"] = ac_folder, name

    part_row = {"source_pdf": name, "pages": n_pages, "parsed_voters": n, "parsed_male": gcount["Male"],
                "parsed_female": gcount["Female"], "parsed_third": gcount["Third Gender"],
                "cmap_fixes": cmap_fixes, "actualtext_repairs": dup_removed}
    part_row.update({k: v for k, v in cover.items() if k != "sections"})
    part_row["sections"] = "; ".join(f"{k}-{v}" for k, v in sorted(cover["sections"].items()))
    part_row.update(summary)
    part_row.update(checks)
    part_row["crosscheck_pages"] = f"{cross_pages} checked, {cross_skipped} skipped"
    part_row["status"] = "PASS" if all(v == "OK" for v in checks.values()) and \
        not any(i["severity"] == "ERROR" for i in issues) else "FAIL"
    return voters, part_row, issues


# ----------------------------------------------------------------------------
# Excel output
# ----------------------------------------------------------------------------
PART_COLS = [
    "source_pdf", "status", "state_code", "state_name", "roll_year", "pages",
    "parsed_voters", "cover_total", "sum_total", "sumrow_total",
    "parsed_male", "cover_male", "parsed_female", "cover_female", "parsed_third", "cover_third",
    "assembly_no", "assembly_name", "reservation", "parliament_no", "parliament_name",
    "parliament_reservation", "part_no", "start_sl", "end_sl", "sections", "main_village", "ward",
    "post_office", "police_station", "panchayat", "block", "tehsil", "district", "pincode",
    "polling_station_type", "polling_station_no", "polling_station_name", "polling_station_address",
    "auxiliary_ps", "revision_year", "qualifying_date", "revision_type", "publication_date",
    "roll_identification", "count_vs_cover", "count_vs_summary", "sl_sequence", "male_count",
    "female_count", "third_count", "part_no_footer", "part_no_filename", "assembly_no_folder",
    "cover_totals", "summary_vs_cover", "page_footers", "crosscheck", "glyph_check", "crosscheck_pages",
    "cmap_fixes", "actualtext_repairs", "sections_from_extra_page",
]
ISSUE_COLS = ["severity", "source_pdf", "sl_no", "check", "detail"]
SUMMARY_COLS = ["ac_folder", "status", "assembly_no", "assembly_name", "pdfs", "pdfs_pass", "pdfs_fail",
                "voters", "cover_total", "male", "female", "third", "errors", "warnings", "seconds", "file"]


def write_xlsx(path, sheets):
    import xlsxwriter
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp.xlsx"
    wb = xlsxwriter.Workbook(tmp, {"constant_memory": True, "strings_to_urls": False,
                                   "strings_to_formulas": False, "strings_to_numbers": False})
    hdr = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
    for sname, cols, rows in sheets:
        ws = wb.add_worksheet(sname)
        widths = [len(c) for c in cols]
        for r in rows[:300]:
            for i, c in enumerate(cols):
                v = r.get(c)
                widths[i] = max(widths[i], len(str(v)) if v is not None else 0)
        for i, w in enumerate(widths):
            ws.set_column(i, i, min(max(8, w + 2), 50))
        ws.freeze_panes(1, 0)
        ws.write_row(0, 0, cols, hdr)
        for ri, r in enumerate(rows, 1):
            ws.write_row(ri, 0, ["" if r.get(c) is None else r.get(c) for c in cols])
        if rows:
            ws.autofilter(0, 0, len(rows), len(cols) - 1)
    wb.close()
    for _ in range(5):
        try:
            os.replace(tmp, path)
            return path
        except PermissionError:
            time.sleep(2)
    alt = os.path.splitext(path)[0] + time.strftime("_%Y%m%d_%H%M%S") + ".xlsx"
    os.replace(tmp, alt)
    print(f"[warn] {path} is open/locked (close it in Excel); saved as {alt}", flush=True)
    return alt


def _safe_pdf(path, ac):
    try:
        return process_pdf(path, ac)
    except Exception as e:
        name = os.path.basename(path)
        return [], {"source_pdf": name, "status": "FAIL"}, [
            {"severity": "ERROR", "source_pdf": name, "sl_no": None, "check": "exception", "detail": repr(e)}]


def list_pdfs(ac_dir):
    pdfs = sorted(glob.glob(os.path.join(ac_dir, "*.pdf")), key=natural_key)
    return sorted(pdfs, key=lambda p: int(m.group(1)) if (m := re.search(r"HIN-(\d+)", p)) else 0)


def process_ac(ac_dir, out_root, pdf_workers=1):
    t0 = time.time()
    ac = os.path.basename(os.path.normpath(ac_dir))
    pdfs = list_pdfs(ac_dir)
    out_path = os.path.join(out_root, f"{ac}.xlsx")
    if pdf_workers > 1 and len(pdfs) > 1:
        with ProcessPoolExecutor(pdf_workers) as ex:
            results = list(ex.map(_safe_pdf, pdfs, [ac] * len(pdfs), chunksize=4))
    else:
        results = [_safe_pdf(p, ac) for p in pdfs]
    all_v, parts, issues = [], [], []
    for v, row, iss in results:
        all_v += v
        parts.append(row)
        issues += iss
    epics = Counter(v["epic"] for v in all_v if v["epic"])
    for e, c in epics.items():
        if c > 1:
            where = [f"{v['source_pdf']}#{v['sl_no']}" for v in all_v if v["epic"] == e]
            issues.append({"severity": "WARN", "source_pdf": "", "sl_no": None, "check": "duplicate_epic",
                           "detail": f"{e} x{c}: {', '.join(where)}"})
    acs = sorted({r.get("assembly_no") for r in parts if r.get("assembly_no") is not None})
    if len(acs) > 1:
        issues.append({"severity": "ERROR", "source_pdf": "", "sl_no": None, "check": "assembly_mixed",
                       "detail": f"assembly numbers {acs}"})
    out_path = write_xlsx(out_path, [("Voters", COLUMNS, all_v), ("Parts", PART_COLS, parts),
                                     ("Issues", ISSUE_COLS, issues)])
    g = Counter(v["gender"] for v in all_v)
    n_err = sum(i["severity"] == "ERROR" for i in issues)
    return {"ac_folder": ac, "status": "PASS" if n_err == 0 else "FAIL",
            "assembly_no": acs[0] if len(acs) == 1 else ",".join(map(str, acs)),
            "assembly_name": next((r.get("assembly_name") for r in parts if r.get("assembly_name")), ""),
            "pdfs": len(pdfs), "pdfs_pass": sum(r.get("status") == "PASS" for r in parts),
            "pdfs_fail": sum(r.get("status") != "PASS" for r in parts), "voters": len(all_v),
            "cover_total": sum(r.get("cover_total") or 0 for r in parts),
            "male": g["Male"], "female": g["Female"], "third": g["Third Gender"], "errors": n_err,
            "warnings": sum(i["severity"] == "WARN" for i in issues),
            "seconds": round(time.time() - t0, 1), "file": out_path}


def load_previous_summary(path):
    if not os.path.exists(path):
        return {}
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        try:
            rows = list(wb["Summary"].iter_rows(values_only=True))
        finally:
            wb.close()
        return {r[0]: dict(zip(rows[0], r)) for r in rows[1:] if r and r[0]}
    except Exception:
        return {}


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", nargs="?", default=os.path.join(here, "input", "S24"))
    ap.add_argument("-o", "--output", default=os.path.join(here, "output_up"))
    ap.add_argument("-w", "--workers", type=int, default=max(1, min(4, os.cpu_count() or 1)))
    ap.add_argument("--force", action="store_true", help="redo ACs whose Excel already exists")
    a = ap.parse_args()
    try:
        import xlsxwriter  # noqa: F401
    except ImportError:
        sys.exit("Missing package: run   pip install xlsxwriter")
    inp = a.input.strip().strip('"').rstrip(">")
    if glob.glob(os.path.join(inp, "*.pdf")):
        ac_dirs = [inp]
    else:
        ac_dirs = sorted((d for d in glob.glob(os.path.join(inp, "*"))
                          if os.path.isdir(d) and glob.glob(os.path.join(d, "*.pdf"))), key=natural_key)
    if not ac_dirs:
        sys.exit(f"No PDFs found in {inp}")
    os.makedirs(a.output, exist_ok=True)
    for f in glob.glob(os.path.join(a.output, "*.tmp.xlsx")):
        os.remove(f)
    sum_path = os.path.join(a.output, "RUN_SUMMARY.xlsx")
    summary = load_previous_summary(sum_path)
    todo = [d for d in ac_dirs if a.force or not os.path.exists(
        os.path.join(a.output, os.path.basename(os.path.normpath(d)) + ".xlsx"))]
    print(f"{len(ac_dirs)} AC folders, {len(ac_dirs) - len(todo)} already done (skipped), "
          f"{len(todo)} to process with {a.workers} workers -> {os.path.abspath(a.output)}")
    t0 = time.time()

    def record(res, k):
        summary[res["ac_folder"]] = res
        print(f"[{k:>3}/{len(todo)}] {res['ac_folder']:<12} {res['status']}  {res['voters']:>7,} voters "
              f"(cover {res['cover_total']:,})  PDFs {res['pdfs_pass']}/{res['pdfs']} pass  err {res['errors']}  "
              f"{res['seconds']}s   elapsed {(time.time() - t0) / 60:.1f} min", flush=True)
        write_xlsx(sum_path, [("Summary", SUMMARY_COLS, [summary[x] for x in sorted(summary, key=natural_key)])])

    def failed(d, e):
        return {"ac_folder": os.path.basename(os.path.normpath(d)), "status": "FAIL", "voters": 0,
                "cover_total": 0, "pdfs": 0, "pdfs_pass": 0, "errors": 1, "seconds": 0, "file": f"crashed: {e!r}"}

    try:
        if len(todo) == 1:
            try:
                record(process_ac(todo[0], a.output, a.workers), 1)
            except Exception as e:
                record(failed(todo[0], e), 1)
        elif todo:
            with ProcessPoolExecutor(a.workers) as ex:
                futs = {ex.submit(process_ac, d, a.output, 1): d for d in todo}
                for k, f in enumerate(as_completed(futs), 1):
                    try:
                        record(f.result(), k)
                    except Exception as e:
                        record(failed(futs[f], e), k)
    except KeyboardInterrupt:
        print("\nInterrupted. Finished ACs are saved; re-run the same command to continue.")
        sys.exit(1)
    fails = [s for s in summary.values() if s.get("status") != "PASS"]
    tot = sum(int(s.get("voters") or 0) for s in summary.values())
    print(f"\nDone in {(time.time() - t0) / 60:.1f} min.  {len(summary)} ACs, {tot:,} voters, "
          f"{len(fails)} AC(s) with errors.")
    print(f"Summary: {sum_path}")
    if fails:
        print("Check the 'Issues' sheet of:", ", ".join(str(s["ac_folder"]) for s in fails))
    sys.exit(0 if not fails else 1)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
