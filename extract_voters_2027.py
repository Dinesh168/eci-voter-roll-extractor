#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Punjab (S19) Draft Electoral Roll 2027 PDF -> Excel, with full validation.

Built for the 2027 ECI "EROLLGEN-S19-<AC>-SSR-DraftRoll-Revision1-PAN-<part>-WI.pdf"
layout (Punjabi, text-based PDF, 3 voter cards per row). Tested on 234 PDFs
(2 random parts x all 117 Punjab ACs, 186,238 voters): all checks pass.

USAGE
    python extract_voters_2027.py input\\S19 -o output -w 8   # folder of AC1..AC117 folders
    python extract_voters_2027.py input\\S19\\AC2              # one AC folder
    python extract_voters_2027.py D:\\2027_voter_data\\S19 -o D:\\excel -w 8

LAYOUT VARIANTS HANDLED
    * ward / pincode / polling-station type present or blank on the cover
    * section list wrapping, empty section names ("2-"), and a cover list that is
      truncated with the full list on an extra page after the summary page
    * summary table not on the last page
    * names / relative names / house addresses wrapped onto 2+ lines, and lines
      PyMuPDF splits mid-word ("ਹਰਪ੍ਰੀ" + "ਤ ਕੌਰ")
    * EPICs in odd formats (aBC1234567, A1B2345678, ABCO123456, ABCDO12345) kept as printed
    * relation ਪਿਤਾ/ਪਤੀ/ਮਾਤਾ/ਪਤਨੀ/ਹੋਰ, gender ਪੁਰਸ਼/ਇਸਤਰੀ/ਤੀਜਾ ਲਿੰਗ
    * blank relative name / house number in the PDF (kept blank, flagged)
    * Hindi-script names and legacy-font glyphs (Ã Ç) kept exactly as printed, flagged
    * PDF text glitches fixed: doubled marks (ਲਿੰੰਗ), misplaced sihari (ਿਪਤਾ -> ਪਿਤਾ)

OUTPUT (Excel only, in <output>\\)
    <AC>.xlsx              sheets: Voters, Parts (cover metadata + checks), Issues
    RUN_SUMMARY.xlsx       one row per AC: PASS/FAIL, voters vs cover total, errors

    Several ACs run in parallel (-w N, default up to 4). Re-running skips ACs whose
    <AC>.xlsx already exists (resume after a stop); --force redoes them.

VALIDATION (per PDF)  - ERROR makes the PDF/AC FAIL, WARN is informational
    * voter cards found == cover-page total == summary-page total        (ERROR)
    * serial numbers contiguous from cover start SL to end SL            (ERROR)
    * male / female / third-gender counts == cover page counts           (ERROR)
    * CROSS-CHECK: every page is read twice - (A) as a text stream, (B) card by card
      using the drawn card boxes - and every field of every voter must be identical
                                                                           (ERROR)
    * every card: EPIC present, name, relation, age 18-125, gender        (ERROR)
    * unparsed text inside a card                                          (ERROR)
    * part number in footers / file name == cover                         (ERROR)
    * odd EPIC format, blank relative/house no., unusual script, sections  (WARN)
  and across the AC: duplicate EPICs (WARN), single assembly number (ERROR).

Exit code 0 = all checks passed, 1 = at least one ERROR (see Issues sheet).
Requires: pip install PyMuPDF xlsxwriter openpyxl
"""

import argparse
import glob
import os
import re
import sys
from collections import Counter

import fitz  # PyMuPDF
import time
from concurrent.futures import ProcessPoolExecutor, as_completed

# ----------------------------------------------------------------------------
# Text normalisation
# ----------------------------------------------------------------------------
GURMUKHI_CONS = "ਕ-ਹਖ਼-ਫ਼"        # consonants (incl. ਖ਼ ਗ਼ ...)
DUP_MARKS_RE = re.compile(r"([ੰੱਂ਼੍])\1+")   # ੰੰ ੱੱ ਂਂ ਼਼ ੍੍
# a word that starts with sihari (ਿ) is a visual-order glyph bug: "ਿਪਤਾ" -> "ਪਿਤਾ"
SIHARI_FIX_RE = re.compile(r"(?<![਀-੿])ਿ([" + GURMUKHI_CONS + r"]਼?)")


def norm(s):
    if not s:
        return ""
    s = DUP_MARKS_RE.sub(r"\1", s)
    s = SIHARI_FIX_RE.sub(r"\1ਿ", s)
    s = s.replace(" ", " ")
    return re.sub(r"[ \t]+", " ", s).strip()


# ----------------------------------------------------------------------------
# Patterns (applied to normalised text)
# ----------------------------------------------------------------------------
EPIC_RE = re.compile(r"^[A-Z]{3}[0-9]{7}$")
# EPICs seen in real rolls: ABC1234567, aBC1234567, A1B2345678, AbC1234567, ABCDO12345
EPIC_LOOSE_RE = re.compile(r"^(?=(?:[^0-9]*[0-9]){4})[A-Za-z0-9/]{6,16}$")
SL_RE = re.compile(r"^\d{1,5}$")
NAME_RE = re.compile(r"^ਨਾਮ\s*:\s*(.*)$")
REL_RE = re.compile(r"^(ਪਿਤਾ|ਪਤੀ|ਮਾਤਾ|ਪਤਨੀ|ਹੋਰ)(?:\s*ਦਾ ਨਾਮ)?\s*:\s*(.*)$")
HOUSE_RE = re.compile(r"^ਮਕਾਨ ਨੰ\.?\s*:\s*(.*)$")
AGE_RE = re.compile(r"^ਉਮਰ\s*:\s*(\d*)\s*ਲਿੰਗ\s*:\s*(.*)$")
PHOTO_RE = re.compile(r"^ਫੋਟੋ")
SECTION_RE = re.compile(r"^ਸੈਕਸ਼ਨ ਨੰ\.? ਅਤੇ ਨਾਮ\s+(\d+)\s*-\s*(.*)$")
FOOTER_RE = re.compile(r"ਕੁੱਲ ਪੰਨੇ\s*(\d+)\s*-\s*ਪੰਨਾ\s*(\d+)")
PART_FOOT_RE = re.compile(r"^ਭਾਗ ਨੰਬਰ\s*:\s*(\d+)")
AC_FOOT_RE = re.compile(r"^ਵਿਧਾਨ ਸਭਾ ਚੋਣ ਹਲਕੇ ਦਾ ਨੰ\.?, ਨਾਮ\s*:\s*(\d+)\s*-\s*(.*)$")
PAGE_NOISE = ("ਉਮਰ 01-01", "ਪ੍ਰਕਾਸ਼ਨਾ ਦੀ ਮਿਤੀ", "ਵਿਧਾਨ ਸਭਾ ਚੋਣ ਹਲਕੇ", "ਭਾਗ ਨੰਬਰ", "ਕੁੱਲ ਪੰਨੇ")

REL_EN = {"ਪਿਤਾ": "Father", "ਪਤੀ": "Husband", "ਮਾਤਾ": "Mother", "ਪਤਨੀ": "Wife", "ਹੋਰ": "Other"}
GENDER_EN = {"ਪੁਰਸ਼": "Male", "ਇਸਤਰੀ": "Female", "ਤੀਜਾ ਲਿੰਗ": "Third Gender"}

COLUMNS = [
    "ac_folder", "source_pdf", "state_code", "state_name", "assembly_no", "assembly_name",
    "reservation", "parliament_no", "parliament_name", "part_no", "district", "tehsil",
    "police_station", "main_village", "ward", "pincode", "polling_station_no",
    "polling_station_name", "section_no", "section_name", "page_no",
    "sl_no", "epic", "voter_name", "relation_type", "relation_type_pa",
    "relative_name", "house_no", "age", "gender", "gender_pa", "photo_status",
    "warnings",
]


def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


# ----------------------------------------------------------------------------
# Line reconstruction: PyMuPDF sometimes splits one visual line into two
# (e.g. "ਹਰਪ੍ਰੀ" + "ਤ ਕੌਰ"). Re-join pieces on the same baseline that touch.
# ----------------------------------------------------------------------------
def page_lines(page, clip=None):
    return [t[0] for t in page_items(page, clip)]


def page_items(page, clip=None):
    """[(normalised_text, x0, y0, x1, y1)] in reading order, split lines re-joined."""
    out = []  # [text, x0, y0, x1, y1]
    for b in page.get_text("dict", clip=clip)["blocks"]:
        for ln in b.get("lines", []):
            txt = "".join(s["text"] for s in ln["spans"])
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
    return res


# ----------------------------------------------------------------------------
# Cover page (page 1) and summary page (last page)
# ----------------------------------------------------------------------------
COVER_LABELS = (
    "ਸੁਧਾਈ ਦਾ ਸਾਲ", "ਯੋਗਤਾ ਮਿਤੀ", "ਸੁਧਾਈ ਦੀ ਕਿਸਮ", "ਪ੍ਰਕਾਸ਼ਨਾ ਦੀ ਮਿਤੀ", "ਵੋਟਰ ਸੂਚੀ ਦੀ ਪਹਿਚਾਣ",
    "2. ", "3. ", "4. ", "ਭਾਗ ਵਿੱਚ ਸੈਕਸ਼ਨਾਂ", "ਮੁੱਖ ਪਿੰਡ", "ਵਾਰਡ", "ਪੁਲਿਸ ਸਟੇਸ਼ਨ", "ਤਹਿਸੀਲ",
    "ਜ਼ਿਲ੍ਹਾ", "ਪਿੰਨ ਕੋਡ", "ਪੋਲਿੰਗ ਸਟੇਸ਼ਨ", "ਇਸ ਭਾਗ ਵਿੱਚ", "ਦੀ ਗਿਣਤੀ", "ਕੁੱਲ ਵੋਟਰ",
)


def _is_label(s):
    return any(s.startswith(x) for x in COVER_LABELS)


def _after(lines, label, default=""):
    """Value on the line following a label line (strip a leading ':').
    A blank value (next line is ':' or another label) returns ''."""
    for i, l in enumerate(lines):
        if l.startswith(label):
            rest = l[len(label):].strip()
            # same-line value only when it follows a ':'  (else it is label text)
            if (rest.startswith(":") or label.rstrip().endswith(":")) and rest.lstrip(":").strip():
                return rest.lstrip(":").strip()
            if i + 1 < len(lines):
                v = lines[i + 1]
                if v.startswith(":"):
                    return v[1:].strip()
                return "" if _is_label(v) else v.strip()
            return ""
    return default


def parse_cover(lines):
    c = {}
    txt = "\n".join(lines)
    m = re.search(r"^ਵੋਟਰ ਸੂਚੀ\s+(\d{4})\s+(S\d+)\s+(.+)$", txt, re.M)
    c["roll_year"], c["state_code"], c["state_name"] = (
        (int(m.group(1)), m.group(2), m.group(3).strip()) if m else (None, "", ""))
    ident = []
    try:
        i = next(k for k, l in enumerate(lines) if l.startswith("ਵੋਟਰ ਸੂਚੀ ਦੀ ਪਹਿਚਾਣ"))
        for l in lines[i + 1:]:
            if _is_label(l):
                break
            ident.append(l)
    except StopIteration:
        pass
    c["roll_identification"] = " ".join(ident)
    m = re.search(r"ਰਾਖਵੇਂਕਰਨ ਦੀ ਸਥਿਤੀ\s*:\s*(\d+)\s*-\s*([^\n(]+?)\s*(\([^)]*\))?\s*$", txt, re.M)
    if m:
        c["assembly_no"] = int(m.group(1))
        c["assembly_name"] = m.group(2).strip()
        c["reservation"] = (m.group(3) or "").strip("() ")
    m = re.search(r"ਸਥਿਤ ਹੈ\s*:\s*(\d+)\s*-\s*(.+)$", txt, re.M)
    if m:
        c["parliament_no"] = int(m.group(1))
        c["parliament_name"] = m.group(2).strip()
    m = re.search(r"^ਭਾਗ ਨੰਬਰ\s*:\s*(\d+)", txt, re.M)
    c["part_no"] = int(m.group(1)) if m else None
    c["revision_year"] = _after(lines, "ਸੁਧਾਈ ਦਾ ਸਾਲ")
    c["qualifying_date"] = _after(lines, "ਯੋਗਤਾ ਮਿਤੀ")
    c["revision_type"] = _after(lines, "ਸੁਧਾਈ ਦੀ ਕਿਸਮ")
    c["publication_date"] = _after(lines, "ਪ੍ਰਕਾਸ਼ਨਾ ਦੀ ਮਿਤੀ")

    c["sections"] = parse_section_list(lines)

    c["main_village"] = _after(lines, "ਮੁੱਖ ਪਿੰਡ/ ਸ਼ਹਿਰ") or _after(lines, "ਮੁੱਖ ਪਿੰਡ/ਸ਼ਹਿਰ")
    c["ward"] = _after(lines, "ਵਾਰਡ")
    c["police_station"] = _after(lines, "ਪੁਲਿਸ ਸਟੇਸ਼ਨ")
    c["tehsil"] = _after(lines, "ਤਹਿਸੀਲ")
    c["district"] = _after(lines, "ਜ਼ਿਲ੍ਹਾ")
    c["pincode"] = _after(lines, "ਪਿੰਨ ਕੋਡ")
    for k in ("ward", "police_station", "tehsil", "district", "pincode", "main_village"):
        # an empty value makes _after return the next label; reject those
        if c[k] and not c[k][0].isdigit() and k == "pincode":
            c[k] = ""
        if c[k] in ("ਪੁਲਿਸ ਸਟੇਸ਼ਨ", "ਤਹਿਸੀਲ", "ਜ਼ਿਲ੍ਹਾ", "ਪਿੰਨ ਕੋਡ", "ਵਾਰਡ"):
            c[k] = ""
    c["polling_station_type"] = _after(lines, "ਪੋਲਿੰਗ ਸਟੇਸ਼ਨ ਦੀ ਕਿਸਮ")
    ps = _after(lines, "ਪੋਲਿੰਗ ਸਟੇਸ਼ਨ ਦਾ ਨੰ. ਅਤੇ ਨਾਮ")
    mm = re.match(r"^(\d+)\s*-\s*(.*)$", ps)
    c["polling_station_no"] = int(mm.group(1)) if mm else None
    c["polling_station_name"] = mm.group(2).strip() if mm else ps
    # address may wrap over several lines
    addr = []
    try:
        i = next(k for k, l in enumerate(lines) if l.startswith("ਪੋਲਿੰਗ ਸਟੇਸ਼ਨ ਦਾ ਪਤਾ"))
        for l in lines[i + 1:]:
            if l.startswith("ਇਸ ਭਾਗ ਵਿੱਚ"):
                break
            addr.append(l)
    except StopIteration:
        pass
    c["polling_station_address"] = " ".join(addr)
    aux = _after(lines, "ਦੀ ਗਿਣਤੀ :")
    c["auxiliary_ps"] = int(aux) if aux.isdigit() else None

    # counts: the 6 numbers after "ਜੋੜ": start SL, end SL, male, female, third, total
    nums = []
    try:
        i = next(k for k, l in enumerate(lines) if l == "ਜੋੜ")
        for l in lines[i + 1:]:
            if re.fullmatch(r"\d+", l):
                nums.append(int(l))
                if len(nums) == 6:
                    break
            elif nums:
                break
    except StopIteration:
        pass
    keys = ("start_sl", "end_sl", "cover_male", "cover_female", "cover_third", "cover_total")
    for k, v in zip(keys, nums + [None] * 6):
        c[k] = v
    return c


SECTION_LIST_HDR = "ਭਾਗ ਵਿੱਚ ਸੈਕਸ਼ਨਾਂ"
SUMMARY_HDR = "ਵੋਟਰਾਂ ਦੀ ਸਾਰਣੀ"


def parse_section_list(lines):
    """'N-name' lines after the section-list header (cover, or the extra full-list page
    that follows the summary when the cover box overflows)."""
    sections = {}
    try:
        i = next(k for k, l in enumerate(lines) if l.startswith(SECTION_LIST_HDR))
    except StopIteration:
        return sections
    last = None
    for l in lines[i + 1:]:
        if l.startswith("ਮੁੱਖ ਪਿੰਡ") or l.startswith(PAGE_NOISE):
            break
        mm = re.match(r"^(\d+)\s*-\s*(.*)$", l)
        if mm:                                   # "2-" (empty name) is valid
            last = int(mm.group(1))
            sections[last] = mm.group(2).strip()
        elif last is not None:                   # wrapped section name
            sections[last] = (sections[last] + " " + l).strip()
    return sections


def parse_summary(lines):
    """Last page: first 4 numbers after 'ਜੋੜ' = male, female, third, total."""
    nums = []
    try:
        i = next(k for k, l in enumerate(lines) if l == "ਜੋੜ")
        for l in lines[i + 1:]:
            if re.fullmatch(r"\d+", l):
                nums.append(int(l))
                if len(nums) == 8:
                    break
    except StopIteration:
        pass
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
    """Separate page header/footer from the card body."""
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
    """Method A: read the page as one text stream and cut it at each card start."""
    body, info = split_page(lines)
    starts = [i for i in range(len(body)) if is_card_start(body, i)]
    cards = [parse_card(body[s:e], page_no)
             for s, e in zip(starts, starts[1:] + [len(body)])]
    return cards, info


def card_boxes(page):
    """Card rectangles from the drawn grid (vertical border lines of each card)."""
    import fitz as _f
    bands = {}
    for g in page.get_drawings():
        r = g["rect"]
        if r.width < 2.5 and 55 < r.height < 140:
            bands.setdefault((round(r.y0), round(r.y1)), set()).add(round((r.x0 + r.x1) / 2, 1))
    boxes = []
    for (y0, y1), xs in sorted(bands.items()):
        xs = sorted(xs)
        merged = []
        for x in xs:                           # merge near-duplicate lines
            if not merged or x - merged[-1] > 2:
                merged.append(x)
        for a, b in zip(merged[0::2], merged[1::2]):
            if b - a > 100:
                boxes.append(_f.Rect(a, y0, b, y1))
    return boxes


def box_cards(page, items, page_no):
    """Method B: assign every text line to the card box that contains it (by position),
    then read each box on its own - independent of the stream order used by method A."""
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
        top = sorted([z for z in b if z[0] - top_y < 3], key=lambda q: q[1])   # SL ... EPIC
        rest = [z for z in b if z[0] - top_y >= 3]
        # card layout: text column on the left, photo placeholder column on the right
        split_x = r.x0 + 0.70 * r.width
        photo = [z for z in rest if z[1] >= split_x]
        text = [z for z in rest if z[1] < split_x]
        lines = [t for _, _, t in top] + [t for _, _, t in photo] + [t for _, _, t in text]
        if lines and SL_RE.match(lines[0]):
            out.append(parse_card(lines, page_no))
    return out


def parse_card(seg, page_no):
    """seg = lines of one card, first line is the serial number."""
    cur = {"sl_no": int(seg[0]), "epic": "", "photo_status": "", "voter_name": "",
           "relation_type_pa": "", "relative_name": "", "house_no": "",
           "age": None, "gender_pa": "", "page_no": page_no, "_extra": []}
    last = None
    for l in seg[1:]:
        if not cur["epic"] and not cur["voter_name"] and EPIC_LOOSE_RE.match(l):
            cur["epic"] = l
            last = None
        elif PHOTO_RE.match(l):
            cur["photo_status"] = l
            last = "photo_status"
        elif (m := NAME_RE.match(l)):
            cur["voter_name"] = m.group(1)
            last = "voter_name"
        elif (m := REL_RE.match(l)):
            cur["relation_type_pa"], cur["relative_name"] = m.group(1), m.group(2)
            last = "relative_name"
        elif (m := HOUSE_RE.match(l)):
            cur["house_no"] = m.group(1)
            last = "house_no"
        elif (m := AGE_RE.match(l)):
            cur["age"] = int(m.group(1)) if m.group(1) else None
            cur["gender_pa"] = m.group(2).strip()
            last = None
        elif last:  # wrapped continuation of previous field
            cur[last] = (cur[last] + " " + l).strip()
            if last != "photo_status":
                cur["_extra"].append(f"wrapped {last}")
        else:
            cur["_extra"].append(f"unparsed line: {l}")
    return cur


CROSS_FIELDS = ("epic", "voter_name", "relation_type_pa", "relative_name",
                "house_no", "age", "gender_pa")


# ----------------------------------------------------------------------------
# One PDF
# ----------------------------------------------------------------------------
def process_pdf(path, ac_folder):
    name = os.path.basename(path)
    issues = []  # (severity, pdf, sl_no, check, detail)

    def issue(sev, check, detail, sl=None):
        issues.append({"severity": sev, "source_pdf": name, "sl_no": sl,
                       "check": check, "detail": detail})

    doc = fitz.open(path)
    items = [page_items(p) for p in doc]
    pages = [[t[0] for t in it] for it in items]
    cover = parse_cover(pages[0])
    # summary table: page with its heading (normally last, but an extra section-list
    # page may follow it)
    sum_idx = next((k for k in range(len(pages) - 1, 0, -1)
                    if any(l.startswith(SUMMARY_HDR) for l in pages[k])), len(pages) - 1)
    summary = parse_summary(pages[sum_idx])
    # full section list on a later page (when the cover box overflows)
    for k in range(1, len(pages)):
        if k != sum_idx and any(l.startswith(SECTION_LIST_HDR) for l in pages[k]) \
                and not any(l.startswith("ਉਮਰ :") for l in pages[k]):
            extra = parse_section_list(pages[k])
            if extra:
                cover["sections"].update(extra)
                cover["sections_from_extra_page"] = k + 1

    voters, seen_sections = [], set()
    footer_totals, footer_parts = set(), set()
    cross_pages = cross_skipped = cross_mismatch = 0
    for pno in range(1, doc.page_count):
        lines = pages[pno]
        if not any(l.startswith("ਉਮਰ :") for l in lines):
            continue  # map/photo page or summary page
        cards, info = parse_voter_page(lines, pno + 1)
        sec_no, sec_name = info["section_no"], info["section_name"]
        if info["footer_total"]:
            footer_totals.add(info["footer_total"])
        if info["part_no"]:
            footer_parts.add(info["part_no"])
        if info["footer_page"] and info["footer_page"] != pno + 1:
            issue("WARN", "footer_page", f"page {pno+1} footer says page {info['footer_page']}")
        if sec_no is None:
            issue("ERROR", "section_missing", f"page {pno+1} has no section header")
        else:
            seen_sections.add(sec_no)
            if cover["sections"] and sec_no not in cover["sections"]:
                issue("WARN", "section_not_on_cover",
                      f"page {pno+1} section {sec_no} not in cover list (cover list is truncated)")

        # ---- independent cross-check: card-box segmentation must give identical data
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
                issue("ERROR", "crosscheck_cards",
                      f"page {pno+1}: stream SL {sorted(set(amap) - set(bmap))} vs "
                      f"box SL {sorted(set(bmap) - set(amap))}")
            for sl in set(amap) & set(bmap):
                for fld in CROSS_FIELDS:
                    if amap[sl][fld] != bmap[sl][fld]:
                        cross_mismatch += 1
                        issue("ERROR", "crosscheck_field",
                              f"page {pno+1} {fld}: stream '{amap[sl][fld]}' vs box '{bmap[sl][fld]}'", sl)
        for c in cards:
            c["section_no"], c["section_name"] = sec_no, sec_name
            voters.append(c)
    doc.close()

    # ---- per-card checks ----
    gcount = Counter()
    for v in voters:
        sl = v["sl_no"]
        w = list(v.pop("_extra"))
        v["relation_type"] = REL_EN.get(v["relation_type_pa"], "")
        v["gender"] = GENDER_EN.get(v["gender_pa"], "")
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
            issue("ERROR", "relation_type", f"unknown relation '{v['relation_type_pa']}'", sl)
        if v["age"] is None or not (18 <= v["age"] <= 125):
            issue("ERROR", "age", f"age invalid: {v['age']}", sl)
        if not v["gender"]:
            issue("ERROR", "gender", f"unknown gender '{v['gender_pa']}'", sl)
        if not v["house_no"]:
            w.append("house_no empty")
        for fld in ("voter_name", "relative_name"):
            if re.search(r"[A-Za-z0-9]", v[fld]):
                w.append(f"latin/digit in {fld}")
            if re.search(r"[ऀ-ॿ]", v[fld]):
                w.append(f"Devanagari (Hindi) script in {fld}")
        for fld in ("voter_name", "relative_name", "house_no"):
            odd = sorted({ch for ch in v[fld] if not ("਀" <= ch <= "੿" or ch.isascii()
                                                     or "ऀ" <= ch <= "ॿ")})
            if odd:  # printed like this in the PDF (legacy-font glyphs such as Ã Ç)
                w.append(f"unusual character {''.join(odd)} in {fld} (as printed)")
        for x in w:
            if x.startswith("unparsed"):
                issue("ERROR", "unparsed", x, sl)
        v["warnings"] = "; ".join(w)

    # ---- per-PDF totals ----
    n = len(voters)
    sls = [v["sl_no"] for v in voters]
    exp_total = cover.get("cover_total")
    checks = {}

    def chk(key, ok, detail):
        checks[key] = "OK" if ok else "FAIL"
        if not ok:
            issue("ERROR", key, detail)

    chk("count_vs_cover", n == exp_total, f"parsed {n} voters, cover says {exp_total}")
    chk("count_vs_summary", n == summary["sum_total"], f"parsed {n}, summary page says {summary['sum_total']}")
    s0, s1 = cover.get("start_sl"), cover.get("end_sl")
    exp_sl = list(range(s0, s1 + 1)) if s0 and s1 else []
    missing = sorted(set(exp_sl) - set(sls))
    dup = sorted(k for k, c in Counter(sls).items() if c > 1)
    chk("sl_sequence", sls == exp_sl,
        f"SL not contiguous {s0}-{s1}; missing={missing[:20]} dup={dup[:20]}")
    chk("male_count", gcount["Male"] == cover.get("cover_male"),
        f"male {gcount['Male']} vs cover {cover.get('cover_male')}")
    chk("female_count", gcount["Female"] == cover.get("cover_female"),
        f"female {gcount['Female']} vs cover {cover.get('cover_female')}")
    chk("third_count", gcount["Third Gender"] == cover.get("cover_third"),
        f"third gender {gcount['Third Gender']} vs cover {cover.get('cover_third')}")
    chk("part_no_footer", footer_parts <= {cover.get("part_no")},
        f"footer part numbers {footer_parts} vs cover {cover.get('part_no')}")
    chk("crosscheck", cross_mismatch == 0,
        f"{cross_mismatch} differences between stream and card-box extraction")
    checks["crosscheck_pages"] = f"{cross_pages} checked, {cross_skipped} skipped"
    if cover["sections"] and not set(cover["sections"]) <= seen_sections:
        issue("WARN", "sections_seen", f"cover sections {sorted(set(cover['sections']) - seen_sections)} "
              f"have no voter pages")
    # part number in file name: original ECI name (...-PAN-18-WI.pdf) or organised (18.pdf)
    m = re.search(r"PAN-(\d+)", name) or re.fullmatch(r"(\d+)\.pdf", name, re.I)
    if m and cover.get("part_no") is not None:
        chk("part_no_filename", int(m.group(1)) == cover["part_no"],
            f"filename part {m.group(1)} vs cover part {cover['part_no']}")
    m = re.fullmatch(r"AC0*(\d+)", ac_folder, re.I)
    if m and cover.get("assembly_no") is not None:
        chk("assembly_no_folder", int(m.group(1)) == cover["assembly_no"],
            f"folder {ac_folder} vs cover assembly {cover['assembly_no']}")

    # ---- metadata consistency
    tot = cover.get("cover_total")
    chk("cover_totals", tot is not None and
        (cover.get("cover_male") or 0) + (cover.get("cover_female") or 0) + (cover.get("cover_third") or 0) == tot
        and cover.get("start_sl") is not None and cover.get("end_sl") is not None
        and cover["end_sl"] - cover["start_sl"] + 1 == tot,
        f"cover counts inconsistent: male+female+third / SL range vs total {tot}")
    chk("summary_vs_cover",
        (summary["sum_male"], summary["sum_female"], summary["sum_third"], summary["sum_total"]) ==
        (cover.get("cover_male"), cover.get("cover_female"), cover.get("cover_third"), tot) and
        (summary["sumrow_male"], summary["sumrow_female"], summary["sumrow_third"], summary["sumrow_total"]) ==
        (summary["sum_male"], summary["sum_female"], summary["sum_third"], summary["sum_total"]),
        f"summary page {summary} vs cover")
    bad_foot = []
    for k in range(1, len(pages)):
        t = "\n".join(pages[k])
        fm = re.search(r"ਵਿਧਾਨ ਸਭਾ ਚੋਣ ਹਲਕੇ ਦਾ ਨੰ\.?, ਨਾਮ\s*:\s*(\d+)\s*-\s*(.+)", t)
        pm = re.search(r"ਪ੍ਰਕਾਸ਼ਨਾ ਦੀ ਮਿਤੀ\s*:-\s*(\S+)", t)
        qm = re.search(r"ਉਮਰ (\d\d-\d\d-\d{4})", t)
        gm = FOOTER_RE.search(t)
        if not (fm and int(fm.group(1)) == cover.get("assembly_no")
                and fm.group(2).strip() == cover.get("assembly_name")
                and pm and pm.group(1) == cover.get("publication_date")
                and qm and qm.group(1) == cover.get("qualifying_date")
                and gm and int(gm.group(1)) == len(pages) and int(gm.group(2)) == k + 1):
            bad_foot.append(k + 1)
    chk("page_footers", not bad_foot,
        f"pages {bad_foot[:15]}: footer AC/date/page-number differs from cover")
    for fld in ("assembly_no", "assembly_name", "part_no", "publication_date", "qualifying_date",
                "state_code"):
        if cover.get(fld) in (None, ""):
            issue("ERROR", "metadata_missing", f"cover field '{fld}' not found")
    for fld in ("district", "tehsil", "police_station", "main_village", "pincode",
                "polling_station_name", "polling_station_address", "polling_station_type"):
        if cover.get(fld) in (None, ""):
            issue("WARN", "metadata_blank", f"cover field '{fld}' is blank in the PDF")
    if cover.get("pincode") and not re.fullmatch(r"\d{6}", cover["pincode"]):
        issue("WARN", "pincode_format", f"pincode '{cover['pincode']}'")
    if cover.get("polling_station_no") not in (None, cover.get("part_no")):
        issue("WARN", "ps_no_vs_part", f"polling station {cover.get('polling_station_no')} "
              f"!= part {cover.get('part_no')}")

    # ---- attach metadata ----
    for v in voters:
        v.update({
            "ac_folder": ac_folder, "source_pdf": name, "state_code": cover.get("state_code"),
            "state_name": cover.get("state_name"), "ward": cover.get("ward"),
            "assembly_no": cover.get("assembly_no"), "assembly_name": cover.get("assembly_name"),
            "reservation": cover.get("reservation"), "parliament_no": cover.get("parliament_no"),
            "parliament_name": cover.get("parliament_name"), "part_no": cover.get("part_no"),
            "district": cover.get("district"), "tehsil": cover.get("tehsil"),
            "police_station": cover.get("police_station"), "main_village": cover.get("main_village"),
            "pincode": cover.get("pincode"), "polling_station_no": cover.get("polling_station_no"),
            "polling_station_name": cover.get("polling_station_name"),
        })

    part_row = {"source_pdf": name, "pages": len(pages), "parsed_voters": n,
                "parsed_male": gcount["Male"], "parsed_female": gcount["Female"],
                "parsed_third": gcount["Third Gender"]}
    part_row.update({k: v for k, v in cover.items() if k != "sections"})
    part_row["sections"] = "; ".join(f"{k}-{v}" for k, v in sorted(cover["sections"].items()))
    part_row.update(summary)
    part_row.update(checks)
    part_row["status"] = "PASS" if all(v == "OK" for k, v in checks.items()
                                       if k != "crosscheck_pages") and \
        not any(i["severity"] == "ERROR" for i in issues) else "FAIL"
    return voters, part_row, issues


# ----------------------------------------------------------------------------
# Excel writer (xlsxwriter, constant memory = fast for 200k+ rows)
# ----------------------------------------------------------------------------
PART_COLS = [
    "source_pdf", "status", "state_code", "state_name", "roll_year", "pages",
    "parsed_voters", "cover_total", "sum_total", "sumrow_total",
    "parsed_male", "cover_male", "parsed_female", "cover_female", "parsed_third", "cover_third",
    "assembly_no", "assembly_name", "reservation", "parliament_no", "parliament_name",
    "part_no", "start_sl", "end_sl", "sections", "main_village", "ward", "police_station",
    "tehsil", "district", "pincode", "polling_station_type", "polling_station_no",
    "polling_station_name", "polling_station_address", "auxiliary_ps", "revision_year",
    "qualifying_date", "revision_type", "publication_date", "roll_identification",
    "count_vs_cover", "count_vs_summary", "sl_sequence", "male_count", "female_count",
    "third_count", "part_no_footer", "part_no_filename", "assembly_no_folder", "cover_totals",
    "summary_vs_cover", "page_footers", "crosscheck", "crosscheck_pages", "sections_from_extra_page",
]
ISSUE_COLS = ["severity", "source_pdf", "sl_no", "check", "detail"]
SUMMARY_COLS = ["ac_folder", "status", "assembly_no", "assembly_name", "pdfs", "pdfs_pass",
                "pdfs_fail", "voters", "cover_total", "male", "female", "third",
                "errors", "warnings", "seconds", "file"]


def write_xlsx(path, sheets):
    """sheets: [(name, columns, rows)] where rows are dicts. Atomic: temp file + rename."""
    import xlsxwriter
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp.xlsx"
    wb = xlsxwriter.Workbook(tmp, {"constant_memory": True, "strings_to_urls": False,
                                   "strings_to_formulas": False, "strings_to_numbers": False})
    hdr = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
    for name, cols, rows in sheets:
        ws = wb.add_worksheet(name)
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
    for attempt in range(5):                # target may be open in Excel
        try:
            os.replace(tmp, path)
            return path
        except PermissionError:
            time.sleep(2)
    alt = os.path.splitext(path)[0] + time.strftime("_%Y%m%d_%H%M%S") + ".xlsx"
    os.replace(tmp, alt)
    print(f"[warn] {path} is open/locked (close it in Excel); saved as {alt}", flush=True)
    return alt


# ----------------------------------------------------------------------------
# One AC folder  ->  <out>\<AC>.xlsx
# ----------------------------------------------------------------------------
def _safe_pdf(path, ac):
    try:
        return process_pdf(path, ac)
    except Exception as e:  # keep going, report it
        name = os.path.basename(path)
        return [], {"source_pdf": name, "status": "FAIL"}, [
            {"severity": "ERROR", "source_pdf": name, "sl_no": None,
             "check": "exception", "detail": repr(e)}]


def list_pdfs(ac_dir):
    pdfs = sorted(glob.glob(os.path.join(ac_dir, "*.pdf")), key=natural_key)
    return sorted(pdfs, key=lambda p: int(m.group(1)) if (m := re.search(r"PAN-(\d+)", p)) else 0)


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

    # cross-PDF checks
    epics = Counter(v["epic"] for v in all_v if v["epic"])
    dups = {e for e, c in epics.items() if c > 1}
    if dups:
        where = {}
        for v in all_v:
            if v["epic"] in dups:
                where.setdefault(v["epic"], []).append(f"{v['source_pdf']}#{v['sl_no']}")
        for e, w in where.items():
            issues.append({"severity": "WARN", "source_pdf": "", "sl_no": None,
                           "check": "duplicate_epic", "detail": f"{e} x{len(w)}: {', '.join(w)}"})
    acs = sorted({r.get("assembly_no") for r in parts if r.get("assembly_no") is not None})
    if len(acs) > 1:
        issues.append({"severity": "ERROR", "source_pdf": "", "sl_no": None,
                       "check": "assembly_mixed", "detail": f"assembly numbers {acs}"})

    out_path = write_xlsx(out_path, [("Voters", COLUMNS, all_v),
                          ("Parts", PART_COLS, parts),
                          ("Issues", ISSUE_COLS, issues)])

    g = Counter(v["gender"] for v in all_v)
    n_err = sum(i["severity"] == "ERROR" for i in issues)
    return {
        "ac_folder": ac, "status": "PASS" if n_err == 0 else "FAIL",
        "assembly_no": acs[0] if len(acs) == 1 else ",".join(map(str, acs)),
        "assembly_name": next((r.get("assembly_name") for r in parts if r.get("assembly_name")), ""),
        "pdfs": len(pdfs), "pdfs_pass": sum(r.get("status") == "PASS" for r in parts),
        "pdfs_fail": sum(r.get("status") != "PASS" for r in parts),
        "voters": len(all_v), "cover_total": sum(r.get("cover_total") or 0 for r in parts),
        "male": g["Male"], "female": g["Female"], "third": g["Third Gender"],
        "errors": n_err, "warnings": sum(i["severity"] == "WARN" for i in issues),
        "seconds": round(time.time() - t0, 1), "file": out_path,
    }


def load_previous_summary(path):
    """Keep rows of already-finished ACs when resuming."""
    if not os.path.exists(path):
        return {}
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        try:
            rows = list(wb["Summary"].iter_rows(values_only=True))
        finally:
            wb.close()                      # release the file handle (Windows)
        return {r[0]: dict(zip(rows[0], r)) for r in rows[1:] if r and r[0]}
    except Exception:
        return {}


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", nargs="?", default=os.path.join(here, "input"),
                    help="folder of PDFs, or folder of AC sub-folders (default: ./input)")
    ap.add_argument("-o", "--output", default=os.path.join(here, "output"))
    ap.add_argument("-w", "--workers", type=int, default=max(1, min(4, os.cpu_count() or 1)),
                    help="parallel processes (default: up to 4)")
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
    for f in glob.glob(os.path.join(a.output, "*.tmp.xlsx")):   # leftovers of an interrupted run
        os.remove(f)

    sum_path = os.path.join(a.output, "RUN_SUMMARY.xlsx")
    summary = load_previous_summary(sum_path)
    todo = [d for d in ac_dirs
            if a.force or not os.path.exists(
                os.path.join(a.output, os.path.basename(os.path.normpath(d)) + ".xlsx"))]
    print(f"{len(ac_dirs)} AC folders, {len(ac_dirs) - len(todo)} already done (skipped), "
          f"{len(todo)} to process with {a.workers} workers -> {os.path.abspath(a.output)}")

    t0 = time.time()

    def record(res, k):
        summary[res["ac_folder"]] = res
        print(f"[{k:>3}/{len(todo)}] {res['ac_folder']:<12} {res['status']}  "
              f"{res['voters']:>7,} voters (cover {res['cover_total']:,})  "
              f"PDFs {res['pdfs_pass']}/{res['pdfs']} pass  err {res['errors']}  "
              f"{res['seconds']}s   elapsed {(time.time() - t0) / 60:.1f} min", flush=True)
        write_xlsx(sum_path, [("Summary", SUMMARY_COLS,
                               [summary[x] for x in sorted(summary, key=natural_key)])])

    def failed(d, e):
        return {"ac_folder": os.path.basename(os.path.normpath(d)), "status": "FAIL",
                "voters": 0, "cover_total": 0, "pdfs": 0, "pdfs_pass": 0, "errors": 1,
                "seconds": 0, "file": f"crashed: {e!r}"}

    try:
        if len(todo) == 1:                       # one AC: parallelise over its PDFs
            try:
                record(process_ac(todo[0], a.output, a.workers), 1)
            except Exception as e:
                record(failed(todo[0], e), 1)
        elif todo:                               # many ACs: one AC per process
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
