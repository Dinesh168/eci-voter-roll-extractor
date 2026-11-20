#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Download ECI 2027 draft-roll PDFs into  input\\<STATE>\\AC<ac>\\<part>.pdf

URL pattern
  https://voters.eci.gov.in/eroll/2027/<state>/ssr-draftroll/<ac>/
      2027-EROLLGEN-<STATE>-<ac>-SSR-DraftRoll-Revision1-<LANG>-<part>-WI.pdf

USAGE
  python download_rolls.py                       # S19 (Punjab), AC 1-117, 2 random parts each
  python download_rolls.py --per-ac 5 --seed 7   # 5 random parts per AC
  python download_rolls.py --acs 1-10,25 --parts all          # every part of those ACs
  python download_rolls.py --acs 2 --parts 1-10               # fixed parts
  python download_rolls.py --state S24 --lang HIN --acs 86 --parts 18

HOW IT WORKS
  * Finds how many parts each AC has with cheap HEAD requests (binary search),
    then picks random parts (reproducible with --seed) unless --parts is given.
  * Each PDF is verified before it is kept: HTTP size == Content-Length, starts with
    %PDF, contains %%EOF, opens in PyMuPDF. Written as .partial then renamed.
  * Already-downloaded valid files are skipped -> safe to re-run / resume.
  * Log: input\\<STATE>\\_download_log.xlsx
"""

import argparse
import os
import random
import sys
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = ("https://voters.eci.gov.in/eroll/2027/{st}/ssr-draftroll/{ac}/"
        "2027-EROLLGEN-{ST}-{ac}-SSR-DraftRoll-Revision1-{lang}-{part}-WI.pdf")
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
           "Accept": "application/pdf,*/*"}
_print_lock = threading.Lock()


def log(*a):
    with _print_lock:
        print(*a, flush=True)


def url_for(state, lang, ac, part):
    return BASE.format(st=state.lower(), ST=state.upper(), ac=ac, lang=lang, part=part)


def exists(url, retries=4):
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers=HEADERS, method="HEAD")
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.status == 200
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return False
        except Exception:
            pass
        time.sleep(1.5 * (i + 1))
    raise RuntimeError(f"cannot reach {url}")


def max_part(state, lang, ac, cap=2000):
    """Highest existing part number (parts are numbered 1..N)."""
    if not exists(url_for(state, lang, ac, 1)):
        return 0
    lo, hi = 1, 64
    while hi <= cap and exists(url_for(state, lang, ac, hi)):
        lo, hi = hi, hi * 2
    hi = min(hi, cap + 1)
    while hi - lo > 1:                       # invariant: lo exists, hi doesn't
        mid = (lo + hi) // 2
        if exists(url_for(state, lang, ac, mid)):
            lo = mid
        else:
            hi = mid
    return lo


def valid_pdf(path, expected_size=None):
    try:
        size = os.path.getsize(path)
        if expected_size and size != expected_size:
            return False, f"size {size} != {expected_size}"
        with open(path, "rb") as f:
            head = f.read(5)
            f.seek(max(0, size - 2048))
            tail = f.read()
        if head != b"%PDF-":
            return False, "not a PDF header"
        if b"%%EOF" not in tail:
            return False, "truncated (no %%EOF)"
        import fitz
        with fitz.open(path) as d:
            if d.page_count < 1:
                return False, "no pages"
            return True, f"{d.page_count} pages"
    except Exception as e:
        return False, repr(e)


def download(url, dest, retries=4):
    if os.path.exists(dest):
        ok, info = valid_pdf(dest)
        if ok:
            return "exists", os.path.getsize(dest), info
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    tmp = dest + ".partial"
    last = ""
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=120) as r:
                expected = int(r.headers.get("Content-Length") or 0) or None
                with open(tmp, "wb") as f:
                    while chunk := r.read(1 << 16):
                        f.write(chunk)
            ok, info = valid_pdf(tmp, expected)
            if ok:
                os.replace(tmp, dest)
                return "downloaded", os.path.getsize(dest), info
            last = info
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code}"
            if e.code == 404:
                break
        except Exception as e:
            last = repr(e)
        time.sleep(2 * (i + 1))
    if os.path.exists(tmp):
        os.remove(tmp)
    return "failed", 0, last


def parse_range(spec):
    out = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            out += range(int(a), int(b) + 1)
        elif part:
            out.append(int(part))
    return out


def process_ac(a, ac, rng_seed):
    st = a.state.upper()
    try:
        n = max_part(st, a.lang, ac)
    except Exception as e:
        return [{"ac": ac, "max_part": None, "part": None, "status": "failed", "detail": str(e)}]
    if n == 0:
        return [{"ac": ac, "max_part": 0, "part": None, "status": "no_parts", "detail": "part 1 not found"}]
    if a.parts == "random":
        rng = random.Random(f"{rng_seed}-{ac}")
        parts = sorted(rng.sample(range(1, n + 1), min(a.per_ac, n)))
    elif a.parts == "all":
        parts = list(range(1, n + 1))
    else:
        parts = [p for p in parse_range(a.parts) if p <= n]
    rows = []
    for p in parts:
        u = url_for(st, a.lang, ac, p)
        dest = os.path.join(a.out, st, f"AC{ac}", f"{p}.pdf")
        status, size, info = download(u, dest)
        rows.append({"ac": ac, "max_part": n, "part": p, "status": status, "size": size,
                     "detail": info, "file": dest, "url": u})
    log(f"AC{ac:<4} parts 1-{n:<4} -> " + ", ".join(f"{r['part']}:{r['status']}" for r in rows))
    return rows


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--state", default="S19")
    ap.add_argument("--lang", default="PAN", help="language code in file name (PAN, HIN, ...)")
    ap.add_argument("--acs", default="1-117", help="e.g. 1-117 or 1,5,9-12")
    ap.add_argument("--parts", default="random", help="'random' (default), 'all', or e.g. 1-10")
    ap.add_argument("--per-ac", type=int, default=2, help="random parts per AC (default 2)")
    ap.add_argument("--seed", default="2027", help="random seed (same seed = same parts)")
    ap.add_argument("--threads", type=int, default=6)
    ap.add_argument("--out", default=os.path.join(here, "input"))
    a = ap.parse_args()

    acs = parse_range(a.acs)
    print(f"{a.state.upper()}  ACs {a.acs} ({len(acs)})  parts={a.parts}"
          f"{' x' + str(a.per_ac) if a.parts == 'random' else ''}  -> {a.out}")
    t0 = time.time()
    rows = []
    with ThreadPoolExecutor(a.threads) as ex:
        futs = [ex.submit(process_ac, a, ac, a.seed) for ac in acs]
        for f in as_completed(futs):
            rows += f.result()
    rows.sort(key=lambda r: (r["ac"], r.get("part") or 0))

    # log workbook
    try:
        import xlsxwriter
        os.makedirs(os.path.join(a.out, a.state.upper()), exist_ok=True)
        path = os.path.join(a.out, a.state.upper(), "_download_log.xlsx")
        cols = ["ac", "max_part", "part", "status", "size", "detail", "file", "url"]
        wb = xlsxwriter.Workbook(path)
        ws = wb.add_worksheet("Downloads")
        ws.write_row(0, 0, cols, wb.add_format({"bold": True}))
        for i, r in enumerate(rows, 1):
            ws.write_row(i, 0, ["" if r.get(c) is None else r.get(c) for c in cols])
        ws.autofilter(0, 0, len(rows), len(cols) - 1)
        ws.freeze_panes(1, 0)
        wb.close()
        print("Log:", path)
    except Exception as e:
        print("[warn] could not write log workbook:", e)

    cnt = {}
    for r in rows:
        cnt[r["status"]] = cnt.get(r["status"], 0) + 1
    print(f"Done in {(time.time() - t0) / 60:.1f} min: {cnt}")
    bad = [r for r in rows if r["status"] in ("failed", "no_parts")]
    for r in bad:
        print(f"  [!] AC{r['ac']} part {r.get('part')}: {r['status']} {r.get('detail')}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
