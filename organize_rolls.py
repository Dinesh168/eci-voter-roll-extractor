#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAFE organiser for ECI electoral-roll PDFs  ->  <DEST>\\<STATE>\\AC<ac>\\<part>.pdf

    2027-EROLLGEN-S24-86-SSR-DraftRoll-Revision1-HIN-18-WI.pdf
        -> <DEST>\\S24\\AC86\\18.pdf

USAGE
    python organize_rolls.py SOURCE DEST --dry-run     # preview only, changes nothing
    python organize_rolls.py SOURCE DEST               # move
    python organize_rolls.py SOURCE DEST --copy        # copy, originals untouched
    python organize_rolls.py --undo  DEST\\_organize_logs\\manifest_<time>.csv
    python organize_rolls.py --verify DEST\\_organize_logs\\manifest_<time>.csv

NO-DATA-LOSS GUARANTEES
    * A file is NEVER overwritten. Same-drive moves use os.rename, which fails if
      the target exists; copies go to "<target>.partial" and are renamed the same way.
    * Cross-drive move / --copy: bytes are compared after copying; the source is
      deleted only after the copy is verified identical.
    * Target already exists:
        - byte-identical  -> source is LEFT IN PLACE (status duplicate_identical)
        - different       -> source goes to DEST\\_conflicts\\...\\<original name>
                             (numbered __2, __3 ... if that name exists too)
    * Names that don't match the pattern are not touched.
    * Every action is appended + flushed to the manifest immediately, so the log
      survives a crash / power cut. --undo moves everything back from it.
    * Ctrl+C stops cleanly after the current file. Re-run the same command to resume
      (already-moved files are no longer in SOURCE, so nothing is done twice).
    * A final verification checks every moved file exists at its target with the
      right size and that no source remains where a move was recorded.
"""

import argparse
import csv
import filecmp
import os
import re
import shutil
import sys
import time

NAME_RE = re.compile(
    r"EROLLGEN-(?P<state>S\d{1,3})-(?P<ac>\d{1,4})-.*-(?P<part>\d{1,5})(?:-[A-Z]{1,4})?\.pdf$",
    re.IGNORECASE,
)
FIELDS = ["time", "source", "target", "state", "ac", "part", "size", "status", "detail"]


# ----------------------------------------------------------------------------
# primitive, never-overwrite file operations
# ----------------------------------------------------------------------------
def rename_no_overwrite(src, dst):
    """Atomic rename on the same volume; raises FileExistsError if dst exists."""
    if os.name == "nt":
        os.rename(src, dst)                 # Windows: fails if dst exists
    else:
        os.link(src, dst)                   # POSIX: link fails if dst exists
        os.unlink(src)


def same_volume(a, b):
    try:
        return os.stat(a).st_dev == os.stat(b).st_dev
    except OSError:
        return os.path.splitdrive(os.path.abspath(a))[0].lower() == \
            os.path.splitdrive(os.path.abspath(b))[0].lower()


def safe_transfer(src, dst, copy, same_vol):
    """Move/copy src -> dst without ever overwriting and without losing src.
    Returns 'moved' / 'copied'. Raises on any problem (src is then still intact)."""
    if os.path.lexists(dst):
        raise FileExistsError(dst)
    if not copy and same_vol:
        rename_no_overwrite(src, dst)
        return "moved"
    tmp = dst + ".partial"
    if os.path.lexists(tmp):
        os.remove(tmp)                      # leftover from an interrupted run (ours)
    try:
        shutil.copy2(src, tmp)
        if os.path.getsize(tmp) != os.path.getsize(src) or not filecmp.cmp(src, tmp, shallow=False):
            raise IOError("copy verification failed (bytes differ)")
        rename_no_overwrite(tmp, dst)
    except BaseException:
        if os.path.lexists(tmp):
            os.remove(tmp)
        raise
    if copy:
        return "copied"
    os.remove(src)                          # only after verified copy is in place
    return "moved"


def unique_path(path):
    if not os.path.lexists(path):
        return path
    root, ext = os.path.splitext(path)
    i = 2
    while os.path.lexists(f"{root}__{i}{ext}"):
        i += 1
    return f"{root}__{i}{ext}"


# ----------------------------------------------------------------------------
# scanning
# ----------------------------------------------------------------------------
def scan(src, skip_dir):
    stack = [src]
    skip = os.path.normcase(os.path.abspath(skip_dir))
    while stack:
        d = stack.pop()
        try:
            with os.scandir(d) as it:
                for e in it:
                    if e.is_dir(follow_symlinks=False):
                        if os.path.normcase(os.path.abspath(e.path)) != skip:
                            stack.append(e.path)
                    elif e.is_file(follow_symlinks=False) and e.name.lower().endswith(".pdf"):
                        yield e.path, e.stat().st_size
        except OSError as ex:
            print(f"[warn] cannot read {d}: {ex}")


def build_plan(src, dest):
    jobs, bad = [], []
    for p, size in scan(src, dest):
        m = NAME_RE.search(os.path.basename(p))
        if not m:
            bad.append((p, size))
            continue
        st, ac, part = m["state"].upper(), int(m["ac"]), int(m["part"])
        jobs.append({"source": p, "size": size, "state": st, "ac": ac, "part": part,
                     "target": os.path.join(dest, st, f"AC{ac}", f"{part}.pdf")})
    jobs.sort(key=lambda j: (j["state"], j["ac"], j["part"], j["source"]))
    return jobs, bad


# ----------------------------------------------------------------------------
# manifest (append + flush per row => crash safe)
# ----------------------------------------------------------------------------
class Manifest:
    def __init__(self, path):
        self.path = path
        self.f = open(path, "w", newline="", encoding="utf-8-sig")
        self.w = csv.DictWriter(self.f, fieldnames=FIELDS)
        self.w.writeheader()
        self.n = 0

    def log(self, **row):
        row.setdefault("time", time.strftime("%H:%M:%S"))
        self.w.writerow({k: row.get(k, "") for k in FIELDS})
        self.f.flush()
        self.n += 1
        if self.n % 500 == 0:
            os.fsync(self.f.fileno())

    def close(self):
        self.f.flush()
        os.fsync(self.f.fileno())
        self.f.close()


# ----------------------------------------------------------------------------
# run
# ----------------------------------------------------------------------------
def run(a):
    src, dest = os.path.abspath(a.source), os.path.abspath(a.dest)
    t0 = time.time()
    print(f"Scanning {src} ...")
    jobs, bad = build_plan(src, dest)
    total_bytes = sum(j["size"] for j in jobs)
    print(f"  {len(jobs):,} matching PDFs ({total_bytes / 1e9:.2f} GB), "
          f"{len(bad):,} unrecognised names  ({time.time() - t0:.1f}s)")

    # predicted collisions (by target name) for the preview
    seen, collide = set(), 0
    for j in jobs:
        k = os.path.normcase(j["target"])
        if k in seen or os.path.lexists(j["target"]):
            collide += 1
        seen.add(k)
    print(f"  {collide:,} files map to an already-used target (checked byte-by-byte when run)")
    for j in jobs[:3]:
        print(f"    {j['source']}\n      -> {j['target']}")

    if a.dry_run:
        if bad:
            print(f"\nUnrecognised names (first 20 of {len(bad):,}) - will NOT be touched:")
            for p, _ in bad[:20]:
                print("  ", p)
        print("\nDRY RUN - nothing changed. Re-run without --dry-run to apply.")
        return 0

    os.makedirs(dest, exist_ok=True)
    same_vol = same_volume(src, dest)
    mode = "copy" if a.copy else ("move (rename, same drive)" if same_vol else "move (copy+verify+delete)")
    if a.copy or not same_vol:
        free = shutil.disk_usage(dest).free
        if free < total_bytes * 1.05 + 1e9:
            print(f"ABORT: need ~{total_bytes / 1e9:.1f} GB on destination, only {free / 1e9:.1f} GB free.")
            return 2
    print(f"Mode: {mode}")

    logdir = os.path.join(dest, "_organize_logs")
    os.makedirs(logdir, exist_ok=True)
    man = Manifest(os.path.join(logdir, f"manifest_{time.strftime('%Y%m%d_%H%M%S')}.csv"))
    print(f"Manifest (written live): {man.path}\n")

    for p, size in bad:
        man.log(source=p, size=size, status="unrecognised_name", detail="not touched")

    made_dirs = set()
    stats = {}
    interrupted = False
    try:
        for i, j in enumerate(jobs, 1):
            s, t = j["source"], j["target"]
            base = dict(source=s, state=j["state"], ac=j["ac"], part=j["part"], size=j["size"])
            try:
                if not os.path.exists(s):
                    status, detail, t = "source_missing", "gone before processing", ""
                else:
                    if os.path.lexists(t):
                        if filecmp.cmp(s, t, shallow=False):
                            status, detail = "duplicate_identical", f"identical to {t}; source left in place"
                            t = ""
                        else:
                            t = unique_path(os.path.join(dest, "_conflicts", j["state"], f"AC{j['ac']}",
                                                         os.path.basename(s)))
                            status = None
                    else:
                        status = None
                    if status is None:
                        d = os.path.dirname(t)
                        if d not in made_dirs:
                            os.makedirs(d, exist_ok=True)
                            made_dirs.add(d)
                        action = safe_transfer(s, t, a.copy, same_vol)
                        status = "conflict_" + action if "_conflicts" in t else action
                        detail = ""
            except Exception as e:                       # file stays in SOURCE
                status, detail, t = "error", f"{type(e).__name__}: {e}", ""
            man.log(target=t, status=status, detail=detail, **base)
            stats[status] = stats.get(status, 0) + 1
            if i % 5000 == 0 or i == len(jobs):
                el = time.time() - t0
                print(f"  {i:,}/{len(jobs):,}  {i / max(el, 1e-9):.0f} files/s  {stats}")
    except KeyboardInterrupt:
        interrupted = True
        print("\nInterrupted by user - stopping safely (current file finished).")
    finally:
        man.close()

    print(f"\nResult: {stats}")
    ok = verify(man.path)
    if interrupted:
        print("Re-run the same command to continue; finished files will not be touched again.")
    print(f"Manifest: {man.path}")
    return 0 if ok and not stats.get("error") and not interrupted else 1


# ----------------------------------------------------------------------------
# verify / undo
# ----------------------------------------------------------------------------
def read_manifest(path):
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def verify(path):
    rows = read_manifest(path)
    moved = [r for r in rows if r["status"] in ("moved", "conflict_moved")]
    copied = [r for r in rows if r["status"] in ("copied", "conflict_copied")]
    problems = []
    for r in moved + copied:
        t = r["target"]
        if not os.path.isfile(t):
            problems.append(f"target missing: {t}")
        elif str(os.path.getsize(t)) != r["size"]:
            problems.append(f"size mismatch: {t}")
        if r["status"].endswith("moved") and os.path.exists(r["source"]):
            problems.append(f"source still exists after move: {r['source']}")
    errors = [r for r in rows if r["status"] == "error"]
    print(f"Verification: {len(moved):,} moved + {len(copied):,} copied checked, "
          f"{len(problems)} problems, {len(errors)} files left in source due to errors")
    for p in problems[:20]:
        print("  [PROBLEM]", p)
    for r in errors[:20]:
        print("  [error, still in source]", r["source"], "-", r["detail"])
    return not problems


def undo(path):
    rows = [r for r in read_manifest(path) if r["status"] in ("moved", "conflict_moved")]
    n = errs = 0
    for r in rows:
        s, t = r["source"], r["target"]
        try:
            os.makedirs(os.path.dirname(s), exist_ok=True)
            safe_transfer(t, s, copy=False, same_vol=same_volume(t, os.path.dirname(s)))
            n += 1
        except Exception as e:
            errs += 1
            print(f"[error] {t} -> {s}: {type(e).__name__}: {e}")
    print(f"Undo: restored {n:,} of {len(rows):,} files, {errs:,} errors (nothing overwritten).")
    return 1 if errs else 0


# ----------------------------------------------------------------------------
def clean(p):
    return p.strip().strip('"').rstrip(">").strip() if p else p


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", nargs="?")
    ap.add_argument("dest", nargs="?")
    ap.add_argument("--dry-run", action="store_true", help="show the plan, change nothing")
    ap.add_argument("--copy", action="store_true", help="copy instead of move")
    ap.add_argument("--undo", metavar="MANIFEST", help="move files back using a manifest")
    ap.add_argument("--verify", metavar="MANIFEST", help="re-check a previous run")
    a = ap.parse_args()
    if a.undo:
        return undo(clean(a.undo))
    if a.verify:
        return 0 if verify(clean(a.verify)) else 1
    if not (a.source and a.dest):
        ap.error("SOURCE and DEST are required")
    a.source, a.dest = clean(a.source), clean(a.dest)
    if re.search(r'[<>:"|?*]', os.path.splitdrive(a.dest)[1]):
        ap.error(f"DEST contains an invalid character: {a.dest!r}")
    if not os.path.isdir(a.source):
        ap.error(f"SOURCE folder not found: {a.source!r}")
    if os.path.normcase(os.path.abspath(a.source)) == os.path.normcase(os.path.abspath(a.dest)):
        ap.error("SOURCE and DEST must be different folders")
    return run(a)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
