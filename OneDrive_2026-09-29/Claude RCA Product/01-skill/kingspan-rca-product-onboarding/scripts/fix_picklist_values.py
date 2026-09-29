#!/usr/bin/env python3
"""
fix_picklist_values.py
----------------------
Fix the two recurring AttributePicklistValue import errors:

  DUPLICATE_VALUE   – Code must be globally unique across the whole org
  STRING_TOO_LONG   – Name must be <= 80 characters

Usage:
    python fix_picklist_values.py INPUT_APV.csv OUTPUT_FIXES.csv [--collide CODES.txt]

    --collide CODES.txt   optional list of codes (one per line) that already exist
                          in the org, or that the import rejected with
                          DUPLICATE_VALUE. These are prefixed as well.

A Code is prefixed when it appears in more than one picklist in the input file,
or when it is listed in --collide. Numeric size codes get a bare prefix
("10" -> "TW10"); everything else gets PREFIX_ + upper-case code
("X" -> "RL_X", "No" -> "LE_NO"). Name / DisplayValue over 80 characters are
shortened. Only the changed rows are written to OUTPUT_FIXES.csv.

The picklist-facing `Value` and `Abbreviation` stay the original code (used in
the product code string); only the internal `Code` key is prefixed. Adjust
ABBREV below to control the prefixes.
"""
import sys, csv, re
from collections import defaultdict

MAX_NAME = 80

# Explicit abbreviations for known picklists (extend as needed).
# Key = a substring of the picklist code; first match wins, so list more
# specific keys first ("burglar guards" before "guards").
ABBREV = {
    "throat width": "TW", "throat length": "TL",
    "removable louvre": "RL", "insulated adapter": "IA", "insulation": "INS",
    "burglar": "BGA", "guards": "GD", "finish": "FIN", "coverage": "COV",
    "upstand": "UD", "assembly": "AS", "lifting": "LE",
    "width": "W", "length": "L", "panel": "PN", "control": "CT",
    "base": "BS", "release": "RL2", "version": "VR",
}


def prefix_for(picklist_code):
    pc = (picklist_code or "").lower()
    for key, ab in ABBREV.items():
        if key in pc:
            return ab
    # fallback: initials of the picklist words, skipping the product prefix
    words = re.sub(r"[^A-Za-z ]", "", picklist_code or "").split()[1:]
    return "".join(w[0] for w in words[:3]).upper() or "PL"


def prefixed(pfx, code):
    return f"{pfx}{code}" if code.isdigit() else f"{pfx}_{code.upper()}"


# Common abbreviations to shorten long Names
SHORTEN = [
    ("pneumatic cylinders", "pneum. cyl."),
    ("pneumatic cylinder", "pneum. cyl."),
    ("electric motors", "elec. motors"),
    ("electric motor", "elec. motor"),
    ("pressure", "pres."),
    ("assisting spring", "assist. spring"),
    ("with sprinkler shield", "w/ sprinkler shield"),
    ("positions", "pos."),
]


def shorten(text):
    t = text
    for long, short in SHORTEN:
        if len(t) <= MAX_NAME:
            break
        t = t.replace(long, short)
    return t[:MAX_NAME]


def main():
    args = sys.argv[1:]
    collide = set()
    if "--collide" in args:
        i = args.index("--collide")
        with open(args[i + 1], encoding="utf-8") as f:
            collide = {line.strip() for line in f if line.strip()}
        del args[i:i + 2]
    if len(args) < 2:
        print(__doc__)
        sys.exit(1)
    src, dst = args[0], args[1]

    with open(src, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    header = rows[0]

    def col(*names):
        for n in names:
            if n in header:
                return header.index(n)
        sys.exit(f"Column {names} not found in header {header}")

    ci_pl = col("Picklist.Code", "Picklist:AttributePicklist:Code")
    ci_code = col("Code")
    ci_name = col("Name")
    ci_disp = col("DisplayValue")

    picklists = defaultdict(set)
    for row in rows[1:]:
        picklists[row[ci_code]].add(row[ci_pl])
    clash = {c for c, pls in picklists.items() if len(pls) > 1} | collide

    fixes = [header]
    n_dupe = n_long = 0
    for row in rows[1:]:
        new = list(row)
        if row[ci_code] in clash:
            new[ci_code] = prefixed(prefix_for(row[ci_pl]), row[ci_code])
            n_dupe += 1
        for ci in (ci_name, ci_disp):
            if len(row[ci]) > MAX_NAME:
                new[ci] = shorten(row[ci])
                n_long += 1
        if new != row:
            fixes.append(new)

    with open(dst, "w", newline="", encoding="utf-8") as f:
        csv.writer(f, lineterminator="\r\n").writerows(fixes)

    print(f"Wrote {len(fixes)-1} fixed rows to {dst}")
    print(f"  Code prefixes applied : {n_dupe}")
    print(f"  Names shortened       : {n_long}")


if __name__ == "__main__":
    main()
