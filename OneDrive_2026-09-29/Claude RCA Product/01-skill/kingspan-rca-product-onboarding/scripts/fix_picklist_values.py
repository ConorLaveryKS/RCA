#!/usr/bin/env python3
"""
fix_picklist_values.py
----------------------
Fix the two recurring AttributePicklistValue import errors:

  DUPLICATE_VALUE   – Code must be globally unique across the whole org
  STRING_TOO_LONG   – Name must be <= 80 characters

Usage:
    python fix_picklist_values.py INPUT_APV.csv OUTPUT_FIXES.csv

Reads an AttributePicklistValue CSV, applies:
  1. A per-picklist Code prefix so codes are globally unique
     (derived from the picklist code initials, e.g. "ECO Throat Width" -> "TW").
  2. Name / DisplayValue truncation to 80 chars.
Writes only the rows that CHANGED to OUTPUT_FIXES.csv, ready for re-import.

The picklist-facing `Value` field is preserved unchanged; only the internal
`Code` key is prefixed. Adjust ABBREV below to control the prefixes.
"""
import sys, csv, re

MAX_NAME = 80

# Explicit abbreviations for known picklists (extend as needed).
# Key = a substring that appears in the picklist code; value = the Code prefix.
ABBREV = {
    "throat width": "TW", "throat length": "TL",
    "removable louvre": "RL", "insulation": "INS", "insulated adapter": "IA",
    "guards": "GD", "burglar": "BGA", "finish": "FIN", "coverage": "COV",
    "upstand": "UD", "assembly": "AS", "lifting": "LE",
    "width": "W", "length": "L", "panel": "PN", "control": "CT",
    "base": "BS", "release": "RL2", "version": "VR",
}


def prefix_for(picklist_code):
    pc = (picklist_code or "").lower()
    for key, ab in ABBREV.items():
        if key in pc:
            return ab
    # fallback: initials of the picklist code words
    words = re.sub(r"[^A-Za-z ]", "", picklist_code or "").split()
    return "".join(w[0] for w in words[:3]).upper() or "PL"


# Common abbreviations to shorten long Names
SHORTEN = [
    ("pneumatic cylinder", "pneum. cyl."),
    ("pneumatic cylinders", "pneum. cyl."),
    ("electric motors", "elec. motors"),
    ("electric motor", "elec. motor"),
    ("pressure", "pres."),
    ("spring return", "spring ret."),
    ("with sprinkler shield", "w/ sprinkler shield"),
    ("assisting spring", "assist. spring"),
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
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]

    with open(src, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)

    header = rows[0]
    # locate columns (tolerant of header naming)
    def col(*names):
        for i, h in enumerate(header):
            for n in names:
                if n.lower() in h.lower():
                    return i
        return None

    ci_pl   = col("Picklist")
    ci_code = col("Code")
    ci_name = col("Name")
    ci_disp = col("DisplayValue")
    ci_val  = col("Value")

    fixes = [header]
    n_dupe, n_long = 0, 0

    for row in rows[1:]:
        changed = False
        new = list(row)

        # 1. prefix Code for global uniqueness
        if ci_pl is not None and ci_code is not None:
            pl = row[ci_pl]; code = row[ci_code]
            pfx = prefix_for(pl)
            if code and not code.startswith(pfx + "_") and not code.startswith(pfx):
                new[ci_code] = f"{pfx}_{code}"
                # preserve original in Value if present and empty-safe
                if ci_val is not None and not row[ci_val]:
                    new[ci_val] = code
                changed = True
                n_dupe += 1

        # 2. truncate Name / DisplayValue to 80
        for ci in [ci_name, ci_disp]:
            if ci is not None and row[ci] and len(row[ci]) > MAX_NAME:
                new[ci] = shorten(row[ci])
                changed = True
                n_long += 1

        if changed:
            fixes.append(new)

    with open(dst, "w", newline="", encoding="utf-8") as f:
        csv.writer(f, quoting=csv.QUOTE_ALL).writerows(fixes)

    print(f"Wrote {len(fixes)-1} fixed rows to {dst}")
    print(f"  Code prefixes applied : {n_dupe}")
    print(f"  Names truncated       : {n_long}")


if __name__ == "__main__":
    main()
