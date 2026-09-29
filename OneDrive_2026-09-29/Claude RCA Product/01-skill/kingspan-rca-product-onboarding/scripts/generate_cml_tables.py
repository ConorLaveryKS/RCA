#!/usr/bin/env python3
"""
generate_cml_tables.py
----------------------
Emit pre-computed CML table constraints for decimal output fields, so the
constraint engine never has to search the decimal domain (the fix for the
"Reached limit of tree nodes for search" error).

Two generators are included:

  1. area_table()   – geometric area  Av = width_mm * length_mm / 1e6   (rounded 2dp)
  2. aa_tables()    – aerodynamic area Aa = Av * factor, one table per factor group

Edit the WIDTHS, LENGTHS, and GROUPS lists to match your product, then run:

    python generate_cml_tables.py > area_and_aa_constraints.cml

Paste the output into your CML model (sections 3 and 4 by convention).
"""

# ── Product dimensions (edit these) ──────────────────────────────────
WIDTHS = [426, 576, 726, 876, 1026, 1176, 1326, 1476, 1626, 1776,
          1926, 2000, 2126, 2326]
LENGTHS = [705, 838, 971, 1104, 1237, 1370, 1503, 1636, 1769, 1902,
           2035, 2168, 2301, 2434, 2567, 2700, 2833, 2966, 3099, 3232,
           3365, 3498]

# Variable names in your CML (edit to match your prefix)
WIDTH_VAR  = "ECO_ThroatWidth_mm"
LENGTH_VAR = "ECO_ThroatLength_mm"
AV_VAR     = "ECO_GeomAreaAv"
AA_VAR     = "ECO_AeroAreaAa"

# Aa factor groups: (factor, condition_string, comment)
# factor = Cv_panel * guard_factor, already multiplied out.
GROUPS = [
    (0.60, '(ECO_LouvreType == "A1B" || ECO_LouvreType == "A1X") && ECO_Guards == "X"',
     "Cv=0.60 A1B/A1X no guard"),
    (0.57, '(ECO_LouvreType == "A1B" || ECO_LouvreType == "A1X") && ECO_Guards == "BG"',
     "Cv=0.57 A1B/A1X bird guard"),
    (0.65, '(ECO_LouvreType == "A2B" || ECO_LouvreType == "A2X") && ECO_Guards == "X"',
     "Cv=0.65 A2 no guard"),
    # ... add the rest of your factor groups here
]


def av(w, l):
    return round((w * l) / 1_000_000, 2)


def area_table():
    print("    // Geometric area Av  (pre-computed, all size combinations)")
    print(f"    constraint(table({WIDTH_VAR}, {LENGTH_VAR}, {AV_VAR},")
    rows = [f"        {{{w}, {l}, {av(w, l):.2f}}}" for w in WIDTHS for l in LENGTHS]
    print(",\n".join(rows))
    print('        ), "Geometric area Av from dimensions");')
    print()


def aa_tables():
    for factor, cond, comment in GROUPS:
        print(f"    // Aerodynamic area Aa  —  {comment}")
        print(f"    constraint(")
        print(f"        ({cond}) ->")
        print(f"        table({WIDTH_VAR}, {LENGTH_VAR}, {AA_VAR},")
        rows = [f"        {{{w}, {l}, {round(av(w, l) * factor, 2):.2f}}}"
                for w in WIDTHS for l in LENGTHS]
        print(",\n".join(rows))
        print(f'        ),')
        print(f'        "Aa {comment}");')
        print()


if __name__ == "__main__":
    n = len(WIDTHS) * len(LENGTHS)
    print(f"// Auto-generated CML table constraints")
    print(f"// {len(WIDTHS)} widths x {len(LENGTHS)} lengths = {n} size combinations")
    print(f"// {len(GROUPS)} Aa factor groups -> {len(GROUPS) * n} Aa entries")
    print()
    area_table()
    aa_tables()
