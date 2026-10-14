# Airlite — re-import set "2" (AIR2)

A clean re-import of Airlite after the first import (`09-airlite-product-template`) partly failed.
Everything carries the suffix **2** so nothing clashes with records left in the org by the first attempt.

| File | What |
|---|---|
| `Airlite2_ProductPriceTemplate.xlsx` | Template (source of truth): product AIR2, 16 picklists / 112 values |
| `Airlite2_RCA_Import_Inspector.xlsx` | Salesforce Inspector Reloaded import, one tab per step + Import Steps tab with queries |
| `Airlite2_ConstraintModel.cml` | Constraint model: 17 constraints, 20 messages, Av table (279) + Aa tables (4 × 279) |

## Why the first import was missing data
- 15 picklist attributes failed on step 02 (and so did their steps 03 / 07 / 08).
- RAL Colour, Snow Load, Wind Load and Aerodynamic Area Aa failed because their **names** matched ECO attributes
  (the `AIR ` name prefix was added afterwards).

## What changed in "2"
- Suffix 2 on picklist codes (`AIR Version2`), attribute names / codes (`AIR Version2` / `AIR_Version2`),
  value codes (`AIR_Version2_1E`), classification (`AIR_AIR2_ProductClassification`) and product (`AIR2`).
- Width is a picklist in 50 mm steps (500–2000) so Av and Aa can be exact pre-computed tables
  (Aa = configurator Aerodynamik sheet incl. its width interpolation). Other widths = special request.
- Picklist defaults are set only on step 08 (ProductClassificationAttr), not on step 02.

Rebuild:
```
python ../09-airlite-product-template/build_airlite_template.py ../02-product-template/ECOTech_ProductPriceTemplate.xlsx "Airlite Tech. Configurator(1).xlsm" Airlite2_ProductPriceTemplate.xlsx 2
python ../01-skill/kingspan-rca-product-onboarding/scripts/generate_import_csvs.py Airlite2_ProductPriceTemplate.xlsx <csv_dir> AIR
python ../01-skill/kingspan-rca-product-onboarding/scripts/build_inspector_import.py <csv_dir> Airlite2_RCA_Import_Inspector.xlsx AIR --cml Airlite2_ConstraintModel.cml --test "AIR2, Width 1200, Length 2100, Installation A0, Wind baffle WM → Av 2.52 m², Aa 1.42 m²"
python build_airlite_cml.py Airlite2_ProductPriceTemplate.xlsx "Airlite Tech. Configurator(1).xlsm" Airlite2_ConstraintModel.cml 2
```
