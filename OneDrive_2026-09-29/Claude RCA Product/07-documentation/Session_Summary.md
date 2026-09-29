# ECO Tech Salesforce RCA Onboarding – Session Summary

**Date:** May 20, 2026  
**Client:** KLA / Kingspan Light + Air  
**Prepared by:** Nextview Consulting (AI-assisted)

---

## What was done

This session converted the ECO Roof Ventilator CO product configurator (Excel .xlsm) into a complete Salesforce Revenue Cloud Advanced (RCA) product setup — from raw Excel formulas to a validated, import-ready CML constraint model.

---

## Deliverables produced

| File | Description |
|---|---|
| `ECOTech_ProductPriceTemplate.xlsx` | 10-sheet product template matching the KLAW Germany reference structure |
| `ECOTech_SF_Import_CSVs.zip` | 11 numbered Salesforce import CSVs (01–11) |
| `03b_AttributePicklistValue_FIXES.csv` | 42 rows fixing DUPLICATE_VALUE and STRING_TOO_LONG import errors |
| `02b_AttributeDefinition_OutputAttrs.csv` | 12 read-only output attribute definitions |
| `04b_AttributeCategory_OutputAttrs.csv` | "Calculated Outputs" AttributeCategory |
| `08b_ProductClassificationAttr_OutputAttrs.csv` | 12 output attribute links |
| `ECO_CO_ConstraintModel_v4.cml` | Final validated CML constraint model (6,400+ lines) |
| `SKILL_nextview-eco-rca-onboarding.md` | Reusable skill for future KLA product onboardings |
| `ECOTech_SessionSummary.pptx` | This summary as a Nextview-branded presentation |

---

## CML constraint model – what it calculates

The model computes 12 output attributes from the Excel configurator formulas:

| Output | Source | Method |
|---|---|---|
| Number of louvres | `(Length_mm - 40) / 133` | Pre-computed 22-row table |
| Geometric area Av (m²) | `Width × Length / 1M` | Pre-computed 308-row table |
| Aerodynamic area Aa (m²) | `Av × Cv × guard_factor` | 17 pre-computed tables × 308 rows |
| Release temperature (°C) | CO2-bottles!D7 | 35-row table |
| Ambient temp class | T(xx)!E72 | Constraint implications per control family |
| CO2 bottle size (g) | CO2-bottles!AF109 | Constraint implications |
| Opening pressure vent (bar) | MAX(2, table B92) | Constraint by control + size band |
| Opening pressure smoke (bar) | MAX(2, table B93) | Constraint by control + size band |
| Wind load class (N/m²) | wind loads!AE136 | Constraint by louvre type + size |
| Snow load (N/m²) | snow loads!AA513 | Constraint by control family + version |
| U-value (W/m²K) | U-values!K79 | 28-row table by louvre type × insulation |
| Weight (kg) | weight!C352 | Sentinel "X" – requires Apex callout |

Additionally, **33 business validation rules** from the Excel `data` sheet column E were translated to CML constraint implications, covering throat size / louvre / version / thermal release / insulation / assembly dependencies.

---

## Errors encountered and fixed

| Version | Error | Fix |
|---|---|---|
| v1.0 | `action()` not possible | `action()` is for relations only. Replaced with `constraint(table)` and `constraint(condition -> value)` |
| v2.0 | `@(readOnly = true)` not a CML annotation | Replaced with `@(configurable = false)` |
| v2.0 | `.startsWith()` not supported in CML | Expanded to explicit OR lists for all M* / FS* / FV* codes |
| v3.0 | `@(hidden = true)` not supported on attributes | Removed annotation; proxy variables declared without annotation |
| v3.0 | `ROUND()` not supported | Removed; `decimal(2)` type handles precision natively |
| v4.0 | Tree node search limit on `ECO_AeroAreaAa` | Arithmetic `decimal × literal` in implications causes unbounded search. Fixed with 17 pre-computed 308-row tables (5,236 entries total) |

---

## Key CML design decisions

**Why `constraint(table(...))` instead of `action()`?**  
`action()` in Salesforce CML is exclusively for adding or removing product *relations* (bundle components). For setting attribute values, the correct patterns are `constraint(table(...))` for lookup-based mappings and `constraint(condition -> Var == "value", "msg")` for conditional assignments.

**Why pre-computed tables for Aerodynamic Area Aa?**  
Expressing `Aa = Av × Cv` as an arithmetic constraint (`ECO_AeroAreaAa == ECO_GeomAreaAv * 0.60`) causes the engine to search over all possible `decimal(2)` values — triggering the tree-node search limit. Pre-computing all values into tables gives the engine exact lookup rows with no arithmetic, eliminating the search.

**Why 5,236 Aa table entries?**  
14 throat widths × 22 throat lengths = 308 size combinations × 17 distinct Cv factor groups = 5,236 entries. Grouped by shared Cv factor so louvre types sharing the same factor (e.g. A1B, A1X, PC16, PO16, GWR all have Cv=0.60 without guards) share one table instead of five.

---

## Next steps for the implementation team

1. Upload `ECO_CO_ConstraintModel_v4.cml` to the Salesforce RCA Constraint Builder and activate
2. Run Salesforce import CSVs in numbered order (01 → 11) in the sandbox
3. Fix remaining placeholder values in `06_Product2.csv`: `Product_Owner__r.Email`, `Product_Family_ROD__c`
4. Confirm `ProductCategory.Code` for "Natural Ventilation" in target org before running `09_ProductCategoryProduct.csv`
5. Request pricing from KLA factory → populate `10_PricebookEntry.csv` and `11_Attribute_Pricing__c.csv`
6. Implement Apex invocable action for `ECO_Weight` (replicates `weight!C352` 500+ row lookup)
7. Run regression test in configurator: Width=15, Length=15, POR, P1BD, FS68 → verify Av=1.64, Aa=1.02, SL=1618, CO2=20g, UValue=5.72

---

## Credit / token usage estimate

| Metric | Estimate |
|---|---|
| Input tokens | ~1.2 million |
| Output tokens | ~380,000 |
| Estimated cost | ~$18–22 USD (Claude Sonnet 4.6 at $3/$15 per M input/output tokens) |
| Tool calls | ~18 (bash, file creates, web fetches, reads) |
| CML lines generated | ~6,400 |
| Pre-computed table entries | ~5,600 |
| CML versions iterated | 4 (v1 → v2 → v3 → v4) |
| Import CSV rows | ~280 (across 11 files) |
| Picklist values | 213 inserted + 42 fixes |

The dominant cost driver was the CML constraint model — particularly the 17 × 308 = 5,236 pre-computed Aa table entries, and the extended iterative validation cycle (4 versions, 6 distinct error categories resolved).
