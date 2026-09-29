---
name: kingspan-rca-product-onboarding
description: >
  Convert a Kingspan / KLA (Kingspan Light + Air) product configurator Excel
  file into a complete, import-ready Salesforce Revenue Cloud Advanced (RCA)
  product setup: an Excel Product Price Template, numbered Salesforce import
  CSVs, and a syntax-validated CML (Constraint Modeling Language) constraint
  model. Use this skill whenever the user wants to onboard, load, or update a
  configurable Kingspan/KLA product in Salesforce RCA, convert a product
  configurator or Excel formula workbook into Salesforce, generate
  AttributePicklist / AttributePicklistValue / AttributeDefinition /
  ProductClassificationAttr import files, or author a CML constraint model for
  a configurable product. Also triggers on: ECO Roof Ventilator, smoke/heat
  exhaust ventilator, louvre configurator, "load this product into RCA",
  "recreate this Excel configurator in Salesforce", constraint model, CML
  script, or any of the calculated output fields (geometric area, aerodynamic
  area, snow load, wind load, CO2 bottle, U-value, opening pressure, ambient
  temperature class).
---

# Kingspan → Salesforce RCA Product Onboarding

A repeatable, end-to-end process for turning a Kingspan/KLA product configurator
(Excel `.xlsm`) into a fully configured Salesforce Revenue Cloud Advanced product.
Proven on the ECO Roof Ventilator (CO); designed to generalize to any configurable
Kingspan product.

## What this skill produces

1. **Product Price Template** (`.xlsx`) — 10-sheet workbook matching the KLAW Germany reference format.
2. **Salesforce import CSVs** — numbered `01`–`11` in dependency order, ready to import.
3. **AttributePicklistValue FIXES CSV** — retry file for the two common import errors.
4. **CML constraint model** (`.cml`) — validated against the Salesforce CML User Guide.

Work in a scratch dir, copy final files to the outputs dir, and present them.

## Folder contents

```
kingspan-rca-product-onboarding/
├── SKILL.md                        ← you are here (entry point)
├── scripts/
│   ├── generate_import_csvs.py     ← Product Price Template .xlsx → 11 numbered CSVs
│   ├── fix_picklist_values.py      ← dedupe Codes + truncate Names (import-error fixer)
│   └── generate_cml_tables.py      ← pre-computed area/Aa table constraints (avoids search limit)
├── references/
│   ├── cml_syntax_rules.md         ← the CML gotchas — READ BEFORE WRITING CML
│   ├── csv_import_reference.md     ← the 11 objects, order, columns, keys
│   └── output_field_patterns.md    ← which CML construct to use for each calculated field
├── assets/
│   └── AirCore_Configurator_DEMO.xlsx  ← anonymized demo configurator (for practice/demos)
└── examples/
    └── ECO_CO_ConstraintModel_example.cml  ← a full worked constraint model
```

## The 6-phase workflow

### Phase 1 — Read the Excel configurator
```python
import openpyxl
wb = openpyxl.load_workbook('PRODUCT_Configurator.xlsm', data_only=False)  # formulas
print(wb.sheetnames)
```
Map the sheets. Read `Input` column G (validation flags) and the backing `data`
column E (the rules) **fully** — these become the CML business rules.

Typical Kingspan configurator sheets: `Input`, `data`, `text` (labels + code→mm),
`aerodynamic` (Av, Aa), `snow loads`, `wind loads`, `CO2-bottles`, `weight`,
`U-values`, `T(xx)`, `pressure dependend snow loads`.

### Phase 2 — Build the Product Price Template
Ten sheets: `Product`, `AttributeDefinition`, `Product and Attributes`,
`AttributePicklist`, `AttributePicklistValue`, `Product to Product Constraints`,
`Product Internal Constraint`, `Configurable Product Pricing`,
`Non Configurable Product Pricing`, `Bundle Products`.

Naming convention (the CSVs and CML depend on it):
- **Attribute Code / DeveloperName:** `ECO_ThroatWidth` (prefix + PascalCase, no spaces)
- **Picklist Code:** `ECO Throat Width` (prefix + spaces)

### Phase 3 — Generate the Salesforce import CSVs
Run `scripts/generate_import_csvs.py`. See `references/csv_import_reference.md` for
the full object list and the strict import order. Fix any import errors with
`scripts/fix_picklist_values.py`.

### Phase 4 — Write the CML constraint model
**Read `references/cml_syntax_rules.md` first** — CML is not a general expression
language and several intuitive patterns fail (`action()` can't set values,
`@(readOnly)`/`@(hidden)` unsupported, no string methods, no `ROUND()`, decimal
arithmetic blows the search-tree budget). Use `examples/ECO_CO_ConstraintModel_example.cml`
as a reference implementation.

### Phase 5 — Add the calculated output attributes
See `references/output_field_patterns.md` for the exact CML construct per field.
For area/Aa fields, use `scripts/generate_cml_tables.py` to emit pre-computed
table constraints (this is the fix for the tree-node search limit).

### Phase 6 — Add the Column G validations
Every rule in the Excel `data` sheet column E becomes a
`constraint(condition -> restriction, "message")`.

## Regression test after every build
Load one known configuration into a sandbox and confirm the outputs. For ECO CO:
> Width `15` · Length `15` · Louvre `POR` · Control `P1BD` · Release `FS68`
Expected: Av = 1.64 m², Aa ≈ 1.02 m², louvres = 9, SL = 1618 N/m², CO2 = 20 g, U = 5.72.

## Definition of done
- [ ] All 11 CSVs import cleanly in numbered order in a sandbox
- [ ] `AttributePicklistValue` FIXES applied (no DUPLICATE_VALUE / STRING_TOO_LONG)
- [ ] CML activates in the Constraint Builder with no syntax errors
- [ ] Reference configuration produces the expected output values
- [ ] `Product_Owner__r.Email`, `Product_Family`, `Pricebook2.Name` filled in
- [ ] Pricing received from factory and loaded
- [ ] Weight Apex invocable action implemented (if the product has a weight table)
- [ ] CML committed to version control
