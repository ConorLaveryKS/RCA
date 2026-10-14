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
Finished templates go to `C:\Users\conor.lavery\OneDrive - Kingspan\Commercial\RCA\Claud\Configurator templates`
(save there directly when running locally; from a cloud session commit to the repo, send the file, and ask the user to save it there; see the repo `CLAUDE.md`).

## Folder contents

```
kingspan-rca-product-onboarding/
├── SKILL.md                        ← you are here (entry point)
├── scripts/
│   ├── generate_import_csvs.py     ← Product Price Template .xlsx → 11 numbered CSVs
│   ├── fix_picklist_values.py      ← dedupe Codes + truncate Names (import-error fixer)
│   ├── build_inspector_import.py   ← CSVs → Salesforce Inspector Reloaded import workbook (1 tab per step)
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
Run `scripts/generate_import_csvs.py TEMPLATE.xlsx OUT_DIR PREFIX`. Add
`--outputs OUTPUT_ATTRS.csv` (calculated attributes in AttributeDefinition-sheet layout)
to also emit 02b/04b/08b. See `references/csv_import_reference.md` for
the full object list and the strict import order. Fix any import errors with
`scripts/fix_picklist_values.py` (use `--collide` for codes the org rejected).
Names and codes are made **org-unique by default** (see the rule below); `--legacy`
reproduces the ECO CO files byte-for-byte as they were first imported.

> **Rule – org-unique keys (always, for every new product).** Salesforce enforces these
> across the WHOLE org, and earlier products already use the plain values:
> - `AttributeDefinition.Name` → `<PREFIX> <Name>` (e.g. `AIR Version`); `Label` stays clean.
> - `AttributePicklistValue.Code` → `<PREFIX>_<Picklist>_<Value>` (e.g. `AIR_Coating_S`); `Value` stays the real code.
> - `AttributePicklist.Code` and `AttributeDefinition.Code` already carry the prefix (`AIR Coating`, `AIR_Coating`).
> - `AttributeCategory` codes are shared between products: a DUPLICATE_VALUE there is expected and fine.
> Skipping this caused DUPLICATE_VALUE on Airlite (`Name duplicates value on record …`) and a
> knock-on `Foreign key external ID … not found` on every later step that referenced those attributes.

For import with **Salesforce Inspector Reloaded**, build the workbook with
`scripts/build_inspector_import.py <csv_dir> <Product>_RCA_Import_Inspector.xlsx PREFIX` and follow
"Importing with Salesforce Inspector Reloaded" in `references/csv_import_reference.md`.

> **When giving the user import instructions, list the steps in sequence and, for each step that
> needs an Id or org value, give the exact Data Export (SOQL) query to find it**: owner User Id and
> product family (step 06), attribute check (07/08), Product2 Id and ProductCategory Id **with catalog**
> (09), Pricebook2 Id (10). The queries are in `references/csv_import_reference.md` → "Data Export
> queries to run before each step", and on the workbook's Import Steps tab. Also remind the user to
> change the Inspector **Object** before every run.
> Number the steps exactly as the workbook tabs (01–11). Steps that are not imports get a label,
> never a number: **CML** (constraint model, Constraint Builder) and **TEST** (quote) come after
> tab 09 and before the pricing tabs 10–11 (which wait for factory prices).

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

### Phase 7 — Import, load the constraint model, test (sequence)
Build the import workbook with the model file and the reference test so the Import Steps tab
carries the full sequence:
```
python scripts/build_inspector_import.py <csv_dir> <Product>_RCA_Import_Inspector.xlsx PREFIX \
    --cml <Product>_ConstraintModel.cml --test "W 1200 / L 2100 / A0 → Av 2.52, Aa 1.42"
```
The sequence (always give it to the user in this order and with these labels):

| Step | What | Where |
|---|---|---|
| **01 – 09** | Data imports, one tab per object (queries for Ids on the Import Steps tab) | Inspector → Data Import |
| **CML-1** | Check every attribute in the .cml exists and is linked (tabs 02 + 08 all succeeded) | Data Export |
| **CML-2** | Create the constraint model + version | Constraint Builder |
| **CML-3** | Paste the whole .cml into the CML editor, save | Constraint Builder |
| **CML-4** | Associate the CML `type` with the product | Constraint Builder |
| **CML-5** | Validate; fix errors via `references/cml_syntax_rules.md` | Constraint Builder |
| **CML-6** | Activate | Constraint Builder |
| **TEST** | Quote with the reference configuration, check outputs | Quote |
| **10 – 11** | Pricing – only once factory prices exist | Inspector → Data Import |

Never number CML / TEST like tabs: tab 10 is PricebookEntry.

## Regression test after every build
Load one known configuration into a sandbox and confirm the outputs. For ECO CO:
> Width `15` · Length `15` · Louvre `POR` · Control `P1BD` · Release `FS68`
Expected: Av = 1.64 m², Aa ≈ 1.02 m², louvres = 9, SL = 1618 N/m², CO2 = 20 g, U = 5.72.

## Definition of done
- [ ] All 11 CSVs import cleanly in numbered order in a sandbox
- [ ] `AttributePicklistValue` FIXES applied (no DUPLICATE_VALUE / STRING_TOO_LONG)
- [ ] CML-1…CML-6 done: attributes checked, model created, pasted, linked to the product, validated, activated
- [ ] Reference configuration produces the expected output values
- [ ] `Product_Owner__r.Email`, `Product_Family`, `Pricebook2.Name` filled in
- [ ] Pricing received from factory and loaded
- [ ] Weight Apex invocable action implemented (if the product has a weight table)
- [ ] CML committed to version control
