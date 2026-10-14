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

## The 7-phase workflow

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
| **CML-4** | **Required before activation**: association CML `type` → product: Inspector Insert **`ExpressionSetConstraintObj`** (ExpressionSetId 9QL…, ReferenceObjectId = product, ConstraintModelTag = CML type, ConstraintModelTagType = Type) – workbook tab "CML-4 ExpressionSetConstrObj". Not the "New Type" button | Constraint Builder |
| **CML-5** | Validate; fix errors via `references/cml_syntax_rules.md` | Constraint Builder |
| **CML-6** | Activate | Constraint Builder |
| **10** | PricebookEntry in **both org price books** – `01s0Y000004WPyFQAW` (standard, first) and `01sAP000003qiEXYAY` (second) – with **UnitPrice 0** placeholder – standard price book first, then the quote's price book, per currency. **Required before TEST**: without it the product is not in the quote's product picker | Inspector → Data Import |
| **TEST** | Quote with the **complete** reference configuration (every input, taken from the configurator's own example), check outputs and that no warnings appear | Quote |
| **11** | Attribute pricing – once factory prices exist (then also update the 10 UnitPrice) | Inspector → Data Import |

Never number CML / TEST like tabs: tab 10 is PricebookEntry. Order: 01–09 → CML-1…6 → 10 → TEST → 11.

## Troubleshooting – error → fix (from the Airlite import, Oct 2026)

Look the message up here first. Fix the **earliest** failing step before re-running later ones:
a failure in one step makes every later step that looks it up fail too.

| Step | Error message (Inspector `__Errors`) | Cause | Fix |
|---|---|---|---|
| 02 | `DUPLICATE_VALUE: duplicate value found: Name duplicates value on record with id …` | `AttributeDefinition.Name` is unique org-wide; earlier products use plain names (Version, Insulation, RAL Colour, Snow Load, Wind Load, Aerodynamic Area Aa …) | `<PREFIX> <Name>` – default in `generate_import_csvs.py` |
| 03 | `DUPLICATE_VALUE` on Code | Value `Code` is unique org-wide (S, X, FS68, N5 …) | `<PREFIX>_<Picklist>_<Value>` – default in the generator |
| 01 re-run | `DUPLICATE_VALUE … Name` | The picklists already exist from the first run | Skip the tab; nothing to fix |
| 04 | `DUPLICATE_VALUE` | Category already created by another product | Expected – categories are shared |
| 06 | `'01/01/2026' is not a valid value for the type xsd:dateTime` | `AvailabilityDate` is a dateTime; Excel reformatted it | Type `2026-01-01T00:00:00.000Z` in the Text cell |
| 06 | `''2026-…' is not a valid value for the type xsd:dateTime` | Apostrophe typed into a Text cell is kept | Remove the apostrophe |
| 06 | `MALFORMED_ID: Product Owner: id value of incorrect type: <alias>` | Lookup needs the User Id | `SELECT Id FROM User WHERE Alias = '…'` (005…) |
| 07 / 08 | `INVALID_FIELD: Foreign key external ID: <code> not found for field Code in entity AttributeDefinition` | That attribute was never created in step 02 | Fix step 02 first, then re-run only the failed rows |
| 08 | `INVALID_OR_NULL_FOR_RESTRICTED_PICKLIST: Display Type: bad value … Dropdown` | `DisplayType` is a restricted picklist | Leave blank (builder does this) |
| 08 | `INVALID_INPUT: Enter a default value for an attribute that's visible and read-only` | Calculated outputs are read-only + visible | DefaultValue `0` (Number) / `X` (Text) (builder does this) |
| 09 | `INVALID_FIELD: No such column 'ProductId' on entity 'ProductClassificationAttr'` | Inspector kept the previous object selected | Set Object = ProductCategoryProduct |
| 09 | `MALFORMED_ID: Category ID: id value of incorrect type: <name>` | Name given instead of Id | `SELECT Id, Name, Catalog.Name FROM ProductCategory WHERE Name LIKE '%…%'` (0ZG…) |
| CML-6 | `We couldn't find any associations for constraint version 9QL…` | No `ExpressionSetConstraintObj` linking the CML type to the product | Insert one (tab "CML-4 ExpressionSetConstrObj"). The `9QL…` Id in the message **is** the ExpressionSetId of the model |
| CML-4 | Builder only offers "New Type" (Name + Parent type) / "Add Labels" | Neither creates the product association | Use the Inspector insert above |
| Quote | Product not found in the quote's product picker | No active PricebookEntry in the quote's price book / currency | Tab 10 with UnitPrice 0 (standard price book first); check product IsActive and the tab 09 category |
| Quote | Many attributes / options missing on the configured product | Steps 02/03/07/08 partly failed | Check with the CML-1 query; if many records are half-created, rebuild with a suffix (re-import set, e.g. `AIR2`) |

**Delivering files to the user**
- Re-sending a file with the same name: the user may open the old download. Use a new
  name for every re-send (`…_v2.xlsx`, `…_v3.xlsx`).
- Excel sheet names: max 31 characters (the builder asserts this).
- Verify a calculated output against the configurator's own example before handing over
  (Airlite: 1200 × 2100, A0 → Av 2.52, Aa 1.42).

**Worked examples**
- ECO CO: `../../03-salesforce-import-csvs/`, `../../04-cml-constraint-model/`, `../../08-eco-import-csvs-regenerated/`
- Airlite (first attempt, partly failed): `../../09-airlite-product-template/` – template builder `build_airlite_template.py`
- Airlite re-import set AIR2 (all fixes applied): `../../10-airlite2-product-template/` – includes
  `build_airlite_cml.py`, a CML generator (picklists, Av/Aa tables with interpolation, rules,
  numeric messages, flange limits) to copy for the next product.

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
