# Salesforce import CSV reference

Import in this exact order — each object depends on the ones above it.

| # | File | Object | Notes |
|---|---|---|---|
| 01 | AttributePicklist | `AttributePicklist` | One per configurable attribute |
| 02 | AttributeDefinition | `AttributeDefinition` | DataType = Picklist / Number / Text |
| 03 | AttributePicklistValue | `AttributePicklistValue` | **Code globally unique** (see below) |
| 04 | AttributeCategory | `AttributeCategory` | Logical UI groups |
| 05 | ProductClassification | `ProductClassification` | One per product type |
| 06 | Product2 | `Product2` | `RCA_Product__c=true`, `Configurable_Product__c=true` |
| 07 | AttributeCategoryAttribute | `AttributeCategoryAttribute` | Category → attribute links |
| 08 | ProductClassificationAttr | `ProductClassificationAttr` | Classification → attribute links |
| 09 | ProductCategoryProduct | `ProductCategoryProduct` | Product → catalog category |
| 10 | PricebookEntry | `PricebookEntry` | Placeholder until factory pricing |
| 11 | Attribute_Pricing__c | `Attribute_Pricing__c` | Attribute-level pricing |

Write CSVs UTF-8, minimal quoting, CRLF line endings (this is what
`generate_import_csvs.py` produces and what imported cleanly for ECO). Column headers
must match the field names in the org's query set; ECO used dot notation for
lookups (`Picklist.Code`). The ECO FIXES file used `Picklist:AttributePicklist:Code`
and also imported, so either form works.

## Key column sets

**01 AttributePicklist:** `Name, Code, Description, Status, DataType, CurrencyIsoCode`

**02 AttributeDefinition:** `Name, CurrencyIsoCode, UnitOfMeasureId, Label, Description,
DataType, IsActive, IsRequired, DefaultValue, SourceSystemIdentifier, ValueDescription,
Picklist.Code, DefaultHelpText, Code, DeveloperName` (02b output attributes use the same columns)

**03 AttributePicklistValue:** `Name, CurrencyIsoCode, Picklist.Code,
Abbreviation, Status, Code, DisplayValue, IsDefault, Sequence, Value`

**06 Product2:** `Name, IsActive, DisplayUrl, ConfigureDuringSale, Product_Family_ROD__c,
ProductCode, Type, RCA_Product__c, Internal_Product_Description__c, Description,
Configurable_Product__c, AvailabilityDate, Product_Owner__r.Email, BasedOn.Code`
(dates in `YYYY-MM-DD`)

**08 ProductClassificationAttr:** `Name, AttributeCategory.Code, IsPriceImpacting,
IsHidden, IsReadOnly, IsRequired, DisplayType, AttributeDefinition.Code, DefaultValue,
ProductClassification.Code, Sequence`

## Common import errors (AttributePicklistValue)

### DUPLICATE_VALUE
`Code` is unique across the **entire org**, not per picklist. Two collisions recur:
- **Numeric size codes** (`10`, `13`, `15`…) clash between Width and Length →
  prefix: `TW10`, `TL10`. Keep the real value in the `Value` field.
- **Single-character codes** (`X`, `S`, `R`, `I`, `No`, `Yes`) clash across
  attributes → prefix with the attribute abbreviation: `RL_X`, `INS_X`, `GD_X`,
  `BGA_X`, `FIN_X`, `COV_X`, `IA_X`, `UD_X`, `AS_S`, `LE_NO`, `LE_YES`.

The picklist-facing `Value` stays the original code (used in the product code
string); only the internal `Code` key gets prefixed.

### STRING_TOO_LONG
`Name` is max **80 characters**. Abbreviate long descriptions: `pneumatic cylinder`
→ `pneum. cyl.`, `pressure` → `pres.`, `electric motor` → `elec. motor`. Do not
touch `Code` or `Value`.

Produce a separate `*_FIXES.csv` of only the corrected failed rows for re-import.
Use `scripts/fix_picklist_values.py`. It prefixes codes that appear in more than one
picklist in the file. Codes the org rejects because they already exist elsewhere
can't be seen from the file, so pass those with `--collide codes.txt` (one per line).

## Manual fields to complete before go-live
- `06_Product2.csv` → `Product_Owner__r.Email`, `Product_Family_ROD__c`
- `09_ProductCategoryProduct.csv` → confirm `ProductCategory.Code` in the target org
- `10_PricebookEntry.csv` → `Pricebook2.Name`, `UnitPrice` (from factory)
- `08_ProductClassificationAttr.csv` → set `IsPriceImpacting=true` on price-driving attrs

## Org-unique keys – apply to every new product

`generate_import_csvs.py` does this by default (use `--legacy` only to reproduce ECO):

| Field | Unique across | Format | Example |
|---|---|---|---|
| `AttributePicklist.Code` | org | `<PREFIX> <Picklist>` | `AIR Coating` |
| `AttributeDefinition.Name` | org | `<PREFIX> <Name>` (Label unchanged) | `AIR Version` |
| `AttributeDefinition.Code` | org | `<PREFIX>_<Name>` | `AIR_Version` |
| `AttributePicklistValue.Code` | org | `<PREFIX>_<Picklist>_<Value>` (Value unchanged) | `AIR_Coating_S` |
| `AttributeCategory.Code` | org, **shared** | plain | `Device_Configuration` – reuse; duplicate error is fine |

## Importing with Salesforce Inspector Reloaded

Build the workbook: `scripts/build_inspector_import.py <csv_dir> <Product>_RCA_Import_Inspector.xlsx PREFIX`.
One object per run – copy one tab (with header), Data Import → object, Insert, paste, check mapping, Run;
re-run only failed rows before moving to the next tab.

Lessons from the Airlite import:
- **Lookup headers** use `Field:Object:ExternalId`, e.g. `Picklist:AttributePicklist:Code`,
  `AttributeDefinition:AttributeDefinition:Code`, `BasedOn:ProductClassification:Code`.
- **`Product2.AvailabilityDate` is a dateTime** in this org: `2026-01-01T00:00:00.000Z`. Excel turns
  dates into `01/01/2026` (rejected); the workbook cells are Text, so type the value **without** an
  apostrophe (in a Text cell the apostrophe is kept and also rejected).
- **`Product_Owner__c` needs the User Id** (starts `005`), not an alias or e-mail:
  `SELECT Id FROM User WHERE Alias = '...'` (MALFORMED_ID otherwise).
- **`Product_Family_ROD__c`**: optional; use an existing value
  (`SELECT Product_Family_ROD__c, COUNT(Id) FROM Product2 GROUP BY Product_Family_ROD__c`).
- **`ProductCategoryProduct`** needs `ProductId` (returned by the Product2 run) and `ProductCategoryId`.
- If one step fails, every later step that looks it up fails with
  `Foreign key external ID … not found` – fix the earliest failing step first, then re-run the rest.
- Picklist-attribute `DefaultValue`s point at values created in step 03; if step 02 rejects them,
  clear the default and rely on `ProductClassificationAttr.DefaultValue` (step 08).
