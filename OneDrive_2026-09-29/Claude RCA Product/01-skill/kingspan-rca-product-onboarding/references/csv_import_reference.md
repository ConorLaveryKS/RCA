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

Write CSVs with `csv.QUOTE_ALL`, UTF-8. Column headers must match the SOQL field
names in the org's query set (e.g. `Picklist:AttributePicklist:Code`).

## Key column sets

**01 AttributePicklist:** `Name, Code, Description, Status, DataType, CurrencyIsoCode`

**02 AttributeDefinition:** `Name, CurrencyIsoCode, Label, Description, DataType,
IsActive, IsRequired, DefaultValue, Picklist.Code, Code, DeveloperName`

**03 AttributePicklistValue:** `Name, CurrencyIsoCode, Picklist:AttributePicklist:Code,
Abbreviation, Status, Code, DisplayValue, IsDefault, Sequence, Value`

**06 Product2:** `Name, IsActive, DisplayUrl, ConfigureDuringSale, ProductCode, Type,
RCA_Product__c, Description, Configurable_Product__c, AvailabilityDate`
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
Use `scripts/fix_picklist_values.py`.

## Manual fields to complete before go-live
- `06_Product2.csv` → `Product_Owner__r.Email`, `Product_Family_ROD__c`
- `09_ProductCategoryProduct.csv` → confirm `ProductCategory.Code` in the target org
- `10_PricebookEntry.csv` → `Pricebook2.Name`, `UnitPrice` (from factory)
- `08_ProductClassificationAttr.csv` → set `IsPriceImpacting=true` on price-driving attrs
