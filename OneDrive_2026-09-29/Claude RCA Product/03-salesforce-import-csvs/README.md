# Salesforce Import CSVs

Import in numbered order (01 → 11) — each depends on the ones above it.

## Base files (01–11)
The 11 core objects for the ECO CO product. See
`../01-skill/kingspan-rca-product-onboarding/references/csv_import_reference.md`
for column definitions and the strict import order.

## Supplementary output-attribute files (import after 02/04/08)
- `04b_AttributeCategory_OutputAttrs.csv`      → the "Calculated Outputs" category
- `02b_AttributeDefinition_OutputAttrs.csv`    → 12 read-only calculated attributes
- `08b_ProductClassificationAttr_OutputAttrs.csv` → links them (all IsReadOnly=true)

## Error fixes
- `03b_AttributePicklistValue_FIXES.csv` → 42 corrected rows for DUPLICATE_VALUE
  and STRING_TOO_LONG errors. Re-import after the base 03 file.

## Before go-live, fill in:
- 06_Product2.csv → Product_Owner__r.Email, Product_Family_ROD__c
- 09_ProductCategoryProduct.csv → confirm ProductCategory.Code in target org
- 10_PricebookEntry.csv → Pricebook2.Name, UnitPrice (from factory)
