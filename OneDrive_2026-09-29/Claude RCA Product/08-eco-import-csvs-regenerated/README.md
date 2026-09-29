# ECO CO — Regenerated Salesforce Import CSVs

Regenerated from `ECOTech_ProductPriceTemplate_1.xlsx` (same data as
`02-product-template/ECOTech_ProductPriceTemplate.xlsx`) with:

```
python ../01-skill/kingspan-rca-product-onboarding/scripts/generate_import_csvs.py \
    ECOTech_ProductPriceTemplate_1.xlsx . ECO \
    --outputs ../03-salesforce-import-csvs/02b_AttributeDefinition_OutputAttrs.csv
```

Import in numbered order (01 → 11), each depends on the ones above it.

## Compared with `03-salesforce-import-csvs/`

| File | Status |
|---|---|
| 01–11 base files | Byte-identical to the files already imported |
| `04b`, `08b` | Same values (minimal quoting instead of quote-all) |
| `02b_AttributeDefinition_OutputAttrs.csv` | **Corrected.** The original used the template's column names (`In RCA`, `Name ENG`, `DeveloperName (API Name)`, `Name GER`…), which do not map to `AttributeDefinition` fields. This version uses the same Salesforce columns as `02`. |
| `03b_AttributePicklistValue_FIXES.csv` | Copied unchanged. It records the codes the org actually rejected, which cannot be derived from the template alone. |

## Supplementary output-attribute files (import after 02/04/08)
- `04b_AttributeCategory_OutputAttrs.csv` → the "Calculated Outputs" category
- `02b_AttributeDefinition_OutputAttrs.csv` → 12 calculated attributes
- `08b_ProductClassificationAttr_OutputAttrs.csv` → links them (all `IsReadOnly=true`)

## Before go-live, fill in
- `06_Product2.csv` → `Product_Owner__r.Email`, `Product_Family_ROD__c`
- `09_ProductCategoryProduct.csv` → confirm `Natural_Ventilation` exists in the target org
- `10_PricebookEntry.csv` → `Pricebook2.Name`, `UnitPrice` (from factory)
- `11_Attribute_Pricing__c.csv` → `Attribute__c` and costs (from factory)
