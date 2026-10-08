# Airlite (AIR) — Product Price Template

`Airlite_ProductPriceTemplate.xlsx` is the Airlite Tech. Configurator (v1.54) converted into the
same 10-sheet KLAW Germany Product Price Template structure used for ECO CO, plus an
`Airlite Notes` sheet listing sources, assumptions and open points.

| Sheet | Content |
|---|---|
| Product | AIR – Airlite Smoke and Heat Ventilator (owner + availability date to fill, yellow) |
| AttributeDefinition | 24 configurable attributes + 12 calculated (read-only) outputs, EN + DE names |
| Product and Attributes | Links, categories, defaults and allowed values |
| AttributePicklist / AttributePicklistValue | 15 picklists, 81 values, EN + DE labels from the configurator's GPNDS table |
| Product Internal Constraint | 60 rules from the configurator's 1E/1S validation columns |
| Configurable Product Pricing | **No prices in the configurator** – pre-filled price-request list, yellow cells to fill |

Rebuild: `python build_airlite_template.py ../02-product-template/ECOTech_ProductPriceTemplate.xlsx "Airlite Tech. Configurator(1).xlsm" Airlite_ProductPriceTemplate.xlsx`

Generate import CSVs: `python ../01-skill/kingspan-rca-product-onboarding/scripts/generate_import_csvs.py Airlite_ProductPriceTemplate.xlsx <out_dir> AIR`
