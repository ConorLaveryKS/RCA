#!/usr/bin/env python3
"""
generate_import_csvs.py
-----------------------
Read a Kingspan Product Price Template (.xlsx) and emit the numbered Salesforce
import CSVs (01-09; pricing files 10-11 are stubbed for factory data).

Usage:
    python generate_import_csvs.py TEMPLATE.xlsx OUTPUT_DIR [PREFIX]

    TEMPLATE.xlsx   the 10-sheet Product Price Template
    OUTPUT_DIR      directory to write the CSVs into
    PREFIX          attribute prefix, e.g. "ECO" (default: inferred from template)

The template is expected to follow the KLAW Germany reference structure with the
sheets: Product, AttributeDefinition, AttributePicklist, AttributePicklistValue,
Product and Attributes. Column positions are read by header name where possible.

This is a scaffold: review the output and adjust org-specific column headers to
match your SOQL query set before importing.
"""
import sys, os, csv
import openpyxl


def read_rows(ws, skip=2):
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < skip:
            continue
        if all(v is None for v in row):
            continue
        rows.append(row)
    return rows


def write_csv(out_dir, filename, headers, rows):
    path = os.path.join(out_dir, filename)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, quoting=csv.QUOTE_ALL)
        w.writerow(headers)
        w.writerows(rows)
    print(f"  {filename}: {len(rows)} rows")


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    template = sys.argv[1]
    out_dir = sys.argv[2]
    prefix = sys.argv[3] if len(sys.argv) > 3 else None
    os.makedirs(out_dir, exist_ok=True)

    wb = openpyxl.load_workbook(template, data_only=True)

    def sheet(name_contains):
        for s in wb.sheetnames:
            if name_contains.lower() in s.lower():
                return wb[s]
        raise KeyError(f"No sheet matching '{name_contains}' in {wb.sheetnames}")

    # ---- 01 AttributePicklist ----
    apl = read_rows(sheet("AttributePicklist"))
    rows = []
    for r in apl:
        name, code = (r + (None,) * 8)[0], (r + (None,) * 8)[1]
        if not code:
            continue
        desc = (r + (None,) * 8)[2] or ""
        rows.append([name, code, desc, "Active", "Text", "EUR"])
    write_csv(out_dir, "01_AttributePicklist.csv",
              ["Name", "Code", "Description", "Status", "DataType", "CurrencyIsoCode"], rows)

    # ---- 02 AttributeDefinition ----
    ad = read_rows(sheet("AttributeDefinition"))
    rows = []
    for r in ad:
        rr = (r + (None,) * 13)
        # expected order: InRCA, NameENG, Label, Desc, Seq, DataType, IsActive,
        #                 IsRequired, Default, PicklistCode, Code, DeveloperName, NameGER
        name_eng, label, desc, dtype = rr[1], rr[2], rr[3], rr[5]
        is_active, is_req, default_v = rr[6], rr[7], rr[8]
        pl_code, code, dev = rr[9], rr[10], rr[11]
        if not code:
            continue
        rows.append([name_eng, "EUR", label, desc or "", dtype, is_active, is_req,
                     default_v or "", pl_code or "", code, dev])
    write_csv(out_dir, "02_AttributeDefinition.csv",
              ["Name", "CurrencyIsoCode", "Label", "Description", "DataType",
               "IsActive", "IsRequired", "DefaultValue", "Picklist.Code",
               "Code", "DeveloperName"], rows)

    # ---- 03 AttributePicklistValue ----
    apv = read_rows(sheet("AttributePicklistValue"))
    # default map from AttributeDefinition (picklist_code|value -> True)
    default_map = {}
    for r in ad:
        rr = (r + (None,) * 13)
        pl_code, default_v = rr[9], rr[8]
        if pl_code and default_v:
            default_map[f"{pl_code}|{default_v}"] = True
    rows = []
    for r in apv:
        rr = (r + (None,) * 6)
        pl_code, val_code, label_en = rr[0], rr[1], rr[2]
        seq = rr[4]
        if not pl_code or not val_code:
            continue
        is_def = "true" if default_map.get(f"{pl_code}|{val_code}") else "false"
        rows.append([label_en, "EUR", pl_code, val_code, "Active", val_code,
                     label_en, is_def, seq or "", val_code])
    write_csv(out_dir, "03_AttributePicklistValue.csv",
              ["Name", "CurrencyIsoCode", "Picklist:AttributePicklist:Code",
               "Abbreviation", "Status", "Code", "DisplayValue", "IsDefault",
               "Sequence", "Value"], rows)

    # ---- 04 AttributeCategory + 05 ProductClassification + 07/08 links ----
    pa = read_rows(sheet("Product and Attributes"))
    cats, classifs = {}, {}
    for r in pa:
        rr = (r + (None,) * 8)
        cat, classif = rr[4], rr[7]
        if cat and cat not in cats:
            cats[cat] = cat.replace(" ", "_").replace("&", "and")
        if classif and classif not in classifs:
            classifs[classif] = classif.replace(" ", "_").replace("-", "_").replace(".", "")
    write_csv(out_dir, "04_AttributeCategory.csv", ["Name", "Code"],
              [[n, c] for n, c in cats.items()])
    write_csv(out_dir, "05_ProductClassification.csv", ["Name", "Status", "Code"],
              [[n, "Active", c] for n, c in classifs.items()])

    rows_cat, rows_pca = [], []
    for r in pa:
        rr = (r + (None,) * 8)
        seq, attr_code, default, cat, classif = rr[1], rr[2], rr[3], rr[4], rr[7]
        if cat and attr_code:
            rows_cat.append([cats.get(cat, cat), attr_code, ""])
        if classif and attr_code:
            cc = classifs.get(classif, classif)
            rows_pca.append([f"{cc}_{attr_code}", cats.get(cat, "") if cat else "",
                             "false", "false", "false", "false", "Dropdown",
                             attr_code, default or "", cc, seq or ""])
    write_csv(out_dir, "07_AttributeCategoryAttribute.csv",
              ["AttributeCategory.Code", "AttributeDefinition.Code",
               "AttributeDefinition.Name"], rows_cat)
    write_csv(out_dir, "08_ProductClassificationAttr.csv",
              ["Name", "AttributeCategory.Code", "IsPriceImpacting", "IsHidden",
               "IsReadOnly", "IsRequired", "DisplayType", "AttributeDefinition.Code",
               "DefaultValue", "ProductClassification.Code", "Sequence"], rows_pca)

    # ---- 06 Product2 ----
    prod = read_rows(sheet("Product"), skip=5)
    rows = []
    for r in prod:
        rr = (r + (None,) * 23)
        name, prod_code = rr[4], rr[7]
        if not name or str(name).startswith("Steps"):
            continue
        rows.append([name, "TRUE", "", "Allowed", prod_code or "", "", "true",
                     rr[11] or "", "true", "2025-01-01"])
    write_csv(out_dir, "06_Product2.csv",
              ["Name", "IsActive", "DisplayUrl", "ConfigureDuringSale",
               "ProductCode", "Type", "RCA_Product__c", "Description",
               "Configurable_Product__c", "AvailabilityDate"], rows)

    # ---- 09 ProductCategoryProduct (stub) / 10 / 11 (pricing placeholders) ----
    write_csv(out_dir, "09_ProductCategoryProduct.csv",
              ["Product.Name", "ProductCategory.Code"],
              [["<PRODUCT NAME>", "<CATEGORY CODE>"]])
    write_csv(out_dir, "10_PricebookEntry.csv",
              ["Pricebook2.Name", "Product2.Name", "CurrencyIsoCode", "UnitPrice",
               "IsActive", "UseStandardPrice"],
              [["<PRICEBOOK>", "<PRODUCT NAME>", "EUR", "<PRICE>", "true", "false"]])
    write_csv(out_dir, "11_Attribute_Pricing__c.csv",
              ["Product__r.Name", "Attribute__c", "Cost__c", "Factor__c",
               "Attribute_Value__c"],
              [["<PRODUCT NAME>", "<ATTRIBUTE>", "", "", ""]])

    print(f"\nDone. CSVs written to {out_dir}")
    print("Review 06/09/10/11 placeholders and org-specific headers before import.")


if __name__ == "__main__":
    main()
