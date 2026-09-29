#!/usr/bin/env python3
"""
generate_import_csvs.py
-----------------------
Read a Kingspan Product Price Template (.xlsx) and emit the numbered Salesforce
import CSVs (01-11; pricing files 10-11 are placeholders until factory data).

Usage:
    python generate_import_csvs.py TEMPLATE.xlsx OUTPUT_DIR [PREFIX] [--outputs OUTPUT_ATTRS.csv]

    TEMPLATE.xlsx   the 10-sheet Product Price Template
    OUTPUT_DIR      directory to write the CSVs into
    PREFIX          attribute prefix, e.g. "ECO" (default: inferred from the
                    AttributeDefinition codes, e.g. "ECO_Version" -> "ECO")
    --outputs FILE  optional CSV of calculated output attributes, laid out like
                    the template's AttributeDefinition sheet (header row "In RCA",
                    "Name ENG", "Label", ...). Emits 02b / 04b / 08b.

The template is expected to follow the KLAW Germany reference structure with the
sheets: Product, AttributeDefinition, AttributePicklist, AttributePicklistValue,
Product and Attributes. Columns are read by header name.

Output matches the files imported for ECO CO: comma-separated, minimal quoting,
CRLF line endings, UTF-8. ProductClassification codes follow
"<PREFIX>_<ProductCode>_ProductClassification" (e.g. ECO_CO_ProductClassification).
Placeholders to fill before go-live are written as "TO BE FILLED" or left blank.
"""
import sys, os, csv, datetime
import openpyxl

TBF = "TO BE FILLED"
CATEGORY_OUTPUTS = ("Calculated Outputs", "Calculated_Outputs")


def s(v):
    """Cell value -> CSV string."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def to_code(name):
    """'Control & Drive' -> 'Control_and_Drive'."""
    return s(name).replace("&", "and").replace(" ", "_").replace("__", "_")


def to_date(v):
    """'01.01.2025' / datetime -> '2025-01-01'."""
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%Y-%m-%d")
    t = s(v)
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.datetime.strptime(t, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return t


def table(ws, header_row):
    """Rows below header_row (1-based) as dicts keyed by header text."""
    rows = list(ws.iter_rows(values_only=True))
    headers = [s(h) for h in rows[header_row - 1]]
    out = []
    for r in rows[header_row:]:
        if all(v is None or s(v) == "" for v in r):
            continue
        out.append({h: r[i] if i < len(r) else None
                    for i, h in enumerate(headers) if h})
    return out


def write_csv(out_dir, filename, headers, rows):
    path = os.path.join(out_dir, filename)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(headers)
        w.writerows(rows)
    print(f"  {filename}: {len(rows)} rows")


def attr_def_row(r):
    return [s(r.get("Name ENG")), "EUR", "", s(r.get("Label")), s(r.get("Description")),
            s(r.get("DataType")), s(r.get("IsActive")), s(r.get("IsRequired")),
            s(r.get("DefaultValue")), "", "", s(r.get("Picklist.Code")), "",
            s(r.get("Code")), s(r.get("DeveloperName (API Name)")) or s(r.get("Code"))]


ATTR_DEF_HEADERS = ["Name", "CurrencyIsoCode", "UnitOfMeasureId", "Label", "Description",
                    "DataType", "IsActive", "IsRequired", "DefaultValue",
                    "SourceSystemIdentifier", "ValueDescription", "Picklist.Code",
                    "DefaultHelpText", "Code", "DeveloperName"]
PCA_HEADERS = ["Name", "AttributeCategory.Code", "IsPriceImpacting", "IsHidden",
               "IsReadOnly", "IsRequired", "DisplayType", "AttributeDefinition.Code",
               "DefaultValue", "ProductClassification.Code", "Sequence"]


def main():
    args = sys.argv[1:]
    outputs_csv = None
    if "--outputs" in args:
        i = args.index("--outputs")
        outputs_csv = args[i + 1]
        del args[i:i + 2]
    if len(args) < 2:
        print(__doc__)
        sys.exit(1)

    template, out_dir = args[0], args[1]
    prefix = args[2] if len(args) > 2 else None
    os.makedirs(out_dir, exist_ok=True)

    wb = openpyxl.load_workbook(template, data_only=True)

    def sheet(name):
        for sn in wb.sheetnames:
            if sn.strip().lower() == name.lower():
                return wb[sn]
        raise KeyError(f"No sheet named '{name}' in {wb.sheetnames}")

    ad = [r for r in table(sheet("AttributeDefinition"), 2) if s(r.get("Code"))]
    if not prefix:
        prefix = s(ad[0]["Code"]).split("_")[0]

    # ---- 01 AttributePicklist ----
    apl = [r for r in table(sheet("AttributePicklist"), 2) if s(r.get("Code"))]
    write_csv(out_dir, "01_AttributePicklist.csv",
              ["Name", "Code", "Description", "Status", "DataType", "CurrencyIsoCode"],
              [[s(r.get("Name ENG")), s(r["Code"]), s(r.get("Description")),
                s(r.get("Status")) or "Active", s(r.get("DataType")) or "Text", "EUR"]
               for r in apl])

    # ---- 02 AttributeDefinition ----
    write_csv(out_dir, "02_AttributeDefinition.csv", ATTR_DEF_HEADERS,
              [attr_def_row(r) for r in ad])

    # ---- 03 AttributePicklistValue ----
    defaults = {(s(r.get("Picklist.Code")), s(r.get("DefaultValue"))) for r in ad}
    rows = []
    for r in table(sheet("AttributePicklistValue"), 2):
        pl, val, label = s(r.get("Picklist.Code")), s(r.get("Value (Code)")), s(r.get("Label ENG"))
        if not pl or not val:
            continue
        is_def = "true" if (pl, val) in defaults else "false"
        rows.append([label, "EUR", pl, val, "Active", val, label, is_def,
                     s(r.get("Sequence")), val])
    write_csv(out_dir, "03_AttributePicklistValue.csv",
              ["Name", "CurrencyIsoCode", "Picklist.Code", "Abbreviation", "Status",
               "Code", "DisplayValue", "IsDefault", "Sequence", "Value"], rows)

    # ---- 04 AttributeCategory + 05 ProductClassification ----
    pa = [r for r in table(sheet("Product and Attributes"), 2)
          if s(r.get("AttributeDefinition Code"))]
    cats, classifs = {}, {}
    for r in pa:
        cat, cl = s(r.get("Attribute Category")), s(r.get("Product Classification"))
        if cat and cat not in cats:
            cats[cat] = to_code(cat)
        if cl and cl not in classifs:
            classifs[cl] = f"{prefix}_{s(r.get('Product Code'))}_ProductClassification"
    write_csv(out_dir, "04_AttributeCategory.csv", ["Name", "Code"],
              [[n, c] for n, c in cats.items()])
    write_csv(out_dir, "05_ProductClassification.csv", ["Name", "Status", "Code"],
              [[n, "Active", c] for n, c in classifs.items()])

    # ---- 06 Product2 ----
    products = [r for r in table(sheet("Product"), 5) if s(r.get("Name"))]
    rows = []
    for r in products:
        rows.append([s(r["Name"]), s(r.get("IsActive")) or "TRUE", s(r.get("DisplayUrl")),
                     s(r.get("ConfigureDuringSale")) or "Allowed", "",
                     s(r.get("ProductCode")), s(r.get("Type")), "true",
                     s(r.get("Internal Product Description")),
                     s(r.get("Product Description")),
                     s(r.get("Configurable Product")).lower() or "true",
                     to_date(r.get("Availability Date")),
                     s(r.get("Product Owner")), s(r.get("BasedOn.Code"))])
    write_csv(out_dir, "06_Product2.csv",
              ["Name", "IsActive", "DisplayUrl", "ConfigureDuringSale",
               "Product_Family_ROD__c", "ProductCode", "Type", "RCA_Product__c",
               "Internal_Product_Description__c", "Description",
               "Configurable_Product__c", "AvailabilityDate",
               "Product_Owner__r.Email", "BasedOn.Code"], rows)

    # ---- 07 AttributeCategoryAttribute + 08 ProductClassificationAttr ----
    write_csv(out_dir, "07_AttributeCategoryAttribute.csv",
              ["AttributeCategory.Code", "AttributeDefinition.Code", "AttributeDefinition.Name"],
              [[cats[s(r["Attribute Category"])], s(r["AttributeDefinition Code"]), ""]
               for r in pa if s(r.get("Attribute Category"))])
    rows = []
    for r in pa:
        cl = s(r.get("Product Classification"))
        if not cl:
            continue
        cc, attr = classifs[cl], s(r["AttributeDefinition Code"])
        rows.append([f"{cc}_{attr}", cats.get(s(r.get("Attribute Category")), ""),
                     "false", "false", "false", "false", "Dropdown", attr,
                     s(r.get("Default Picklist Value")), cc, s(r.get("Sequence"))])
    write_csv(out_dir, "08_ProductClassificationAttr.csv", PCA_HEADERS, rows)

    # ---- 09 ProductCategoryProduct / 10 PricebookEntry / 11 Attribute_Pricing__c ----
    write_csv(out_dir, "09_ProductCategoryProduct.csv", ["Product.Name", "ProductCategory.Code"],
              [[s(r["Name"]), to_code(r.get("Category")) or TBF] for r in products])
    write_csv(out_dir, "10_PricebookEntry.csv",
              ["Pricebook2.Name", "Product2.Name", "CurrencyIsoCode", "UnitPrice",
               "IsActive", "UseStandardPrice", "Country_Product_Markup__c"],
              [[TBF, s(r["Name"]), "EUR", "", "true", "false", ""] for r in products])
    write_csv(out_dir, "11_Attribute_Pricing__c.csv",
              ["Product__r.Name", "Attribute__c", "Cost__c", "Factor__c", "Attribute_Value__c"],
              [[s(r["Name"]), TBF, "", "", ""] for r in products])

    # ---- 02b / 04b / 08b calculated output attributes (optional) ----
    if outputs_csv:
        with open(outputs_csv, newline="", encoding="utf-8-sig") as f:
            outs = [r for r in csv.DictReader(f) if s(r.get("Code"))]
        write_csv(out_dir, "04b_AttributeCategory_OutputAttrs.csv", ["Name", "Code"],
                  [list(CATEGORY_OUTPUTS)])
        write_csv(out_dir, "02b_AttributeDefinition_OutputAttrs.csv", ATTR_DEF_HEADERS,
                  [attr_def_row(r) for r in outs])
        rows = []
        for cc in classifs.values():
            for r in outs:
                attr = s(r["Code"])
                rows.append([f"{cc}_{attr}", CATEGORY_OUTPUTS[1], "false", "false",
                             "true", "false", "Text", attr, "", cc, s(r.get("Sequence"))])
        write_csv(out_dir, "08b_ProductClassificationAttr_OutputAttrs.csv", PCA_HEADERS, rows)

    print(f"\nDone. CSVs written to {out_dir}")
    print(f"Fill in '{TBF}' / blank go-live fields in 06, 09, 10, 11 before import.")


if __name__ == "__main__":
    main()
