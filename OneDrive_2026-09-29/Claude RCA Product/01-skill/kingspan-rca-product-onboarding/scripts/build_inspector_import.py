"""Turn generate_import_csvs.py output into a Salesforce Inspector Reloaded import workbook.
Usage: python build_inspector_import.py CSV_DIR OUT.xlsx PREFIX
One tab per import step, lookups as Field:Object:ExternalId, picklist value codes made org-unique.
"""
import sys, csv, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

csv_dir, out, P = sys.argv[1:4]
HDR = PatternFill("solid", fgColor="1F4E79")
YEL = PatternFill("solid", fgColor="FFFF00")
F = Font(name="Arial", size=10)


def read(name):
    with open(os.path.join(csv_dir, name), newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    return rows[0], rows[1:]


def keep(header, rows, cols, rename=None):
    rename = rename or {}
    idx = [header.index(c) for c in cols]
    return [rename.get(c, c) for c in cols], [[r[i] for i in idx] for r in rows]


wb = openpyxl.Workbook()
steps = []


def tab(title, header, rows, action, notes, fill_cols=()):
    ws = wb.create_sheet(title)
    ws.append(header)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = HDR
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = F
            c.number_format = "@"
            if ws.cell(1, c.column).value in fill_cols:
                c.fill = YEL
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(60, max(12, *(len(str(c.value or "")) + 2 for c in col)))
    ws.freeze_panes = "A2"
    steps.append((title, action, len(rows), notes))


# 01 AttributePicklist
h, r = read("01_AttributePicklist.csv")
tab("01 AttributePicklist", h, r, "Insert", "")

# 02 AttributeDefinition
h, r = read("02_AttributeDefinition.csv")
h, r = keep(h, r, ["Name", "CurrencyIsoCode", "Label", "Description", "DataType", "IsActive", "IsRequired",
                   "DefaultValue", "Picklist.Code", "Code", "DeveloperName"],
            {"Picklist.Code": "Picklist:AttributePicklist:Code"})
for row in r:  # Name must be unique org-wide (ECO already has Version, Insulation, ...)
    if not row[0].startswith(P + " "):
        row[0] = f"{P} {row[0]}"
tab("02 AttributeDefinition", h, r, "Insert",
    f"Name is prefixed '{P} ' because AttributeDefinition Name must be unique across the org; Label (shown to users) "
    "is unchanged. Includes the 12 calculated output attributes (Description starts 'Calculated').")

# 03 AttributePicklistValue – Code made unique across the org: <PREFIX>_<picklist>_<value>
h, r = read("03_AttributePicklistValue.csv")
ipl, icode = h.index("Picklist.Code"), h.index("Code")
for row in r:
    pl = row[ipl].split(" ", 1)[1].replace(" ", "")
    if not row[icode].startswith(P + "_"):
        row[icode] = f"{P}_{pl}_{row[icode]}".replace(" ", "")
h[ipl] = "Picklist:AttributePicklist:Code"
tab("03 AttributePicklistValue", h, r, "Insert",
    f"Code is prefixed ({P}_<Picklist>_<Value>) so it can't clash with ECO codes already in the org; Value is unchanged.")

# 04 AttributeCategory
h, r = read("04_AttributeCategory.csv")
r.append(["Calculated Outputs", "Calculated_Outputs"])
tab("04 AttributeCategory", h, r, "Insert",
    "Several categories may already exist from earlier products, e.g. ECO (Device_Configuration, Filling, Control_and_Drive, Base_and_Flange, "
    "Options, Finish, Calculated_Outputs). A DUPLICATE_VALUE error on those rows is fine – they are reused.")

# 05 ProductClassification
h, r = read("05_ProductClassification.csv")
cls_code = r[0][h.index("Code")]
tab("05 ProductClassification", h, r, "Insert", "")

# 06 Product2
h, r = read("06_Product2.csv")
cols = ["Name", "IsActive", "ConfigureDuringSale", "Product_Family_ROD__c", "ProductCode", "RCA_Product__c",
        "Internal_Product_Description__c", "Description", "Configurable_Product__c", "AvailabilityDate",
        "Product_Owner__r.Email", "BasedOn.Code"]
h, r = keep(h, r, cols, {"Product_Owner__r.Email": "Product_Owner__c", "BasedOn.Code": "BasedOn:ProductClassification:Code"})
for row in r:
    row[h.index("BasedOn:ProductClassification:Code")] = cls_code
    row[h.index("Product_Owner__c")] = "<User Id of product owner>"
    row[h.index("AvailabilityDate")] = row[h.index("AvailabilityDate")] or "<YYYY-MM-DDT00:00:00.000Z>"
tab("06 Product2", h, r, "Insert",
    "Fill yellow cells. Product_Owner__c needs the owner's User Id starting 005, not an alias or e-mail (Data Export: SELECT Id FROM User WHERE Email = '...'). "
    "AvailabilityDate is a date/time field: enter it as text, e.g. 2026-01-01T00:00:00.000Z (type a leading apostrophe in Excel so it is not reformatted to 01/01/2026). BasedOn links the product to the classification so it inherits the attributes.",
    fill_cols=("Product_Owner__c", "AvailabilityDate", "Product_Family_ROD__c"))

# 07 AttributeCategoryAttribute
h, r = read("07_AttributeCategoryAttribute.csv")
h, r = keep(h, r, ["AttributeCategory.Code", "AttributeDefinition.Code"],
            {"AttributeCategory.Code": "AttributeCategory:AttributeCategory:Code",
             "AttributeDefinition.Code": "AttributeDefinition:AttributeDefinition:Code"})
outs = {row[h.index("AttributeDefinition:AttributeDefinition:Code")]
        for row in r if row[0] == "Calculated_Outputs"}
tab("07 AttributeCategoryAttr", h, r, "Insert", "")

# 08 ProductClassificationAttr
h, r = read("08_ProductClassificationAttr.csv")
h2, r2 = read("02_AttributeDefinition.csv")
dtype = {row[h2.index("Code")]: row[h2.index("DataType")] for row in r2}
ia = h.index("AttributeDefinition.Code")
for row in r:
    code = row[ia]
    if code in outs:
        row[h.index("IsReadOnly")] = "true"
    if dtype.get(code) != "Picklist":
        row[h.index("DisplayType")] = ""
h = [{"AttributeCategory.Code": "AttributeCategory:AttributeCategory:Code",
      "AttributeDefinition.Code": "AttributeDefinition:AttributeDefinition:Code",
      "ProductClassification.Code": "ProductClassification:ProductClassification:Code"}.get(c, c) for c in h]
tab("08 ProductClassificationAttr", h, r, "Insert",
    "Calculated outputs are IsReadOnly = true. DisplayType is only set (Dropdown) for picklists.")

# 09 ProductCategoryProduct
tab("09 ProductCategoryProduct", ["ProductId", "ProductCategoryId"],
    [["<Product2 Id from step 06>", "<ProductCategory Id>"]], "Insert",
    "ProductId = the Id Inspector returns in step 06. ProductCategoryId: Data Export "
    "SELECT Id, Name FROM ProductCategory WHERE Name = '<category from the Product sheet>'.",
    fill_cols=("ProductId", "ProductCategoryId"))

# 10 / 11 pricing – not yet
tab("10 PricebookEntry", ["Pricebook2Id", "Product2Id", "UnitPrice", "IsActive", "CurrencyIsoCode"],
    [["<Standard Price Book Id>", "<Product2 Id>", "<price>", "true", "EUR"]], "Insert – LATER",
    "Skip until the factory supplies prices. Standard Price Book entry first, then the KLAW price book.",
    fill_cols=("Pricebook2Id", "Product2Id", "UnitPrice"))
tab("11 Attribute_Pricing__c", ["Product__c", "Attribute__c", "Cost__c", "Factor__c", "Attribute_Value__c"],
    [["<Product2 Id>", "<attribute>", "<cost>", "", "<value>"]], "Insert – LATER",
    "Skip until prices are known (see the template's Configurable Product Pricing tab).",
    fill_cols=("Product__c", "Attribute__c", "Cost__c", "Attribute_Value__c"))

# Overview tab
ws = wb["Sheet"]
ws.title = "Import Steps"
ws.append(["Step", "Inspector action", "Rows", "Notes"])
for c in ws[1]:
    c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF"); c.fill = HDR
for s in steps:
    ws.append(list(s))
ws.append([])
ws.append(["How", "Salesforce Inspector Reloaded → Data Import. For each tab in order: pick the object, "
           "action Insert, copy the whole tab (incl. header row) and paste into the data box, check the field "
           "mapping, Run. Fix any red rows, re-run only those, then move to the next tab."])
for row in ws.iter_rows(min_row=2):
    for c in row:
        c.font = F; c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 8; ws.column_dimensions["D"].width = 110
wb.save(out)
print("\n".join(f"{t}: {n} rows" for t, _, n, _ in steps))
