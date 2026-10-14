"""Turn generate_import_csvs.py output into a Salesforce Inspector Reloaded import workbook.
Usage: python build_inspector_import.py CSV_DIR OUT.xlsx PREFIX [--cml MODEL.cml] [--test "expected result"]
One tab per import step, lookups as Field:Object:ExternalId, picklist value codes made org-unique.
The Import Steps tab lists the tabs 01-09, then the labelled non-import steps CML-1..CML-6 (constraint
model) and TEST, then the pricing tabs 10-11. --cml names the model file and reads its `type` line;
--test is the reference configuration and expected outputs shown on the TEST step.
"""
import sys, csv, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

args = sys.argv[1:]
def opt(name):
    if name in args:
        i = args.index(name); v = args[i + 1]; del args[i:i + 2]; return v
cml_file, test_text = opt("--cml"), opt("--test")
csv_dir, out, P = args[0:3]
cml_type = None
if cml_file and os.path.exists(cml_file):
    import re
    m = re.findall(r"^type\s+(\w+)\s*:\s*LineItem", open(cml_file, encoding="utf-8").read(), re.M)
    cml_type = m[0] if m else None
cml_name = os.path.basename(cml_file) if cml_file else "<Product>_ConstraintModel.cml"
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
    "Fill yellow cells. Product_Owner__c = owner's User Id (starts 005, not alias/e-mail). "
    "AvailabilityDate is a date/time: type 2026-01-01T00:00:00.000Z exactly, NO apostrophe (cells are Text). "
    "Product_Family_ROD__c optional – use an existing value. BasedOn links the product to the classification so it inherits the attributes.",
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
    if code in outs:  # visible + read-only needs a default value (INVALID_INPUT otherwise)
        row[h.index("IsReadOnly")] = "true"
        if not row[h.index("DefaultValue")]:
            row[h.index("DefaultValue")] = "0" if dtype.get(code) == "Number" else "X"
    row[h.index("DisplayType")] = ""  # restricted picklist; "Dropdown" is rejected – blank = type default
h = [{"AttributeCategory.Code": "AttributeCategory:AttributeCategory:Code",
      "AttributeDefinition.Code": "AttributeDefinition:AttributeDefinition:Code",
      "ProductClassification.Code": "ProductClassification:ProductClassification:Code"}.get(c, c) for c in h]
tab("08 ProductClassificationAttr", h, r, "Insert",
    "Calculated outputs are IsReadOnly = true with a default (0 for numbers, X for text) – required for "
    "visible read-only attributes; the constraint model overwrites them. DisplayType left blank "
    "(restricted picklist – 'Dropdown' is rejected; blank uses the attribute type's default).")

# 09 ProductCategoryProduct
tab("09 ProductCategoryProduct", ["ProductId", "ProductCategoryId"],
    [["<Product2 Id from step 06>", "<ProductCategory Id>"]], "Insert",
    "Object = ProductCategoryProduct (Inspector keeps the last object selected – change it!). "
    "Both columns need record Ids (01t… / 0ZG…), not names – see the Data Export queries on the Import Steps tab.",
    fill_cols=("ProductId", "ProductCategoryId"))

# CML-4 association (only when --cml is given) – confirmed against ECO's ExpressionSetConstraintObj rows
if cml_type:
    tab("CML-4 ExpressionSetConstraintObj",
        ["ExpressionSetId", "ReferenceObjectId", "ConstraintModelTag", "ConstraintModelTagType"],
        [["<constraint model Id 9QL… (shown in the activation error)>", "<Product2 Id 01t…>", cml_type, "Type"]],
        "Insert",
        "Run at step CML-4, after the .cml is pasted and saved. Object = ExpressionSetConstraintObj. "
        "ExpressionSetId starts 9QL (the Id in 'We couldn't find any associations for constraint version …').",
        fill_cols=("ExpressionSetId", "ReferenceObjectId"))

# 10 / 11 pricing – not yet
tab("10 PricebookEntry", ["Pricebook2Id", "Product2Id", "UnitPrice", "IsActive", "CurrencyIsoCode"],
    [["<Standard Price Book Id>", "<Product2 Id>", "<price>", "true", "EUR"]], "Insert – LATER",
    "Skip until the factory supplies prices. Standard Price Book entry first, then the KLAW price book.",
    fill_cols=("Pricebook2Id", "Product2Id", "UnitPrice"))
tab("11 Attribute_Pricing__c", ["Product__c", "Attribute__c", "Cost__c", "Factor__c", "Attribute_Value__c"],
    [["<Product2 Id>", "<attribute>", "<cost>", "", "<value>"]], "Insert – LATER",
    "Skip until prices are known (see the template's Configurable Product Pricing tab).",
    fill_cols=("Product__c", "Attribute__c", "Cost__c", "Attribute_Value__c"))

# Data Export (SOQL) queries to find the Ids / values each step needs
Q = {
    "06 Product2": [
        ("Product owner User Id (005…)", "SELECT Id, Name, Alias, Email FROM User WHERE Alias = '<alias>' OR Email = '<e-mail>'"),
        ("Existing product family values", "SELECT Product_Family_ROD__c, COUNT(Id) FROM Product2 WHERE Product_Family_ROD__c != null GROUP BY Product_Family_ROD__c"),
    ],
    "07 AttributeCategoryAttr": [
        ("Check all attributes from step 02 exist", f"SELECT Code, Name, DataType FROM AttributeDefinition WHERE Code LIKE '{P}_%' ORDER BY Code"),
    ],
    "09 ProductCategoryProduct": [
        ("ProductId (01t…) – the product from step 06", "SELECT Id, Name, ProductCode FROM Product2 WHERE ProductCode = '<product code>'"),
        ("ProductCategoryId (0ZG…) with its catalog", "SELECT Id, Name, Catalog.Name, CatalogId FROM ProductCategory WHERE Name LIKE '%<category>%' ORDER BY Catalog.Name, Name"),
        ("Which category/catalog a similar product uses", "SELECT ProductCategoryId, ProductCategory.Name, ProductCategory.Catalog.Name FROM ProductCategoryProduct WHERE Product.ProductCode = '<similar product code>'"),
    ],
    "10 PricebookEntry": [
        ("Pricebook2Id", "SELECT Id, Name, IsStandard FROM Pricebook2 WHERE IsActive = true"),
    ],
}
pc = r_prod_code = None
for ws_ in wb.worksheets:
    if ws_.title == "06 Product2":
        hdr = [c.value for c in ws_[1]]
        r_prod_code = ws_.cell(2, hdr.index("ProductCode") + 1).value
for k in Q:
    Q[k] = [(d, q.replace("<product code>", r_prod_code or "<product code>")) for d, q in Q[k]]

# Overview tab
ws = wb["Sheet"]
ws.title = "Import Steps"
ws.append(["Step", "Inspector action", "Rows", "Notes", "Data Export query to find the Ids for this step"])
for c in ws[1]:
    c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF"); c.fill = HDR
# Non-import steps get letters, never a tab number, and sit between tab 09 and the pricing tabs
pcode = r_prod_code or "<product code>"
ctype = cml_type or "<type name from the .cml file>"
EXTRA = [
    ("CML-1 Check attributes", "Data Export", "",
     "Before loading the model, every attribute named in the .cml must exist and be linked to the classification "
     "(tabs 02 and 08 fully succeeded). Expect all configurable + output attributes.",
     f"SELECT AttributeDefinition.Code, IsReadOnly, DefaultValue FROM ProductClassificationAttr "
     f"WHERE ProductClassification.Code = '{cls_code}' ORDER BY Sequence"),
    ("CML-2 Create model", "Constraint Builder", "",
     f"Not an Inspector import. App Launcher → Constraint Models → New. Name it after the product "
     f"(e.g. {pcode} Constraint Model) and create / open its first version in the Constraint Builder.", ""),
    ("CML-3 Paste CML", "Constraint Builder", "",
     f"Open the CML editor of that version, select all and replace with the full contents of {cml_name}. "
     f"The file declares 'type {ctype} : LineItem'. Save.", ""),
    ("CML-4 Link product", "Constraint Builder", "",
     f"REQUIRED before activation (else: 'We couldn't find any associations for constraint version 9QL…'). "
     f"Not done in the Constraint Builder UI ('New Type' only adds a CML type; 'Add Labels' is unrelated). "
     f"Inspector → Data Import, Object ExpressionSetConstraintObj, Insert the tab 'CML-4 ExpressionSetConstraintObj': "
     f"ExpressionSetId = model Id (9QL…), ReferenceObjectId = product {pcode} Id (01t…), ConstraintModelTag = {ctype}, "
     f"ConstraintModelTagType = Type. Then refresh the builder and continue.",
     f"How a working model (ECO) is linked:\nSELECT Id, ExpressionSet.Name, ReferenceObjectId, ReferenceObject.Name, "
     f"ConstraintModelTag, ConstraintModelTagType FROM ExpressionSetConstraintObj WHERE ExpressionSet.Name LIKE '%ECO%'\n\n"
     f"Model Id (9QL…):\nSELECT Id, Name FROM ExpressionSet WHERE Name LIKE '%{pcode}%'\n\n"
     f"Product Id:\nSELECT Id, Name FROM Product2 WHERE ProductCode = '{pcode}'"),
    ("CML-5 Validate", "Constraint Builder", "",
     "Validate / compile. On an error, note the message and line number – common causes are in "
     "references/cml_syntax_rules.md (unsupported annotation, string method, ROUND, decimal arithmetic → tree-node "
     "search limit). Fix the .cml, paste again, re-validate.", ""),
    ("CML-6 Activate", "Constraint Builder", "",
     "Activate the version. Only one active version per model.", ""),
    ("TEST – Quote", "Quote", "",
     "Create a quote, add the product, configure. Check: every attribute shows, dropdown values are right, "
     "invalid combinations are blocked with the rule message, calculated outputs fill in. "
     + (f"Reference configuration: {test_text}" if test_text else "Use the configurator's own example as reference."),
     ""),
]
for s in steps:
    if s[0].startswith("CML"):
        continue  # its own tab, described by the CML-4 row
    if s[0].startswith("10 "):
        for e in EXTRA:
            ws.append(list(e))
    qs = Q.get(s[0], [])
    ws.append(list(s) + ["\n\n".join(f"{d}:\n{q}" for d, q in qs)])
ws.append([])
ws.append(["How", "Salesforce Inspector Reloaded → Data Import. For each tab in order: pick the object, "
           "action Insert, copy the whole tab (incl. header row) and paste into the data box, check the field "
           "mapping, Run. Fix any red rows, re-run only those, then move to the next tab."])
for row in ws.iter_rows(min_row=2):
    for c in row:
        c.font = F; c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 8; ws.column_dimensions["D"].width = 70
ws.column_dimensions["E"].width = 90
wb.save(out)
print("\n".join(f"{t}: {n} rows" for t, _, n, _ in steps))
