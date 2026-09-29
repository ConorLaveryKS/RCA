# CML syntax rules — read before writing CML

All discovered by iterating against real Salesforce validation errors during the
ECO Roof Ventilator onboarding. Each is a rule that will fail at compile or run time.

## ❌ `action()` cannot set attribute values

`action()` is **only** for adding/removing product *relations* (bundle components).
It cannot assign a value to an attribute.

```
// WRONG — "action operation is not possible"
action(true, ECO_NumLouvres = (ECO_ThroatLength_mm - 40) / 133);

// RIGHT — table constraint for a value mapping
constraint(table(ECO_ThroatLength_mm, ECO_NumLouvres,
    {705, 5}, {838, 6}, ...), "Louvre count");

// RIGHT — implication for a conditional value
constraint(ECO_Version == "2V" -> ECO_SnowLoad == "1500", "2V snow load");
```

## ❌ Annotations NOT supported on attributes

| Annotation | Status | Use instead |
|---|---|---|
| `@(readOnly = true)` | ✗ not valid | `@(configurable = false)` |
| `@(hidden = true)` | ✗ not valid on attributes | `rule(true, "hide", "attribute", "AttrName")` |

```
rule(true, "disable", "attribute", "ECO_NumLouvres");
decimal(2) ECO_NumLouvres;
```

## ❌ No string methods

`.startsWith()`, `.contains()`, etc. do not exist. Expand to explicit `||` lists:
```
// WRONG
Control.startsWith("M")
// RIGHT
ECO_Control == "M1B24" || ECO_Control == "M2B24" || ...
```

## ❌ `ROUND()` is not supported

Remove it; declare the variable as `decimal(2)` and the engine keeps 2 dp natively.
```
// WRONG
ECO_AeroAreaAa == ROUND(ECO_GeomAreaAv * 0.60, 2)
// (but see the next rule — even without ROUND, arithmetic here is risky)
```

## ❌ Decimal arithmetic in implications → "Reached limit of tree nodes for search"

Setting a free `decimal` variable via arithmetic (`Aa == Av * 0.60`) forces the
engine to search the whole decimal domain and it blows the search-tree budget.

**Fix: pre-compute every value into a lookup table**, grouped by shared factor.
Use `scripts/generate_cml_tables.py` to emit these.

```
// WRONG — unbounded decimal search
constraint((ECO_LouvreType == "A1B") && ECO_Guards == "X" ->
           ECO_AeroAreaAa == ECO_GeomAreaAv * 0.60, "Aa A1B");

// RIGHT — pre-computed table (value already multiplied out)
constraint(
    (ECO_LouvreType == "A1B" || ECO_LouvreType == "A1X") && ECO_Guards == "X" ->
    table(ECO_ThroatWidth_mm, ECO_ThroatLength_mm, ECO_AeroAreaAa,
        {426, 705, 0.18}, {426, 838, 0.22}, ... ),   // all size combinations
    "Aa Cv=0.60 no guard");
```
On ECO CO this was 17 factor groups × 308 size combinations = 5,236 entries.

## ❌ `exclude()` / `require()` take relations, not attribute expressions

Their second argument must be `relation[Type]` (a child product). For
attribute-to-attribute rules use implication:
```
// WRONG
exclude(ECO_Version != "3S", ECO_Guards == "BG", "msg");
require(ECO_Coverage != "X", ECO_Finish, "msg");
// RIGHT
constraint(ECO_Version != "3S" -> ECO_Guards != "BG", "msg");
constraint(ECO_Coverage != "X" -> ECO_Finish != "X", "msg");
```

## ✅ Other notes
- Use ASCII `->` for implication, never a Unicode arrow.
- `message(condition, "text", "Warning"|"Error"|"Info")` — text is a single string
  literal (no `&` concatenation).
- Proxy variables (code→mm, release-temp int) are declared plainly:
  `decimal(0) ECO_ThroatWidth_mm;` with no annotation.
- A code→value mapping and a two-input→output calculation are both best done as
  `constraint(table(...))`; a conditional string assignment is best as implication.
