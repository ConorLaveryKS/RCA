# Output-attribute patterns

Which CML construct to use for each calculated output field. Rule of thumb:

- **one input → output** or a **value-to-value mapping** → 2-column `constraint(table(...))`
- **two inputs → output** → 3-column `constraint(table(...))`
- **conditional string** → `constraint(condition -> Var == "value", "msg")`
- **never** express a decimal output as arithmetic on a free variable (search-limit)

| Output field | Source (ECO) | CML pattern |
|---|---|---|
| Number of louvres | `(Length_mm - 40) / 133` | 2-col table: `Length_mm → NumLouvres` |
| Geometric area Av | `Width_mm × Length_mm / 1e6` | 3-col table: `Width_mm, Length_mm → Av` (all combos) |
| Aerodynamic area Aa | `Av × Cv × guard_factor` | Pre-computed 3-col tables grouped by Cv factor |
| Release temperature | code → °C lookup | 2-col table `ThermalRelease → TempInt`, then bind output |
| Ambient temp class | control-family logic | implication per branch |
| CO2 bottle size | control + release | implication `-> Bottle == "X/20/28/40"` |
| Opening pressure (vent) | `MAX(2, table)` by size | implication `-> Press == "2/4.9/9.8/X"` |
| Opening pressure (smoke) | `MAX(2, table)` by size | implication per size band |
| Wind load | louvre + size | implication `-> WL == "1500/3000/4000"` |
| Snow load | control family + version | implication `-> SL == "…"` |
| U-value | LouvreType × Insulation | 3-col table; CB* base → `"X"` via separate implication |
| Weight | large factory table | sentinel `"X"` + **Apex invocable action** (too large for CML) |

## Binding an output to a proxy variable

Some outputs are computed via a hidden proxy (e.g. release temperature as an int),
then copied to the visible output attribute:

```
// proxy declared plainly, no annotation
decimal(0) ECO_ReleaseTemp_int;

// 1. map code → proxy
constraint(table(ECO_ThermalRelease, ECO_ReleaseTemp_int,
    {"RX",0},{"R68",68},{"R93",93},{"R141",141},{"R182",182}), "release temp");

// 2. bind proxy → visible output
constraint(true -> ECO_ReleaseTemp == ECO_ReleaseTemp_int, "release temp output");
```

## Conditional table (guarded by a condition)

For U-value, the table only applies to non-conical bases; conical bases return "X":
```
constraint(BaseType == "CB2" || BaseType == "CB3" || BaseType == "CB4"
           -> ECO_UValue == "X", "no U-value for conical bases");

constraint(BaseType != "CB2" && BaseType != "CB3" && BaseType != "CB4" ->
    table(ECO_LouvreType, ECO_Insulation, ECO_UValue,
        {"A1B","X","5.51"}, {"A1B","I","2.43"}, ...), "U-value by panel & insulation");
```

## Fields that need Apex, not CML

If a lookup table is very large (hundreds of rows depending on several dimensions,
e.g. the weight table), do **not** inline it in CML. Set a sentinel `"X"` in the
model and compute the real value with an Apex `@InvocableMethod` called from a
Price Rule / action on the quote line.
