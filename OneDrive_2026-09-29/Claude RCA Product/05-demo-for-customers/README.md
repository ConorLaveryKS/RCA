# Demo Configurator (for customer demos)

`AirCore_Configurator_DEMO.xlsx` — a fully functional, **fictional** product
configurator that mirrors the structure and complexity of a real Kingspan
configurator. Safe to show to any prospect.

- Product: AirCore Natural Ventilation Unit (fictional)
- Vendor: Meridian Building Systems (fictional)
- 3 sheets: Input (configure + calculated outputs), Reference (lookups),
  Validation (12 business rules)
- 28 live formulas; change a dropdown and outputs recalculate; pick an invalid
  combination and the validity flag flips.

## How to use in a live demo
1. Open the file and show it behaving like a real configurator.
2. Hand it to Claude with the skill (`../01-skill/...`).
3. Watch Claude generate the attributes, CSVs, and CML model from it — timed.

Pairs with the customer-facing deck in `../06-presentations/`.
