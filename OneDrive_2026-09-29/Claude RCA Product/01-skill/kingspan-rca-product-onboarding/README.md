# Kingspan → Salesforce RCA Product Onboarding (Skill)

A reusable skill for converting a Kingspan/KLA product configurator Excel into a
complete Salesforce Revenue Cloud Advanced (RCA) product: Product Price Template,
numbered import CSVs, and a validated CML constraint model.

**Start with [`SKILL.md`](SKILL.md).**

## Layout
- `SKILL.md` — the process (entry point)
- `references/` — CML syntax rules, CSV import reference, output-field patterns
- `scripts/` — CSV generator, picklist-value fixer, CML table generator
- `assets/` — anonymized demo configurator for practice & customer demos
- `examples/` — a full worked CML constraint model

## Install
Drop this folder into your skills directory, e.g.
`/mnt/skills/organization/kingspan-rca-product-onboarding/`, then give Claude a
product configurator `.xlsm` plus a reference Product Price Template.
