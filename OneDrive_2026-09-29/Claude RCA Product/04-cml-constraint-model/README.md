# CML Constraint Model

## Use this one
**`ECO_CO_ConstraintModel_FINAL.cml`** — the validated, working model.
Upload to the Salesforce RCA Constraint Builder and activate.

14 sections · 5,236 pre-computed table entries · 33 business-rule validations ·
12 calculated output fields.

## version-history/  (for learning, not for use)
Shows how the model evolved and which error each version fixed:
- `v1_initial.cml`             — first draft (used action(), @readOnly)
- `v2_annotations-fixed.cml`   — fixed annotations & string methods
- `v3_action-removed.cml`      — replaced action() with constraint(table)/implication
- `v3b_pre-search-fix.cml`     — before the tree-node search-limit fix
- (FINAL adds pre-computed Aa tables + column-G validations)

See `../01-skill/kingspan-rca-product-onboarding/references/cml_syntax_rules.md`
for the full explanation of every fix.

## Regression test
Width 15 · Length 15 · Louvre POR · Control P1BD · Release FS68
→ Av=1.64 m², Aa≈1.02 m², louvres=9, SL=1618 N/m², CO2=20g, U=5.72
