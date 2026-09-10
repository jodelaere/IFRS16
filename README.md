# IFRS16
Claude Repository IFRS16

## Monthly roll-forward automation

Design and first-draft Office Script for automating the monthly IFRS16 Input
Board Pack update (Anaplan export → input tabs → refresh → movement schedule →
new lease detection).

- [`docs/maandprocedure.md`](docs/maandprocedure.md) — **the monthly procedure**, in Dutch, for running the close by hand. Start here.
- [`docs/automation-design.md`](docs/automation-design.md) — process mapping, SharePoint folder structure, the template changes and why each one exists.
- [`scripts/ifrs16-template-setup.ts`](scripts/ifrs16-template-setup.ts) — one-time Office Script that applies the template changes to a workbook. Safe to re-run.
- [`scripts/ifrs16-monthly-rollforward.ts`](scripts/ifrs16-monthly-rollforward.ts) — Office Script for the monthly run via Power Automate's "Run script" action. Not in use: the flow was not built, the close is run by hand per the procedure above.
- [`scripts/power-query/`](scripts/power-query) — the corrected M for both output queries, to paste via the Advanced Editor.
