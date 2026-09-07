# IFRS16
Claude Repository IFRS16

## Monthly roll-forward automation

Design and first-draft Office Script for automating the monthly IFRS16 Input
Board Pack update (Anaplan export → input tabs → refresh → movement schedule →
new lease detection).

- [`docs/automation-design.md`](docs/automation-design.md) — process mapping, SharePoint folder structure, and open validation points.
- [`scripts/ifrs16-monthly-rollforward.ts`](scripts/ifrs16-monthly-rollforward.ts) — Office Script draft (to be run via Power Automate's "Run script" action). Several ranges are marked `CONFIRM-ME` pending validation against the live workbook.
