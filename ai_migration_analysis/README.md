# Legacy DAS Modernization Analysis Pack

Created: 2026-07-02

This folder is a continuation kit for future AI/human reverse engineering of the DAS / Visual Effmcg 4.0 legacy system.

The user goal is to modernize the legacy desktop application into a fully web-based application while preserving business logic. No implementation has been started. These docs capture what has been verified so far and what must still be extracted before a safe rewrite.

## Source Of Truth

- Guidance first read: `../claudeguidance.md`
- Legacy application root: `../LegacyCodebase`
- Current generated Postgres target: `../MigrateDest_v1/init_postgres.sql`
- Main VB6 project: `../LegacyCodebase/Effmcg.vbp`

## Important Scope Note

The legacy tree is large:

- 2,722 files under `LegacyCodebase`
- 743 DBF files
- 214 VB6 `.frm` forms
- 216 `.frx` form/resource binaries
- 57 FoxPro `.prg` programs
- 26 `.bas` modules
- 6 `.cls` classes
- 93 `.exe` files
- 57 `.dll` files
- 240 `.rp1`, 37 `.rp5`, and 27 `.rrw` report files

This pack documents the architecture and the most business-critical logic found so far, but the project still needs staged line-by-line extraction for every high-risk form before migration code is written. Do not claim full functional parity until the extraction queue in `09_continuation_queue.md` is completed.

## Docs

1. `01_project_inventory.md` - file inventory, folder map, main project dependencies.
2. `02_architecture_overview.md` - verified legacy architecture and modernization implications.
3. `03_data_model.md` - table/domain model from the generated Postgres DDL and source usage.
4. `04_business_workflows.md` - current business workflows and transaction flows.
5. `05_business_logic_extraction.md` - extracted business rules, formulas, and source anchors.
6. `06_integrations_and_devices.md` - Tally, HHT, SMS, barcode, reports, licensing.
7. `07_migration_stack_and_roadmap.md` - recommended stack and staged migration roadmap.
8. `08_risks_reuse_redesign.md` - risks, reusable assets, redesign candidates.
9. `09_continuation_queue.md` - exact next work for future AI models.

## How Future Models Should Continue

1. Read this `README.md`.
2. Read `../claudeguidance.md`.
3. Read docs 01 through 09 in order.
4. Pick one module from `09_continuation_queue.md`.
5. Extract that module into a new Markdown file in this folder.
6. For every rule, cite file and line references.
7. Keep assumptions clearly marked.

## Non-Negotiable Migration Principle

The new web system must not be a blind port of VB6 screens or raw DBF tables. Preserve behavior first, then redesign architecture around explicit domain services, relational integrity, audit logs, and testable business rules.
