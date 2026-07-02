# 08 Risks, Reuse, And Redesign

## Highest Migration Risks

### 1. Scheme Engine

Risk:

`M_FRM_SAL.frm` contains many scheme validators and handlers. Auto-apply dispatch is known, but individual handlers still need full extraction.

Impact:

Wrong free goods or discounts directly affect revenue, claims, and customer trust.

Mitigation:

- extract every scheme handler
- create golden-master tests by scheme type
- build versioned rules engine
- include an explainability trace for every applied scheme

### 2. Quantity Conversion

Risk:

Legacy quantities use decimal-looking values for case/piece representation.

Impact:

Using decimal math naively will corrupt stock and invoice values.

Mitigation:

- implement `MTS`, `STM`, `P_SUB_UNT`, `P_QTY_UNT`, `RATCAL`, `WGTCAL`, `UNTCAL` exactly
- test with many `SUB_UNT` values

### 3. Stock Movement Semantics

Risk:

Transaction types appear with different signs in different routines.

Impact:

Stock ledger and current stock may diverge.

Mitigation:

- build transaction-type dictionary by form/workflow
- derive movements from posted transaction intent, not table name alone
- reconcile against historical stock reports

### 4. Duplicate Tables

Risk:

`INL/PNL` and `INV/PNV` semantics are not yet fully proven.

Impact:

Wrong source of truth selection can lose transactions or duplicate postings.

Mitigation:

- analyze all reads/writes to each pair
- inspect `dsync.PRG`, print/report flows, and HHT flows
- stage raw tables separately before canonical mapping

### 5. Tally Integration

Risk:

Tally payloads are string-built and response parsing is brittle.

Impact:

Financial postings may fail silently or partially.

Mitigation:

- create integration outbox
- persist payload/response
- retry safely
- parse XML formally
- reconcile exported vouchers

### 6. HHT Sync

Risk:

Flat file/serial sync updates many core tables.

Impact:

Mobile/field-force migration can break route sales and van reconciliation.

Mitigation:

- preserve compatibility importer/exporter first
- replace with offline-first sync later
- audit every sync session

### 7. Reports As Business Logic

Risk:

Legacy reports may contain implicit filters/calculations.

Impact:

Users may treat report totals as truth.

Mitigation:

- catalog reports
- prioritize high-use reports
- compare new report totals against legacy outputs

## Reusable Assets

Reusable directly as source material:

- DBF data
- generated Postgres DDL as raw import reference
- legacy IDs and prefixes
- HHT file examples
- Tally XML generation behavior
- report names and report library organization
- product/customer/scheme examples
- existing transaction histories for golden-master tests

Reusable conceptually:

- business workflows
- transaction types
- pricing rules
- scheme rules
- stock conversion rules
- route/market/salesman structure
- ledger posting patterns

Not reusable directly:

- VB6 forms
- ActiveX controls
- FoxPro ODBC access layer
- report binaries
- serial modem implementation
- global mutable state
- raw DBF schema as final schema

## Components To Redesign

### Data Model

Move from raw DBF tables to a normalized relational model with:

- explicit primary keys
- foreign keys
- check constraints
- immutable posting tables
- audit tables
- raw import tables retained separately

### Scheme Engine

Move to a versioned rules engine with:

- conditions
- benefits
- eligibility explanation
- deterministic calculation
- simulation before activation

### Stock Ledger

Move to immutable stock movements:

- receipt/issue/adjustment movement types
- product/batch/location/van dimensions
- source document references
- no direct mutable `CUR_LEV` as source of truth

### Finance Ledger

Move to double-entry ledger postings:

- debit/credit entries
- source document references
- posting status
- Tally export status

### Field Force

Short-term:

- compatibility HHT import/export.

Long-term:

- offline-first mobile/PWA with conflict resolution.

### Reporting

Replace report binaries with:

- Postgres views/materialized views
- web report screens
- PDF/Excel exports
- BI-compatible schema

## Missing Documentation

Still needed:

- exact user roles and permission model
- current deployment environment
- actual HHT hardware/file protocol specification
- Tally version and target company setup
- report usage priority
- current accounting reconciliation process
- backup/restore operating procedure
- branch/multi-company behavior
- tax/GST evolution requirements
