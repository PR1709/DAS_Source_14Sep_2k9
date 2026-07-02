# 02 Architecture Overview

## System Identity

The system is DAS / Visual Effmcg 4.0, a legacy FMCG/direct-store-distribution ERP for distributor operations. The data and branding indicate beverage distribution, with Coca-Cola product assets and SKU names.

## Current Runtime Architecture

The application is a VB6 desktop executable that talks directly to Visual FoxPro DBF/DBC files over ODBC.

```text
VB6 Forms
  -> shared .bas/.cls modules
  -> global ADODB/DAO connections
  -> Visual FoxPro ODBC driver
  -> EFF.DBC + DBF/CDX/FPT files
  -> report engines / Tally / HHT / SMS / BarTender
```

There is no clean separation between UI, service logic, and persistence. Business rules live inside form event handlers, grid events, and shared utility modules.

## Database Access Pattern

`modGeneral.bas` opens FoxPro using:

```vb
Driver={Microsoft Visual FoxPro Driver};SourceDB=<CURDIR>\EFF.DBC;SourceType=DBC;Exclusive=No;...
```

The code mostly uses:

- global `ADODB.Connection` variables
- global recordsets
- dynamic SQL string concatenation
- `LIKE` against exact IDs
- pessimistic/keyset recordsets
- `AddNew`, field assignment, `Update`
- direct `con.Execute` updates/deletes

## Global State

`EFF_LIB.bas` declares many public variables:

- `CON`, `PNV`, `PNL`, `RRI`, `CN`, `CN1`, `cn2`
- invoice/date/report globals: `INVOICE_DATE`, `FRM_INV_DAT`, `TO_INV_DAT`, `FRM_INV`, `END_INV`, `INV_TYP`
- user/session globals: `USERNAME`
- route/customer/sales globals: `ROU_IDY`, `SAL_QRY`

This means the current behavior may depend on form activation order and global mutations.

## UI Architecture

The UI is form-heavy and grid-heavy. Common patterns:

- `SSTab` tabs for multi-step screens
- `MSFlexGrid` editable rows
- `ListView` lookup lists
- toolbar buttons for new/save/cancel/print
- keyboard-driven workflow using Enter/arrow keys
- uppercase conversion and blocked characters in keypress handlers
- modal lookup lists for product/customer/account/salesman

## Domain Boundaries In The Existing App

The codebase does not enforce explicit bounded contexts, but the functional clusters are clear:

- Master data: accounts, products, market, route, salesman, category, company, groups.
- Sales: sales invoice, retailer invoice, sales order, schemes, free goods.
- Van/field force: HHT sync, van load/return, mobile billing, route loads.
- Purchase: order, purchase invoice, receipt.
- Inventory: stock in/out, return, adjustment, stock ledger, product/batch quantities.
- Finance: receipt/payment/journal/voucher/ledger/Tally export.
- Empties/deposits: shells, crates, bottles, deposits, breakage/leakage.
- Reporting: operational and financial report definitions.
- Admin: login, permissions, setup, backup, restore, indexing, year close.

## Modernization Implication

The migration must extract and stabilize business behavior before replacing UI. A safe target architecture is:

```text
Web UI
  -> typed REST API
  -> domain services
  -> PostgreSQL with constraints/audit
  -> background workers
  -> integrations
```

Business logic should move out of screens into testable services:

- pricing service
- scheme/rules service
- invoice posting service
- stock ledger service
- van reconciliation service
- finance ledger/Tally sync service
- HHT/mobile sync service

## Why Not A Direct Rewrite

A direct screen-by-screen rewrite would preserve accidental complexity:

- UI event order as business logic
- raw DBF table shape
- duplicate transactional tables
- no transaction boundaries
- no audit trail
- fragile external integrations

The correct path is behavior-preserving, architecture-redesigning migration.
