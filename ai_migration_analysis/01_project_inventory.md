# 01 Project Inventory

## Verified Counts

Full `LegacyCodebase` inventory:

| Type | Count |
|---|---:|
| all files | 2,722 |
| `.dbf` | 743 |
| `.cdx` | 410 |
| `.frm` | 214 |
| `.frx` | 216 |
| `.prg` | 57 |
| `.bas` | 26 |
| `.cls` | 6 |
| `.vbp` | 11 |
| `.pjx` | 44 |
| `.pjt` | 44 |
| `.exe` variants | 100 |
| `.dll` | 57 |
| `.ocx` | 10 |
| `.rp1` | 240 |
| `.rp5` | 37 |
| `.rrw` | 27 |
| `.txt` | 69 |
| `.log` | 111 |

## Major Folders

- `LegacyCodebase/` - main VB6/VFP application, DBF tables, report files, utilities, binaries.
- `LegacyCodebase/HHT/` - handheld terminal flat files for stock, party, load sheet, salesman, empties, invoice import/export.
- `LegacyCodebase/INDEXFILES/` - index/HHT/SMS/Tally/barcode-focused VB6 project and support files.
- `LegacyCodebase/Reports/` - report-side DBFs/indexes/report definitions.
- `LegacyCodebase/MMEXTRACT/` and `CRSEXTRACT/` - extraction projects and duplicated/autobackup DBF/report sets.
- `LegacyCodebase/DASXtract/` and `FASTBKPEXE/` - backup/extract/index-related app copies.
- `LegacyCodebase/SetDefaults/` - separate VB project for defaults.

## Main Application

Main project file: `../LegacyCodebase/Effmcg.vbp`

Project identity:

- Type: VB6 executable
- Startup: `M_FRM_STR`
- Title: `Effmcg`
- Product name: `Visual Effmcg 4.0`
- Company: `Consoft Systems Pvt Ltd, INDIA`
- Description: Fast Moving Consumer Goods software

## Core VB6 References And Objects

From `Effmcg.vbp`, the application depends on:

- Microsoft DAO 3.51
- Microsoft ADO 2.5
- Microsoft Data Report Designer
- Microsoft Data Formatting Object Library
- Microsoft Data Environment
- Microsoft Scripting Runtime
- Microsoft Script Control
- VBScript regex/globals
- Excel automation
- MSXML 2.0
- RTSlink DLL for Tally integration
- BarTender 8.01
- `CL32.DLL` licensing/dongle style library
- `ENC.DLL` / `DEC.DLL`
- ActiveX controls:
  - `MSADODC.OCX`
  - `MSFLXGRD.OCX`
  - `MSCOMCT2.OCX`
  - `MSMASK32.OCX`
  - `TABCTL32.OCX`
  - `MSCOMCTL.OCX`
  - `RRW32.OCX`
  - `COMDLG32.OCX`
  - `SYSINFO.OCX`
  - `MSCOMM32.OCX`
  - `MSMAPI32.OCX`
  - `MSINET.OCX`
  - `RICHTX32.OCX`

## Top Business-Weight Forms By Line Count

These forms deserve first extraction priority:

| Form | Lines | Likely Area |
|---|---:|---|
| `M_FRM_MBL.frm` | 11,549 | mobile/HHT/market billing |
| `M_FRM_SAL.frm` | 11,009 | sales invoice, schemes, pricing |
| `M_FRM_SOR.frm` | 10,895 | sales order |
| `M_FRM_SSO.frm` | 10,733 | secondary sales/order |
| `M_FRM_RSI1.frm` | 9,112 | retailer sales invoice |
| `M_FRM_EVP.frm` | 8,920 | empties/van/payment style workflow |
| `M_FRM_SAL_PSH.frm` | 8,140 | sales variant |
| `M_FRM_RSI.frm` | 6,067 | retailer sales invoice variant |
| `M_FRM_ACT.frm` | 5,848 | account/customer master |
| `M_FRM_SCH1.frm` | 5,097 | scheme variant |
| `M_FRM_POR.frm` | 5,077 | purchase order |
| `M_FRM_PUR.frm` | 5,072 | purchase |
| `M_FRM_PPO.frm` | 5,071 | purchase order variant |
| `M_FRM_STO.frm` | 5,032 | stock out |
| `M_FRM_SRT.frm` | 5,029 | sales return/stock return |
| `M_FRM_STI.frm` | 5,021 | stock in |
| `M_FRM_PRT.frm` | 4,966 | party/print/product related |
| `M_FRM_PRD.frm` | 4,862 | product master |
| `M_VAN_SAL.frm` | 4,477 | van sales |

## High-Value Shared Modules

- `EFF_LIB.bas` and `EFF_LIB1.bas` - shared globals, ID generation, formatting, quantity conversions, stock updates, reports, ledger routines.
- `modGeneral.bas` - FoxPro connection helper.
- `ModTallyInterface.bas` - RTSlink/Tally XML send/get wrapper.
- `modProcess.bas` - Tally group/account/UOM/stock item/voucher export.
- `modXml.bas` - Tally XML payload builders.
- `modInnComm.bas` - serial SMS send/receive/delete and comm settings.
- `modSMSSend.bas` - SMS master record selection helpers.
- `modMessageTemplate.bas` - SMS template files.

## Current Destination

`../MigrateDest_v1/init_postgres.sql` is a generated Postgres migration with 96 `CREATE TABLE` statements and data inserts. It is useful as an extraction artifact, not a final schema.

Problems:

- raw DBF shape
- no reliable primary keys/foreign keys
- Visual FoxPro `_nullflags`
- legacy `VARCHAR(10)` IDs everywhere
- duplicate/sync tables retained as-is
- binary/control character artifacts in inserts
