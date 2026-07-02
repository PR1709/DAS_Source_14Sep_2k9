# 04 Business Workflows

## Sales Invoice Workflow

Primary source: `../LegacyCodebase/M_FRM_SAL.frm`

Likely user flow:

1. Start new sale.
2. Enter/select invoice number/date.
3. Select customer/retailer (`ACT`), salesman (`SMN`), route.
4. Optionally reference van challan (`VAN`) or sales order.
5. Enter salable SKU lines in grid `MS`.
6. Enter free/takeback/replacement SKU lines in grid `MSF`.
7. System calculates:
   - line rate
   - product amount
   - discounts
   - tax
   - coupon/CSR discount
   - free quantity amounts
   - gross amount
   - net amount
   - route-off rounding
8. System validates stock unless disabled by system settings or van reference.
9. Auto-apply schemes if enabled.
10. Save to invoice/header tables and print/report.

Important source anchors:

- `P_FRE_CAL` calculates free-grid totals.
- `V_PRD_LST` controls product selection and stock lookup.
- `AutoApplyScheme` validates and applies schemes.
- `PNV_PNL_SaveProcedure` writes salable/free/scheme lines to `PNV`.
- `ROU_OFF` handles rounding/route offset behavior.

## Product Selection Workflow

Source: `M_FRM_SAL.frm`, `V_PRD_LST`.

Rules:

- Requires `prd.dbf`, `brd.dbf`, `prt.dbf`, `dir.dbf`.
- Product lookup can be by name or short name.
- If a van challan is selected, products are filtered through `VAN` and available net load.
- Without van reference, product availability depends on `BRD.PRD_QTY` unless system stock-check setting disables it.
- Only certain `STK_TYP` are valid:
  - normal salable grid: `Salable`
  - free grid: `Free`, `Salable`, `Promotional Item`

## Scheme Workflow

Primary source: `M_FRM_SAL.frm`, `AutoApplyScheme`.

Flow:

1. Clear current scheme grids.
2. Validate eligible schemes:
   - free quantity
   - bill amount
   - multiple quantity
   - multiple quantity OR
   - multiple quantity OR mix
   - brand
   - pack size
   - pack type
   - segment
   - category
3. Mark all eligible schemes as checked.
4. For each checked scheme, dispatch based on scheme type and benefit type.
5. Add scheme benefit lines or discount amounts.
6. If scheme amount is positive:
   - set amount-less descriptor to `Scheme Amount`
   - set freight/amount-less field to scheme amount
   - lock those fields
7. Recalculate gross/net.

Scheme dispatch matrix:

| Scope | Benefit | Handler |
|---|---|---|
| SKU Level | Free Quantity | `ADDSKUQUANTIY` |
| Bill Level | Free Quantity | `ADDBILLQUANTIY` |
| Multiple SKU Level | Free Quantity | `ADDMLTQTYAND` |
| SKU Level | Cash Off | `ADDSKUAMOUNT` |
| SKU Level | Discount Percent | `ADDSKUPCG` |
| Bill Level | Cash Off | `ADDBILAMT` |
| Bill Level | Discount Percent | `ADDBILPCG` |
| Multiple SKU Level | Cash Off / Discount Percent | `MLTBILPCGAMT` |
| Brand | Free Quantity | `AddBrand` |
| Pack Size | Free Quantity | `AddPackSize` |
| Pack Type | Free Quantity | `AddPackType` |
| Segment | Free Quantity | `AddSegment` |
| Category | Free Quantity | `AddCategory` |
| Brand/Pack/Category/Segment | Cash Off / Discount Percent | `AddBillAmtGroupType` |

## Stock Update Workflow

Primary source: `EFF_LIB.bas`

Core idea:

- Stock can be represented as packed quantity with sub-units.
- Utility functions convert packed quantity to total sub-units and back.
- Receipt-like transaction types increase stock.
- Issue-like transaction types decrease stock.

Receipt-like transaction types found:

- `OPS`
- `PUR`
- `SRN`
- `STO`
- `RSR`
- `PCR`
- `STI`
- `DMR`

Issue-like transaction types found:

- `LSL`
- `SAL`
- `PRT`
- `STI`
- `RSI`
- `PCS`
- `STO`

Warning: transaction type semantics vary by routine. `STI` and `STO` appear in both contexts depending on report/update routine, so each workflow must be verified.

## HHT Import / Export Workflow

Primary source: `M_FRM_MIBS.frm`, `HHT/*.TXT`

Flow:

1. Open serial COM port via `MSCOMM32.OCX`.
2. Read HHT identity.
3. Validate HHT against `HHT` table.
4. Write/read flat files:
   - `INL_STR.TXT`
   - `INV_STR.TXT`
   - master files like `STOCK.TXT`, `PARTY.TXT`, `LOAD_SHEET.TXT`
5. Parse imported transaction lines.
6. Update:
   - `VAN`
   - `INL`
   - `IRT`
   - `PNL`
   - `INV`
   - `PNV`
   - `SNV`
   - `SHL`
   - `LOG`
   - `IDY`

## Tally Export Workflow

Sources:

- `ModTallyInterface.bas`
- `modProcess.bas`
- `modXml.bas`
- `ETLInterfaceForm.frm`

Flow:

1. Connect to Tally through `RTSLINK.DLL`.
2. Read data from DAS tables.
3. Convert to Tally XML:
   - groups
   - ledgers/accounts
   - UOM
   - stock groups
   - stock items
   - accounting vouchers
   - inventory vouchers
4. Send XML payload.
5. Parse response errors and display/log them.

## SMS Workflow

Sources:

- `modInnComm.bas`
- `modSMSSend.bas`
- `frmSMS.frm`
- `INDEXFILES/Setup/SMSMessageTemplate.txt`

Flow:

1. Load communication settings.
2. Use serial modem through `MSComm`.
3. Break long messages into chunks.
4. Send AT commands.
5. Read responses.
6. Write send/receive/delete logs.

## Report Workflow

Sources:

- `EFF_LIB.bas` report helpers
- `.RP1`, `.RP5`, `.RRW` files
- `RRI.DBF`

The application opens report libraries and documents with scope/filter/index parameters. Some print flows update `DIR` values and report runtime rows before shelling `pri.exe` or opening report documents.

Modern replacement should be report specs over database views, not report binary ports.
