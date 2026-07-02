# 05 Business Logic Extraction

This document captures concrete rules found in source. It is not yet the complete rulebook; see `09_continuation_queue.md`.

## Quantity Representation

Source: `../LegacyCodebase/EFF_LIB.bas`

The system represents case/unit quantities as decimal-looking values, where the decimal portion often means sub-units/pieces, not a mathematical fraction.

Example:

- `SUB_UNT = 24`
- quantity `1.005` means 1 case + 5 pieces
- total pieces = `1 * 24 + 5 = 29`

### `MTS(M_QTY, S_QTY)`

Meaning: convert mixed case-piece quantity to total sub-units.

Rule:

- if `S_QTY <> 1`:
  - `whole_cases = Int(M_QTY)`
  - `pieces = (M_QTY - Int(M_QTY)) * 1000`
  - total = `whole_cases * S_QTY + pieces`
- else total = `M_QTY`

Migration note:

Do not treat legacy quantity decimals as ordinary decimal units without checking `SUB_UNT`.

### `STM(M_QTY, S_QTY)`

Meaning: convert total sub-units back to mixed case-piece style quantity.

The function starts by tracking negative quantities and then formats quotient/remainder back into decimal-looking quantity.

Migration note:

This needs exact reimplementation with golden-master tests because stock quantities, invoice line quantities, and free goods all depend on it.

### `P_SUB_UNT(QTY, SUBQTY)` and `P_QTY_UNT(M_QTY, S_QTY)`

These are alternate helper routines for converting between packed quantity and total sub-units. They are used in newer stock update routines.

## Stock Quantity Helpers

Source: `EFF_LIB.bas`

### `CONQTY(EQTY, prdid)`

Subtracts entered quantity from current product level:

1. Load `sub_unt`, `cur_lev` from `PRD`.
2. Convert entered quantity to pieces using `MTS`.
3. Convert current level to pieces using `MTS`.
4. Return `STM(current_pieces - entered_pieces, sub_unt)`.

### `ADDQTY(EQTY, prdid)`

Adds entered quantity to current product level:

1. Load `sub_unt`, `cur_lev` from `PRD`.
2. Convert both quantities with `MTS`.
3. Return `STM(entered_pieces + current_pieces, sub_unt)`.

### `CONBRDQTY(EQTY, BRDID)` / `ADDBRDQTY(EQTY, BRDID)`

Same subtract/add behavior, but against batch quantity `BRD.PRD_QTY` joined through `PRD`.

### `VANADDQTY` / `VANCONQTY`

Adjusts return quantities for van/RSI flows by reading existing `INV.RET_QTY` for product/invoice/type `RSI`.

## Unit/Rate/Weight Calculations

Source: `EFF_LIB.bas`

### `RATCAL(UNT, RAT, SUBUNT)`

Calculates amount for a mixed quantity:

1. Split quantity into whole units and piece remainder.
2. Piece amount = `pieces * rate / subunit`.
3. Total amount = `whole_units * rate + piece_amount`.
4. Format with `convert(AMT(...))`.

Used in sales/free line value calculations.

### `WGTCAL(UNT, CONFAC, SUBUNT)`

Calculates weight/UOM:

`whole_units * conversion_factor + (pieces * conversion_factor / subunit)`

### `UNTCAL(UNT, prdid)`

Converts a UOM quantity back to case-piece quantity using `PRD.CON_FAC` and `PRD.SUB_UNT`. It warns if conversion factor and sub-units mismatch.

## Stock Rebuild Logic

Source: `EFF_LIB.bas`, `N_UPD_STK`, `O_UPD_STK`, `STL_LED`

### `N_UPD_STK`

Resets `PRD.CUR_LEV = 0`, then:

- Adds quantities for transaction types:
  - `OPS`
  - `PUR`
  - `SRN`
  - `STO`
  - `RSR`
  - `PCR`
- Subtracts quantities for:
  - `LSL`
  - `SAL`
  - `PRT`
  - `STI`
  - `RSI`
  - `PCS`

Uses `P_SUB_UNT` and `P_QTY_UNT`.

### `O_UPD_STK`

Older stock update:

- Adds `OPS`, `PUR`, `SRN`, `STI`, `RSR`
- Subtracts `LSL`
- Subtracts `SAL`, `RSI`, `STO`, `PRT`

Important ambiguity:

`STI` and `STO` can shift category depending on routine. Treat transaction type semantics as workflow-specific until fully extracted.

## Stock Ledger Logic

Source: `EFF_LIB.bas`, `STL_LED`

`STL_LED(FRM_DAT, to_dat, qry)`:

1. Deletes all rows from `STL`.
2. Inserts receipt-side ledger rows from joined `INL`, `INV`, `PRD` where:
   - `STK_UPD = 'Y'`
   - transaction type in `PUR`, `SRN`, `STI`, `DMR`, `OPS`
   - date between range
3. Inserts issue-side ledger rows from `INV`, `PRD`, `PGR` where:
   - `STK_UPD = 'Y'`
   - transaction type in `SAL`, `PRT`, `STO`
   - date between range
4. For salable/free lines, `ISS_QTY = PRD_QTY`.
5. For `TakeBack` or `Damage`, `RCT_QTY = PRD_QTY`.
6. Builds opening stock table `OPS` from movements before from-date.
7. Opens stock ledger report.

## Product Line Pricing In Sales

Source: `M_FRM_SAL.frm`, `V_PRD_LST`, `P_FRE_CAL`, `PNV_PNL_SaveProcedure`

When a product/batch is selected:

- Product line gets:
  - `PRD_IDY`
  - stock type
  - quantity initialized to zero
  - rate application from customer/account (`M_DBT_RAP`)
  - base/sale/MRP rate based on `M_DBT_RAP` and piece-mode setting
  - sales discount `SAL_DIS`
  - sales tax `SAL_TAX`
  - product name/short name
  - `SUB_UNT`
  - `CON_FAC`
  - `BAT_IDY`
  - manufacturing/expiry dates from `BRD`
  - group/category/class from `PGR`
  - batch name

Rate selection:

- If CSR is not applied:
  - `M_DBT_RAP = 'A'` uses `BAS_RAT`
  - `M_DBT_RAP = 'C'` uses `PRD_MRP`
  - else uses `PRD_SDR`
- If CSR is applied:
  - rate comes from `CSR.PRD_RAT` for customer and base batch.

Piece-mode (`pcs_chk.Value = 1`) affects whether rate is per case/unit or divided/multiplied by `SUB_UNT`.

## Free Grid Calculation

Source: `M_FRM_SAL.frm`, `P_FRE_CAL`

For each free grid row:

- If piece mode:
  - line amount = `RATCAL(quantity, rate, subunit)`
- Else:
  - line amount = `quantity * rate`
- Totals:
  - `F_UNT_TOT`
  - `F_AMT_TOT`

## Stock Check For Free/Replacement Lines

Source: `M_FRM_SAL.frm`, `TEXT5_KeyDown`

When entering free/replacement quantity:

- Validate sub-unit remainder is below `SUB_UNT`.
- Find existing invoice total pieces for same batch/invoice/type `SAL` and stock type in `Free`, `Salable`.
- Get already-entered sales stock with `GETSALSTK`.
- Convert candidate quantity to pieces with `MTS`.
- Load available batch quantity from `BRD.PRD_QTY`.
- If system setting `DIR.SYS/STK` is not set to disable stock check and no van reference is present:
  - reject if available stock minus current quantity would be negative.

## Scheme Auto-Apply

Source: `M_FRM_SAL.frm`, `AutoApplyScheme`

Algorithm:

1. Reset scheme grids and amount.
2. Run validators:
   - `VLD_FRE_QTY`
   - `VLD_BIL_AMT`
   - `VLD_MLT_QTY`
   - `VLD_MLT_QTY_OR`
   - `VLD_MLT_QTY_OR_MIX`
   - `VLD_BRAND`
   - `VLD_PackSize`
   - `VLD_PackType`
   - `VLD_Segment`
   - `VLD_Category`
3. Check all eligible list items.
4. Dispatch to handler according to scheme scope and benefit.
5. Mark `APPSCH = True`.
6. If `SCH_AMT > 0`, set amount-less field to `Scheme Amount`, lock it, and recalculate gross/net.

Important:

The individual scheme handlers still need full extraction. They likely contain the exact slab/pro-rata/free-line rules.

## Sales Save To PNV/PNL

Source: `M_FRM_SAL.frm`, `PNV_PNL_SaveProcedure`

The save procedure writes three categories of `PNV` rows:

1. Salable rows from `MS`
2. Free/takeback/replacement rows from `MSF`
3. Scheme-generated free/discount rows from `MSS`

For salable rows:

- `txn_typ = 'SAL'`
- `stk_upd = 'Y'`
- `stk_typ = 'Salable'`
- calculates line discount and tax
- calculates coupon/CSR amount
- stores batch, product group/category/class, MRP, product UOM, sub-units, total pieces
- stores `INC_TTL`, `LST_RAT`, `LST_DIS`, `CPN_AMT`

Tax rule:

- If `DIR SYS/DPV` setting is `T`:
  - trade tax base = `PRDNET - INCAMT`
- Else:
  - trade tax base = `PRDNET - (ADD_AMT + DIS_AMT + INCAMT)`
- tax amount = trade tax + coupon amount

For free rows:

- `stk_typ` comes from grid (`Free`, `Replc.`, etc.)
- tax and discounts are zero
- rate app is `B`
- stock update is `Y`

For scheme rows:

- `stk_typ = 'Free'`
- `sch_idy` is set
- amount is stored in `COL_AMT`
- `INC_TTL` forced to `0` to avoid ODBC error

## Route Offset / Rounding

Source: `M_FRM_SAL.frm`, `ROU_OFF`

Enabled only if `DIR` setting `SYS/ROU` has `DIR_MSG = 'R'`.

Behavior:

- Reads previous route offset from `INL` for same sale invoice.
- Inspects decimal paise part of freight/net amount.
- If paise is between 1 and 49, sets negative offset.
- If paise is 50 or above, rounds up and sets positive offset.
- Otherwise offset is zero.

This needs exact reproduction because it affects invoice totals.

## Scheme Reporting / LNV

Source: `M_FRM_SAL.frm`, `Fill_LNV_Batch`, `Scheme_Price_Update`, `LNV_Rpt_Update`

Rules:

- `Fill_LNV_Batch` fills blank `LNV.LEV_DES` with `INV.BAT_IDY` for sale invoice.
- `Scheme_Price_Update` computes purchase/sale free amount:
  - `P_FRE_AMT = LNV.INV_GRS * BRD.U_PUR_LPP`
  - `S_FRE_AMT = LNV.INV_GRS * BRD.U_SAL_LPP`
- `LNV_Rpt_Update` enriches `LNV` rows with:
  - `SCH.MER_NUM`
  - product `SUB_UNT`
  - product pack
  - group category

## ID Generation

Source: `EFF_LIB.bas`, `increment`, `T7increment`, `Tincrement`, `Sincrement`

Patterns:

- `IDY` stores numeric counters.
- `PRE` stores prefixes.
- `increment(fldval)`:
  - increments IDY field
  - loads prefix from `PRE`
  - pads with zeros to 10 characters
- `T7increment` pads to 7 digits.
- `Tincrement` pads transaction numbers to 7 digits.
- `Sincrement` pads to 10 digits.

Migration note:

Preserve legacy IDs as external/display IDs, but use database-generated surrogate IDs internally.
