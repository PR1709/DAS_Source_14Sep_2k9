# 10 Sales And Scheme Spec

Created: 2026-07-02

Scope: partial, source-cited extraction for P0 Sales Invoice and Scheme Engine. This is not a full parity claim. It captures the inspected behavior from `../LegacyCodebase/M_FRM_SAL.frm` first, with variant/scheme-master coverage queued for continuation.

## Coverage Status

Inspected deeply:

- `../LegacyCodebase/M_FRM_SAL.frm`
  - `GRS_CAL`
  - `saverecord`
  - `P_DIS_PRD`
  - `P_DIS_FPRD`
  - `P_DIS_CAL`
  - `P_FRE_CAL`
  - `V_PRD_LST`
  - `TEXT5_KeyDown`
  - toolbar save/print dispatch
  - van reference lookup helpers
  - stock helper functions used by schemes
  - primary `ADD*` and `VLD_*` scheme routines at summary/detail level
  - `AutoApplyScheme`
  - `PNV_PNL_SaveProcedure`
  - `Fill_LNV_Batch`
  - `Scheme_Price_Update`
  - `LNV_Rpt_Update`

Identified but not fully extracted in this pass:

- `../LegacyCodebase/M_FRM_SAL_CHG.frm` appears to be a sales change/rate variant with analogous routines starting near line 2105.
- `../LegacyCodebase/M_FRM_SAL_RAT.frm` appears to be another sales rate variant with analogous routines starting near line 2105.
- `../LegacyCodebase/M_FRM_SAL_PSH.frm` includes the same scheme engine family, including extra UOM validators, around lines 7225-8067.
- `../LegacyCodebase/M_FRM_SAL(CHG).frm` has now been source-diffed against `M_FRM_SAL_CHG.frm`; despite the filename, it declares itself as `M_FRM_SAL` and follows normal-sale `ord_idy`, print key, stock validation, and discount-total behavior in several important paths.
- `../LegacyCodebase/M_FRM_SCH.frm` scheme master starts event/procedure code around line 1057; save/copy/display routines are known at lines 1855, 1860, 1941, 2029, 2530, 2593, 2676, 2835, 2910, 2985, 3079, 3325, 3392.
- `../LegacyCodebase/M_FRM_SCH1.frm` and `../LegacyCodebase/M_FRM_SCHE.frm` have now been reachability-checked; SCH1 has real but uncompiled alternate scheme logic, and SCHE appears to be a control-only unused shell in this project.

Additional coverage added after first pass:

- `../LegacyCodebase/M_FRM_SCH.frm`
  - `SchemeCopy`
  - `save`
  - `DisplayScheme`
  - `P_DIS_PRD`
  - `SAVESKUFREEQUANTITY`
  - `SAVEMULTIPLEFREEQUANTITY`
  - `saveskucashordiscount`
  - `savemultiplecashordiscount`
  - `savebillcashordiscount`
  - `SAVEGRPTYP`
  - `SAVEMULTIPLECASHORPCG`
  - `SID`
- Variant scan:
  - `M_FRM_SAL_CHG.frm`, `M_FRM_SAL_RAT.frm`, `M_FRM_SAL(CHG).frm`, and `M_FRM_SAL_PSH.frm` have distinct SHA1 hashes and must not be assumed identical to `M_FRM_SAL.frm`.
  - `M_FRM_SAL_PSH.frm` contains scheme validators/applicators through `MLTBILPCGAMT` but no `AutoApplyScheme`, group validators/applicators, `PNV_PNL_SaveProcedure`, or LNV enrichment helper names by `rg`.

## Core Tables And Grids

Primary source: `../LegacyCodebase/M_FRM_SAL.frm`.

- Header source/write tables: `INL`, `PNL`, `IRT`; backups `BNL`, `XNL`.
- Line source/write tables: `INV`, `PNV`; backups `BNV`, `XNV`.
- Scheme reporting/write table: `LNV`; backup `CNV`.
- Master/lookups: `ACT`, `SMN`, `MAR`, `LOC`, `CAT`, `PRD`, `BRD`, `PGR`, `DIR`, `PRT`, `CSR`, `SCH`, `SID`, `VAN`, `RRI`.
- Main product grid `MS`: salable invoice lines.
- Free/replacement grid `MSF`: free, replacement, damage/takeback/breakage/leakage lines.
- Scheme-generated grid `MSS`: auto/manual scheme free/amount rows.
- Scheme list `SCHLST`: eligible scheme list.
- Scheme audit grid `SC`: scheme base/free/amount rows later written to `LNV`.

## `GRS_CAL`

Source: `../LegacyCodebase/M_FRM_SAL.frm:2911`.

Trigger/user action:

- Called after product/free-grid calculations, discount fields, freight/amount fields, route offset validation, and before save.

Tables read/written:

- Reads no tables directly.
- Calls `ROU_OFF`, which reads `DIR` and `INL`.

Validations:

- None directly.

Formulas:

- Product value: `G_PRD_VAL = AMT_TOT`; `T_PRD_VAL = AMT_TOT`.
- Product discount display: `G_PRD_DIS = DIS_TOT + ADD_TOT`.
- Taxable/product net basis: `T_PRD_DIS = AMT_TOT - (DIS_TOT + ADD_TOT)`.
- Tax display: `G_PRD_TAX = TAX_TOT`; `T_PRD_TAX = TAX_TOT + T_PRD_DIS`.
- Gross: `G_GRS_AMT = T_PRD_TAX`; `T_GRS_AMT = T_PRD_TAX`.
- Cash discount: `T_CSD_AMT = T_GRS_AMT - M_CSD_AMT`.
- Special discount: `T_SPD_AMT = T_CSD_AMT - M_SPD_AMT`.
- Tax/surcharge/add/less cascade:
  - `T_TAX_AMT = M_TAX_AMT + T_SPD_AMT`
  - `T_SCT_PCG = M_EXD_AMT + T_TAX_AMT`
  - `T_PKG_CST = M_PKG_CST + T_SCT_PCG`
  - `T_FRT_AMT = T_PKG_CST - M_FRT_AMT`
  - `T_FRE_AMT = T_FRT_AMT - Abs(F_AMT_TOT)`
- Net amount: after `ROU_OFF`, `T_NET_AMT = T_FRE_AMT + M_ROU_OFF`.

Side effects:

- Mutates many visible total controls.
- Recalculates route offset.

Assumptions/unresolved:

- Control names `G_*` and `T_*` are presentation totals, but they also feed posting fields.

Golden-master candidates:

- One-line invoice with no discounts/tax/free goods.
- Invoice with cash discount, special discount, surcharge, package cost, amount-less field, and free-grid amount.
- Route offset enabled/disabled.

## `saverecord`

Source: `../LegacyCodebase/M_FRM_SAL.frm:2948`.

Trigger/user action:

- Save toolbar button (`Toolbar1_ButtonClick` key `s`, lines 7599-7604).
- Ctrl+S (`Form_KeyDown`, line 5188).
- Some Enter-key paths call `saverecord` after final field entry.

Tables read:

- Presence check: `ACT`, `SMN`, `MAR`, `LOC`, `CAT`, `PRD`, `BRD`, `INL`, `INV`, `LNV`, `PNL`, `PNV`, `IRT`, `PGR`, `DIR`, `BNV`, `XNV`, `BNL`, `XNL`, `CBK` at lines 2950-2954.
- `DIR` for improper shutdown `SYS/IND`, GCC `SYS/GCC`, day closure, day closure tolerance `SYS/DCD`, purchase-rate check setting `SYS/PSW`, print/report settings, and report document selection.
- `INV` duplicate invoice check and old-line restore/backup reads.
- `VAN` for reference challan closed status.
- `ACT`, `SMN` validation.
- `BRD` purchase rate and batch stock/rates.
- `CSR` for customer-specific coupon/rate discount during line writes.
- `PRD`, `BRD`, `PGR`, `MAR`, `LOC`, `CAT`, `SMN` through direct lookups and posting enrichment.
- `RRI` for print runtime scope.

Tables written:

- `DIR`: creates day closure row if missing; updates print scope values (`INV/NUM`, `DOC/TYP`) before report.
- `INL`: sale header.
- `INV`: salable/free/scheme line rows; deletes existing sale rows on modify.
- `IRT`: invoice receipt/aging row.
- `PRD`, `BRD`: current quantity/pieces updates.
- `PNL`, `PNV`: print/report header and line mirror tables.
- `LNV`: scheme ledger/report rows.
- backup tables through `TxnMoveBackup`: `INV->BNV`, `LNV->CNV`, `INL->BNL`, `PNV->XNV`, `PNL->XNL`.
- `SCH_SKP.TXT`: append text log for skipped van scheme rows.
- Shell/deposit side effects through `AddShells` and `LOADSHELLS`.

Validations:

- User must confirm save.
- `DIR SYS/IND` must not equal `IMPROPER SHUTDOWN`.
- GCC validation must pass when `DIR SYS/GCC = T`.
- If system numbering is off and this is not modify mode, `INV` must not already contain same `inv_idy` and `txn_typ = 'SAL'`.
- Invoice number is required unless toolbar auto-number button 11 is pressed.
- Referenced van challan cannot be closed (`VAN.VAN_CLS = 'C'`).
- Customer/party must exist in `ACT` where `ACT_TYP` is `SAL` or `BOT`.
- Salesman must exist in `SMN`.
- Sales account must exist in `ACT` where `GRP_CLF like 'Sales'` and blank `ACT_TYP`.
- Net amount cannot be negative.
- Day closure blocks saving for closed dates; missing day row is inserted with `DIR_MSG = 'N'`.
- Day closure tolerance checks latest closed day plus `DIR SYS/DCD`.
- If old scheme quantity flags indicate schemes were cleared and `APPSCH = False`, user is prompted to reapply schemes.
- A free-only invoice is rejected: at least one salable grid line must have quantity if free grid has quantity.
- When `DIR SYS/PSW = Y`, sale rate cannot be below purchase rate. Piece mode compares grid rate to `BRD.PRD_PDR`; non-piece mode compares against `BRD.PRD_PDR / SUB_UNT`.
- Route is mandatory.
- Scheme stock check is skipped for 10-character van reference, otherwise `STOCKCHECK` must return zero.

Formulas:

- Header fields mirror `GRS_CAL` outputs into `INL` and `PNL`.
- Due date in `INL`: `DateAdd("d", m_crt_dat, M_INV_DAT)`; `PNL.due_dat` is written as current `Date`, not the same formula.
- Salable line `PRDNET`:
  - piece mode: `RATCAL(qty, rate, subunit)`.
  - non-piece mode: `qty * rate`.
- Trade discount: `(PRDNET - ADD_AMT) * discount_percent / 100`.
- Tax:
  - if `DIR SYS/DPV = T`: `(PRDNET - INCAMT) * tax_percent / 100`.
  - else: `(PRDNET - (ADD_AMT + DIS_AMT + INCAMT)) * tax_percent / 100`.
- Tax stored includes coupon/list discount amount: `tax_amt = TRDTAX + M_CPN_AMT`.
- Quantities are converted through `STM`, `MTS`, and `P_SUB_UNT` depending on piece mode.
- Scheme free purchase/sale value in `LNV`: free quantity times `BRD.U_PUR_LPP` and `BRD.U_SAL_LPP`.

Side effects:

- Modify mode first reverses prior stock for old `INV` rows, moves prior rows to backups, deletes old `INV`/`LNV`, and later rewrites.
- For normal salable/free stock types, save decrements `PRD.CUR_LEV`, `PRD.PRD_PCS`, and `BRD.PRD_QTY` when no van reference is present.
- For `Damage`, `TakeBack`, `Breakage`, `Leakage`, save increments stock.
- Scheme-generated free lines can be skipped against van reference free stock; skipped rows are logged to `SCH_SKP.TXT` and user gets a message.
- On successful save, `AddShells` is called with stock update `Y` when no van reference, otherwise `N`.
- Print/report may run immediately if `PRT_NME = M_FRM_SAL`.

Report/print calls:

- If print mode uses DOS setting (`DIR SYS/DOS = D`), row 7 of `RRI` is configured and `pri.exe` is shelled.
- Else `RRW_OpenDocument` is used when `DIR SYS/RRW = T`; otherwise `OpenDocument CURDIR\DOS_VCH.RP1` is used.
- `DayCls_PrintBill` uses the same print/report flow when save is blocked by closed day or closed van challan.

Assumptions/unresolved:

- `stk_upd` is referenced at lines 3176 and 3180 without qualifying `REC_SET!stk_upd`; source may rely on a global or implicit field binding. Verify exact VB6 runtime behavior.
- `m_crt_dat` source is outside the inspected slice.
- `PNL.due_dat = Date` while `INL.due_dat = DateAdd(...)` may be intentional report behavior or a bug.
- There is no explicit transaction boundary across `INL`, `INV`, `IRT`, `PRD`, `BRD`, `PNL`, `PNV`, `LNV`.

Golden-master candidates:

- New save with auto-number enabled and disabled.
- Modify existing invoice: verify stock reversal, backups, delete/rewrite behavior.
- Closed day save attempt still prints but does not alter data.
- Van challan closed status blocks save.
- Sale rate below purchase rate in piece and non-piece mode.
- Sale with van reference and scheme free stock shortage logs `SCH_SKP.TXT`.
- Damage/takeback free-grid line increments stock instead of decrementing.

## Product Selection: `P_DIS_PRD`, `P_DIS_FPRD`, `V_PRD_LST`

Sources:

- `P_DIS_PRD`: `../LegacyCodebase/M_FRM_SAL.frm:6745`.
- `P_DIS_FPRD`: `../LegacyCodebase/M_FRM_SAL.frm:6883`.
- `V_PRD_LST`: `../LegacyCodebase/M_FRM_SAL.frm:7053`.

Trigger/user action:

- Typing/selecting product in salable grid `MS` or free grid `MSF`.
- Toolbar edit/search key for product selection calls `V_PRD_LST` when `Text3` is active or `MS.Col = 0` (lines 7637-7639).

Tables read:

- `PRD`, `BRD`, `PGR` for product/batch/group data.
- `CSR` for customer-specific rates/discounts.
- `ACT` for `CST_TYP` HVO behavior.
- `VAN` for van-reference filtering and net load display.
- `DIR SYS/STK` for stock-check/filter setting.
- `PRT` for product number behavior.

Tables written:

- None directly. Writes grid state only.

Validations:

- Salable duplicate batch in `MS` is rejected.
- If a van reference exists but selected batch is not in `VAN`, the cell is colored red; direct selection can show "Product, SKU Not Available for Entry, Check Van Challan!"
- Expired/BBD batch shows a critical warning when `BRD_DOE <= Date`.
- Without `DIR SYS/STK = U` and without van reference, lookup filters `BRD.PRD_QTY > 0`.

Formulas/rate selection:

- Stock type defaults to `Salable` for `MS`.
- If `APPLYCSR = False`:
  - rate app `A`: `BAS_RAT`.
  - rate app `C`: `PRD_MRP`.
  - otherwise `PRD_SDR`.
  - non-piece mode divides selected case rate by `SUB_UNT`.
- HVO customer types `HVO-P` and `HVO-P-WO` force `APPLYCSR = False`.
- If `APPLYCSR = True`, rate comes from `CSR.B_PRD_RAT` or `CSR.PRD_RAT`; on-spot CSR rows set grid amount discount and percent discount.
- Product grid stores product, batch, MRP, rates, sale discount/tax, `SUB_UNT`, `CON_FAC`, `GRP_IDY`, `GRP_CAT`, `GRP_CLS`, batch name, `S_INC_DIS`, `LST_RAT`, `LST_DIS`.
- Van lookup net load display: `MTS(TOT_QTY) - (MTS(RET_QTY) + MTS(LKG_QTY) + MTS(BRK_QTY))`, then `STM` back for display.

Side effects:

- Disables piece-mode checkbox after a product is selected.
- Moves focus to quantity columns.

Golden-master candidates:

- Rate application A/C/default, with piece and non-piece mode.
- CSR on-spot discount rate and HVO customer exceptions.
- Van-referenced product selection shows net load and filters salable only.
- Expired batch warning.

## Calculation Routines: `P_DIS_CAL` And `P_FRE_CAL`

Sources:

- `P_DIS_CAL`: `../LegacyCodebase/M_FRM_SAL.frm:6947`.
- `P_FRE_CAL`: `../LegacyCodebase/M_FRM_SAL.frm:7033`.

Trigger/user action:

- Quantity/rate/discount changes in grids.
- Tab/focus transitions and Enter-key handlers call these before `GRS_CAL`.

Tables read:

- `DIR SYS/DPV`.
- `ACT.CST_TYP`.
- `CSR.DIS_AMT` for HVO-NP/HVO-NP-WO when CSR applies.

Tables written:

- None directly.

Validations:

- None directly.

Formulas:

- Salable line amount in piece mode: `RATCAL(qty, rate, subunit)`.
- Salable line UOM: `WGTCAL(qty, con_fac, subunit)` then `QTY`.
- Included amount: piece mode `RATCAL(qty, S_INC_DIS, subunit)`; non-piece `qty * S_INC_DIS`.
- List/coupon amount:
  - `LST_TOT = total_pieces_or_qty * (LST_RAT / SUB_UNT)`
  - `M_CPN_AMT = LST_TOT * LST_DIS / 100`
- If `DIR SYS/DPV = T`, product tax base excludes only included amount.
- Else product tax base excludes add amount and included amount, then percent discount is subtracted.
- Tax amount: rounded percent of tax base; total tax adds `M_CPN_AMT`.
- HVO-NP/HVO-NP-WO with CSR applies per-case `CSR.DIS_AMT / SUB_UNT` discount against entered pieces.
- Free grid amount:
  - piece mode: `RATCAL(free_qty, free_rate, subunit)`.
  - non-piece: `free_qty * free_rate`.

Side effects:

- Mutates grid amount, UOM, included amount, discount amount columns and total controls.
- Calls `GRS_CAL`.

Assumptions/unresolved:

- `ADD_TOT` actually accumulates percent discount amount (`ADDTOT`) despite the name suggesting additive amount.
- `DIS_TOT` is only incremented in specific modify/CSR branches; the exact meaning should be validated with fixtures.

Golden-master candidates:

- DPV `T` vs non-`T`.
- HVO-NP CSR discount per piece.
- `LST_RAT/LST_DIS` coupon/list amount affecting tax and total.
- Free grid amount in piece/non-piece modes.

## Free Quantity Entry: `TEXT5_KeyDown`

Source: `../LegacyCodebase/M_FRM_SAL.frm:7501`.

Trigger/user action:

- Press Enter in free-grid edit textbox `Text5`, usually for `MSF` quantity or UOM columns.

Tables read:

- `INV` existing same invoice/batch `tot_pcs`.
- `BRD.PRD_QTY`.
- `DIR SYS/STK`.

Tables written:

- None directly.

Validations:

- In piece mode, decimal remainder must be less than row `SUB_UNT`.
- For `MSF` stock type `Free` or `Replc.`, stock is checked unless `DIR SYS/STK = U` or a 10-character van reference exists.
- Available formula: `(existing_invoice_pieces + current_batch_pieces) - (entered_salable_pieces + candidate_free_pieces) >= 0`.

Formulas:

- Non-piece quantity is coerced to integer and converted with `STM`.
- UOM column entry calls `UNTCAL`, then re-enters quantity validation.
- Free UOM display uses `WGTCAL`.

Side effects:

- Updates `MSF` quantity/UOM cells, may append a blank row, recalculates free totals.

Golden-master candidates:

- Subunit overflow rejection, e.g. quantity `1.024` with `SUB_UNT=24`.
- Free/replacement stock shortage without van reference.
- Same data with van reference bypassing local stock check.

## Scheme Eligibility Validators

Sources:

- `VLD_FRE_QTY`: `../LegacyCodebase/M_FRM_SAL.frm:8537`.
- `VLD_BIL_AMT`: `../LegacyCodebase/M_FRM_SAL.frm:8592`.
- `VLD_MLT_QTY`: `../LegacyCodebase/M_FRM_SAL.frm:8631`.
- `VLD_MLT_QTY_OR`: `../LegacyCodebase/M_FRM_SAL.frm:8677`.
- `VLD_MLT_QTY_OR_MIX`: `../LegacyCodebase/M_FRM_SAL.frm:8725`.
- `VLD_MLT_UOM`: `../LegacyCodebase/M_FRM_SAL.frm:8775` (present but commented out in auto-apply).
- `VLD_MLT_UOM_OR`: `../LegacyCodebase/M_FRM_SAL.frm:8823` (present but commented out in auto-apply).
- Group validators `VLD_BRAND`, `VLD_PackSize`, `VLD_PackType`, `VLD_Segment`, `VLD_Category`: start at `../LegacyCodebase/M_FRM_SAL.frm:8872`, `8988`, `9103`, `9218`, `9333`.

Trigger/user action:

- Manual scheme button path around lines 4120-4129.
- `AutoApplyScheme` at lines 10584-10593.

Tables read:

- `SCH` for active date-bounded schemes.
- `SID` for scheme restriction by salesman, route, or customer.
- Product quantities from `MS` grid through helper functions.

Tables written:

- None. Populates `SCHLST`.

Validations/formulas:

- All validators require active `SCH.SCH_STS = 'Y'` and date range `SCH_BDT <= M_INV_DAT <= SCH_EDT`.
- SKU level free quantity:
  - scheme type `SKU Level`, `TXN_TYP='BAS'`.
  - line batch must match `SCH.BAT_IDY`.
  - entered quantity or total pieces must be `>= B_PRD_QTY`.
- Bill level:
  - scheme type `Bill Level`.
  - `B_BIL_AMT <= T_GRS_AMT`.
- Multiple SKU AND:
  - scheme type `Multiple SKU Level`, `B_CON_APP='AND'`.
  - every base row must satisfy `CHECKBATCH(BAT_IDY) >= B_PRD_QTY`.
- Multiple SKU OR:
  - `B_CON_APP='ORO'`.
  - sum of `CHECKBATCH` across base rows must be at least max base quantity for scheme.
- Multiple SKU OR mix:
  - `B_CON_APP='ORM'`.
  - same aggregate threshold pattern as OR.
- UOM product group validators use `CHECKUOM` instead of `CHECKBATCH`, but auto-apply currently comments them out.
- Group validators inspect `SCH_TYP` for Brand/Pack Size/Pack Type/Segment/Category and compare group values against grid product group/attributes. In the inspected validators, `AND` requires every `SCH` base group value to meet its `B_PRD_QTY`; `ORO` requires at least one group value to meet `MAX(B_PRD_QTY)`; `ORM` accumulates group values across rows until the max threshold is met (`../LegacyCodebase/M_FRM_SAL.frm:8872`, `../LegacyCodebase/M_FRM_SAL.frm:8988`, `../LegacyCodebase/M_FRM_SAL.frm:9103`, `../LegacyCodebase/M_FRM_SAL.frm:9218`, `../LegacyCodebase/M_FRM_SAL.frm:9333`).
- For `SID`-restricted schemes, eligibility is allowed if any matching `SID.SMN_IDY`, `SID.ROU_IDY`, or `SID.CST_IDY` exists for the current salesman, route, or customer. This gate appears in SKU, bill, multiple-SKU, UOM, and group validators in current main `M_FRM_SAL.frm` (`../LegacyCodebase/M_FRM_SAL.frm:8557`-`../LegacyCodebase/M_FRM_SAL.frm:8571`, `../LegacyCodebase/M_FRM_SAL.frm:8606`-`../LegacyCodebase/M_FRM_SAL.frm:8620`, `../LegacyCodebase/M_FRM_SAL.frm:8653`-`../LegacyCodebase/M_FRM_SAL.frm:8670`, `../LegacyCodebase/M_FRM_SAL.frm:8894`-`../LegacyCodebase/M_FRM_SAL.frm:8913`).

Side effects:

- `SCHLST` is cleared by `VLD_FRE_QTY`, then appended by each validator.

Golden-master candidates:

- Scheme active date edge cases on start/end dates.
- SID unrestricted vs restricted to salesman/route/customer.
- SKU-level threshold with piece-mode conversion.
- Multiple SKU AND all items satisfied vs one missing.
- OR/ORM aggregate threshold.

## Scheme Quantity And Stock Helpers

Sources:

- `GETSALSTK`: `../LegacyCodebase/M_FRM_SAL.frm:7716`.
- `CHECKBATCH`: `../LegacyCodebase/M_FRM_SAL.frm:7730`.
- `GETFRESTK`: `../LegacyCodebase/M_FRM_SAL.frm:7742`.
- `CHECKBATCHValue`: `../LegacyCodebase/M_FRM_SAL.frm:7756`.
- `GETSCHVAL`: `../LegacyCodebase/M_FRM_SAL.frm:8428`.
- `TXT_SCH_Keypress`: `../LegacyCodebase/M_FRM_SAL.frm:8437`.
- `STOCKCHECK`: `../LegacyCodebase/M_FRM_SAL.frm:8493`.
- `SCHEMEQTY` / `getSCHEMEQTY`: `../LegacyCodebase/M_FRM_SAL.frm:8515`, `../LegacyCodebase/M_FRM_SAL.frm:8521`.

Trigger/user action:

- Used by scheme validators/applicators, manual applied-scheme grid edits, and save-time stock validation.

Tables read:

- `BRD` for current batch stock in `STOCKCHECK`.
- `INV` for prior invoice stock when modifying.
- Otherwise these helpers read `MS`, `MSF`, and `MSS` grids.

Tables written:

- None directly.

Validations/formulas:

- `GETSALSTK(batch)` returns the first matching sale-grid `MS[2]` quantity for `MS[15]=batch`; in pieces mode it converts through `MTS(MS[2], MS[13])`, otherwise returns entered quantity (`../LegacyCodebase/M_FRM_SAL.frm:7716`-`../LegacyCodebase/M_FRM_SAL.frm:7727`).
- `CHECKBATCH(batch)` is similar but requires positive `MS[2]` and returns `Double`; it is the core quantity source for SKU/multiple-SKU thresholds (`../LegacyCodebase/M_FRM_SAL.frm:7730`-`../LegacyCodebase/M_FRM_SAL.frm:7739`).
- `GETFRESTK(batch, stock_type)` scans `MSF` for matching `MSF[11]` and `MSF[1]`; in pieces mode it converts `MSF[2]` through `MTS(MSF[2], MSF[9])` (`../LegacyCodebase/M_FRM_SAL.frm:7742`-`../LegacyCodebase/M_FRM_SAL.frm:7753`).
- `CHECKBATCHValue(batch)` returns sale value for matching `MS[15]`: in pieces mode `MTS(qty, subunit) * Round(rate / subunit, 2)`, otherwise `qty * rate` (`../LegacyCodebase/M_FRM_SAL.frm:7756`-`../LegacyCodebase/M_FRM_SAL.frm:7767`).
- `GETSCHVAL(batch)` returns displayed sale amount `MS[9]` for matching positive-quantity `MS[15]`; this feeds multiple-SKU amount/percent rules (`../LegacyCodebase/M_FRM_SAL.frm:8428`-`../LegacyCodebase/M_FRM_SAL.frm:8434`).
- `TXT_SCH_Keypress` limits manual applied-scheme quantity edits. Normal rows cannot exceed original calculated quantity mirror `MSS[15]`; `ORO`/`ORM` combinations compare against remaining scheme quantity through `SCHEMEQTY` and `getSCHEMEQTY`. Amount edits in `MSS[4]` subtract old value from `SCH_AMT`, replace it, then add the new value (`../LegacyCodebase/M_FRM_SAL.frm:8437`-`../LegacyCodebase/M_FRM_SAL.frm:8488`).
- `STOCKCHECK` walks applied scheme grid `MSS`; for each scheme free batch, it compares `BRD.PRD_QTY` converted through `MTS` plus old invoice stock on modify against free-grid demand, sale-grid demand, and scheme quantity. It returns the failing `MSS` row or `0` (`../LegacyCodebase/M_FRM_SAL.frm:8493`-`../LegacyCodebase/M_FRM_SAL.frm:8513`).
- `SCHEMEQTY` sums current applied quantity `MSS[2]` for a scheme; `getSCHEMEQTY` returns the positive `MSS[11]` value for a scheme, not a sum (`../LegacyCodebase/M_FRM_SAL.frm:8515`-`../LegacyCodebase/M_FRM_SAL.frm:8525`).

Side effects:

- `TXT_SCH_Keypress` mutates `MSS` and `SCH_AMT`; the rest are read-only helpers.

Assumptions/unresolved:

- `GETSALSTK` and `CHECKBATCH` return only the first matching sale batch row, not an aggregate. If duplicate sale-grid batch rows are possible, scheme thresholds may undercount.

Golden-master candidates:

- Duplicate batch rows in `MS` to verify first-row vs aggregate behavior.
- Manual edit of `MSS[2]` for normal, `ORO/ORO`, `ORO/ORM`, and `ORO/AND` row combinations.
- Stock-check fixture with sale, free, and scheme demand all on the same batch, both new and modify modes.

## Scheme Applicators

Sources:

- `ADDSKUQUANTIY`: `../LegacyCodebase/M_FRM_SAL.frm:7788`.
- `ADDBILLQUANTIY`: `../LegacyCodebase/M_FRM_SAL.frm:7860`.
- `ADDMLTQTYAND`: `../LegacyCodebase/M_FRM_SAL.frm:7892`.
- `ADDSKUMIX`: `../LegacyCodebase/M_FRM_SAL.frm:8098`.
- `ADDSKUAMOUNT`: `../LegacyCodebase/M_FRM_SAL.frm:8205`.
- `ADDSKUPCG`: `../LegacyCodebase/M_FRM_SAL.frm:8237`.
- `ADDBILAMT`: `../LegacyCodebase/M_FRM_SAL.frm:8271`.
- `ADDBILPCG`: `../LegacyCodebase/M_FRM_SAL.frm:8299`.
- `MLTBILPCGAMT`: `../LegacyCodebase/M_FRM_SAL.frm:8316`.
- Group applicators `AddBrand`, `AddPackSize`, `AddPackType`, `AddSegment`, `AddCategory`, `ADDSCITEMS`, `AddBillAmtGroupType`, `GETSCHGRPVAL`: start at `../LegacyCodebase/M_FRM_SAL.frm:9495`, `9676`, `9862`, `10039`, `10217`, `10404`, `10476`, `10563`.

Trigger/user action:

- Manual scheme dispatch around lines 4059-4073.
- `AutoApplyScheme` dispatch at lines 10603-10618.

Tables read:

- `SCH`, joined with `BRD` and `PRD` for free SKU details.
- `PRD`, `BRD`, `PGR` for group/batch/product enrichment.

Tables written:

- None directly. Populates `MSS` and `SC`; later `saverecord` writes `INV`, `PNV`, and `LNV`.

Validations/formulas:

- SKU free quantity:
  - Base row (`TXN_TYP='BAS'`) is matched against entered batch.
  - Free row (`TXN_TYP='FRE'`) provides free product/batch.
  - If `PRO_RAT='Y'`, applies proportional quantity using full multiples plus residual formula around lines 7807-7820.
  - Else free quantity is `Int(CHECKBATCH(base_batch) / base_qty) * free_qty`.
  - Adds `SC` row with base quantity consumed and `TXN_TYP='BAS'`.
- Bill free quantity:
  - Free quantity is `Int(T_GRS_AMT / B_BIL_AMT) * B_PRD_QTY`.
- Multiple SKU quantity:
  - AND branch computes `CTR` as number of complete sets across all base rows, then gives each free row `free_qty * (CTR - 1)`.
  - `ORO`/`AND`: each qualifying base batch whose `CHECKBATCH(batch) >= MAX(B_PRD_QTY)` writes one `SC` base row with consumed quantity `Int(batch_qty / max_base_qty) * max_base_qty`; every `F_CON_APP='AND'` free row receives `free_qty * Int(batch_qty / max_base_qty)` in `MSS[2]` and `MSS[15]`.
  - `ORO`/`ORO`: sums `GETSALSTK` across base rows into `STK`, converts to full sets with `Int(STK / MAX(B_PRD_QTY))`, and puts `MAX(free B_PRD_QTY) * STK` only on the first free row; later free rows receive zero.
  - `ORM` handling is delegated to `ADDSKUMIX` after `ADDMLTQTYAND` closes its recordset.
- SKU cash-off:
  - Base quantity from `CHECKBATCH`.
  - If `PRO_RAT='Y'`, amount is `(B_DIS_AMT / B_PRD_QTY) * sale_qty` when threshold met.
  - Else amount is `Int(sale_qty / base_qty) * B_DIS_AMT`.
  - Adds `SC` row with less amount and `TXN_TYP='FRE'`.
- SKU discount percent:
  - Uses `CHECKBATCHValue` as sale value and `PERCENT(sale_value, B_DIS_PCG)`.
  - Adds to `SCH_AMT` without checking `PRO_RAT` in active code; older proportional/non-proportional branch is commented out.
- Bill cash-off:
  - If `PRO_RAT='Y'`, full multiples plus residual prorated amount:
    - full `SCAMT = Int(SALAMT / BILAMT) * SCHAMT`
    - residual `BILMT = ((SALAMT - full_base) / BILAMT) * SCHAMT`
  - Else only full multiples.
- Bill discount percent:
  - `SCH_AMT += PERCENT(T_GRS_AMT, B_DIS_PCG)`.
- Multiple bill percent/amount:
  - Reads `CAS` or `PCG` rows for `DISAMT`/`DISPCG`.
  - AND branch computes complete base sets and applies percent to summed product values or `DISAMT * (CTR - 1)`.
  - ORO branch iterates base rows and applies percent to product value or amount per `MAXQTY` multiples.
  - ORM branch uses `CHECKBATCHValue`, `SKU_SBU`, and `PRD_SDR` in an unusual formula: `((Int(b) / Int(PNR_DRT)) * disamt)`.
- Group bill amount/percent:
  - `AddBillAmtGroupType` computes group value through `GETSCHGRPVAL`, then applies percent or cash-off for inspected `AND` and `ORO` branches. The deeper group applicator section below documents exact branch behavior and remaining `ORM` runtime gap.

Side effects:

- Adds lines to `MSS` for free SKU benefits.
- Adds rows to `SC` for scheme audit/amount rows.
- Mutates `SCH_AMT`.

Assumptions/unresolved:

- `ADDSKUQUANTIY` proportional residual formula at lines 7814-7818 is suspicious: `(ACTQTY / base_qty) / free_qty`. Preserve exactly until domain owners validate.
- Main `MLTBILPCGAMT` stores positive local `LSCH_AMT` into `SC[3]` for most branches, while PSH stores negative running deltas; see the main-vs-PSH delta section below.
- Some variables are reused globally (`STK`, `SCHQTY`, `PRDVAL`, `LSCH_AMT`) and may retain values if not reset in every branch.

Golden-master candidates:

- SKU free quantity with and without `PRO_RAT`.
- Bill free quantity.
- Multiple SKU AND with two required bases.
- Multiple SKU `AND/AND`, `AND/ORO`, `ORO/AND`, `ORO/ORO`, `ORM/AND`, and `ORM/ORO`.
- SKU cash-off and SKU percent discount.
- Bill cash-off with `PRO_RAT=Y` residual amount.
- Multiple bill percent and cash-off.
- Brand/pack/category/segment cash-off vs percent.

### Main Multiple-SKU Free Applicator Detail

Sources:

- `ADDMLTQTYAND`: `../LegacyCodebase/M_FRM_SAL.frm:7892`.
- `ADDSKUMIX`: `../LegacyCodebase/M_FRM_SAL.frm:8098`.
- Manual apply dispatch: `../LegacyCodebase/M_FRM_SAL.frm:4049`.
- PSH comparator: `../LegacyCodebase/M_FRM_SAL_PSH.frm:7552`.

Trigger/user action:

- Manual apply button `Command4_Click` dispatches checked `SCHLST` rows where `f_sch_typ='Free Quantity'` and `SCH_TYP='Multiple SKU Level'` to `ADDMLTQTYAND` (`../LegacyCodebase/M_FRM_SAL.frm:4056`-`../LegacyCodebase/M_FRM_SAL.frm:4064`).
- `AutoApplyScheme` dispatches the same scheme type to `ADDMLTQTYAND` (`../LegacyCodebase/M_FRM_SAL.frm:10603`-`../LegacyCodebase/M_FRM_SAL.frm:10609`).

Tables read:

- `SCH`, `BRD`, and `PRD` in one joined recordset for base/free rows, product names, batch names, MRP, sale rate, product group, `SUB_UNT`, and `PRD.CON_FAC` (`../LegacyCodebase/M_FRM_SAL.frm:7898`-`../LegacyCodebase/M_FRM_SAL.frm:7900`).
- Helper functions read `MS` for sale quantities/amounts.

Tables written:

- None directly. Writes `MSS` free rows and `SC` base rows.

Validations/formulas:

- `AND` base plus `AND` free:
  - Base rows are filtered with `B_CON_APP='AND'`; free rows use `F_CON_APP='AND'`.
  - `CTR` starts at `1`, loops repeatedly across all base rows, and increments after each complete pass. The first failing base threshold exits the loop.
  - Free quantity is `free B_PRD_QTY * (CTR - 1)` in `MSS[2]`; the same value is mirrored to `MSS[15]`; UOM display `MSS[16] = MSS[2] * PRD.CON_FAC / SUB_UNT`.
  - Each base row writes one `SC` row: `SC[0]=scheme`, `SC[1]=base BAT_IDY`, `SC[2]=base B_PRD_QTY * (CTR - 1)`, `SC[4]=BAS` (`../LegacyCodebase/M_FRM_SAL.frm:7902`-`../LegacyCodebase/M_FRM_SAL.frm:7952`).
- `AND` base plus `ORO` free:
  - After the `AND` base pass, the routine reopens all rows and adds every `F_CON_APP='ORO'` free row.
  - Unlike PSH, main currently gives every `ORO` free row `free B_PRD_QTY * (CTR - 1)` and sets `MSS[9]='AND'`, `MSS[14]='ORO'`; the older first-row-only logic is commented out (`../LegacyCodebase/M_FRM_SAL.frm:7954`-`../LegacyCodebase/M_FRM_SAL.frm:7993`).
- `ORO` base plus `AND` free:
  - If no `AND` base branch applies, `B_CON_APP='ORO'` rows are checked one by one against `MAX(B_PRD_QTY)`.
  - For each qualifying base batch, one `SC` row records consumed base quantity `Int(CHECKBATCH(batch) / max_base_qty) * max_base_qty`.
  - Each `F_CON_APP='AND'` free row receives `free B_PRD_QTY * Int(CHECKBATCH(base_batch) / max_base_qty)` in both `MSS[2]` and `MSS[15]`; `MSS[11]` is populated from the current base row's `B_PRD_QTY`, not the free row's quantity (`../LegacyCodebase/M_FRM_SAL.frm:7995`-`../LegacyCodebase/M_FRM_SAL.frm:8037`).
- `ORO` base plus `ORO` free:
  - The routine sums `GETSALSTK` across base rows into `STK`, then converts to full sets with `Int(STK / MAX(base B_PRD_QTY))`.
  - It writes `SC` base rows with actual `GETSALSTK(batch)` for each base row while accumulating.
  - Only the first free row gets `MAX(free B_PRD_QTY) * STK` in `MSS[2]` and `MSS[15]`; subsequent free rows get zero. `MSS[9]='ORO'`, `MSS[14]='ORO'`, and `MSS[11]` stores the total allocated free quantity (`../LegacyCodebase/M_FRM_SAL.frm:8039`-`../LegacyCodebase/M_FRM_SAL.frm:8093`).
- `ORM` base plus `AND` free:
  - `ADDMLTQTYAND` always calls `ADDSKUMIX` after completing the earlier branches.
  - `ADDSKUMIX` filters `B_CON_APP='ORM'`, sums `CHECKBATCH` over all base rows, and if `Int(STK / MAX(B_PRD_QTY)) > 0`, writes one `SC` base row per base batch with `Int(CHECKBATCH(batch))`.
  - Every `F_CON_APP='AND'` free row receives `free B_PRD_QTY * Int(STK / MAX(B_PRD_QTY))`; `MSS[9]='NotAPPLICABLE'` (`../LegacyCodebase/M_FRM_SAL.frm:8095`-`../LegacyCodebase/M_FRM_SAL.frm:8150`).
- `ORM` base plus `ORO` free:
  - `ADDSKUMIX` sums `GETSALSTK` across `ORM` base rows, writes `SC` base rows if at least one full set exists, then converts `STK` to full sets.
  - Only the first free row gets `MAX(free B_PRD_QTY) * STK`; later free rows get zero. `MSS[9]='ORO'`, `MSS[14]='ORO'`, and `MSS[11]` stores total allocated free quantity (`../LegacyCodebase/M_FRM_SAL.frm:8152`-`../LegacyCodebase/M_FRM_SAL.frm:8204`).

Side effects:

- Adds free-benefit rows to `MSS`; all line amounts are `0.00`.
- Adds base audit rows to `SC` for most branches; `AND`/`ORO` adds `SC` for base rows before adding `ORO` free rows, while PSH lacks comparable `SC` writes in its earlier multiple-free branch.

Assumptions/unresolved:

- `STK` is not locally declared or reset at entry in `ADDMLTQTYAND`/`ADDSKUMIX`; if it is module-level/global, prior branch values could leak. Golden-master fixtures should apply multiple schemes sequentially to verify.
- In `ORO`/`AND`, `MSS[11]` uses the base row quantity from `rs`, while `MSS[10]` uses the free row quantity from `RSTMP`; preserve this source behavior until runtime proves intent.
- `MSS[16]` UOM calculation sometimes reads `rs!CONFAC` after `rs` has moved to EOF, guarded only by `If rs.EOF = False` in some branches. Capture whether UI displays blank/stale UOM in these cases.

Golden-master candidates:

- Multiple-SKU free with `AND` base and `AND` free, exactly one complete set and multiple complete sets.
- `AND` base plus two `ORO` free rows: main should allocate quantity to every `ORO` free row; PSH older branch appears first-row-only.
- `ORO` base plus `AND` free with two qualifying base batches to confirm duplicated free rows per qualifying base batch.
- `ORO` base plus `ORO` free with multiple free rows to verify first-row-only allocation and zero later rows.
- `ORM` base plus `AND` free and `ORM` base plus `ORO` free; verify `STK` reset when two schemes are applied in one invoice.

### Main Group Scheme Applicators: Brand / Pack Size / Pack Type / Segment / Category

Sources:

- `AddBrand`: `../LegacyCodebase/M_FRM_SAL.frm:9495`.
- `AddPackSize`: `../LegacyCodebase/M_FRM_SAL.frm:9676`.
- `AddPackType`: `../LegacyCodebase/M_FRM_SAL.frm:9862`.
- `AddSegment`: `../LegacyCodebase/M_FRM_SAL.frm:10039`.
- `AddCategory`: `../LegacyCodebase/M_FRM_SAL.frm:10217`.
- `ADDSCITEMS`: `../LegacyCodebase/M_FRM_SAL.frm:10404`.
- `AddBillAmtGroupType`: `../LegacyCodebase/M_FRM_SAL.frm:10476`.
- `GETGRPVALUE` / `GETSCHGRPVAL`: `../LegacyCodebase/M_FRM_SAL.frm:9448`, `../LegacyCodebase/M_FRM_SAL.frm:10563`.

Trigger/user action:

- Manual scheme application dispatches selected `SCHLST` rows to these handlers for group-type `Free Quantity`, `Cash Off`, or `Discount Percent`.
- `AutoApplyScheme` reaches the same handlers for Brand, Pack Size, Pack Type, Segment, and Category rows after group validators list them (`../LegacyCodebase/M_FRM_SAL.frm:10620`-`../LegacyCodebase/M_FRM_SAL.frm:10630`).

Tables read:

- `SCH` for base/free scheme rows and group condition fields.
- `PRD`, `BRD`, and sometimes `PGR` for free-product display fields populated into `MSS`.
- `PRD` lookups through `GetValue` for brand name, pack size, and pack type comparisons.

Tables written:

- None directly. Free-benefit handlers populate `MSS`; `ADDSCITEMS` and `AddBillAmtGroupType` populate `SC`. Save later persists `MSS` into `INV`/`PNV` and `SC` into `LNV`.

Validations/formulas:

- `GETGRPVALUE` scans current salable grid `MS` and returns quantity for a group value. In piece mode it converts `MS[2]` through `MTS(..., MS[13])`; otherwise it uses entered quantity. Brand, pack size, and pack type compare against `PRD.brd_nme`, `PRD.prd_pck`, and `PRD.prd_pkt`; segment and category compare against `MS[20]` and `MS[19]` (`../LegacyCodebase/M_FRM_SAL.frm:9448`-`../LegacyCodebase/M_FRM_SAL.frm:9493`).
- For group free-quantity `AND`, the applicator cycles all `BAS` rows until one base group cannot satisfy `CTR * B_PRD_QTY`. Free rows receive `B_PRD_QTY * (CTR - 1)` in `MSS[2]` and original mirror `MSS[15]` (`../LegacyCodebase/M_FRM_SAL.frm:9501`-`../LegacyCodebase/M_FRM_SAL.frm:9527`, `../LegacyCodebase/M_FRM_SAL.frm:9682`-`../LegacyCodebase/M_FRM_SAL.frm:9709`, `../LegacyCodebase/M_FRM_SAL.frm:10239`-`../LegacyCodebase/M_FRM_SAL.frm:10249`).
- For group free-quantity `ORO`, the applicator uses `MAX(B_PRD_QTY)` as the denominator; each base group whose quantity is at least that max contributes `Int(group_qty / max_qty)` to `CTR`. Free rows receive `B_PRD_QTY * CTR` in `MSS[2]` and `MSS[15]` (`../LegacyCodebase/M_FRM_SAL.frm:9559`-`../LegacyCodebase/M_FRM_SAL.frm:9583`, `../LegacyCodebase/M_FRM_SAL.frm:9741`-`../LegacyCodebase/M_FRM_SAL.frm:9765`, `../LegacyCodebase/M_FRM_SAL.frm:10283`-`../LegacyCodebase/M_FRM_SAL.frm:10307`).
- For group free-quantity `ORM`, the applicator sums quantities across base group values, computes `CTR = Int(total_qty / MAX(B_PRD_QTY))`, then writes free rows as `B_PRD_QTY * CTR` (`../LegacyCodebase/M_FRM_SAL.frm:9618`-`../LegacyCodebase/M_FRM_SAL.frm:9640`, `../LegacyCodebase/M_FRM_SAL.frm:9800`-`../LegacyCodebase/M_FRM_SAL.frm:9828`, `../LegacyCodebase/M_FRM_SAL.frm:10345`-`../LegacyCodebase/M_FRM_SAL.frm:10367`).
- `AddPackSize` has a unique zero-denominator guard in the `ORM` branch: if `MAX(B_PRD_QTY)=0`, it forces `CTR=1`; the other group applicators do not show that guard in the inspected lines (`../LegacyCodebase/M_FRM_SAL.frm:9809`-`../LegacyCodebase/M_FRM_SAL.frm:9813`).
- Free rows written to `MSS` consistently use: product name, MRP, computed free quantity, sale rate, amount `0.00`, product ID, batch name, batch ID, `SCH_IDY` in `MSS[8]`, base quantity in `MSS[10]`, free/base quantity in `MSS[11]`, `SUB_UNT` in `MSS[12]`, product group in `MSS[13]`, and `F_CON_APP` in `MSS[14]` (`../LegacyCodebase/M_FRM_SAL.frm:9520`-`../LegacyCodebase/M_FRM_SAL.frm:9539`, `../LegacyCodebase/M_FRM_SAL.frm:9757`-`../LegacyCodebase/M_FRM_SAL.frm:9777`, `../LegacyCodebase/M_FRM_SAL.frm:10359`-`../LegacyCodebase/M_FRM_SAL.frm:10380`).
- When a scheme free condition is `ORO`, handlers overwrite `MSS[11]`, and Category also overwrites `MSS[15]`, with the last positive `SCHQTY` value for the scheme (`../LegacyCodebase/M_FRM_SAL.frm:9543`-`../LegacyCodebase/M_FRM_SAL.frm:9549`, `../LegacyCodebase/M_FRM_SAL.frm:10265`-`../LegacyCodebase/M_FRM_SAL.frm:10274`, `../LegacyCodebase/M_FRM_SAL.frm:10384`-`../LegacyCodebase/M_FRM_SAL.frm:10394`).
- `ADDSCITEMS` writes one `SC` base row per matching salable grid line: `SC[0]=scheme ID`, `SC[1]=MS[15]` base batch ID, `SC[2]=matched quantity`, `SC[4]=BAS`. It does not write `SC[3]` for base rows (`../LegacyCodebase/M_FRM_SAL.frm:10404`-`../LegacyCodebase/M_FRM_SAL.frm:10474`).
- `GETSCHGRPVAL` sums sale amount `MS[9]` for matching group values and feeds group cash/percent calculation (`../LegacyCodebase/M_FRM_SAL.frm:10563`-`../LegacyCodebase/M_FRM_SAL.frm:10578`).
- `AddBillAmtGroupType` handles only `AND` and `ORO` group cash/percent branches in the inspected source. It computes the same `CTR` pattern as free quantity, then applies either `PERCENT(sum_group_sale_value, B_DIS_PCG)` or `B_DIS_AMT * (CTR - 1)` for `AND`, and `B_DIS_AMT * CTR` for `ORO`; it writes one benefit row with `SC[0]=scheme ID`, `SC[3]=SCH_AMT - previous_SCH_AMT`, `SC[4]=FRE` (`../LegacyCodebase/M_FRM_SAL.frm:10488`-`../LegacyCodebase/M_FRM_SAL.frm:10523`, `../LegacyCodebase/M_FRM_SAL.frm:10525`-`../LegacyCodebase/M_FRM_SAL.frm:10556`).

Side effects:

- Free-quantity group handlers add `MSS` rows and base audit `SC` rows; they do not directly mutate `SCH_AMT` because free items carry amount `0.00`.
- Group cash/percent handler mutates cumulative `SCH_AMT` and adds an `SC` benefit row that later becomes `LNV.LNV_AMT`.

Assumptions/unresolved:

- Several SQL strings in group free handlers contain `WHER` instead of `WHERE` in the inspected source. The active runtime behavior must be verified because these paths may be protected by legacy data/provider tolerance or may fail at runtime for some group types (`../LegacyCodebase/M_FRM_SAL.frm:9517`, `../LegacyCodebase/M_FRM_SAL.frm:9884`, `../LegacyCodebase/M_FRM_SAL.frm:10297`).
- `SCHQTY`, `PRDVAL`, `stkcount`, `CTR`, and `INC` are not locally declared in these routines; preserve observed state-reset behavior until runtime fixtures prove whether stale values can leak between schemes.
- `AddBillAmtGroupType` has no inspected `ORM` cash/percent branch, even though validators can list group `ORM` schemes.

Golden-master candidates:

- Brand, Pack Size, Pack Type, Segment, and Category free-quantity schemes for `AND`, `ORO`, and `ORM`, including piece mode and non-piece mode.
- Category `ORO` free scheme to verify the extra `MSS[15]` overwrite compared with Brand/Pack Size.
- Pack Size `ORM` with `MAX(B_PRD_QTY)=0` to verify forced `CTR=1`.
- Group cash and group percent for `AND` and `ORO`; capture `SCH_AMT`, `SC`, `LNV`, and invoice net amount.
- Group `ORM` cash/percent fixture to confirm whether active code silently does nothing or another path applies it.

### Main vs PSH Multiple-SKU Amount/Percent Delta

Sources:

- Main `MLTBILPCGAMT`: `../LegacyCodebase/M_FRM_SAL.frm:8316`.
- PSH `MLTBILPCGAMT`: `../LegacyCodebase/M_FRM_SAL_PSH.frm:7898`.

Verified differences:

- Main resets `LSCH_AMT = 0` at entry, increments both cumulative `SCH_AMT` and local `LSCH_AMT`, and writes `SC[3]=LSCH_AMT` for `AND` and `ORO`. In `ORM`, it writes `SC[3]=LSCH_AMT` unless `SCAMT >= SCH_AMT`, where it writes `SCAMT - LSCH_AMT`; `SC[4]=FRE` is always set (`../LegacyCodebase/M_FRM_SAL.frm:8322`-`../LegacyCodebase/M_FRM_SAL.frm:8369`, `../LegacyCodebase/M_FRM_SAL.frm:8370`-`../LegacyCodebase/M_FRM_SAL.frm:8393`, `../LegacyCodebase/M_FRM_SAL.frm:8394`-`../LegacyCodebase/M_FRM_SAL.frm:8425`).
- PSH does not use `LSCH_AMT`. Its `AND` branch stores `SC[3]=previous_SCH_AMT - new_SCH_AMT`, and its `ORO` and `ORM` branches store `SC[3]=SCHAMT - SCH_AMT`. With normal positive discounts and `SCHAMT` not refreshed in the shown routine, these `SC[3]` values are negative while `SCH_AMT` itself increases (`../LegacyCodebase/M_FRM_SAL_PSH.frm:7934`-`../LegacyCodebase/M_FRM_SAL_PSH.frm:7942`, `../LegacyCodebase/M_FRM_SAL_PSH.frm:7944`-`../LegacyCodebase/M_FRM_SAL_PSH.frm:7962`, `../LegacyCodebase/M_FRM_SAL_PSH.frm:7963`-`../LegacyCodebase/M_FRM_SAL_PSH.frm:7981`).
- Main `ORM` cash calculation converts `MAXQTY` through `STM` but then uses `((Int(value) / Int(PRD_SDR)) * disamt)` for cash, not the commented `Int(value / max_qty) * disamt`; PSH `ORM` cash uses `Int(value / MAXQTY) * disamt` without the sale-rate denominator rewrite (`../LegacyCodebase/M_FRM_SAL.frm:8407`-`../LegacyCodebase/M_FRM_SAL.frm:8413`, `../LegacyCodebase/M_FRM_SAL_PSH.frm:7971`-`../LegacyCodebase/M_FRM_SAL_PSH.frm:7975`).
- Main `GETSCHVAL` reads main grid batch ID `MS[15]`, quantity `MS[2]`, and line amount `MS[9]`; PSH `GETSCHVAL` reads PSH grid batch ID `MS[20]`, quantity `MS[7]`, and line amount `MS[14]` (`../LegacyCodebase/M_FRM_SAL.frm:8428`-`../LegacyCodebase/M_FRM_SAL.frm:8435`, `../LegacyCodebase/M_FRM_SAL_PSH.frm:7986`-`../LegacyCodebase/M_FRM_SAL_PSH.frm:7993`).

Golden-master candidates:

- Same multiple-SKU cash and percent scheme in main and PSH with `AND`, `ORO`, and `ORM`; capture `SCH_AMT`, every `SC[3]`, saved `LNV.LNV_AMT`, and printed invoice totals.
- PSH fixture with pre-existing positive `SCH_AMT` before `MLTBILPCGAMT` to confirm the negative delta magnitude.
- Main `ORM` cash fixture where `CHECKBATCHValue`, `MAXQTY`, `SKU_SBU`, and `PRD_SDR` differ enough to expose the sale-rate-denominator behavior.

### Main vs PSH Eligibility And Stock Delta

Sources:

- Main group/SKU validators: `../LegacyCodebase/M_FRM_SAL.frm:8537`-`../LegacyCodebase/M_FRM_SAL.frm:9446`.
- PSH scheme trigger and validators: `../LegacyCodebase/M_FRM_SAL_PSH.frm:3926`, `../LegacyCodebase/M_FRM_SAL_PSH.frm:7225`-`../LegacyCodebase/M_FRM_SAL_PSH.frm:7426`.
- Main `STOCKCHECK`: `../LegacyCodebase/M_FRM_SAL.frm:8493`.
- PSH `STOCKCHECK`: `../LegacyCodebase/M_FRM_SAL_PSH.frm:8042`.

Verified differences:

- Main auto-apply runs SKU, bill, multiple-SKU, and all five group validators, then applies matching group free/cash/percent schemes. UOM validators exist but are commented out of auto-apply (`../LegacyCodebase/M_FRM_SAL.frm:10579`-`../LegacyCodebase/M_FRM_SAL.frm:10630`).
- PSH's observed scheme-search button runs only SKU, bill, and multiple-SKU validators; UOM validators are present but commented out, and group validators/applicators were not found by `rg` in PSH (`../LegacyCodebase/M_FRM_SAL_PSH.frm:3926`-`../LegacyCodebase/M_FRM_SAL_PSH.frm:3938`).
- Main validators gate restricted schemes through `SID`: if any `SID` row exists for the scheme, current salesman, route, or customer must match one of `SID.SMN_IDY`, `SID.ROU_IDY`, or `SID.CST_IDY`. This pattern is visible in group validators and earlier SKU/multiple validators (`../LegacyCodebase/M_FRM_SAL.frm:8899`-`../LegacyCodebase/M_FRM_SAL.frm:8918`, `../LegacyCodebase/M_FRM_SAL.frm:9014`-`../LegacyCodebase/M_FRM_SAL.frm:9033`, `../LegacyCodebase/M_FRM_SAL.frm:9359`-`../LegacyCodebase/M_FRM_SAL.frm:9378`).
- The inspected PSH validators contain no equivalent `SID` lookups, so a customer/salesman/route-restricted scheme can be eligible in PSH while hidden in main sales. This is a verified source delta, not a migration recommendation.
- Main `STOCKCHECK` iterates applied scheme rows `MSS` and therefore validates scheme free-batch stock demand in addition to salable/free-grid demand (`../LegacyCodebase/M_FRM_SAL.frm:8493`-`../LegacyCodebase/M_FRM_SAL.frm:8513`).
- PSH `STOCKCHECK` iterates sale grid `MS`, checks only each selected sale batch against free-grid demand and salable demand, and does not add separate `MSS` scheme-free demand in the inspected function (`../LegacyCodebase/M_FRM_SAL_PSH.frm:8048`-`../LegacyCodebase/M_FRM_SAL_PSH.frm:8062`).

Golden-master candidates:

- SID-restricted SKU, bill, and multiple-SKU schemes in main vs PSH: unmatched customer/salesman/route should be absent in main and present in PSH if thresholds match.
- UOM product group scheme in main and PSH manual/auto search: verify the commented validators are not invoked by the observed buttons.
- Scheme free batch with enough salable stock but insufficient scheme-free stock: main should fail through `MSS`-based `STOCKCHECK`; PSH behavior should be captured separately because its function checks sale-grid batches.

## `AutoApplyScheme`

Source: `../LegacyCodebase/M_FRM_SAL.frm:10579`.

Trigger/user action:

- Called after credit-day entry when scheme settings require auto-apply (lines 4890, 4896, 4919, 4925).
- Can also be reached through manual command path.

Tables read:

- Indirectly all tables read by validators and applicators: mainly `SCH`, `SID`, `PRD`, `BRD`, `PGR`.

Tables written:

- None directly.

Validations:

- Runs validators in order:
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
- UOM validators exist but are commented out in this routine.

Dispatch:

- Every eligible `SCHLST` item is checked, then dispatched:
  - `SKU Level` + `Free Quantity`: `ADDSKUQUANTIY`
  - `Bill Level` + `Free Quantity`: `ADDBILLQUANTIY`
  - `Multiple SKU Level` + `Free Quantity`: `ADDMLTQTYAND`
  - `SKU Level` + `Cash Off`: `ADDSKUAMOUNT`
  - `SKU Level` + `Discount Percent`: `ADDSKUPCG`
  - `Bill Level` + `Cash Off`: `ADDBILAMT`
  - `Bill Level` + `Discount Percent`: `ADDBILPCG`
  - `Multiple SKU Level` + `Cash Off` or `Discount Percent`: `MLTBILPCGAMT`
  - group `Free Quantity`: `AddBrand`, `AddPackSize`, `AddPackType`, `AddSegment`, or `AddCategory`
  - group `Cash Off` or `Discount Percent`: `AddBillAmtGroupType`
- The multiple-SKU and group dispatch cases are source-backed at `../LegacyCodebase/M_FRM_SAL.frm:10618`-`../LegacyCodebase/M_FRM_SAL.frm:10631`.

Side effects:

- Clears/initializes scheme grids and `SCH_AMT`.
- Sets `APPSCH = True`.
- If `SCH_AMT > 0`, sets `M_AMT_LES = "Scheme Amount"`, sets `M_FRT_AMT = SCH_AMT`, locks those fields, then calls `GRS_CAL`.

Golden-master candidates:

- Auto-apply with multiple simultaneously eligible schemes to capture ordering and stacking.
- Positive cash/percent scheme amount locks amount-less fields.
- Free-quantity-only scheme leaves amount-less behavior unchanged.

## `PNV_PNL_SaveProcedure`

Source: `../LegacyCodebase/M_FRM_SAL.frm:10683`.

Trigger/user action:

- Called by `saverecord` after writing `INL`, `INV`, `IRT`, stock updates, and `LNV` rows.

Tables read:

- `CSR` for coupon/rate discount fields.
- `PGR` for scheme-line group/category/class.
- `PRD` for scheme-line list rates/discounts.
- `DIR SYS/DPV` for tax base.

Tables written:

- `PNV` only. `PNL` header is written directly inside `saverecord`.

Validations:

- Only writes rows with nonblank product and positive quantity or amount.
- Scheme rows require `APPSCH = True`.

Formulas:

- Salable rows:
  - piece mode stores entered qty/rate/MRP directly.
  - non-piece mode stores `PRD_QTY = STM(qty, SUB_UNT)` and multiplies rate/MRP by `SUB_UNT`.
  - `M_CPN_AMT` repeats list/coupon amount formula from `P_DIS_CAL`.
  - `INCAMT` and `PRDNET` use `RATCAL` in piece mode.
  - Tax base follows `DIR SYS/DPV` exactly as in `saverecord`.
  - `COL_AMT = line_amount + Round(LST_TOT, 2)`.
  - `INC_TTL = INCAMT`, `STK_TYP='Salable'`, `STK_UPD='Y'`.
- Free rows:
  - `RAT_APP='B'`, discounts/tax zero, `STK_TYP` from `MSF`, `STK_UPD='Y'`.
  - `COL_AMT = free_line_amount + Round(LST_TOT, 2)`.
  - ASSUMPTION: `LST_TOT` is not reset in the free loop; verify whether this leaks previous salable line list amount.
- Scheme rows:
  - `RAT_APP='B'`, `STK_TYP='Free'`, `SCH_IDY` from `MSS`.
  - `COL_AMT = MSS amount`.
  - `INC_TTL = 0#` due to noted ODBC error.

Side effects:

- Mirrors invoice lines for print/report database.

Golden-master candidates:

- Compare `INV` vs `PNV` row fields for same invoice.
- Scheme free line with `INC_TTL=0`.
- DPV `T` vs non-`T` tax in `PNV`.

## Scheme Reporting Helpers

Sources:

- `Fill_LNV_Batch`: `../LegacyCodebase/M_FRM_SAL.frm:10925`.
- `Scheme_Price_Update`: `../LegacyCodebase/M_FRM_SAL.frm:10946`.
- `LNV_Rpt_Update`: `../LegacyCodebase/M_FRM_SAL.frm:10968`.

Trigger/user action:

- Called by `saverecord` immediately after `PNV_PNL_SaveProcedure`.

Tables read:

- `LNV`, `INV`, `BRD`, `PRD`, `PGR`, `SCH`.

Tables written:

- `LNV`.

Validations:

- None beyond record presence.

Formulas:

- `Fill_LNV_Batch`: if `LNV.LEV_DES` is blank, sets it to first `INV.BAT_IDY` for invoice and `TXN_TYP='SAL'`.
- `Scheme_Price_Update`: for free rows with nonblank `LEV_DES` and nonzero `INV_GRS`:
  - `P_FRE_AMT = LNV.INV_GRS * BRD.U_PUR_LPP`
  - `S_FRE_AMT = LNV.INV_GRS * BRD.U_SAL_LPP`
- `LNV_Rpt_Update`:
  - copies `SCH.MER_NUM`.
  - finds product through `BRD.PRD_IDY`, copies `PRD.SUB_UNT` and `PRD.PRD_PCK`.
  - copies `PGR.GRP_CAT`.

Side effects:

- Enriches scheme report rows after line posting.

Assumptions/unresolved:

- `Fill_LNV_Batch` uses the first matching `INV` row, not a row matching scheme product; this may be only a fallback for amount-only schemes.

Golden-master candidates:

- Amount-only scheme with blank `LEV_DES`.
- Free scheme with batch LPP values.
- Report row enrichment for product pack/category.

## Print Helpers

Sources:

- `Printrecord`: `../LegacyCodebase/M_FRM_SAL.frm:7294`.
- `DayCls_PrintBill`: `../LegacyCodebase/M_FRM_SAL.frm:10645`.
- Save-time print block: `../LegacyCodebase/M_FRM_SAL.frm:3767`.

Trigger/user action:

- Toolbar print key `p`.
- Ctrl+P.
- Save-time auto-print when `PRT_NME = M_FRM_SAL`.
- Day closure/closed van blocked save may still call `DayCls_PrintBill`.

Tables read/written:

- Reads `PRT`.
- Updates `DIR` print scope/doc type and `RRI` row 7 for DOS print.

Report calls:

- `M_FRM_PRV.Show 1` for preview.
- `pri.exe` for DOS print.
- `RRW_OpenDocument` for Report Writer path.
- `OpenDocument CURDIR\DOS_VCH.RP1` for RP1 path.

Golden-master candidates:

- `DIR SYS/DOS = D` with/without `pri.exe`.
- `DIR SYS/RRW = T`.
- Preview with `PRT_NME=M_FRM_SAL` true/false.

## Scheme Master: `M_FRM_SCH.frm`

Primary source: `../LegacyCodebase/M_FRM_SCH.frm`.

Purpose:

- Defines rows consumed by the sales invoice scheme engine.
- One logical scheme ID (`SCH_IDY`) is stored as multiple `SCH` rows.
- Row role is encoded by `TXN_TYP`:
  - `BAS`: base/eligibility leg.
  - `FRE`: free-product benefit leg.
  - `CAS`: cash-off benefit leg.
  - `PCG`: percent-discount benefit leg.
- Restriction rows are stored in `SID` and later checked by invoice validators against customer, salesman, or route.

### `SchemeCopy`

Source: `../LegacyCodebase/M_FRM_SCH.frm:1860`.

Trigger/user action:

- Scheme cascade/copy action on an existing scheme.

Tables read:

- `SCH` rows for selected `SCH_IDY`.
- `SID` rows for selected `SCH_IDY`.

Tables written:

- New `SCH` rows with a new `"SCH" & T7increment("idy_sch")`.
- New `SID` rows copied to the new scheme ID.

Validations:

- Requires a nonblank selected `Sch_Idy`.
- User must confirm copy/save.

Formulas/field mapping:

- Copies most `SCH` fields as-is: `SCH_NME`, `SCH_TYP`, `F_SCH_TYP`, batch/base/free product fields, quantity/amount/discount fields, `TXN_TYP`, `PRO_RAT`, `B_CON_APP`, `F_CON_APP`, `GRP_TYP`, `MER_NUM`, `MER_QTY`, `MER_AMT`.
- Resets `SCH_BDT`, `SCH_EDT`, and `SYS_DAT` to current `Date`; resets `SYS_TIM`; forces `SCH_STS = 'Y'`.

Side effects:

- Displays new scheme ID, then clears the form.

Assumptions/unresolved:

- Copy keeps the same scheme name, so users may rely on ID/date to distinguish copies.

Golden-master candidates:

- Copy unrestricted scheme.
- Copy restricted scheme with customer/salesman/route `SID` rows.

### `save`

Source: `../LegacyCodebase/M_FRM_SCH.frm:1941`.

Trigger/user action:

- Toolbar save or Enter-key path after entering scheme data.

Tables read/written:

- For existing `Sch_Idy`, writes existing `SCH` rows only:
  - `SCH_EDT`
  - `SCH_STS`
  - `CMP_DST`
  - `MER_NUM`
- For new schemes, dispatches to one of the specialized save routines below.

Validations:

- User must confirm save.
- `Sch_Nme` cannot be blank.

Dispatch matrix:

- `Sch_Cmb.ListIndex = 1`, `Sub_Cmb.ListIndex = 1`: `SAVESKUFREEQUANTITY`.
- `Sch_Cmb.ListIndex = 3`, `Sub_Cmb.ListIndex = 1`: `SAVEMULTIPLEFREEQUANTITY`.
- `Sch_Cmb.ListIndex = 3`, `Sub_Cmb.ListIndex > 1`: `SAVEMULTIPLECASHORPCG`.
- `Sch_Cmb.ListIndex = 1`, `Sub_Cmb.ListIndex > 1`: `saveskucashordiscount`.
- `Sch_Cmb.ListIndex = 2`: `savebillcashordiscount`.
- `Sch_Cmb.ListIndex > 3`: `SAVEGRPTYP`.

Assumptions/unresolved:

- Need to map combo indexes to display strings from `Form_Load` for machine-readable docs. Existing invoice code already sees strings like `SKU Level`, `Bill Level`, `Multiple SKU Level`, `Brand`, `Pack Size`, `Pack Type`, `Segment`, `Category`, `Free Quantity`, `Cash Off`, and `Discount Percent`.

Golden-master candidates:

- Existing scheme update only changes status/end date/merchandise number and does not rewrite detail rows.
- Each dispatch branch creates expected `SCH.TXN_TYP` row set.

### `DisplayScheme`

Source: `../LegacyCodebase/M_FRM_SCH.frm:2029`.

Trigger/user action:

- Selecting/viewing an existing scheme.

Tables read:

- `SCH`, `SID`, plus lookup display from `ACT`, `SMN`, `PRD`, `BRD`.

Tables written:

- None.

Behavior:

- Loads `SID` restrictions into `list1` and checks all listed rows.
- Sets scheme header controls from first `SCH` row: name, start/end date, active status, `CMP_DST`, `MER_NUM`, `MER_QTY`, `MER_AMT`.
- Selects `Sch_Cmb` and `Sub_Cmb` by matching stored text values.
- Disables scheme type, benefit type, start date, base/free tabs, and product grids for existing scheme display.
- Dispatches display:
  - group types: `dispschemegrt3`
  - SKU/multiple SKU: `DISPSCHEMESKU`
  - bill cash/percent: loads `TxtValue` and `TxtGrsAmt`
  - bill free quantity: filters `TXN_TYP='FRE'`, loads free grid and free condition combo.

Assumptions/unresolved:

- Several lookup SQL strings contain `WHER` instead of `WHERE`; verify whether this is dead path, tolerated by VFP, or typo masked by `On Error`.

Golden-master candidates:

- Display bill-level free quantity scheme.
- Display multiple SKU scheme with `B_CON_APP='ORM'`.
- Display customer-restricted scheme and route-restricted scheme.

### `P_DIS_PRD`

Source: `../LegacyCodebase/M_FRM_SCH.frm:2530`.

Trigger/user action:

- Product/batch selected for scheme base or free grid.

Tables read:

- `PRD`, `BRD`.

Tables written:

- None; grid only.

Validations:

- Duplicate batch in the target grid is rejected.
- Only `PRD.STK_TYP IN ('Salable')` is selectable.
- Historical comments show a removed `PRD_QTY > 0` filter; current code allows zero-stock products into schemes.

Grid mapping:

- `MSbase`/`MSFree` columns:
  - `0`: product name.
  - `1`: MRP.
  - `3`: batch ID.
  - `4`: product ID.
  - `5`: sub-units.

Golden-master candidates:

- Duplicate base SKU rejected.
- Zero-stock SKU can still be selected for scheme definition.

### `SAVESKUFREEQUANTITY`

Source: `../LegacyCodebase/M_FRM_SCH.frm:2593`.

Trigger/user action:

- New scheme save for SKU-level free quantity.

Tables written:

- `SCH`: two rows, inside `con.BeginTrans`/`CommitTrans`.
- `SID`: restrictions through `SID(schid)`.

Validations:

- Base grid row 1 quantity must be nonzero.
- Free grid row 1 quantity must be nonzero.

Encoding:

- New ID: `"SCH" & T7increment("idy_sch")`.
- Base row:
  - `SCH_TYP = Sch_Cmb`, `F_SCH_TYP = Sub_Cmb`.
  - `B_PRD_IDY = MSbase(1,4)`.
  - `B_PRD_QTY = MTS(MSbase qty, MSbase subunit)`.
  - `BAT_IDY = MSbase(1,3)`.
  - `TXN_TYP = 'BAS'`.
- Free row:
  - `B_PRD_IDY = MSFree(1,4)`.
  - `B_PRD_QTY = MTS(MSFree qty, MSFree subunit)`.
  - `F_PRD_IDY = MSbase(1,3)`.
  - `BAT_IDY = MSFree(1,3)`.
  - `TXN_TYP = 'FRE'`.
  - `PRO_RAT = 'Y'` when proportional checkbox is on, else `N`.
- Both rows include name/type/status/date/company-distributor/merchandise/user/system fields.

Assumptions/unresolved:

- `F_PRD_IDY` stores the base batch ID in this routine, despite the field name implying a product ID. Invoice applicator later filters `F_PRD_IDY` against sale grid batch ID, so preserve this behavior.

Golden-master candidates:

- SKU buy X pieces get Y free pieces, proportional off.
- Same with proportional on.

### `SAVEMULTIPLEFREEQUANTITY`

Source: `../LegacyCodebase/M_FRM_SCH.frm:2676`.

Trigger/user action:

- New scheme save for multiple-SKU free quantity.

Tables written:

- Multiple `SCH` base rows and free rows.
- `SID` restrictions.

Validations:

- At least one base grid row must have positive quantity.
- Free product check appears duplicated against `MSFree.Rows < 2`; verify because a blank two-row grid may pass row-count validation.

Encoding:

- Base rows:
  - one row per nonblank `MSbase`.
  - `TXN_TYP='BAS'`.
  - `B_PRD_QTY` is raw grid quantity, not `MTS`.
  - `B_CON_APP` from `CmbBase`: index `0 -> AND`, `1 -> ORO`, `2 -> ORM`.
- Free rows:
  - one row per nonblank `MSFree`.
  - `TXN_TYP='FRE'`.
  - `B_PRD_QTY` is raw grid quantity.
  - `F_CON_APP` from `CmbFree`: index `0 -> AND`, `1 -> ORO`.
  - `PRO_RAT='N'`.

Golden-master candidates:

- Multiple SKU AND with two base rows and one free row.
- Multiple SKU ORO and ORM base condition encoding.
- Free condition AND vs ORO.

### `saveskucashordiscount`

Source: `../LegacyCodebase/M_FRM_SCH.frm:2835`.

Trigger/user action:

- New scheme save for SKU-level cash-off or percent discount.

Tables written:

- `SCH`: base row plus one benefit row.
- `SID`.

Validations:

- Base quantity must be nonzero.

Encoding:

- Base row:
  - `B_PRD_IDY = MSbase(1,4)`.
  - `B_PRD_QTY = Val(MSbase(1,2))`.
  - `BAT_IDY = MSbase(1,3)`.
  - `TXN_TYP='BAS'`.
- Benefit row:
  - if `Sub_Cmb.ListIndex = 2`, `B_DIS_AMT = TxtValue`, `TXN_TYP='CAS'`.
  - if `Sub_Cmb.ListIndex = 3`, `B_DIS_PCG = TxtValue`, `TXN_TYP='PCG'`.
  - `BAT_IDY=''`, `F_PRD_IDY=''`.
  - `PRO_RAT` from proportional checkbox.

Golden-master candidates:

- SKU cash-off threshold.
- SKU percent-discount threshold.
- Proportional flag on/off.

### `savemultiplecashordiscount` And `SAVEMULTIPLECASHORPCG`

Sources:

- `savemultiplecashordiscount`: `../LegacyCodebase/M_FRM_SCH.frm:2910`.
- `SAVEMULTIPLECASHORPCG`: `../LegacyCodebase/M_FRM_SCH.frm:3325`.

Trigger/user action:

- Multiple-SKU cash-off/percent scheme save. `save` dispatches to `SAVEMULTIPLECASHORPCG`; `savemultiplecashordiscount` exists and is similar but may be older/dead path.

Tables written:

- `SCH` base rows plus one `CAS` or `PCG` benefit row.
- `SID` is called in `SAVEMULTIPLECASHORPCG`; `savemultiplecashordiscount` does not call `SID` in the inspected lines.

Encoding:

- Base rows:
  - `TXN_TYP='BAS'`.
  - `B_CON_APP` from `CmbBase`.
  - In `savemultiplecashordiscount`, `ORO/ORM` rows use `MSbase(1,2)` as every row's `B_PRD_QTY`, while `AND` uses each row's own quantity.
  - In `SAVEMULTIPLECASHORPCG`, every row uses its own quantity.
- Benefit row:
  - `CAS` with `B_DIS_AMT` for cash-off.
  - `PCG` with `B_DIS_PCG` for percent.
  - `BAT_IDY=''`.
  - `PRO_RAT` from proportional checkbox.

Assumptions/unresolved:

- `savemultiplecashordiscount` may be obsolete because `save` dispatch uses `SAVEMULTIPLECASHORPCG`. Keep both documented until call graph is confirmed.
- `SAVEMULTIPLECASHORPCG` does not wrap in `BeginTrans`/`CommitTrans`, unlike several other save routines.

Golden-master candidates:

- Multiple SKU cash-off with `AND`.
- Multiple SKU percent with `ORO`.
- Compare both routines if reachable from UI or shortcuts.

### `savebillcashordiscount`

Source: `../LegacyCodebase/M_FRM_SCH.frm:2985`.

Trigger/user action:

- Bill-level free quantity, cash-off, or percent scheme save.

Tables written:

- `SCH`; `SID`.

Validations:

- `TxtGrsAmt` must be nonzero.

Encoding:

- New scheme ID, one row per free grid row or one amount/percent row.
- All rows store `B_BIL_AMT = TxtGrsAmt`.
- Free quantity (`Sub_Cmb.ListIndex = 1`):
  - `BAT_IDY = MSFree(A,3)`.
  - `B_PRD_IDY = MSFree(A,4)`.
  - `B_PRD_QTY = MSFree(A,2)` as raw text/number.
  - `TXN_TYP='FRE'`.
- Cash-off (`Sub_Cmb.ListIndex = 2`):
  - `B_DIS_AMT = TxtValue`.
  - `TXN_TYP='CAS'`.
- Percent (`Sub_Cmb.ListIndex = 3`):
  - `B_DIS_PCG = TxtValue`.
  - `TXN_TYP='PCG'`.
- `PRO_RAT` from proportional checkbox.

Assumptions/unresolved:

- `rs.Update` is outside the `If MSFree... Or Sub_Cmb...` block; if no row was added for an iteration, this could be fragile. Needs runtime validation.

Golden-master candidates:

- Bill amount free product with multiple free rows.
- Bill cash-off and bill percent.
- Proportional bill cash-off.

### `SAVEGRPTYP`

Source: `../LegacyCodebase/M_FRM_SCH.frm:3079`.

Trigger/user action:

- Group-type scheme save: brand, pack size, pack type, segment, category, or similar `Sch_Cmb.ListIndex > 3`.

Tables written:

- `SCH`; `SID`.

Encoding:

- Base/group rows from `MSPack`:
  - `GRP_TYP = MSPack(A,0)`.
  - `B_PRD_QTY = MSPack(A,1)`.
  - `B_CON_APP` from `CmbPCon`: `AND`, `ORO`, or `ORM`.
  - `TXN_TYP='BAS'`.
  - For cash/percent group schemes, the base row also receives `B_DIS_AMT` or `B_DIS_PCG` and `TXN_TYP` is set to `CAS`/`PCG` before being overwritten to `BAS`; this means discount values may be present on `BAS` rows.
- If benefit type is free quantity, free rows from `MSFree`:
  - `B_PRD_IDY`, `BAT_IDY`, `B_PRD_QTY`.
  - `F_CON_APP` from `CmbFree`.
  - `TXN_TYP='FRE'`.
  - `PRO_RAT='N'`.

Assumptions/unresolved:

- For group cash/percent schemes, invoice `AddBillAmtGroupType` reads `CAS`/`PCG` rows, but `SAVEGRPTYP` as inspected appears to overwrite `TXN_TYP` to `BAS` in the base loop. Verify whether another path writes group `CAS`/`PCG` rows or whether this is a bug masked by copied fields.

Golden-master candidates:

- Brand free quantity scheme.
- Brand cash-off scheme, verifying whether a `CAS` row exists after save.
- Pack/category group scheme with `ORM`.

### `SID`

Source: `../LegacyCodebase/M_FRM_SCH.frm:3392`.

Trigger/user action:

- Called by most new scheme save routines after ID generation.

Tables written:

- `SID`.

Encoding:

- Iterates checked `list1` items.
- `Combo1.ListIndex = 1`: stores `CST_IDY`, blanks `SMN_IDY` and `ROU_IDY`.
- `Combo1.ListIndex = 2`: stores `SMN_IDY`, blanks `CST_IDY` and `ROU_IDY`.
- `Combo1.ListIndex = 3`: stores `ROU_IDY`, blanks `SMN_IDY` and `CST_IDY`.

Golden-master candidates:

- Same scheme restricted to one customer.
- Same scheme restricted to one salesman.
- Same scheme restricted to one route.

## Sales Variant Notes

Files scanned:

- `../LegacyCodebase/M_FRM_SAL_CHG.frm`
- `../LegacyCodebase/M_FRM_SAL_RAT.frm`
- `../LegacyCodebase/M_FRM_SAL_PSH.frm`
- `../LegacyCodebase/M_FRM_SAL(CHG).frm`

Verified observations:

- `M_FRM_SAL_CHG.frm`, `M_FRM_SAL_RAT.frm`, and `M_FRM_SAL(CHG).frm` are smaller sales variants with `GRS_CAL`, `saverecord`, `P_DIS_PRD`, `P_DIS_CAL`, `V_PRD_LST`, `PRINTRECORD`, and `ROU_OFF`, but no `AutoApplyScheme`, `VLD_*`, `ADD*`, `PNV_PNL_SaveProcedure`, `Fill_LNV_Batch`, `Scheme_Price_Update`, or `LNV_Rpt_Update` names found by targeted `rg`.
- `M_FRM_SAL_PSH.frm` includes `VLD_FRE_QTY`, `VLD_BIL_AMT`, `VLD_MLT_QTY`, `VLD_MLT_QTY_OR`, `VLD_MLT_QTY_OR_MIX`, `VLD_MLT_UOM`, `VLD_MLT_UOM_OR`, `ADDSKUQUANTIY`, `ADDBILLQUANTIY`, `ADDMLTQTYAND`, `ADDSKUMIX`, `ADDSKUAMOUNT`, `ADDSKUPCG`, `ADDBILAMT`, `ADDBILPCG`, and `MLTBILPCGAMT`, but no group scheme validators/applicators or `AutoApplyScheme` name by `rg`.
- `M_FRM_SAL_PSH.frm` uses shifted salable grid columns compared with `M_FRM_SAL.frm`: e.g. save/calculation references use `MS.TextMatrix(A,11)` for add amount, `12` for discount percent, `13` for tax, `18`/`20` for subunit/batch-related values in inspected search hits. Do not reuse the `M_FRM_SAL.frm` grid map for PSH.
- `M_FRM_SAL_PSH.frm` salable tax formula in inspected hits is simpler: `TRDTAX = (PRDNET - (ADD_AMT + DIS_AMT)) * tax / 100`, without the `DIR SYS/DPV`/`INCAMT` branch shown in main `M_FRM_SAL.frm`.

Golden-master candidates:

- Same invoice entered in `M_FRM_SAL.frm` and `M_FRM_SAL_PSH.frm` to compare totals and posted fields.
- Same SKU cash-off scheme in main sales and PSH, especially `MLTBILPCGAMT` because PSH appears to store negative/delta `SC` amounts differently in search hits.

## Alternate Scheme Forms

### `M_FRM_SCHE.frm`

Source: `../LegacyCodebase/M_FRM_SCHE.frm:1`.

Observed status:

- `M_FRM_SCHE.frm` defines a "Schemes" form with tabs for scheme details, SKU level, bill level, multiple SKU level, UOM product group, and view.
- It contains controls and a form attribute (`Attribute VB_Name = "M_FRM_SCHE"`), but targeted `rg` found no `Sub`, `Function`, `.Open`, `con.Execute`, `AddNew`, `UPDATE`, or `DELETE` code in this file.

Assumptions/unresolved:

- This may be an abandoned UI shell, a corrupted/truncated export, or behavior may have existed in a paired binary/resource file. Do not treat it as an active source of scheme logic until project membership/runtime usage is confirmed.

Golden-master candidates:

- None until runtime confirms whether this form is reachable.

### `M_FRM_SCH1.frm`

Source: `../LegacyCodebase/M_FRM_SCH1.frm`.

Status:

- Full alternate scheme master, older timestamp than `M_FRM_SCH.frm`.
- Uses a single large `save` routine rather than the decomposed save routines in `M_FRM_SCH.frm`.
- No `SID` write hits found in targeted search; this form appears not to save customer/salesman/route restrictions in the inspected save path.

#### `save`

Source: `../LegacyCodebase/M_FRM_SCH1.frm:1350`.

Trigger/user action:

- `CREATE` sets `MOD_REC = False` and calls `save` at lines 1345-1348.
- Toolbar save also likely calls `save`; full toolbar branch remains to be extracted.

Tables read:

- `SCH`.
- Product/batch grids are UI-derived; product lookup routines elsewhere in the file read product tables.

Tables written:

- `SCH`.

Validations:

- User must confirm save.
- `SCH.DBF` must exist.
- Scheme name cannot be blank.
- End date must be greater than or equal to start date.
- If bill-level tab is enabled, gross amount cannot be negative. Note: code checks `< 0`, not `<= 0`.
- For new SKU-level schemes, base product must be selected if the SKU tab is enabled.

Modify behavior:

- If `MOD_REC = True`, opens all `SCH` rows for the selected scheme ID and updates only:
  - `SCH_EDT`
  - `SCH_STS`
- Does not rewrite detail rows.

New scheme behavior:

- Opens `SCH`, begins a transaction, creates `schid = "SCH" & T7increment("idy_sch")`, and writes rows according to `SCH_CMB.ListIndex` and `SUB_CMB.ListIndex`.
- Commits only if `savecheck = True`; otherwise rolls back.

Encoding:

- `SCH_CMB.ListIndex = 0`, `SUB_CMB.ListIndex = 0`: SKU base quantity to free quantity.
  - Base rows from `BAS_PRD_VIW`: `TXN_TYP='BAS'`, `B_PRD_IDY`, `B_PRD_QTY`, `BAT_IDY`.
  - Free rows from `FRE_PRD_VIW`: `TXN_TYP='FRE'`, `B_PRD_IDY`, `B_PRD_QTY`, `F_PRD_IDY` set to base batch ID, `BAT_IDY`, `PRO_RAT` from checkbox.
- `SCH_CMB.ListIndex = 0`, `SUB_CMB.ListIndex = 1 or 2`: SKU base quantity to cash-off or percent.
  - Base rows: `TXN_TYP='BAS'`.
  - Benefit row: index `1 -> CAS` with `B_DIS_AMT`; index `2 -> PCG` with `B_DIS_PCG`; `PRO_RAT` from checkbox.
- `SCH_CMB.ListIndex = 1`: bill-level scheme.
  - `B_BIL_AMT = GRS_AMT`.
  - `SUB_CMB.ListIndex = 0`: free row with product/batch and `TXN_TYP='FRE'`.
  - `SUB_CMB.ListIndex = 1`: cash-off `TXN_TYP='CAS'`.
  - `SUB_CMB.ListIndex = 2`: percent `TXN_TYP='PCG'`.
- `SCH_CMB.ListIndex = 2`: multiple SKU level.
  - Base rows from `MS_BAS`: `TXN_TYP='BAS'`, `B_CON_APP='AND'` when `BAS_AND` is true, else `ORO` or `ORM` based on `Check1`.
  - Free rows from `MS_FRE`: `TXN_TYP='FRE'`, `F_CON_APP='AND'` when `FRE_AND` is true, else `ORO`.
  - If free grid row has no product but has amount/percent, writes `CAS` or `PCG` benefit row.
- `SCH_CMB.ListIndex = 3`: UOM product group.
  - Base rows from `BAS_MS`: `TXN_TYP='BAS'`, `B_CON_APP='AND'`, `ORO`, or `ORM`.
  - Free rows from `FRE_MS`: `TXN_TYP='FRE'`, `F_CON_APP='AND'` or `ORO`.
  - Empty-product free rows with amount/percent write `CAS`/`PCG`.

Formulas:

- Unlike `M_FRM_SCH.frm`, this save routine stores many quantities directly from grid columns without `MTS`.
- Several `ORO` branches use row 1 quantity/amount columns for all rows in inspected code.

Side effects:

- Calls `cancel`, `Schemelist`, `clrctr M_FRM_SCH`, resets `MOD_REC`, and disables toolbar button 12 after save.

Assumptions/unresolved:

- Combo index meanings differ from `M_FRM_SCH.frm`: in `M_FRM_SCH1.frm`, SKU level appears to be index `0`, bill level `1`, multiple SKU `2`, UOM product group `3`.
- The form resets `clrctr M_FRM_SCH` rather than `M_FRM_SCH1`, which may be a copy/paste bug or shared form naming quirk.
- Need to inspect `setval`, `Form_Load`, and grid headers to map all columns precisely.

Golden-master candidates:

- Same SKU-level free scheme created in `M_FRM_SCH.frm` and `M_FRM_SCH1.frm`, comparing `SCH` rows.
- Bill-level cash and percent schemes in `M_FRM_SCH1.frm`.
- Multiple SKU scheme with `BAS_AND`, `ORO`, and `ORM`.
- UOM product group scheme, since `M_FRM_SAL.frm` has UOM validators but auto-apply comments them out.

## Variant And Scheme Master Continuation Tasks

Status update from continuation pass:

- Partially advanced: `M_FRM_SAL_PSH.frm` save/calculation/grid/scheme behavior now has a deeper delta spec below.
- Later continuation update below deepens `M_FRM_SAL_CHG.frm`, `M_FRM_SAL_RAT.frm`, and `M_FRM_SAL(CHG).frm` save/calculation/report deltas.
- Later continuation update below confirms `M_FRM_SCH1.frm`/`M_FRM_SCHE.frm` reachability and SCH1 display/product-list paths.

## Continuation Pass: `M_FRM_SAL_PSH.frm` Invoice Variant Delta

Source inspected:

- `LegacyCodebase/M_FRM_SAL_PSH.frm`, targeted windows around `GRS_CAL`, `saverecord`, grid initialization, product display, product/free calculations, scheme validators, and scheme applicators.

### `GRS_CAL`

- File/lines: `LegacyCodebase/M_FRM_SAL_PSH.frm:2827`.
- Trigger/user action: Called after line/free-grid edits, header charges/discount/tax changes, route-off changes, save validation, and calculation helpers.
- Tables read: none directly.
- Tables written: none directly.
- Validations: none directly.
- Formula:
  - `G_PRD_VAL = AMT_TOT`, `T_PRD_VAL = AMT_TOT`.
  - `G_PRD_DIS = DIS_TOT + ADD_TOT`.
  - `T_PRD_DIS = AMT_TOT - (DIS_TOT + ADD_TOT)`.
  - `G_PRD_TAX = TAX_TOT`.
  - `T_PRD_TAX = TAX_TOT + T_PRD_DIS`.
  - `G_GRS_AMT = T_PRD_TAX`; `T_GRS_AMT = T_PRD_TAX`.
  - Header cash discount and special discount are treated as absolute amounts: `T_CSD_AMT = T_GRS_AMT - M_CSD_AMT`; `T_SPD_AMT = T_CSD_AMT - M_SPD_AMT`.
  - Header tax and surcharge amounts are additive: `T_TAX_AMT = M_TAX_AMT + T_SPD_AMT`; `T_SCT_PCG = M_EXD_AMT + T_TAX_AMT`.
  - Packing cost adds, freight subtracts, and free amount subtracts: `T_PKG_CST = M_PKG_CST + T_SCT_PCG`; `T_FRT_AMT = T_PKG_CST - M_FRT_AMT`; `T_FRE_AMT = T_FRT_AMT - Abs(F_AMT_TOT)`.
  - Calls `ROU_OFF`, then `T_NET_AMT = T_FRE_AMT + M_ROU_OFF`.
- Side effects: updates many total labels/textboxes; invokes round-off calculation.
- Report/print calls: none.
- Assumptions/unresolved:
  - Names ending in `_PCG` hold absolute amounts in this routine for cash discount/special discount/tax/surcharge.
  - `F_AMT_TOT` free stock value reduces invoice total, unlike ordinary charge rows.
- Golden-master candidates:
  - Invoice with line discount, line tax, free goods, packing, freight, surcharge, and round-off.
  - Header `M_CSD_AMT` and `M_SPD_AMT` nonzero with no percent values, confirming amount-not-percent behavior.

### `saverecord`

- File/lines: `LegacyCodebase/M_FRM_SAL_PSH.frm:2866`.
- Trigger/user action: Confirmed save from toolbar/keyboard; `Ctrl+S` reaches `saverecord` at `LegacyCodebase/M_FRM_SAL_PSH.frm:4796` and toolbar save reaches it at `LegacyCodebase/M_FRM_SAL_PSH.frm:7066`.
- Tables read:
  - Required table check: `act`, `smn`, `mar`, `loc`, `cat`, `prd`, `brd`, `inl`, `inv`, `lnv`, `pnl`, `pnv`, `irt`, `pgr`, `dir` at `LegacyCodebase/M_FRM_SAL_PSH.frm:2868`.
  - `ACT` party/account validation at `LegacyCodebase/M_FRM_SAL_PSH.frm:2889` and `LegacyCodebase/M_FRM_SAL_PSH.frm:2903`.
  - `SMN` salesman validation at `LegacyCodebase/M_FRM_SAL_PSH.frm:2896`.
  - `DIR` day-closure lookup/insert at `LegacyCodebase/M_FRM_SAL_PSH.frm:2919`.
  - Existing `INV`, `INL`, `LNV`, `PNV`, `PNL` rows during modification backup/rewrite at `LegacyCodebase/M_FRM_SAL_PSH.frm:3003`, `LegacyCodebase/M_FRM_SAL_PSH.frm:3019`, and `LegacyCodebase/M_FRM_SAL_PSH.frm:3105`.
  - `BRD`/`PRD_PDR` purchase-rate guard via `check`/`GetValue` when `SYS/PSW/DIR_MSG = Y` at `LegacyCodebase/M_FRM_SAL_PSH.frm:2967`.
- Tables written:
  - `DIR`: inserts day closure control row if missing at `LegacyCodebase/M_FRM_SAL_PSH.frm:2929`.
  - `INL`: invoice header add/update at `LegacyCodebase/M_FRM_SAL_PSH.frm:3026`.
  - `IRT`: receivable tracking row when net amount is positive at `LegacyCodebase/M_FRM_SAL_PSH.frm:3082`.
  - `INV`: salable lines, free/return lines, and scheme free lines at `LegacyCodebase/M_FRM_SAL_PSH.frm:3195`, `LegacyCodebase/M_FRM_SAL_PSH.frm:3269`, and `LegacyCodebase/M_FRM_SAL_PSH.frm:3334`.
  - `LNV`: deletes old rows on modify, inserts scheme benefit lines from `MSS` and scratch `SC` at `LegacyCodebase/M_FRM_SAL_PSH.frm:3024`, `LegacyCodebase/M_FRM_SAL_PSH.frm:3356`, and `LegacyCodebase/M_FRM_SAL_PSH.frm:3367`.
  - `PRD`, `BRD`: stock decrements/increments for salable, free, return/damage/takeback/breakage/leakage rows at `LegacyCodebase/M_FRM_SAL_PSH.frm:3254`, `LegacyCodebase/M_FRM_SAL_PSH.frm:3306`, and `LegacyCodebase/M_FRM_SAL_PSH.frm:3316`.
  - `PNL`, `PNV`: print/report mirror header and line rows at `LegacyCodebase/M_FRM_SAL_PSH.frm:3111`, `LegacyCodebase/M_FRM_SAL_PSH.frm:3381`, `LegacyCodebase/M_FRM_SAL_PSH.frm:3455`, and `LegacyCodebase/M_FRM_SAL_PSH.frm:3525`.
  - Backup tables: `BNV`, `CNV`, `BNL`, `XNV`, `XNL` through `TxnMoveBackup` at `LegacyCodebase/M_FRM_SAL_PSH.frm:3019` and `LegacyCodebase/M_FRM_SAL_PSH.frm:3105`.
- Validations:
  - User confirmation required.
  - Invoice ID required unless toolbar button 11 is pressed to auto-generate.
  - Party must be `ACT_TYP` `SAL` or `BOT`.
  - Salesman must exist.
  - Sales account must be an `act` row with `grp_clf = Sales` and blank `ACT_TYP`.
  - `T_NET_AMT` cannot be negative.
  - Closed day blocks save when matching `DIR` day row has `DIR_MSG = Y`.
  - If old schemes were cleared and not reapplied (`CHGQTY` says `Yes Schemes` or `Yes Schemes Qty` and `APPSCH = False`), prompts to re-apply schemes and exits if accepted.
  - Free rows cannot be saved without any sale products: sale grid total-stock presence is checked via `MS` column `4`, while free grid quantity uses `MSF` column `2`.
  - Optional purchase-rate guard blocks when billing/MRP rate is below purchase rate; source appears to reference `MS` columns `25` and `23`, which conflict with the grid header map below and needs runtime verification.
  - `STOCKCHECK` must return `0`.
- Formulas:
  - Salable `INV` line quantity stored from `MS` column `5` (`Bil.Stk`) and stock decremented by that same billed stock quantity, not visible `Quantity` column `7`.
  - Salable line gross/net basis uses `PRDNET = RATCAL(MS[7], MS[10], MS[18])` in pieces mode, otherwise `MS[7] * MS[10]`; however the line amount column is recalculated from `MS[5]` in `P_DIS_CAL`, creating an unresolved `Quantity` vs `Bil.Stk` discrepancy.
  - Free grid lines use `MSF[2]` quantity and `MSF[4]` rate, with `COL_AMT = MSF[5]`.
  - Scheme `MSS` lines are always written as `STK_TYP = Free`; quantity is `STM(MSS[2], MSS[12])`; line amount is `MSS[4]`; scheme ID is `MSS[8]`.
  - `LNV` receives both free-item rows from `MSS` and discount scratch rows from `SC`; the `SC` insert maps `LEV_IDY=SC[0]`, `LEV_DES=SC[1]`, `INV_GRS=SC[2]`, `LNV_AMT=SC[3]`.
- Side effects:
  - On modify, existing stock effects are reversed before line rewrite.
  - On modify, original transaction rows are moved to backup tables before delete/rewrite.
  - Writes separate print database copies to `PNL`/`PNV`, including customer/market/location/category names and system print remarks from `DIR`.
  - Sets `INL.SCH_CHK = Y`, `INL.SRN_AMT = G_FRE_AMT`, `PNL.SCH_CHK = Y`, `PNL.SRN_AMT = G_FRE_AMT`.
- Report/print calls:
  - Save prepares `PNL`/`PNV` print database rows.
  - If `PRT.PRT_NME = M_FRM_SAL_PSH`, save updates `DIR INV/NUM` to the invoice and `DIR DOC/TYP` to `SAL` (`LegacyCodebase/M_FRM_SAL_PSH.frm:3580`-`LegacyCodebase/M_FRM_SAL_PSH.frm:3582`).
  - If `DIR SYS/DOS = D`, PSH configures row 7 of `RRI` with invoice low/high scope, `PNV.CDX` as main index, one copy, and shells `pri.exe`; if `pri.exe` is missing, it warns instead of printing (`LegacyCodebase/M_FRM_SAL_PSH.frm:3583`-`LegacyCodebase/M_FRM_SAL_PSH.frm:3604`).
  - Otherwise PSH opens `DOS_VCH.RP1` with report title `Push Cart Challan (Cases)` when `pcs_chk.Value = 1`, else `Push Cart Challan (Pieces)`; both calls use `Inv.cdx,c,inv_idy` and invoice scope (`LegacyCodebase/M_FRM_SAL_PSH.frm:3606`-`LegacyCodebase/M_FRM_SAL_PSH.frm:3610`).
  - Ctrl+P dispatches to `Printrecord`; toolbar print dispatch reaches the same print path (`LegacyCodebase/M_FRM_SAL_PSH.frm:4795`, `LegacyCodebase/M_FRM_SAL_PSH.frm:7070`).
- Assumptions/unresolved:
  - `MS[5]` billed stock and `MS[7]` quantity are both used in financial/stock formulas. Need runtime examples to confirm intended meaning.
  - In non-pieces mode, one BRD update uses `MS.TextMatrix(A,18)` as the `BAT_IDY` predicate while passing `MS.TextMatrix(A,20)` to `CONBRDQTY` at `LegacyCodebase/M_FRM_SAL_PSH.frm:3261`; likely a defect or dead mode.
  - The purchase-rate guard references hidden columns that do not match the visible header map: `MS[25]` is `GRP_CLS` and `MS[23]` is `GRP_IDY`.
  - `PNV` free line `PRDNET` non-pieces formula multiplies `MSF[2] * MSF[2]` at `LegacyCodebase/M_FRM_SAL_PSH.frm:3480`; likely defect, but preserve unless proven otherwise.
- Golden-master candidates:
  - New PSH invoice with one salable row and one free row; compare `INL`, `INV`, `IRT`, `PNL`, `PNV`, `PRD`, `BRD`.
  - Modify same invoice and confirm stock is reversed once, backup tables receive old rows, and current rows are rewritten once.
  - Free-only invoice attempt should block.
  - Scheme-applied invoice should create both `INV` free rows and `LNV` rows.
  - PSH print flag on/off, DOS mode on/off, and pieces/cases switch; capture `DIR`, `RRI`, `pri.exe` shell attempt, and `DOS_VCH.RP1` title/scope.

### PSH Grid Field Maps

`MS` sale grid headers are initialized in `Form_Load` at `LegacyCodebase/M_FRM_SAL_PSH.frm:4938`.

| Grid | Column | Header / inferred meaning |
| --- | ---: | --- |
| `MS` | 0 | `Product ID` |
| `MS` | 1 | `Stock Type` hidden |
| `MS` | 2 | `Opg.Stk` |
| `MS` | 3 | `New Stk` |
| `MS` | 4 | `Ttl.Stk` |
| `MS` | 5 | `Bil.Stk` |
| `MS` | 6 | `Ret.Stk` |
| `MS` | 7 | `Quantity` hidden |
| `MS` | 8 | `UOM` hidden |
| `MS` | 9 | `RateID` |
| `MS` | 10 | `Rate` |
| `MS` | 11 | `Dis.Amt` |
| `MS` | 12 | `CD%` |
| `MS` | 13 | `Tax%` |
| `MS` | 14 | `Amount` |
| `MS` | 15 | `MRP` |
| `MS` | 16 | `Product Name` |
| `MS` | 17 | `Short Name` |
| `MS` | 18 | `UPC` / sub-units |
| `MS` | 19 | `uom` / conversion factor |
| `MS` | 20 | `BATIDY` |
| `MS` | 21 | `Mfg.Date` |
| `MS` | 22 | `BBD-Exp.Date` |
| `MS` | 23 | group ID hidden, written by `P_DIS_PRD` |
| `MS` | 24 | group category hidden, written by `P_DIS_PRD` |
| `MS` | 25 | group class hidden, written by `P_DIS_PRD` |
| `MS` | 26 | `Batch` |

`MSF` free/return grid headers are initialized at `LegacyCodebase/M_FRM_SAL_PSH.frm:5019`.

| Grid | Column | Header / inferred meaning |
| --- | ---: | --- |
| `MSF` | 0 | `Product ID` |
| `MSF` | 1 | `Stock Type` |
| `MSF` | 2 | `Quantity` |
| `MSF` | 3 | `UOM` |
| `MSF` | 4 | `Rate` |
| `MSF` | 5 | `Amount` |
| `MSF` | 6 | `MRP` |
| `MSF` | 7 | `Product Name` |
| `MSF` | 8 | `Short Name` |
| `MSF` | 9 | `UPC` / sub-units |
| `MSF` | 10 | `uom` / conversion factor |
| `MSF` | 11 | `BATIDY` |
| `MSF` | 12 | `GRPIDY` |
| `MSF` | 13 | `GRPCAT` |
| `MSF` | 14 | `GRPCLS` |
| `MSF` | 15 | `Batch` |

`MSS` applied scheme/free-benefit grid headers are initialized at `LegacyCodebase/M_FRM_SAL_PSH.frm:5062`.

| Grid | Column | Header / inferred meaning |
| --- | ---: | --- |
| `MSS` | 0 | `Product Name` |
| `MSS` | 1 | `MRP` |
| `MSS` | 2 | `Quantity` |
| `MSS` | 3 | `Rate` |
| `MSS` | 4 | `Amount` |
| `MSS` | 5 | `Product ID` |
| `MSS` | 6 | `Batch` |
| `MSS` | 7 | `batidy` |
| `MSS` | 8 | Header says `subunts`, but save/applicator code uses it as `SCH_IDY`; unresolved header/data mismatch. |
| `MSS` | 9 | Header says `grpid`, but applicator code uses it for base/free condition marker in some branches. |
| `MSS` | 10 | Header says `grpcat`, but applicator code uses it for base quantity. |
| `MSS` | 11 | Header says `grpclf`, but applicator code uses it for free quantity. |
| `MSS` | 12 | Hidden; applicator/save code uses as sub-units. |
| `MSS` | 13 | Hidden; applicator/save code uses as product group. |
| `MSS` | 14 | Hidden; applicator code uses as free condition (`F_CON_APP`) in some branches. |
| `MSS` | 15 | Hidden; duplicate original/free quantity used by scheme quantity editing checks. |

`SC` scratch scheme-discount grid is not visibly initialized in the inspected PSH window. Its write sites imply:

| Grid | Column | Inferred meaning |
| --- | ---: | --- |
| `SC` | 0 | Scheme ID (`SCHIDY`) |
| `SC` | 1 | Benefit batch/free product ID when available |
| `SC` | 2 | Free quantity/gross quantity |
| `SC` | 3 | Less/discount amount |
| `SC` | 4 | Type marker in `M_FRM_SAL.frm`; not populated by the inspected PSH cash/percent write sites |

### `P_DIS_PRD` and `P_DIS_FPRD`

- File/lines: `LegacyCodebase/M_FRM_SAL_PSH.frm:6417` and `LegacyCodebase/M_FRM_SAL_PSH.frm:6535`.
- Trigger/user action: Product/batch selection for sale grid (`P_DIS_PRD`) or free/return grid (`P_DIS_FPRD`).
- Tables read:
  - `PRD`, `BRD`, `PGR` joined by batch ID.
  - `CSR` for customer-specific rates when `APPLYCSR = True`.
- Tables written: none directly.
- Validations:
  - Required table check for `prd`, `brd`, `pgr`, `csr`.
  - `P_DIS_PRD` blocks duplicate batch (`BRDID = MS[20]`) within the sale voucher.
- Formulas:
  - `P_DIS_PRD` chooses rate from `bas_rat`, `prd_mrp`, or `prd_sdr` based on `M_DBT_RAP` (`A`, `C`, or other), optionally replacing with `CSR.PRD_RAT`.
  - In non-pieces mode, rate/MRP are divided by `sub_unt`.
  - `P_DIS_FPRD` defaults free-grid rate to `prd_sdr` or `CSR.PRD_RAT`, with the same sub-unit division behavior.
- Side effects:
  - Populates hidden batch/group metadata and disables `pcs_chk`.
  - `P_DIS_FPRD` moves focus into stock type selection by calling `MSF_keypress Asc("F")`, then showing `List7`.
- Report/print calls: none.
- Assumptions/unresolved:
  - Free-grid rate is still populated and valued even when stock type later becomes `Free`; `GRS_CAL` subtracts `F_AMT_TOT` from invoice totals.
- Golden-master candidates:
  - Same batch selected twice in PSH sale grid should block duplicate.
  - Customer-specific rate enabled/disabled for same customer/batch.

### `P_DIS_CAL` and `P_FRE_CAL`

- File/lines: `LegacyCodebase/M_FRM_SAL_PSH.frm:6596` and `LegacyCodebase/M_FRM_SAL_PSH.frm:6633`.
- Trigger/user action: sale/free grid edits, product select, quantity/rate/discount/tax changes.
- Tables read/written: none directly.
- Validations: none directly.
- Formulas:
  - `P_DIS_CAL` recalculates sale-grid `MS[14]` from billed stock `MS[5]` and rate `MS[10]`; in pieces mode it uses `RATCAL(MS[5], MS[10], MS[18])`.
  - `MS[8]` UOM is recalculated by `WGTCAL(MS[5], MS[19], MS[18])` in pieces mode, otherwise `MS[5] * (MS[19]/MS[18])`.
  - Product total begins as `MS[14] - MS[11]`; `CD%` (`MS[12]`) then subtracts; `Tax%` (`MS[13]`) then adds.
  - `DIS_TOT` accumulates only `MS[11]`; `ADD_TOT` accumulates the calculated `CD%` amount despite the name.
  - `AMT_TOT` accumulates gross line amount `MS[14]`, not tax-included `PRD_TOT`.
  - `P_FRE_CAL` recalculates `MSF[5]` from `MSF[2]` and `MSF[4]`, then accumulates `F_UNT_TOT` and `F_AMT_TOT`.
- Side effects:
  - `P_DIS_CAL` calls `GRS_CAL`; `P_FRE_CAL` does not call `GRS_CAL` in the inspected body, so callers often call both.
- Report/print calls: none.
- Assumptions/unresolved:
  - `ADD_TOT` appears to mean calculated cash discount amount, not add amount, in this PSH variant.
- Golden-master candidates:
  - Compare tax/discount totals where `MS[11]`, `MS[12]`, and `MS[13]` are all nonzero.
  - Confirm changing `MSF[2]` updates `F_AMT_TOT` and then total only after caller invokes `GRS_CAL`.

### PSH Scheme Validators and Applicators

- File/lines:
  - Validators: `VLD_FRE_QTY` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7225`, `VLD_BIL_AMT` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7273`, `VLD_MLT_QTY` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7298`, `VLD_MLT_QTY_OR` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7329`, `VLD_MLT_QTY_OR_MIX` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7362`, `VLD_MLT_UOM` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7395`, `VLD_MLT_UOM_OR` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7426`.
  - Applicators/helpers: `ADDSKUQUANTIY` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7468`, `ADDBILLQUANTIY` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7521`, `ADDMLTQTYAND` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7552`, `ADDSKUMIX` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7715`, `ADDSKUAMOUNT` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7795`, `ADDSKUPCG` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7824`, `ADDBILAMT` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7855`, `ADDBILPCG` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7882`, `MLTBILPCGAMT` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7898`, `GETSCHVAL` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7986`, `TXT_SCH_Keypress` at `LegacyCodebase/M_FRM_SAL_PSH.frm:7995`, `STOCKCHECK` at `LegacyCodebase/M_FRM_SAL_PSH.frm:8042`.
- Trigger/user action:
  - `Command6_Click` is the observed "find/validate schemes" trigger: it clears `MSS`, clears `SC`, resets `SCH_AMT`, resets the select-all caption, then runs `VLD_FRE_QTY`, `VLD_BIL_AMT`, `VLD_MLT_QTY`, `VLD_MLT_QTY_OR`, and `VLD_MLT_QTY_OR_MIX`; `VLD_MLT_UOM` and `VLD_MLT_UOM_OR` are present but commented out in this trigger (`LegacyCodebase/M_FRM_SAL_PSH.frm:3926`-`LegacyCodebase/M_FRM_SAL_PSH.frm:3938`).
  - `Command3_Click` toggles all `SCHLST` checkbox selections and changes caption between `Select A&ll` and `De-Select A&ll` (`LegacyCodebase/M_FRM_SAL_PSH.frm:3866`-`LegacyCodebase/M_FRM_SAL_PSH.frm:3878`).
  - `Command4_Click` clears prior applied rows, then dispatches checked `SCHLST` entries to applicators based on list subtype/type text (`LegacyCodebase/M_FRM_SAL_PSH.frm:3881`-`LegacyCodebase/M_FRM_SAL_PSH.frm:3912`).
  - `Command5_Click` finalizes applied schemes: sets `APPSCH=True`; if `SCH_AMT > 0`, marks `M_AMT_LES = "Scheme Amount"` and copies `SCH_AMT` to `M_FRT_AMT`; then calls `GRS_CAL` and returns to totals/route-offset flow (`LegacyCodebase/M_FRM_SAL_PSH.frm:3915`-`LegacyCodebase/M_FRM_SAL_PSH.frm:3924`).
- Tables read:
  - `SCH` heavily, filtered by active status, invoice date range, scheme type, transaction type, and condition flags.
  - `PRD`, `BRD` for free-product materialization.
  - No `SID` lookup was observed in the inspected PSH validator bodies, unlike current main `M_FRM_SAL.frm` validators.
- Tables written:
  - `MSS` grid for materialized free products.
  - `SC` scratch grid for cash/percent discounts.
  - Later `saverecord` persists `MSS` and `SC` to `INV`, `LNV`, and `PNV`.
- Validations:
  - Active schemes require `sch_sts = 'Y'`, `SCH_BDT <= M_INV_DAT`, and `SCH_EDT >= M_INV_DAT`.
  - SKU-level free quantity validates matching base `BAT_IDY` and quantity threshold, avoiding duplicates in `SCHLST` through `chkidy`.
  - Bill-level validates `b_bil_amt <= T_GRS_AMT`.
  - Multiple SKU `AND` requires every base batch to meet `B_PRD_QTY`.
  - Multiple SKU `ORO`/`ORM` sum available batch quantities and compare to `MAX(B_PRD_QTY)`.
  - UOM validators use `CHECKUOM`, which reads sale grid `MS[8]`; they are defined but not invoked by the observed command-button sequence.
  - PSH validators add eligible active schemes directly to `SCHLST` once quantity/amount thresholds pass; the inspected validator range contains no salesman/route/customer `SID` restriction gate (`LegacyCodebase/M_FRM_SAL_PSH.frm:7225`-`LegacyCodebase/M_FRM_SAL_PSH.frm:7426`).
  - Manual edits in `MSS` are constrained by `TXT_SCH_Keypress`: quantity edits in `MSS[2]` cannot exceed the original calculated quantity mirror `MSS[15]` unless the row is an `ORO` free/base combination, where it checks against remaining scheme quantity using `SCHEMEQTY`; amount edits in `MSS[4]` subtract the old amount from `SCH_AMT`, replace it, then add the new amount (`LegacyCodebase/M_FRM_SAL_PSH.frm:7995`-`LegacyCodebase/M_FRM_SAL_PSH.frm:8039`).
  - `STOCKCHECK` compares each selected sale batch's `BRD.PRD_QTY` plus old invoice stock on modify against current free/sale stock demand from `GETFRESTK` and `GETSALSTK`; it returns the failing grid row number or zero (`LegacyCodebase/M_FRM_SAL_PSH.frm:8042`-`LegacyCodebase/M_FRM_SAL_PSH.frm:8064`).
- Formulas:
  - `ADDSKUQUANTIY` proportional mode (`PRO_RAT='Y'`) computes whole applications plus extra remainder free quantity:
    - `APPQTY = Int(CHECKBATCH(base) / base_qty)`;
    - base consumed quantity `APP1QTY = APPQTY * base_qty`;
    - free whole quantity `APPQTY = APPQTY * free_qty`;
    - remainder `ACTQTY = CHECKBATCH(base) - APP1QTY`;
    - extra free `Int(ACTQTY / Int(base_qty / free_qty))`; final quantity is extra plus whole quantity.
  - Non-proportional SKU quantity is `Int(CHECKBATCH(base) / base_qty) * free_qty`.
  - Bill-level free quantity is `Int(T_GRS_AMT / b_bil_amt) * free_qty`.
  - Multiple SKU AND loops over base rows to compute count `CTR`; free quantity is `free_qty * (CTR - 1)`.
  - Multiple SKU ORO with AND free uses `free_qty * Int(CHECKBATCH(selected_base) / MAX(base_qty))`.
  - ORO/ORO and ORM/ORO branches aggregate stock, divide by max base quantity, and put the whole free quantity on the first free row while later rows receive zero.
  - SKU cash amount uses proportional cash when `PRO_RAT='Y'`, else whole multiples: `Round(Int(salqty / scheme_qty) * discount_amount)`.
  - SKU percent uses `PERCENT(CHECKBATCHValue(base), B_DIS_PCG)` directly; old proportional block is commented out.
  - Bill cash amount with `PRO_RAT='Y'` computes whole multiples plus remainder proportion; otherwise whole multiples only.
  - Bill percent uses `PERCENT(T_GRS_AMT, B_DIS_PCG)`.
  - Multiple-SKU cash/percent (`MLTBILPCGAMT`) first looks for a `CAS` row to get `B_DIS_AMT`, `PRO_RAT`, and max base quantity; if no cash row exists, it looks for a `PCG` row for `B_DIS_PCG` (`LegacyCodebase/M_FRM_SAL_PSH.frm:7904`-`LegacyCodebase/M_FRM_SAL_PSH.frm:7912`).
  - `MLTBILPCGAMT` `AND` branch computes complete base sets with the same `CTR - 1` pattern as free quantity; percent applies to summed `GETSCHVAL(base)` sale values, while cash applies `disamt * (CTR - 1)` (`LegacyCodebase/M_FRM_SAL_PSH.frm:7914`-`LegacyCodebase/M_FRM_SAL_PSH.frm:7942`).
  - `MLTBILPCGAMT` `ORO` branch iterates base rows; percent repeatedly applies to the accumulating `PRDVAL`, while cash applies `disamt` or `Int(batch_qty / max_qty) * disamt` for qualifying rows (`LegacyCodebase/M_FRM_SAL_PSH.frm:7943`-`LegacyCodebase/M_FRM_SAL_PSH.frm:7962`).
  - `MLTBILPCGAMT` `ORM` branch uses `CHECKBATCHValue` instead of `CHECKBATCH` for base amount/value and otherwise mirrors the `ORO` cash/percent loop (`LegacyCodebase/M_FRM_SAL_PSH.frm:7964`-`LegacyCodebase/M_FRM_SAL_PSH.frm:7983`).
  - `GETSCHVAL` returns `MS[14]` for the matching sale grid batch `MS[20]` when visible quantity `MS[7] > 0`; this value feeds percent/cash multiple-SKU amount rules (`LegacyCodebase/M_FRM_SAL_PSH.frm:7986`-`LegacyCodebase/M_FRM_SAL_PSH.frm:7992`).
- Side effects:
  - Materialized free rows store `SCH_IDY` in `MSS[8]`, despite the header saying `subunts`.
  - Cash/percent applicators accumulate global `SCH_AMT` and append discount rows to `SC`.
  - `MLTBILPCGAMT` appends `SC` rows with `SC[0]=scheme ID` and `SC[3]` set from the delta calculation; in the `AND`, `ORO`, and `ORM` branches the stored delta is `SAMT - SCH_AMT` or `SCAMT - SCH_AMT`, which can be negative after `SCH_AMT` increases (`LegacyCodebase/M_FRM_SAL_PSH.frm:7939`-`LegacyCodebase/M_FRM_SAL_PSH.frm:7981`).
- Report/print calls: none directly.
- Assumptions/unresolved:
  - UOM validators exist in PSH but are commented out in the observed validation sequence, matching the earlier unresolved behavior in `M_FRM_SAL.frm`.
  - `ADDSKUQUANTIY` uses `rs.filter = "F_PRD_idy = sale batch id and txn_typ like 'FRE'"`, so `SCH.F_PRD_IDY` appears to link a free row back to its base batch.
  - Several `MSS` header labels are stale relative to actual write/save use.
  - Because PSH does not show `SID` filtering in the inspected validators, customer/salesman/route-restricted schemes may apply differently between `M_FRM_SAL.frm` and `M_FRM_SAL_PSH.frm`; verify with runtime data before unifying eligibility rules.
- Golden-master candidates:
  - SKU free scheme with `PRO_RAT=Y`, sale quantity just above whole threshold plus remainder.
  - Bill free scheme with invoice gross just below, equal to, and above multiple thresholds.
  - Multiple SKU `AND`, `ORO`, and `ORM` with two base batches and two free rows to capture row distribution behavior.
  - SKU cash/percent and bill cash/percent schemes, verifying `SC` to `LNV` persistence.
  - PSH multiple-SKU cash and percent with `AND`, `ORO`, and `ORM`; verify whether `SC[3]` negative deltas are persisted exactly to `LNV.LNV_AMT`.
  - Manual edit of applied scheme quantity and amount in `MSS`; verify `SCH_AMT` recalculation, invalid quantity rejection, and later `INV`/`LNV` save.
  - PSH stock shortage with modify mode old invoice quantity present; verify `STOCKCHECK` returns the expected grid row and save blocks.
  - Same `SID`-restricted scheme in main sale and PSH sale; confirm whether main hides it while PSH lists/applies it.

## Continuation Pass: `M_FRM_SAL_CHG.frm`, `M_FRM_SAL_RAT.frm`, and `M_FRM_SAL(CHG).frm` Targeted Delta

Scope note: this pass inspected targeted save/calculation/product-list hits for the smaller CHG/RAT variants. It does not claim full extraction of all UI event handlers.

### Shared Smaller-Variant `GRS_CAL`

- File/lines:
  - `LegacyCodebase/M_FRM_SAL_CHG.frm:2105`
  - `LegacyCodebase/M_FRM_SAL_RAT.frm:2105`
  - `LegacyCodebase/M_FRM_SAL(CHG).frm:2114` from previous procedure inventory and targeted hits.
- Trigger/user action: Called from line edits, discount/tax/freight/round-off changes, save, and product calculation.
- Tables read/written: none directly.
- Validations: none directly.
- Formula:
  - Gross is `M_PRD_AMT`.
  - Cash discount is percentage-based: `M_CSD_AMT = Round(M_PRD_AMT * M_CSD_PCG / 100, 2)`.
  - Special discount is percentage-based on post-cash-discount amount: `M_SPD_AMT = Round(M_CSD_DED * M_SPD_PCG / 100, 2)`.
  - Header tax is percentage-based on post-special-discount amount.
  - Excise/extra discount amount `M_EXD_AMT` subtracts from tax-included amount.
  - Freight amount is derived from `(taxamt + M_TAX_AMT) * M_FRT_TYP / 100`, then added.
  - Packing cost is added.
  - Calls `ROU_OFF`.
- Side effects: updates header totals and round-off basis.
- Report/print calls: none.
- Assumptions/unresolved:
  - `M_FRT_TYP` acts as a percentage rate here and is stored into `INL.EXD_PCG`, not simply a freight type.
  - This differs materially from `M_FRM_SAL.frm` and `M_FRM_SAL_PSH.frm`, where several header amount fields are treated as amounts.
- Golden-master candidates:
  - Same line amount with nonzero cash discount percent, special discount percent, tax percent, freight type percent, packing, and excise amount across `M_FRM_SAL_CHG.frm`, `M_FRM_SAL_RAT.frm`, and main `M_FRM_SAL.frm`.

### `M_FRM_SAL_CHG.frm` `saverecord`

- File/lines: `LegacyCodebase/M_FRM_SAL_CHG.frm:2123`.
- Trigger/user action: Save confirmation from toolbar/keyboard.
- Tables read:
  - `ACT` dealer/party validation.
  - `SMN` salesman validation.
  - `act` sales-account validation.
  - Existing `INV`/`INL` rows on modify.
  - `PRD.SUB_UNT` while writing each invoice line.
  - `PRT` for print trigger.
- Tables written:
  - `INL` header.
  - `INV` line rows.
  - `PRD.CUR_LEV` decremented via `CONQTY`.
  - `DIR` invoice/doc settings before report print.
- Validations:
  - User confirmation.
  - Dealer must be `ACT_TYP` `BOT` or `SAL`.
  - Salesman must exist.
  - Sales account must have `grp_clf like 'Sale%'` and blank `ACT_TYP`.
  - `M_NET_AMT` must be nonnegative.
- Formulas:
  - Modify reverses prior `INV.PRD_QTY` to `PRD.CUR_LEV` using `ADDQTY`, deletes prior `INV`, then rewrites.
  - New invoice ID is `SAL` + `T7increment("IDY_SAL")`.
  - Header `INL.ord_idy = "2"` in CHG.
  - Line fields map: `MS[0]` product ID, `MS[1]` list rate, `MS[2]` list discount, `MS[3]` quantity, `MS[4]` rate, `MS[6]` rate app, `MS[7]` add discount percent, `MS[8]` trade/cash discount percent, `MS[9]` tax percent, `MS[10]` collected amount.
  - `PRDNET = RATCAL(MS[3], MS[4], PRD.SUB_UNT)`.
  - `ADD_AMT = PRDNET * MS[7] / 100`.
  - `DIS_AMT = (PRDNET - ADD_AMT) * MS[8] / 100`.
  - `TAX_AMT = (PRDNET - (ADD_AMT + DIS_AMT)) * MS[9] / 100`.
- Side effects:
  - On print flag `PRT_NME = M_FRM_SAL_CHG`, updates `DIR` invoice/doc values and opens `VOUCH.RP1`.
  - Calls `Cancelrecord` at end.
- Report/print calls:
  - `OpenDocument CURDIR & "\VOUCH.RP1", "Invoice", 2, Trim(INC), Trim(INC), 2, CURDIR & "\inv.cdx,c,inv_idy"` at `LegacyCodebase/M_FRM_SAL_CHG.frm:2248`.
- Assumptions/unresolved:
  - No `BRD` batch stock update in inspected save path; only `PRD.CUR_LEV` is updated.
  - No free/scheme tables or `LNV` writes found by targeted search.
- Golden-master candidates:
  - CHG invoice save with one row; verify `INL.ord_idy = "2"` and report trigger key uses `M_FRM_SAL_CHG`.
  - Modify CHG invoice; verify prior stock is reversed and line rewritten.

### `M_FRM_SAL_RAT.frm` `saverecord`

- File/lines: `LegacyCodebase/M_FRM_SAL_RAT.frm:2123`.
- Trigger/user action: Save confirmation from toolbar/keyboard.
- Tables read/written, validations, and line formulas: same as `M_FRM_SAL_CHG.frm` in the inspected save window.
- Variant-specific differences:
  - Header `INL.ord_idy = "1"` at `LegacyCodebase/M_FRM_SAL_RAT.frm:2178`.
  - Print flag checks `PRT_NME = M_FRM_SAL_RAT` at `LegacyCodebase/M_FRM_SAL_RAT.frm:2245`.
- Report/print calls:
  - Same `OpenDocument` call to `VOUCH.RP1` as CHG when print flag is present.
- Assumptions/unresolved:
  - Need project/menu inspection to identify business meaning of `ord_idy = "1"` vs `"2"`; current evidence suggests CHG/RAT are two order/report buckets with otherwise near-identical save behavior.
- Golden-master candidates:
  - Same data saved through RAT and CHG forms, diffing only expected `INL.ord_idy`, `PRT_NME` behavior, and any UI default differences.

### Smaller-Variant `P_DIS_CAL` Discrepancy

- File/lines:
  - `LegacyCodebase/M_FRM_SAL_CHG.frm:3967`
  - `LegacyCodebase/M_FRM_SAL_RAT.frm:3933` from targeted hit output.
  - `LegacyCodebase/M_FRM_SAL(CHG).frm:3921` from targeted hit output.
- Trigger/user action: product grid quantity/rate/discount/tax changes and product selection.
- Tables read: `PRD.SUB_UNT` per line.
- Tables written: none directly.
- Formula:
  - Starts line amount from `RATCAL(MS[3], MS[4], SUB_UNT)` or `MS[3] * MS[4]`.
  - Subtracts `MS[7]` add discount percent.
  - In `M_FRM_SAL_CHG.frm`, the inspected body does not subtract `MS[8]` trade/cash discount before tax; it adds tax from `MS[9]` after only the `MS[7]` subtraction.
  - In targeted hits for `M_FRM_SAL_RAT.frm` and `M_FRM_SAL(CHG).frm`, `MS[8]` is subtracted before tax.
  - Recalculates `MS[5]` UOM using `WGTCAL(MS[3], MS[2], SUB_UNT)`, accumulates `units`, `tax`, `taxamt`, and `M_PRD_AMT`, then calls `GRS_CAL`.
- Side effects: updates grid amount column and header totals.
- Report/print calls: none.
- Assumptions/unresolved:
  - CHG's omission of `MS[8]` from `P_DIS_CAL` conflicts with CHG `saverecord`, which still writes `DIS_AMT` from `MS[8]`. This may cause displayed totals and persisted line discount/tax basis to diverge.
- Golden-master candidates:
  - CHG line with both `MS[7]` and `MS[8]` nonzero; compare displayed amount vs persisted `INV.DIS_AMT`, `INV.TAX_AMT`, and `INL.NET_AMT`.
  - Same line in RAT and `SAL(CHG)` to confirm whether they subtract both discounts in display totals.

Next model should continue with these exact commands/passes:

```sh
rg -n "^(Private |Public )?(Sub|Function) " LegacyCodebase/M_FRM_SAL_CHG.frm LegacyCodebase/M_FRM_SAL_RAT.frm LegacyCodebase/M_FRM_SAL_PSH.frm 'LegacyCodebase/M_FRM_SAL(CHG).frm'
rg -n "^(Private |Public )?(Sub|Function) " LegacyCodebase/M_FRM_SCH.frm LegacyCodebase/M_FRM_SCH1.frm LegacyCodebase/M_FRM_SCHE.frm
rg -n "AutoApplyScheme|VLD_|ADD.*|SCH_AMT|PNV_PNL_SaveProcedure|Scheme_Price_Update|LNV_Rpt_Update" LegacyCodebase/M_FRM_SAL_PSH.frm
```

Required next extraction:

- Continue `M_FRM_SAL_PSH.frm` delta after the CHG/RAT/(CHG) deltas below.
- Confirm group cash/percent `SAVEGRPTYP` behavior with actual saved data or runtime test, because the inspected code raises a row-encoding ambiguity.
- Continue exact field maps for columns that remain inferred rather than header-backed.

## Continuation Update: CHG/RAT/(CHG), SCH1, SCHE, Grid Maps

Date: 2026-07-03. This update is still partial and source-cited. It deepens the requested variant and alternate-form areas, but does not complete all P0 sales/scheme extraction.

### Variant Reachability And Project Status

- `Effmcg.vbp` includes `M_FRM_SAL_PSH.frm`, `M_FRM_SCH.frm`, and `M_FRM_SAL.frm`, but not `M_FRM_SCH1.frm`, `M_FRM_SCHE.frm`, `M_FRM_SAL_CHG.frm`, `M_FRM_SAL_RAT.frm`, or `M_FRM_SAL(CHG).frm` (`LegacyCodebase/Effmcg.vbp:117`, `LegacyCodebase/Effmcg.vbp:127`, `LegacyCodebase/Effmcg.vbp:155`).
- The MDI scheme master menu opens `M_FRM_SCH.Show 1`, with no observed `M_FRM_SCH1` or `M_FRM_SCHE` call (`LegacyCodebase/M_MDI.frm:1409`).
- The MDI sales menu opens current sale forms `M_FRM_SAL.Show 1` and `M_FRM_SAL_PSH.Show 1`, not CHG/RAT/(CHG) variants (`LegacyCodebase/M_MDI.frm:1217`, `LegacyCodebase/M_MDI.frm:1221`).
- `M_FRM_ASD.frm` can tag `M_FRM_SAL_PSH` as an invoice flow (`LegacyCodebase/M_FRM_ASD.frm:292`).
- `SYS_PSW.frm` can show `M_FRM_SAL_CHG` after password logic, so CHG is not compiled in the current `.vbp` but may still have a separately-loaded/admin path in some deployments (`LegacyCodebase/SYS_PSW.frm:98`).
- `M_FRM_PRV.frm` maintains print-setup `PRT` rows for `M_FRM_SAL_CHG`, `M_FRM_SAL_RAT`, and `M_FRM_SAL_PSH`; this is report configuration evidence, not proof that CHG/RAT forms are in the active executable (`LegacyCodebase/M_FRM_PRV.frm:141`, `LegacyCodebase/M_FRM_PRV.frm:153`, `LegacyCodebase/M_FRM_PRV.frm:441`).
- `M_FRM_SCHE.frm` contains form/control declarations only and no executable procedures after `Attribute VB_Name = "M_FRM_SCHE"` (`LegacyCodebase/M_FRM_SCHE.frm:1223`). Targeted `rg` found no project/menu references to `M_FRM_SCHE` except the file's own form/resource declarations.
- Conclusion: `M_FRM_SCH1.frm` and `M_FRM_SCHE.frm` appear to be uncompiled/unreachable from this `Effmcg.vbp` build. `M_FRM_SCHE.frm` is an unused shell in this source tree. `M_FRM_SCH1.frm` has real logic but is not in the project file.
- UNRESOLVED: current `Effmcg.vbp` does not compile CHG/RAT/(CHG), but password/report forms reference CHG/RAT names. Treat those variants as legacy or separately-deployed until build artifacts/menu wiring are proven.
- Golden-master candidates: inspect production shortcut/build metadata for any executable that launches CHG/RAT; if reachable, run one invoice through each and compare `INL.ORD_IDY`, `INV`, stock changes, and report flag behavior.

### `M_FRM_SAL_CHG.frm` Deep Delta Against Main `M_FRM_SAL.frm`

- File/lines: `LegacyCodebase/M_FRM_SAL_CHG.frm:2105`, `LegacyCodebase/M_FRM_SAL_CHG.frm:2123`, `LegacyCodebase/M_FRM_SAL_CHG.frm:3066`, `LegacyCodebase/M_FRM_SAL_CHG.frm:3923`, `LegacyCodebase/M_FRM_SAL_CHG.frm:3967`, `LegacyCodebase/M_FRM_SAL_CHG.frm:4001`, `LegacyCodebase/M_FRM_SAL_CHG.frm:4047`.
- Trigger/user action: save from toolbar/keyboard calls `saverecord`; product grid edits call `P_DIS_PRD`/`P_DIS_CAL`; print toolbar calls `PRINTRECORD`.
- Tables read: `ACT`, `SMN`, `PRD`, `INL`, `INV`, `PRT`, `DIR`.
- Tables written: `INL`, `INV`, `PRD.CUR_LEV`, `DIR` print scope. No `PNL`, `PNV`, `LNV`, `BRD`, `IRT`, backup tables, scheme grids, or shell/deposit helpers were observed in the inspected CHG save path.
- Validations: confirmation prompt; dealer must exist as `ACT_TYP` `BOT`/`SAL`; salesman must exist; sales account must have `grp_clf like 'Sale%'` and blank `ACT_TYP`; net amount must be nonnegative (`LegacyCodebase/M_FRM_SAL_CHG.frm:2127`, `LegacyCodebase/M_FRM_SAL_CHG.frm:2133`, `LegacyCodebase/M_FRM_SAL_CHG.frm:2141`, `LegacyCodebase/M_FRM_SAL_CHG.frm:2149`, `LegacyCodebase/M_FRM_SAL_CHG.frm:2156`).
- Formulas:
  - `GRS_CAL` is the older compact cascade: `M_GRS_AMT=M_PRD_AMT`, cash discount amount from `M_PRD_AMT*M_CSD_PCG/100`, special discount from cash-discount-deducted amount, tax from special-discount-deducted amount, net after `M_EXD_AMT`, surcharge/freight from `(taxamt + M_TAX_AMT) * M_FRT_TYP / 100`, plus package cost, then `ROU_OFF` (`LegacyCodebase/M_FRM_SAL_CHG.frm:2105`).
  - `saverecord` writes `INL.ord_idy = "2"` and invoice ID `SAL` + `T7increment("IDY_SAL")` (`LegacyCodebase/M_FRM_SAL_CHG.frm:2175`, `LegacyCodebase/M_FRM_SAL_CHG.frm:2178`).
  - Line save uses one `MS` grid only: product ID, list rate/discount, quantity, rate, rate app, add discount, trade discount, tax percent, collected amount. `PRDNET = RATCAL(MS[3], MS[4], PRD.SUB_UNT)`, `ADD_AMT = PRDNET * MS[7] / 100`, `DIS_AMT = (PRDNET - ADD_AMT) * MS[8] / 100`, and `TAX_AMT = (PRDNET - (ADD_AMT + DIS_AMT)) * MS[9] / 100` (`LegacyCodebase/M_FRM_SAL_CHG.frm:2219`-`LegacyCodebase/M_FRM_SAL_CHG.frm:2245`).
  - Display `P_DIS_CAL` in CHG subtracts `MS[7]` before tax but does not subtract `MS[8]` before tax in the inspected source (`LegacyCodebase/M_FRM_SAL_CHG.frm:3967`). This differs from CHG save formula, which persists `DIS_AMT` and tax after both discounts.
- Side effects: modify mode reverses old `INV.PRD_QTY` into `PRD.CUR_LEV` through `ADDQTY`, deletes old `INV`, reuses old `INL`, then rewrites lines; each new line decrements stock with `CONQTY` (`LegacyCodebase/M_FRM_SAL_CHG.frm:2163`, `LegacyCodebase/M_FRM_SAL_CHG.frm:2247`).
- Report/print calls: if `PRT.PRT_NME = M_FRM_SAL_CHG`, updates `DIR INV/NUM` and `DOC/TYP`, then opens `VOUCH.RP1` (`LegacyCodebase/M_FRM_SAL_CHG.frm:2244`-`LegacyCodebase/M_FRM_SAL_CHG.frm:2248`). `PRINTRECORD` sets preview tag `SLP` (`LegacyCodebase/M_FRM_SAL_CHG.frm:4047`).
- Golden-master candidates: CHG invoice with `MS[7]` and `MS[8]` nonzero to expose display/save mismatch; CHG modify to verify stock reversal and no `BRD`/`LNV`/`PNL` writes; CHG print flag on/off.

### `M_FRM_SAL_RAT.frm` Deep Delta Against CHG/Main

- File/lines: `LegacyCodebase/M_FRM_SAL_RAT.frm:2105`, `LegacyCodebase/M_FRM_SAL_RAT.frm:2123`, `LegacyCodebase/M_FRM_SAL_RAT.frm:3890`, `LegacyCodebase/M_FRM_SAL_RAT.frm:3933`, `LegacyCodebase/M_FRM_SAL_RAT.frm:4013`.
- Trigger/tables/validations: same reduced invoice path as CHG for inspected save behavior.
- Variant-specific formulas and side effects:
  - `INL.ord_idy = "1"` (`LegacyCodebase/M_FRM_SAL_RAT.frm:2178`).
  - Product selection reads `CSR` for `act_idy = M_CST_IDY` and `prd_idy = selected product`; when found, it uses `CSR.prd_rat` as grid rate, otherwise `PRD.SAL_RAT` (`LegacyCodebase/M_FRM_SAL_RAT.frm:3908`-`LegacyCodebase/M_FRM_SAL_RAT.frm:3913`). This is a customer-specific rate variant, not just a cosmetic form copy.
  - `P_DIS_CAL` subtracts both `MS[7]` and `MS[8]` before tax (`LegacyCodebase/M_FRM_SAL_RAT.frm:3933`).
  - Quantity edit enforces stock-out with `CONPRDQTY(REMIND, M_QTY, product) < 0`; CHG has the analogous check commented in the diffed area (`LegacyCodebase/M_FRM_SAL_RAT.frm:3837`).
- Report/print calls: save print flag is `PRT_NME = M_FRM_SAL_RAT`; `PRINTRECORD` preview tag is `SLR` (`LegacyCodebase/M_FRM_SAL_RAT.frm:2245`, `LegacyCodebase/M_FRM_SAL_RAT.frm:4013`).
- Golden-master candidates: same product/customer through CHG and RAT where `CSR.prd_rat` exists; verify RAT uses CSR rate and CHG uses `M_DBT_RAP`/PRD rate. Quantity above stock should block in RAT.

### `M_FRM_SAL(CHG).frm` Deep Delta Against CHG/Main

- File/lines: `LegacyCodebase/M_FRM_SAL(CHG).frm:2114`, `LegacyCodebase/M_FRM_SAL(CHG).frm:2131`, `LegacyCodebase/M_FRM_SAL(CHG).frm:3073`, `LegacyCodebase/M_FRM_SAL(CHG).frm:3921`, `LegacyCodebase/M_FRM_SAL(CHG).frm:3999`.
- Trigger/tables/validations: same reduced invoice save path as CHG for inspected behavior.
- Variant-specific formulas and side effects:
  - The VB form class name is `M_FRM_SAL`, not `M_FRM_SAL_CHG`; because the active project also includes `LegacyCodebase/M_FRM_SAL.frm`, this file cannot be compiled alongside it without a name collision (`LegacyCodebase/M_FRM_SAL(CHG).frm:2101`; current source header begins `Begin VB.Form M_FRM_SAL`).
  - `INL.ord_idy = "0"` and `dispdetails` reloads only `ord_idy='0'`, unlike CHG's `ord_idy='2'` (`LegacyCodebase/M_FRM_SAL(CHG).frm:2185`, `LegacyCodebase/M_FRM_SAL(CHG).frm:3424`-`LegacyCodebase/M_FRM_SAL(CHG).frm:3428`; `LegacyCodebase/M_FRM_SAL_CHG.frm:2178`, `LegacyCodebase/M_FRM_SAL_CHG.frm:3421`-`LegacyCodebase/M_FRM_SAL_CHG.frm:3425`).
  - Print flag checks `PRT_NME = M_FRM_SAL`; preview tag is `SAL` (`LegacyCodebase/M_FRM_SAL(CHG).frm:2252`, `LegacyCodebase/M_FRM_SAL(CHG).frm:4007`).
  - Like RAT, `P_DIS_CAL` subtracts both add discount and trade/cash discount before tax, whereas CHG omits `MS[8]` in its display calculation (`LegacyCodebase/M_FRM_SAL(CHG).frm:3933`-`LegacyCodebase/M_FRM_SAL(CHG).frm:3942`; `LegacyCodebase/M_FRM_SAL_CHG.frm:3980`-`LegacyCodebase/M_FRM_SAL_CHG.frm:3987`).
  - `(CHG)` actively blocks quantity entry when `CONPRDQTY(...) < 0`; the analogous CHG stock-out block is commented out (`LegacyCodebase/M_FRM_SAL(CHG).frm:3839`-`LegacyCodebase/M_FRM_SAL(CHG).frm:3847`, `LegacyCodebase/M_FRM_SAL_CHG.frm:3881`-`LegacyCodebase/M_FRM_SAL_CHG.frm:3889`).
  - CHG permits manual numeric editing of `MS[4]` rate and `MS[5]` UOM and converts UOM through `UNTCAL`; `(CHG)` omits the rate edit branch and uses `MS[5]` Enter only for navigation (`LegacyCodebase/M_FRM_SAL_CHG.frm:3738`-`LegacyCodebase/M_FRM_SAL_CHG.frm:3779`, `LegacyCodebase/M_FRM_SAL_CHG.frm:3899`-`LegacyCodebase/M_FRM_SAL_CHG.frm:3908`; `LegacyCodebase/M_FRM_SAL(CHG).frm:3741`-`LegacyCodebase/M_FRM_SAL(CHG).frm:3758`, `LegacyCodebase/M_FRM_SAL(CHG).frm:3857`-`LegacyCodebase/M_FRM_SAL(CHG).frm:3868`).
- Side effects/report calls:
  - Same reduced `INL`/`INV`/`PRD.CUR_LEV`/`VOUCH.RP1` behavior as CHG, not the full modern `M_FRM_SAL.frm` `PNL`/`PNV`/`LNV` scheme path.
  - CHG uses `On Error Resume Next` in save, detail display, and quantity entry; `(CHG)` removes those guards in corresponding inspected routines (`LegacyCodebase/M_FRM_SAL_CHG.frm:2124`, `LegacyCodebase/M_FRM_SAL_CHG.frm:3418`, `LegacyCodebase/M_FRM_SAL_CHG.frm:3835`; compare `LegacyCodebase/M_FRM_SAL(CHG).frm:2131`, `LegacyCodebase/M_FRM_SAL(CHG).frm:3421`, `LegacyCodebase/M_FRM_SAL(CHG).frm:3797`).
- Assumptions/unresolved:
  - The exact historical executable/project containing `M_FRM_SAL(CHG).frm` was not found. Because it declares `M_FRM_SAL`, treat it as an alternate normal-sale revision, not a callable form named `M_FRM_SAL(CHG)`.
- Golden-master candidates:
  - Compare old `(CHG)` invoice with current `M_FRM_SAL.frm` using identical header/line inputs; expect radically different downstream writes and report mirrors despite the same form class name.
  - Save identical data through CHG and `(CHG)`: expect `ORD_IDY` `2` vs `0`, print key `M_FRM_SAL_CHG` vs `M_FRM_SAL`, and preview tag `SLP` vs `SAL`.
  - Enter quantity above stock: CHG source should allow it because its guard is commented; `(CHG)` should reject it.
  - Attempt manual rate/UOM editing to capture CHG-enabled versus `(CHG)`-restricted grid behavior.

### `M_FRM_SCH1.frm` Reachability, Save, Display, Product Lists

- File/procedure inventory anchors: `save` at `LegacyCodebase/M_FRM_SCH1.frm:1350`; `Form_Load` at `LegacyCodebase/M_FRM_SCH1.frm:2664`; `V_PRD_LST` at `LegacyCodebase/M_FRM_SCH1.frm:4091`; `P_DIS_PRD` at `LegacyCodebase/M_FRM_SCH1.frm:4244`; `DISPLAYALL` at `LegacyCodebase/M_FRM_SCH1.frm:4438`.
- Reachability: not listed in `Effmcg.vbp`; no MDI/menu show call verified. It is real code but not compiled into the current project.
- Trigger/user action:
  - Toolbar `s` calls `save`; toolbar `p` opens `M_FRM_DSN` tagged `"Schemes List"` (`LegacyCodebase/M_FRM_SCH1.frm:3977`).
  - Scheme type selection enables one of four tabs: SKU quantity, bill value, multiple SKU, or UOM; it calls `V_PRD_LST`, `PRODUCTLIST`, or `skulist` according to `SCH_CMB.ListIndex` (`LegacyCodebase/M_FRM_SCH1.frm:3348`-`LegacyCodebase/M_FRM_SCH1.frm:3364`).
- Tables read: `SCH`, `PRD`, `BRD`; missing-table checks call `CHECKTB("sch.dbf")` and `CHECKTB("prd.dbf,brd.dbf")`.
- Tables written: `SCH` only in inspected `save`; no `SID` call was observed in SCH1 save, unlike `M_FRM_SCH.frm`.
- Validations:
  - Confirmation prompt.
  - `SCH_NME` required.
  - end date must be greater than or equal to start date.
  - bill schemes require positive `GRS_AMT`.
  - new SKU/free-quantity schemes require a base product row (`LegacyCodebase/M_FRM_SCH1.frm:1352`-`LegacyCodebase/M_FRM_SCH1.frm:1384`).
- Formulas/encoding:
  - New ID is `SCH` + `T7increment("idy_sch")`; `CON.BeginTrans` is started before row creation (`LegacyCodebase/M_FRM_SCH1.frm:1396`-`LegacyCodebase/M_FRM_SCH1.frm:1397`).
  - SKU quantity to free quantity writes base rows with `txn_typ="BAS"` from `BAS_PRD_VIW`, then free rows with `txn_typ="FRE"` from `FRE_PRD_VIW`; free rows store the free product in `B_PRD_IDY/B_PRD_QTY` and the base batch in `f_prd_idy`, which is a confusing legacy field reuse (`LegacyCodebase/M_FRM_SCH1.frm:1398`-`LegacyCodebase/M_FRM_SCH1.frm:1453`).
  - SKU quantity to cash/percent writes base rows plus a single `CAS` or `PCG` row with `B_DIS_AMT` or `B_DIS_PCG` (`LegacyCodebase/M_FRM_SCH1.frm:1456`-`LegacyCodebase/M_FRM_SCH1.frm:1504`).
  - Bill amount schemes write `B_BIL_AMT = GRS_AMT`; depending on `SUB_CMB`, the row may carry free product/quantity or cash/percent discount (`LegacyCodebase/M_FRM_SCH1.frm:1507`-`LegacyCodebase/M_FRM_SCH1.frm:1544`).
  - Multiple SKU base rows write `B_CON_APP` as `AND`, `ORO`, or `ORM` depending on base AND/OR controls (`LegacyCodebase/M_FRM_SCH1.frm:1548` onward in inspected slice).
- Product-list paths:
  - `V_PRD_LST` filters `BRD,PRD` where batch stock `BRD.PRD_QTY > 0` and product name matches `Text3`, then populates `PRD_LST` with name, MRP, batch quantity, UOM, `BAT_IDY`, conversion factor, and `PRD_IDY` (`LegacyCodebase/M_FRM_SCH1.frm:4091`-`LegacyCodebase/M_FRM_SCH1.frm:4120`).
  - `PRD_LST_V` loads all in-stock `BRD,PRD` rows and targets `PRD_LST`, `PRD_VIW`, `SKU_LST`, or `UOM_LST` based on enabled tab (`LegacyCodebase/M_FRM_SCH1.frm:4127`-`LegacyCodebase/M_FRM_SCH1.frm:4170`).
  - `PRODUCTLIST` repeats the same in-stock batch/product list but searches from `Text1`, `Text2`, or `Text5` for bill/SKU/UOM paths (`LegacyCodebase/M_FRM_SCH1.frm:4270` onward).
  - `P_DIS_PRD` fills the base product view from first matching in-stock `BRD,PRD` row by product ID (`LegacyCodebase/M_FRM_SCH1.frm:4244`-`LegacyCodebase/M_FRM_SCH1.frm:4264`).
  - `DISPLAYALL` bulk-fills all currently in-stock products into base/free grids for SKU or UOM paths (`LegacyCodebase/M_FRM_SCH1.frm:4438`-`LegacyCodebase/M_FRM_SCH1.frm:4505`).
- Side effects:
  - `Form_Load` initializes many alternate grid headers and list data from active schemes; it calls `SETYEAR M_FRM_SCH` and `setval`, despite living in `M_FRM_SCH1.frm` (`LegacyCodebase/M_FRM_SCH1.frm:2664`).
  - `Form_Activate` contains `RS.Open "selec * from prd"` typo; if this code is executed, it should error before product-count validation (`LegacyCodebase/M_FRM_SCH1.frm:2625`).
- Report/print calls: toolbar print path shows `M_FRM_DSN` with tag `"Schemes List"`; no direct `.RP1` report call in inspected SCH1 toolbar path (`LegacyCodebase/M_FRM_SCH1.frm:3990`).
- Golden-master candidates: only if this historical form must be supported. Recreate one SKU-to-free, SKU-to-cash, bill-to-percent, and multiple-SKU scheme from SCH1 and compare rows against current `M_FRM_SCH.frm`.

### `M_FRM_SCHE.frm` Shell Confirmation

- File/lines: form/control body ends at `LegacyCodebase/M_FRM_SCHE.frm:1222`; only VB attributes follow (`LegacyCodebase/M_FRM_SCHE.frm:1223`).
- Trigger/tables/validations/formulas/side effects/report calls: none verified in this source file.
- Conclusion: unused shell for this project. Do not migrate logic from it unless another reachable project copy is found.
- Golden-master candidates: none unless runtime evidence shows this form is launched elsewhere.

### `SAVEGRPTYP` Cash/Percent Ambiguity Status

- Current inspected source: `M_FRM_SCH.frm` remains the active compiled scheme master (`LegacyCodebase/Effmcg.vbp:127`, `LegacyCodebase/M_MDI.frm:1409`).
- Verified by current pass: `SAVEGRPTYP` is in `M_FRM_SCH.frm` and writes group-type schemes into `SCH` using active `MSPack`, `MSFree`, `Sch_Cmb`, `Sub_Cmb`, `CmbPCon`, and `CmbFree` state (`LegacyCodebase/M_FRM_SCH.frm:3079`).
- Trigger/user action: `save` dispatches to `SAVEGRPTYP` when `Sch_Cmb.ListIndex > 3`, i.e. the active group-type scheme tabs rather than SKU or bill scheme modes (`LegacyCodebase/M_FRM_SCH.frm:1978`-`LegacyCodebase/M_FRM_SCH.frm:1980`).
- Tables read/written: opens writable `SCH`; calls `T7increment("idy_sch")`; calls `SID`, which writes the generated scheme ID into `SID`; wraps the group save in an explicit `con.BeginTrans` / `CommitTrans` with rollback on error (`LegacyCodebase/M_FRM_SCH.frm:3082`-`LegacyCodebase/M_FRM_SCH.frm:3144`).
- Exact base row encoding for group type schemes:
  - One `SCH` row is written for each nonblank `MSPack` row (`LegacyCodebase/M_FRM_SCH.frm:3087`).
  - `SCH_IDY`, `SCH_NME`, `SCH_TYP`, `F_SCH_TYP`, `SCH_STS`, `CMP_DST`, `SCH_BDT`, `SCH_EDT`, `MER_NUM`, `MER_QTY`, and `MER_AMT` are copied from the header controls for every row (`LegacyCodebase/M_FRM_SCH.frm:3089`-`LegacyCodebase/M_FRM_SCH.frm:3100`).
  - `BAT_IDY` is blank; `GRP_TYP = MSPack[A,0]`; `B_PRD_QTY = MSPack[A,1]`; `B_CON_APP` is `AND`, `ORO`, or `ORM` from `CmbPCon.ListIndex` 0, 1, or 2 (`LegacyCodebase/M_FRM_SCH.frm:3105`-`LegacyCodebase/M_FRM_SCH.frm:3115`).
  - If `Sub_Cmb.ListIndex = 2`, `B_DIS_AMT` is populated from `TxtValue`; if `Sub_Cmb.ListIndex = 3`, `B_DIS_PCG` is populated from `TxtValue` (`LegacyCodebase/M_FRM_SCH.frm:3101`-`LegacyCodebase/M_FRM_SCH.frm:3102`).
  - The code briefly assigns `TXN_TYP = CAS` or `PCG` for cash/percent, then unconditionally overwrites `TXN_TYP = BAS` before update. Therefore active `SAVEGRPTYP` stores group cash/percent benefit values on the `BAS` rows, not as separate `CAS`/`PCG` rows (`LegacyCodebase/M_FRM_SCH.frm:3103`-`LegacyCodebase/M_FRM_SCH.frm:3117`).
- Exact free row encoding for group type free schemes:
  - Only when `Sub_Cmb.ListIndex = 1`, one `SCH` row is written for each nonblank `MSFree` row (`LegacyCodebase/M_FRM_SCH.frm:3121`-`LegacyCodebase/M_FRM_SCH.frm:3140`).
  - `B_PRD_IDY = MSFree[A,4]`, `BAT_IDY = MSFree[A,3]`, `B_PRD_QTY = MSFree[A,2]`, `F_CON_APP` is `AND` or `ORO` from `CmbFree.ListIndex`, `PRO_RAT = N`, `TXN_TYP = FRE`, plus user/date/time audit fields (`LegacyCodebase/M_FRM_SCH.frm:3133`-`LegacyCodebase/M_FRM_SCH.frm:3140`).
- Display behavior: `dispschemegrt3` reloads `TXN_TYP='BAS'` rows into `MSPack`, restores `TxtValue` from `B_DIS_AMT` or `B_DIS_PCG` based on `Sub_Cmb`, and reloads `TXN_TYP='FRE'` rows into `MSFree` only for free schemes (`LegacyCodebase/M_FRM_SCH.frm:3150`-`LegacyCodebase/M_FRM_SCH.frm:3199`).
- Sales application behavior: `AddBillAmtGroupType` reads the group scheme from `SCH`, filters only `TXN_TYP='BAS'` rows for `B_CON_APP='AND'` or `ORO`, uses `SCH_TYP` as the group dimension, and applies `B_DIS_PCG` as a percent of matched group sale value or `B_DIS_AMT` as a repeated cash amount (`LegacyCodebase/M_FRM_SAL.frm:10476`-`LegacyCodebase/M_FRM_SAL.frm:10556`).
- Group dimension matching:
  - `Brand`: compares `MS` product's `PRD.BRD_NME` via `GetValue("brd_nme", "prd where prd_idy = ...")` (`LegacyCodebase/M_FRM_SAL.frm:10407`, `LegacyCodebase/M_FRM_SAL.frm:10566`).
  - `Pack Size`: compares `PRD.PRD_PCK` (`LegacyCodebase/M_FRM_SAL.frm:10421`, `LegacyCodebase/M_FRM_SAL.frm:10568`).
  - `Pack Type`: compares `PRD.PRD_PKT` (`LegacyCodebase/M_FRM_SAL.frm:10434`, `LegacyCodebase/M_FRM_SAL.frm:10570`).
  - `Segment`: compares `MS[20]` group class (`LegacyCodebase/M_FRM_SAL.frm:10447`, `LegacyCodebase/M_FRM_SAL.frm:10572`).
  - `Category`: compares `MS[19]` group category (`LegacyCodebase/M_FRM_SAL.frm:10460`, `LegacyCodebase/M_FRM_SAL.frm:10574`).
- Side effects: group base matches add `SC` rows with `SC[0]=scheme ID`, `SC[1]=base batch ID`, `SC[2]=matched stock count`, `SC[4]=BAS`; group cash/percent benefit adds an `SC` row with `SC[3]=incremental scheme amount`, `SC[4]=FRE`; sale save later writes these `SC` rows into `LNV` (`LegacyCodebase/M_FRM_SAL.frm:10414`-`LegacyCodebase/M_FRM_SAL.frm:10470`, `LegacyCodebase/M_FRM_SAL.frm:10520`-`LegacyCodebase/M_FRM_SAL.frm:10556`, `LegacyCodebase/M_FRM_SAL.frm:3735`-`LegacyCodebase/M_FRM_SAL.frm:3745`).
- Additional context: SCH1 has analogous cash/percent encoding for SKU and bill paths, where cash and percent are represented as separate `txn_typ` rows (`CAS`/`PCG`) with `B_DIS_AMT`/`B_DIS_PCG`; this supports, but does not prove, the interpretation that group cash/percent should be golden-mastered by inspecting actual `SCH` rows (`LegacyCodebase/M_FRM_SCH1.frm:1456`, `LegacyCodebase/M_FRM_SCH1.frm:1507`).
- UNRESOLVED: no runtime DBF fixture was inspected to confirm whether legacy data contains older `CAS`/`PCG` group rows created by previous versions. Active source path above is verified, but migration must still accept historical row shapes if present.
- Golden-master candidates: create active `M_FRM_SCH.frm` Brand/PackSize/PackType/Segment/Category group schemes for free, cash, and percent; capture every `SCH` row plus `SID`; apply each to invoices with `AND`, `ORO`, and `ORM` base conditions and compare `SC`, `LNV`, `MSS`, `SCH_AMT`, and printed totals.

### Improved Grid Maps

#### Current Main Sales `M_FRM_SAL.frm` `MS`

- Source anchors: grid declares `Cols = 26` (`LegacyCodebase/M_FRM_SAL.frm:1589`); headers and widths initialized in `Form_Load` (`LegacyCodebase/M_FRM_SAL.frm:5309`-`LegacyCodebase/M_FRM_SAL.frm:5386`); save/report mappings at `LegacyCodebase/M_FRM_SAL.frm:3369` and `LegacyCodebase/M_FRM_SAL.frm:10688`; group/add-scheme reads at `LegacyCodebase/M_FRM_SAL.frm:10406` and `LegacyCodebase/M_FRM_SAL.frm:10563`.
- Header-backed and write/read-verified columns:
  - `MS[0]`: `Product ID`; saved to `INV/PNV.PRD_IDY`.
  - `MS[1]`: `Stock Type`; nonblank save qualifier; saved to `INV.STK_TYP`.
  - `MS[2]`: `Quantity`; sale quantity.
  - `MS[3]`: `UOM`; display UOM quantity; saved as `PRD_UOM`.
  - `MS[4]`: `RateID`; rate application; saved as `RAT_APP`.
  - `MS[5]`: `Rate`; sale rate.
  - `MS[6]`: `Dis.Amt`; add/amount discount stored to `ADD_AMT`.
  - `MS[7]`: `CD%`; trade/cash discount percent stored to `TRD_DIS`.
  - `MS[8]`: `Tax%`; tax percent stored to `TRD_TAX`.
  - `MS[9]`: `Amount`; displayed line amount / `COL_AMT` basis.
  - `MS[10]`: `MRP`; report `MRP_RAT`.
  - `MS[11]`: `Product Name`; report product name.
  - `MS[12]`: `Short Name`; header only in inspected code.
  - `MS[13]`: `UPC`; sub-units (`SUB_UNT`).
  - `MS[14]`: `uom`; hidden/reserved, not write-site verified.
  - `MS[15]`: `BATIDY`; batch ID saved to `BAT_IDY`.
  - `MS[16]`: `Mfg.Date`; display manufacturing date.
  - `MS[17]`: `BBD-Exp.Date`; display expiry/BBD date.
  - `MS[18]`: hidden product group ID; saved to `PNV.GRP_IDY`.
  - `MS[19]`: hidden group category; saved to `PNV.GRP_CAT`; used by group `Category` schemes.
  - `MS[20]`: hidden group class/segment; saved to `PNV.GRP_CLS`; used by group `Segment` schemes.
  - `MS[21]`: `Batch`; batch name saved to `BAT_NME`.
  - `MS[22]`: `List Rate`; list/coupon rate.
  - `MS[23]`: `List %+`; list discount/coupon percent.
  - `MS[24]`: `Inc'tve`; inclusive/value basis used for tax exclusion when `DIR SYS/DPV = T`.
  - `MS[25]`: `Inc'TTL`; header initialized, not write-site verified in this pass.

#### Current Main Sales `M_FRM_SAL.frm` `MSF`

- Source anchors: grid declares `Cols = 20` (`LegacyCodebase/M_FRM_SAL.frm:2138`); headers initialized in `Form_Load` (`LegacyCodebase/M_FRM_SAL.frm:5392`-`LegacyCodebase/M_FRM_SAL.frm:5424`); free-grid save/report mappings at `LegacyCodebase/M_FRM_SAL.frm:3471` and `LegacyCodebase/M_FRM_SAL.frm:10787`.
- Verified columns:
  - `MSF[0]`: `Product ID`; saved to `PRD_IDY`.
  - `MSF[1]`: `Stock Type`; observed stock types include `Free`, `Damage`, `TakeBack`, `Breakage`, and `Leakage`.
  - `MSF[2]`: `Quantity`; free/return quantity.
  - `MSF[3]`: `UOM`; display UOM quantity.
  - `MSF[4]`: `Rate`; free/return rate.
  - `MSF[5]`: `Amount`; collected/line amount.
  - `MSF[6]`: `MRP`; report MRP.
  - `MSF[7]`: `Product Name`; report product name.
  - `MSF[8]`: `Short Name`; header only in inspected code.
  - `MSF[9]`: `UPC`; sub-units.
  - `MSF[10]`: `uom`; hidden/reserved, not write-site verified.
  - `MSF[11]`: `BATIDY`; batch ID.
  - `MSF[12]`: `GRPIDY`; product group ID.
  - `MSF[13]`: `GRPCAT`; group category.
  - `MSF[14]`: `GRPCLS`; group class/segment.
  - `MSF[15]`: `Batch`; batch name.
  - `MSF[16]`: `Tax %`; free-grid tax percent.
  - `MSF[17]`: `List Rate`; list/coupon rate.
  - `MSF[18]`: `List %+`; list discount/coupon percent.
  - `MSF[19]`: `Rate`; second rate basis used by free tax amount formula in save.

#### Current Main Sales `M_FRM_SAL.frm` `MSS`

- Source anchors: grid declares `Cols = 17` (`LegacyCodebase/M_FRM_SAL.frm:2168`); headers initialized in `Form_Load` (`LegacyCodebase/M_FRM_SAL.frm:5443`-`LegacyCodebase/M_FRM_SAL.frm:5455`); scheme add rows at `LegacyCodebase/M_FRM_SAL.frm:10360`; save/report mappings at `LegacyCodebase/M_FRM_SAL.frm:3565` and `LegacyCodebase/M_FRM_SAL.frm:10866`.
- Verified columns:
  - `MSS[0]`: `Product Name`; free/scheme product name.
  - `MSS[1]`: `MRP`.
  - `MSS[2]`: `Quantity`; free quantity stored to `INV/PNV.PRD_QTY`.
  - `MSS[3]`: `Rate`; free/scheme rate stored to `INV/PNV.PRD_RAT`.
  - `MSS[4]`: `Amount`; scheme amount stored to `INV/PNV.COL_AMT`.
  - `MSS[5]`: `Product ID`.
  - `MSS[6]`: `Batch`; batch name.
  - `MSS[7]`: `batidy`; batch ID.
  - `MSS[8]`: `subunts` header is misleading in active code; write sites use it as `SCH_IDY`.
  - `MSS[9]`: `grpid` header is misleading/unused in inspected add paths; older comment says possible `TXN_TYP`.
  - `MSS[10]`: `grpcat` header is misleading; add paths use it as base quantity.
  - `MSS[11]`: `grpclf` header is misleading; used as free quantity/available value in scheme edit checks.
  - `MSS[12]`: hidden sub-units used for `STM` and `SUB_UNT`.
  - `MSS[13]`: hidden product group.
  - `MSS[14]`: hidden free condition application (`F_CON_APP`) in add paths.
  - `MSS[15]`: hidden original/calculated scheme quantity mirror.
  - `MSS[16]`: `PRDUOM`; display UOM quantity saved to `PRD_UOM`.
- UNRESOLVED: `MSS[9]` is commented as `TXN_TYP` but not populated in the inspected add paths; verify before depending on it.

#### Current Main Sales `M_FRM_SAL.frm` `SC`

- Source anchors: scratch grid declares `Rows = 0`, `Cols = 7`, no fixed rows/cols (`LegacyCodebase/M_FRM_SAL.frm:211`-`LegacyCodebase/M_FRM_SAL.frm:214`); `ADDSKUAMOUNT`/`ADDSKUPCG`/bill amount rows at `LegacyCodebase/M_FRM_SAL.frm:8231`; group/base rows at `LegacyCodebase/M_FRM_SAL.frm:10414`; group free amount rows at `LegacyCodebase/M_FRM_SAL.frm:10520`; save writes `SC` into `LNV` at `LegacyCodebase/M_FRM_SAL.frm:3735`.
- Verified columns:
  - `SC[0]`: scheme ID.
  - `SC[1]`: batch/base descriptor, usually base `BAT_IDY` for `BAS` rows; blank for several cash/percent benefit rows.
  - `SC[2]`: scheme base quantity/count converted through stock/UOM logic.
  - `SC[3]`: free/cash/discount scheme amount.
  - `SC[4]`: row role, observed values `BAS` and `FRE`.
- Role in side effects: `SC` is the scheme audit/ledger staging grid later written into `LNV` by the main sale save/report helpers.
- Tables written: `LNV.INV_IDY = invoice`, `LEV_IDY = SC[0]`, `LEV_DES = SC[1]`, `INV_GRS = SC[2]`, `LNV_AMT = SC[3]`, `TXN_TYP = SC[4]` (`LegacyCodebase/M_FRM_SAL.frm:3739`-`LegacyCodebase/M_FRM_SAL.frm:3744`).
- UNRESOLVED: `SC[5]` and `SC[6]` are declared but no active write/read site was verified in this pass.

#### Older CHG/RAT/(CHG) `MS`

- Source anchors: `M_FRM_SAL_CHG.frm:3066`, `M_FRM_SAL_CHG.frm:2219`, `M_FRM_SAL_RAT.frm:2212`, `LegacyCodebase/M_FRM_SAL(CHG).frm:2219`.
- Header-backed columns:
  - `MS[0]`: Product ID.
  - `MS[1]`: L Price.
  - `MS[2]`: LD%.
  - `MS[3]`: Units.
  - `MS[4]`: Rate.
  - `MS[5]`: UOM.
  - `MS[6]`: Rate ID.
  - `MS[7]`: AD%.
  - `MS[8]`: CD%.
  - `MS[9]`: Tax%.
  - `MS[10]`: Amount.
  - `MS[11]`: Product Name.
  - `MS[12]`: Short Name.
- Difference from current main `M_FRM_SAL.frm`: no `MSF`, `MSS`, or `SC` scheme/free grids in these older reduced variants.
