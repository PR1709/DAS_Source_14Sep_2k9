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
- `../LegacyCodebase/M_FRM_SAL(CHG).frm` must still be compared against `M_FRM_SAL_CHG.frm`.
- `../LegacyCodebase/M_FRM_SCH.frm` scheme master starts event/procedure code around line 1057; save/copy/display routines are known at lines 1855, 1860, 1941, 2029, 2530, 2593, 2676, 2835, 2910, 2985, 3079, 3325, 3392.
- `../LegacyCodebase/M_FRM_SCH1.frm` and `../LegacyCodebase/M_FRM_SCHE.frm` still require procedure inventory and diff against `M_FRM_SCH.frm`.

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
- Group validators inspect `SCH_TYP` for Brand/Pack Size/Pack Type/Segment/Category and compare group values against grid product group/attributes. Full branch-by-branch extraction remains pending.
- For `SID`-restricted schemes, eligibility is allowed if any matching `SID.SMN_IDY`, `SID.ROU_IDY`, or `SID.CST_IDY` exists for the current salesman, route, or customer.

Side effects:

- `SCHLST` is cleared by `VLD_FRE_QTY`, then appended by each validator.

Golden-master candidates:

- Scheme active date edge cases on start/end dates.
- SID unrestricted vs restricted to salesman/route/customer.
- SKU-level threshold with piece-mode conversion.
- Multiple SKU AND all items satisfied vs one missing.
- OR/ORM aggregate threshold.

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
  - OR/ORM branches aggregate stock and apply by max base quantity. Full branch validation should be revisited in source because some output was truncated during this pass.
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
  - `AddBillAmtGroupType` computes group value through `GETSCHGRPVAL`, then applies percent or cash-off; exact group-type branches still need full manual validation.

Side effects:

- Adds lines to `MSS` for free SKU benefits.
- Adds rows to `SC` for scheme audit/amount rows.
- Mutates `SCH_AMT`.

Assumptions/unresolved:

- `ADDSKUQUANTIY` proportional residual formula at lines 7814-7818 is suspicious: `(ACTQTY / base_qty) / free_qty`. Preserve exactly until domain owners validate.
- `MLTBILPCGAMT` stores `SC.TextMatrix(...,3)=LSCH_AMT` for some branches, not delta from previous scheme amount.
- Some variables are reused globally (`STK`, `SCHQTY`, `PRDVAL`, `LSCH_AMT`) and may retain values if not reset in every branch.

Golden-master candidates:

- SKU free quantity with and without `PRO_RAT`.
- Bill free quantity.
- Multiple SKU AND with two required bases.
- Multiple SKU ORO/ORM.
- SKU cash-off and SKU percent discount.
- Bill cash-off with `PRO_RAT=Y` residual amount.
- Multiple bill percent and cash-off.
- Brand/pack/category/segment cash-off vs percent.

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
  - Group/multiple cash-off/percent handlers are present elsewhere and must be verified against the full dispatch block in a follow-up.

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

Next model should continue with these exact commands/passes:

```sh
rg -n "^(Private |Public )?(Sub|Function) " LegacyCodebase/M_FRM_SAL_CHG.frm LegacyCodebase/M_FRM_SAL_RAT.frm LegacyCodebase/M_FRM_SAL_PSH.frm 'LegacyCodebase/M_FRM_SAL(CHG).frm'
rg -n "^(Private |Public )?(Sub|Function) " LegacyCodebase/M_FRM_SCH.frm LegacyCodebase/M_FRM_SCH1.frm LegacyCodebase/M_FRM_SCHE.frm
rg -n "AutoApplyScheme|VLD_|ADD.*|SCH_AMT|PNV_PNL_SaveProcedure|Scheme_Price_Update|LNV_Rpt_Update" LegacyCodebase/M_FRM_SAL_PSH.frm
```

Required next extraction:

- Diff `M_FRM_SAL_CHG.frm`, `M_FRM_SAL_RAT.frm`, `M_FRM_SAL_PSH.frm`, and `M_FRM_SAL(CHG).frm` against `M_FRM_SAL.frm` for changed formulas and save behavior.
- Deepen `M_FRM_SCH1.frm` display/product-list paths and confirm whether it is reachable in the main VB project.
- Confirm whether `M_FRM_SCHE.frm` is an unused shell or whether logic exists in another project/resource copy.
- Confirm group cash/percent `SAVEGRPTYP` behavior with actual saved data or runtime test, because the inspected code raises a row-encoding ambiguity.
- Build a field map for `MS`, `MSF`, `MSS`, and `SC` columns from design/control headers plus write sites.
