# 09 Continuation Queue

This file is for future AI models. Work through this queue before implementation.

## Rules For Continuation

For each module:

1. Read the full file(s) listed.
2. Extract all public/private subs/functions.
3. For each handler, document:
   - user action that triggers it
   - tables read
   - tables written
   - validations
   - formulas
   - side effects
   - report/print/integration calls
4. Cite file and line references.
5. Mark unclear behavior as `ASSUMPTION` or `UNRESOLVED`.
6. Add golden-master test candidates.

## P0 - Must Extract Before Any Migration Code

### Sales Invoice And Scheme Engine

Files:

- `../LegacyCodebase/M_FRM_SAL.frm`
- `../LegacyCodebase/M_FRM_SAL_CHG.frm`
- `../LegacyCodebase/M_FRM_SAL_RAT.frm`
- `../LegacyCodebase/M_FRM_SAL_PSH.frm`
- `../LegacyCodebase/M_FRM_SAL(CHG).frm`
- `../LegacyCodebase/M_FRM_SCH.frm`
- `../LegacyCodebase/M_FRM_SCH1.frm`
- `../LegacyCodebase/M_FRM_SCHE.frm`

Known functions needing full extraction:

- `GRS_CAL`
- `saverecord`
- `P_DIS_PRD`
- `P_DIS_CAL`
- `P_FRE_CAL`
- `V_PRD_LST`
- `AutoApplyScheme`
- all `VLD_*` scheme validators
- all `ADD*` scheme applicators
- `PNV_PNL_SaveProcedure`
- `Fill_LNV_Batch`
- `Scheme_Price_Update`
- `LNV_Rpt_Update`
- print/report methods

Output doc to create:

- `10_sales_and_scheme_spec.md`

### Stock / Inventory / Batch

Files:

- `../LegacyCodebase/EFF_LIB.bas`
- `../LegacyCodebase/EFF_LIB1.bas`
- `../LegacyCodebase/M_FRM_STI.frm`
- `../LegacyCodebase/M_FRM_STO.frm`
- `../LegacyCodebase/M_FRM_SRT.frm`
- `../LegacyCodebase/M_FRM_PRD.frm`
- `../LegacyCodebase/frmNewRateTemplate.frm`
- `../LegacyCodebase/frmSalesCount.frm`
- `../LegacyCodebase/M_FRM_TRS.frm`

Focus:

- stock transaction sign semantics
- product/batch quantity updates
- conversion functions
- stock ledger
- empties/shells/deposits
- stock report totals

Output doc:

- `11_inventory_stock_spec.md`

### Van Sales / HHT / Field Force

Files:

- `../LegacyCodebase/M_VAN_SAL.frm`
- `../LegacyCodebase/M_FRM_MIBS.frm`
- `../LegacyCodebase/M_FRM_MBL.frm`
- `../LegacyCodebase/M_FRM_HHT.frm`
- `../LegacyCodebase/DAS_XTR.frm`
- `../LegacyCodebase/HHT/*.TXT`
- `../LegacyCodebase/dsync.PRG`

Focus:

- load sheet
- van stock
- return
- shortage/breakage/leakage
- HHT file layouts
- import/export session behavior

Output doc:

- `12_van_hht_spec.md`

### Finance / Ledger / Tally

Files:

- `../LegacyCodebase/M_FRM_RCT.frm`
- `../LegacyCodebase/M_FRM_PMT.frm`
- `../LegacyCodebase/M_FRM_JRN.frm`
- `../LegacyCodebase/M_FRM_IRT.frm`
- `../LegacyCodebase/M_ACT_ RCT.frm`
- `../LegacyCodebase/ModTallyInterface.bas`
- `../LegacyCodebase/modProcess.bas`
- `../LegacyCodebase/modXml.bas`
- `../LegacyCodebase/ETLInterfaceForm.frm`

Focus:

- ledger entries
- voucher types
- receipt/payment posting
- invoice-to-ledger posting
- Tally export payloads

Output doc:

- `13_finance_tally_spec.md`

## P1 - Important After P0

### Purchase

Files:

- `../LegacyCodebase/M_FRM_PUR.frm`
- `../LegacyCodebase/M_FRM_POR.frm`
- `../LegacyCodebase/M_FRM_PPO.frm`
- `../LegacyCodebase/M_PUR_PQY.frm`

Output:

- `14_purchase_spec.md`

### Master Data

Files:

- `../LegacyCodebase/M_FRM_ACT.frm`
- `../LegacyCodebase/M_FRM_PRD.frm`
- `../LegacyCodebase/M_FRM_SMN.frm`
- `../LegacyCodebase/M_FRM_MAR.frm`
- `../LegacyCodebase/M_FRM_LOC.frm`
- `../LegacyCodebase/M_FRM_CAT.frm`
- `../LegacyCodebase/M_FRM_PGR.frm`
- `../LegacyCodebase/M_FRM_GRP.frm`

Output:

- `15_master_data_spec.md`

### Sales Order / Retailer Invoice Variants

Files:

- `../LegacyCodebase/M_FRM_SOR.frm`
- `../LegacyCodebase/M_FRM_SSO.frm`
- `../LegacyCodebase/M_FRM_RSI.frm`
- `../LegacyCodebase/M_FRM_RSI1.frm`
- `../LegacyCodebase/M_FRM_RSI_PRD.frm`

Output:

- `16_sales_order_retailer_invoice_spec.md`

## P2 - Reporting And Admin

### Reports

Files:

- all `.RP1`
- all `.RP5`
- all `.RRW`
- report forms `M_REP_*`
- `RNG_PRT_FRM.frm`
- report helper functions in `EFF_LIB.bas`

Output:

- `17_reporting_catalog.md`

### Admin / Setup / Backup

Files:

- `EFF_LOG.frm`
- `setup.frm`
- `BAC_UP_FRM.frm`
- `M_FRM_RST_BKP.frm`
- `M_FRM_QUK_RST.frm`
- `DBCLEAN.PRG`
- `db_restore.PRG`
- `setdb.PRG`
- `fy_xfr.PRG`
- `sincrement.PRG`

Output:

- `18_admin_setup_backup_spec.md`

## Extraction Aids

Useful commands:

```sh
rg -n "^(Public |Private )?(Sub|Function)" LegacyCodebase/M_FRM_SAL.frm
rg -n "SELECT |INSERT |UPDATE |DELETE |\\.Open |con.Execute" LegacyCodebase/M_FRM_SAL.frm
find LegacyCodebase -maxdepth 1 -type f -iname "*.frm" -exec wc -l {} + | sort -nr
```

## Golden-Master Test Candidates

Create fixtures for:

- simple sale invoice with one salable SKU
- sale invoice with free SKU
- sale invoice with SKU-level free quantity scheme
- sale invoice with bill-level cash off
- sale invoice with discount percent
- sale invoice with CSR/customer-specific rate
- sale invoice with route offset rounding
- purchase invoice updating batch stock
- stock transfer in/out
- van load and return
- HHT imported sale
- Tally voucher export

For each fixture, capture:

- input tables
- user-entered fields
- generated invoice/header rows
- generated invoice line rows
- stock changes
- ledger changes
- printed/report totals
