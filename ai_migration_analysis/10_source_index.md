# 10 Source Index

This is a quick navigation index for future AI models. It does not replace line-by-line extraction.

## Main VB6 Project Membership

The main app `../LegacyCodebase/Effmcg.vbp` includes these forms/modules/classes.

### Startup / Shell / Login / Setup

- `EFF_LOG.frm` - login and startup checks.
- `M_FRM_STR.frm` - project startup form.
- `M_MDI.frm` - MDI shell/menu/status bar.
- `setup.frm` - setup/company creation/default reset logic.
- `company.frm`, `DEL_COM.frm`, `INS_INF.frm` - company/install info/deletion setup.
- `FUN_KEY.frm` - function key setup.
- `SYS_PSW.frm`, `USR_FRM.frm`, `USR_PSW.frm` - password/user management.

### Master Data

- `M_FRM_ACT.frm` - account/customer master.
- `M_FRM_GRP.frm` - account/product group master.
- `M_FRM_MAR.frm` - market master.
- `M_FRM_SMN.frm` - salesman master.
- `M_FRM_PRE.frm` - prefix setup.
- `M_FRM_PGR.frm` - product group/category/class.
- `M_FRM_PRD.frm` - product master.
- `M_FRM_NBC.frm` - new batch creation.
- `M_FRM_ATT.frm` - attributes/attendance style master.
- `M_FRM_LOC.frm` - location.
- `M_FRM_CAT.frm` - category.
- `M_FRM_SHL.frm` - shell/crate master.
- `M_FRM_RUT.frm` - route.
- `M_FRM_SIC.frm` - scheme/item/category support.
- `M_FRM_SSI.frm` - sales/stock setup.
- `frmNewRateTemplate.frm` - rate/batch template.

### Finance / Accounts

- `M_FRM_PMT.frm` - payment.
- `M_FRM_RCT.frm` - receipt.
- `M_FRM_OPN.frm` - opening balance/stock.
- `M_FRM_DBN.frm` - debit note.
- `M_FRM_CRN.frm` - credit note.
- `M_FRM_JRN.frm` - journal.
- `M_FRM_IRT.frm` - invoice receipt/transaction.
- `M_FRM_DBT.frm` - debtor report/filter.
- `M_FRM_PQY.frm` - payment/query.
- `M_FRM_DRP.frm` - daily/report/payment style workflow.
- `M_FRM_PRC.frm` - price/rate processing.

### Purchase

- `M_PUR_PQY.frm` - purchase query.
- `M_FRM_POR.frm` - purchase order.
- `M_FRM_PUR.frm` - purchase invoice.
- `M_FRM_PPO.frm` exists in folder but is not listed in `Effmcg.vbp`; still important.

### Sales / Orders / Invoices

- `M_FRM_ORD.frm` - order.
- `M_FRM_LDSORD.frm` - load sheet order.
- `frmSalesCount.frm` - sales count and stock movement variant.
- `M_FRM_SOR.frm` - sales order.
- `M_FRM_SAL.frm` - main sales invoice.
- `M_FRM_SAL_PSH.frm` - sales variant.
- `M_FRM_RSI1.frm` - retailer sales invoice variant.
- `M_FRM_RSI_PRD.frm` - retailer sales invoice by product.
- `M_FRM_DSN.frm`, `M_FRM_DSN1.frm` - dispatch/sales note variants.
- `M_FRM_SIN.frm` - sales invoice/secondary invoice.
- `M_SAL_PRG.frm` - sales progress/report.

### Schemes / Discounts / Targets

- `M_FRM_SCH.frm` - scheme master/logic.
- `M_FRM_TGT.frm` - target.
- `M_REP_TGT.frm` - target report.
- `M_FRM_QPS.frm` - QPS scheme/quantity promotion support.

### Inventory / Stock

- `M_FRM_ADJ.frm` - adjustment.
- `M_FRM_STI.frm` - stock in.
- `M_FRM_STO.frm` - stock out.
- `M_FRM_SRT.frm` - stock/sales return.
- `M_FRM_TRS.frm` - transaction/shell return summary.
- `M_FRM_OPS.frm` - opening stock/operations.
- `M_FRM_UPD.frm` - stock update.
- `M_FRM_PPQ.frm` - product/stock query/report.
- `M_FRM_PRT.frm` - party/product/print support.
- `PRD_ANA_FRM.frm` - product analysis.
- `M_FRM_AST_OUT.frm`, `M_FRM_AST_INW.frm` - asset out/in.

### Van / HHT / Mobile / Field

- `M_FRM_MBL.frm` - mobile/HHT/market billing.
- `M_REP_MBL.frm` - mobile report.
- `M_REP_VAN.frm` - van report.
- `M_FRM_VVR.frm` - van voucher/report.
- `M_FRM_LDS.frm` - load sheet.
- `M_FRM_RAT.frm` - route/rate.

Important files present but not always listed in main project:

- `M_VAN_SAL.frm`
- `M_FRM_MIBS.frm`
- `M_FRM_HHT.frm`
- `DAS_XTR.frm`

### Empties / Shells / Deposits

- `M_FRM_EVP.frm` exists in folder and is very large; likely high priority even if not listed in main project membership shown.
- `M_REP_EMT.frm` - empties report.
- `M_REP_SHL.frm` - shell report.
- `M_FRM_TRS.frm` - aggregates `SNV` shell/deposit movement.

### Reports

- `M_REP_RET.frm`
- `M_REP_PQY.frm`
- `M_REP_TYP.frm`
- `M_REP_STR.frm`
- `M_REP_ASM.frm`
- `M_REP_SQY.frm`
- `M_REP_SHL.frm`
- `M_REP_TGT.frm`
- `M_REP_SMN.frm`
- `M_REP_VAN.frm`
- `M_REP_EMT.frm`
- `M_RPT_COM.frm`
- `RNG_PRT_FRM.frm`

### Admin / Utility

- `M_FRM_EXT.frm`
- `M_FRM_ASD.frm`
- `M_FRM_CLG_AQY.frm`
- `M_FRM_APD.frm`
- `M_FRM_ABT.frm`
- `M_FRM_ASM.frm`
- `M_FRM_WSR.frm`
- `M_FRM_MAS.frm`
- `M_FRM_COV.frm`
- `M_FRM_EFD.frm`
- `checkfiles.frm`
- `M_FRM_IFD.frm`
- `M_FRM_REF.frm`
- `M_FRM_CHK.frm`
- `M_FRM_SAQ.frm`
- `M_FRM_EXG.frm`
- `BAC_UP_FRM.frm`
- `M_FRM_RST_BKP.frm`
- `M_FRM_QUK_RST.frm`
- `M_LOG_INF.frm`
- `SHT_RSN.frm`
- `M_FRM_XRF.frm`

### Shared Modules

- `EFF_LIB1.bas`
- `modProcess.bas`
- `ModTallyInterface.bas`
- `modXml.bas`
- `modMessageTemplate.bas`
- `modSMSSend.bas`
- `modInnComm.bas`
- `modGeneral.bas`

### Classes

- `Class1.cls`
- `Billsave.cls` exists in folder.
- `encPassword.cls`, `dec.cls` exist in folder.

## FoxPro PRG Programs

High-interest `.PRG` files:

- `dsync.PRG` - synchronizes `INL/PNL` and `INV/PNV`, then indexes.
- `DBCLEAN.PRG` / `dbclean.PRG` - database cleaning.
- `db_restore.PRG` - restore.
- `setdb.PRG`, `setdbold.PRG` - DB reset/staging cleanup.
- `sincrement.PRG` - sequence increments.
- `fy_xfr.PRG` - financial-year transfer.
- `tax_upd.PRG`, `RCALTAX.PRG` - tax recalculation/update.
- `SCH_CHK.PRG` - scheme check.
- `DAY_CLS.PRG` - day close.
- `DAS_DRS.PRG`, `DAS_dBU.PRG` - DAS utilities.
- `IND.PRG`, `CND.PRG`, `PRI.PRG` - indexing/printing/report utilities.

## DBF Duplication Pattern

Many core DBFs appear up to 8 times across:

- root `LegacyCodebase`
- `Reports`
- `MMEXTRACT`
- `MMEXTRACT/AUTOBKP`
- `CRSEXTRACT`
- `CRSEXTRACT/AUTOBKP`
- `DASXtract`
- `FASTBKPEXE`

Core DBFs repeated 8 times include:

`ACT`, `AGD`, `ASM`, `AST`, `ATT`, `BRD`, `CAT`, `CLD`, `CSR`, `DIR`, `GRP`, `IDY`, `INL`, `INV`, `IRT`, `LDS`, `LED`, `LEV`, `LNV`, `LOC`, `LOG`, `MAR`, `PGR`, `PNL`, `PNV`, `PRD`, `PRE`, `PRT`, `RDY`, `RRI`, `SCH`, `SHL`, `SMN`, `SNV`, `STL`, `TGT`, `VAN`, `VCH`, `WEK`, `YER`.

Migration implication:

Identify the live authoritative DBF set before data migration. The generated Postgres script appears sourced from one chosen directory, but duplicated folders may contain alternate schemas or historical states.

## Key Source Search Patterns

Use these to continue extraction:

```sh
rg -n "^(Public |Private )?(Sub|Function)" LegacyCodebase/<file>
rg -n "SELECT |select |INSERT |insert |UPDATE |update |DELETE |delete |\\.Open |con.Execute" LegacyCodebase/<file>
rg -n "MTS\\(|STM\\(|RATCAL\\(|WGTCAL\\(|UNTCAL\\(" LegacyCodebase
rg -n "SCH|scheme|APPSCH|AutoApply|VLD_|ADD.*QTY|DIS_AMT|DIS_PCG" LegacyCodebase/M_FRM_SAL.frm
```
