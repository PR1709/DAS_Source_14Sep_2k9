# 03 Data Model

## Current Generated Postgres Tables

`../MigrateDest_v1/init_postgres.sql` defines these 96 tables:

`act`, `agd`, `asm`, `avt`, `ast`, `att`, `bkp`, `bnl`, `bnv`, `cat`, `cbk`, `cld`, `cnv`, `com`, `csr`, `cur`, `dir`, `dkcontrl`, `dsu`, `fil_lst`, `foxuser`, `fun`, `gcc`, `grp`, `hht`, `idy`, `inc`, `ine`, `irt`, `lds`, `led`, `lev`, `lnv`, `loc`, `log`, `mbr`, `mcl`, `npo`, `odir`, `opa`, `ops`, `optional`, `ord`, `pnl`, `pnv`, `pre`, `prt`, `rdy`, `required`, `rp1`, `rp2`, `rrunout`, `rtg`, `rtp`, `sbk`, `scl`, `scsr`, `shl`, `sic`, `sid`, `sku`, `snv`, `ssch`, `sta`, `str_doc`, `sts`, `tgt`, `top`, `tpd`, `trt`, `typ`, `van`, `vbk`, `vch`, `wek`, `xnl`, `xnv`, `yer`, `aca`, `brd`, `inl`, `inv`, `mar`, `obrd`, `ogcc`, `pgr`, `ppc`, `prd`, `qps`, `reports`, `rri`, `sch`, `smn`, `ssi`, `stl`, `tgb`, `tgc`, `tgs`, `usr`, `wsr`.

## Key Entity Tables

### `ACT` - Account / Retailer / Supplier / Van Master

Purpose:

- customer/retailer master
- supplier/account ledger master
- van pseudo-account entries
- contact/address/tax data
- route/market/salesman/category assignments
- credit/debit controls

Important columns from generated DDL:

- `act_idy`, `act_nme`
- `grp_idy`, `grp_clf`
- address/contact: `act_ad1` to `act_ad4`, `act_cpn`, `act_mbl`, `act_zip`
- tax: `act_cst`, `act_gst`
- type/classification: `act_typ`, `cst_typ`, `ret_typ`
- distribution mapping: `rou_idy`, `loc_idy`, `mar_idy`, `smn_idy`, `cat_idy`
- credit model: `crt_mdl`, `crt_ppd`, `crt_typ`, `crt_dis`, `crt_rap`
- `van_cap`

Observation:

`ACT` mixes multiple concepts and should be decomposed into account, party, customer, supplier, vehicle/van, contact/address, tax profile, and route assignment tables in the new model.

### `PRD` - Product / SKU Master

Purpose:

- SKU identity, brand, pack, rates, taxes, conversion, stock behavior.

Important columns:

- `prd_idy`, `prd_nme`, `prd_prn`
- `grp_idy`, `cmp_idy`
- `brd_nme`, `prd_pck`, `prd_pkt`
- taxes/discounts: `pur_tax`, `sal_tax`, `pur_dis`, `sal_dis`
- rates: `sal_rat`, `pur_rat`, `mrp_rat`, `bas_rat`, `lst_rat`
- stock levels: `cur_lev`, `reo_lev`, `saf_lev`
- conversion: `con_fac`, `prd_uom`, `sub_unt`
- `stk_typ`, `prd_pcs`, `brw_seq`, `tax_typ`

Observation:

SKU, brand, package, batch, price, and stock item are partly blended. `BRD` appears to hold batch/price/quantity details and must be modeled carefully.

### `BRD` - Batch / Brand / Price / Stock Detail

Purpose inferred from source:

- batch ID (`bat_idy`)
- product linkage (`prd_idy`)
- batch/display name (`bat_nme`)
- batch quantity (`prd_qty`)
- purchase/sale rates, MRP, LPP, crate rates/taxes
- expiry/manufacturing fields used in forms

Source usage:

- product lookup joins `PRD` and `BRD`
- stock checks compare `BRD.PRD_QTY`
- scheme free lines store `BAT_IDY`
- price update and new rate templates insert/update `BRD`

### `INL` - Invoice Header

Purpose:

- transaction header for sale, purchase, stock and related invoice/challan workflows.

Important columns:

- `inv_idy`, `inv_dat`, `ord_idy`
- parties/accounts: `cst_idy`, `act_idy`, `smn_idy`
- amounts: `grs_amt`, `spd_pcg`, `spd_amt`, `csd_pcg`, `csd_amt`, `exd_pcg`, `exd_amt`, `tax_pcg`, `tax_amt`, `net_amt`
- transaction: `txn_typ`, `txn_rmk`, `txn_tpt`
- logistics: `dch_idy`, `dch_dat`, `lrn_idy`, `lrn_dat`, `wbl_idy`, `wbl_dat`
- freight/package/cash: `frt_typ`, `frt_amt`, `pkg_cst`, `ttl_cas`
- route/due: `rou_idy`, `rou_off`, `crt_day`, `due_dat`
- scheme fields: `sch_chk`, `sch_idy`, `srn_amt`, `dis_chk`
- audit: `usr_idy`, `sys_dat`, `sys_tim`

### `INV` - Invoice Lines / Stock Movements

Purpose:

- line items for purchases, sales, free goods, stock movement, scheme benefit lines.

Important columns:

- `inv_idy`, `inv_dat`, `prd_idy`, `bat_idy`
- rates/discount/tax: `lst_rat`, `lst_dis`, `prd_rat`, `cpn_rat`, `cpn_amt`, `add_dis`, `add_amt`, `trd_dis`, `dis_amt`, `trd_tax`, `tax_amt`
- quantities: `prd_qty`, `ret_qty`, `opn_qty`, `iss_qty`, `tot_qty`
- piece conversion: `prd_uom`, `sub_unt`, `prd_unt`, `prd_pcs`, `tot_pcs`
- `stk_upd`, `txn_typ`, `stk_typ`
- scheme: `sch_amt`, `sch_idy`, `csr_idy`, `inc_ttl`

Observation:

This is both commercial invoice line and inventory movement. New design should separate invoice line from stock ledger entries, while keeping traceability.

### `SCH` - Scheme / Promotion Definition

Purpose:

- trade promotions, free goods, cash-off, discount percent, bill-level and product/group-level schemes.

Important columns:

- `sch_idy`, `sch_nme`
- `sch_typ`: SKU Level, Bill Level, Multiple SKU Level, Brand, Pack Size, Pack Type, Segment, Category
- `f_sch_typ`: Free Quantity, Cash Off, Discount Percent
- base/free products: `bat_idy`, `b_prd_idy`, `f_prd_idy`
- base/free quantities: `b_prd_qty`, `e_prd_qty`, `f_prd_qty`
- amount/percent: `b_prd_amt`, `b_dis_pcg`, `b_dis_amt`, `b_bil_amt`, `e_bil_amt`
- validity: `sch_bdt`, `sch_edt`
- `txn_typ`: `BAS` or `FRE`
- `sch_sts`
- pro-rata and conversion flags: `pro_rat`, `b_con_app`, `f_con_app`
- group/merchandise: `grp_typ`, `mer_num`, `mer_qty`, `mer_amt`

### `SMN` - Salesman

Purpose:

- salesman identity, address/contact, route/work assignments. The guidance notes day columns/capacity may exist in DBF variants.

### `MAR`, `LOC`, `CAT`, `PGR`, `GRP`

Purpose:

- market, location, category, product group/segment/class, account group.

### `VAN`

Purpose:

- van challan/load/return/reconciliation table.
- Used in product availability when `M_VAN_IDY` is present.
- Query logic computes net load = loaded total - returned/leakage/breakage.

### `SHL` / `SNV`

Purpose:

- shell/crate/empties/deposit movement.
- `M_FRM_TRS.frm` aggregates `SNV` quantities for issued, returned, breakage, deposits.

### `VCH`, `IRT`, `LED`

Purpose:

- vouchers, receipts/payments, ledger/report staging.

## Duplicate / Staging Tables

The system uses pairs like:

- `INL` and `PNL`
- `INV` and `PNV`

`dsync.PRG` compares counts and zaps/appends one side into the other depending on count and date. This suggests one is primary and one is print/report/offline or sync staging, but this must be confirmed before redesign.

## New Model Direction

Suggested canonical model:

- `accounts`
- `parties`
- `customers`
- `suppliers`
- `vehicles`
- `salespeople`
- `routes`
- `markets`
- `products`
- `brands`
- `packages`
- `batches`
- `price_lists`
- `tax_profiles`
- `schemes`
- `scheme_conditions`
- `scheme_benefits`
- `invoices`
- `invoice_lines`
- `stock_movements`
- `van_loads`
- `van_load_lines`
- `empties_movements`
- `payments`
- `ledger_entries`
- `integration_events`
- `audit_log`

Keep legacy IDs as alternate keys during migration.
