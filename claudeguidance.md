
I've extracted the zip and gone through the codebase systematically — file inventory, DBF schema headers (parsed the actual binary structure of 100+ tables), and read the real source in the VB6 forms/modules and FoxPro programs. Before the full plan, an honest scoping note:
What I actually did vs. what "every file" would mean
This is not a small app. It's 2,729 files: 743 .DBF tables, 214 VB6 .frm forms (many 150–460KB each, containing thousands of lines of embedded VB code), 216 .frx binary resource twins, 57 FoxPro .PRG programs, 26 .bas modules, 44 FoxPro project files, plus ~300 legacy Crystal/DataDynamics report definitions (.RP1/.RP5/.RRW) and third-party runtime DLLs. Genuinely reading every line of every .frm in one pass isn't something I can responsibly claim to have done — that's realistically a multi-week reverse-engineering effort even for a human team. What I did do:

Enumerated and categorized all 2,729 files by type/module
Parsed the raw DBF headers of every table in the root folder to get exact field names/types/lengths (not guesses — read directly off disk)
Read the full source of the small .bas/.PRG utility programs
Extracted and read real embedded VB6 code (via strings) from the largest forms (sales invoice, scheme calc) to confirm actual formulas
Found there's already a MigrateDest_v1/ folder someone started, with a Postgres DDL dump generated from the DBFs

For the deep, form-by-form logic audit (all 214 forms), the right tool is Claude Code pointed at this extracted folder, going module-by-module — I'll propose that workflow below. Here's everything I can tell you now.

1. What this system actually is
   This is DAS — a Direct Store Distribution (DSD) ERP for a beverage bottler/distributor, almost certainly a Coca-Cola franchisee distributor in India (Coca-Cola_Bar.bmp, cspl branding assets, logosn.tif). Stack:

Backend: Visual FoxPro 6 database container (EFF.DBC), 743 .DBF tables, accessed via VFP ODBC driver
Front end: VB6 (Effmcg.vbp), ~150+ large forms — classic client-server, single shared DB file, LAN-only
Batch/reporting layer: Native FoxPro .PRG programs + Crystal-Reports-era report definitions
Integrations: Tally ERP (via RTSlink COM + XML over HTTP — confirmed in ModTallyInterface.bas), SMS gateway (frmSMS.frm, SMSReceiveLog.txt), and a Handheld Terminal (HHT) sync layer — flat-file exports (HHT/*.TXT: STOCK, PARTY, LOAD_SHEET, SALESMAN, EMPTIES, etc.) for field sales reps' handheld devices, synced back via dsync.PRG.

2. Confirmed business domain (from actual schema + code)
   Reading the real DBF headers gives an accurate picture — this is a van-sales / pre-sales distribution model with:
   TablePurposeKey fields foundACTCustomer/retailer mastercredit limit (CRT_MDL), route (ROU_IDY), salesman (SMN_IDY), market (MAR_IDY), category, GST/CST tax IDsPRDProduct masterpack size, sale/purchase rate, MRP, sale/purchase tax %, reorder/safety stock levels, UOM conversion factorVANVan/vehicle-level daily stock ledgeropening/issued/returned/billed/free/leakage/breakage qty, cash collected, deposit received, routeSCHTrade schemes/promotionsbuy-qty/free-qty slabs, discount %, date-bound, "merchandise" (MER) qty/amtCSRScheme claim/settlementmerchandiser claims, usage/balance tracking against schemes, base+free product legsINV/SNVInvoice line items / empties (crates, shells, deposits)tax amount, scheme amount, deposit amount, breakageSMNSalesman masterper-day route capacity (Sun–Sat columns), depositMARMarket/beat masterdistributor van type, weekly distribution patternPNL/PNV (from the migration dump)Invoice header/line — full commercial calc: gross, trade discount %, cash discount %, excise, tax %, net, freight, way-bill/L.R. tracking
   From the actual VB code in M_FRM_SAL.frm (sales invoice form — 464KB, the single largest file), the scheme/discount engine is a slab-based calculator:
   SCH_AMT = SCH_AMT + PERCENT(PRDVAL, DISPCG)              ' percentage-slab discount
   SCH_AMT = SCH_AMT + (Int(QTY / MAXQTY) * DISAMT)          ' buy-X-get-Y-style slab, repeating per multiple
   M_TAX_AMT = (line_tax / line_qty) * MTS(...) * tax_pcg / 100
   i.e., schemes stack a percentage-based leg and a quantity-slab leg (repeated for every MAXQTY multiple bought), tax is computed pro-rata per unit off a blended rate, and there's a separate SCAMT/LSCH_AMT running local-vs-cumulative scheme total — this is genuinely intricate and is exactly the kind of logic that needs a dedicated line-by-line extraction pass per form before you'd trust a rewrite of it. I'm flagging it, not fully reverse-engineering it here.
3. Full module map (by file volume — proxy for business weight)
   DomainForms/ProgramsNotesSales / Van SalesM_FRM_SAL*, M_VAN_SAL, M_FRE_SAL*Largest, most complex — invoice, scheme, freight, van reconciliationPurchaseM_FRM_PUR, M_FRM_POR, M_FRM_PPOPO → GRN → purchase invoiceStock/InventoryM_FRM_STI, M_FRM_STO, M_FRM_SRT, SCH_CHK.PRGStock-in/out, transfer, scheme stock checksRoute/Order/Sales OpsM_FRM_SOR, M_FRM_SSO, M_FRM_RSI*, M_FRM_RTGRoute-wise sales order, retailer sales invoiceEmpties/DepositsSNV, SHL, CLD tables, M_FRM_EVPCrate/shell/bottle deposit tracking — common in bottler DSDHHT SyncHHT/*.TXT, dsync.PRGFlat-file push/pull to handheld terminalsTally IntegrationModTallyInterface.bas, frmSMS.frm (SMS)XML-over-COM to Tally accountingMaster DataM_FRM_ACT, M_FRM_PRD, M_FRM_MAR, M_FRM_SMNCustomer, product, market, salesman mastersAdmin/UtilityDBCLEAN.PRG, sincrement.PRG, chg_str.PRGDB maintenance, ID sequencing, year-end rolloverReporting~300 .RP1/.RP5/.RRWLegacy Crystal/DataDynamics — will not port; rebuild as BI views
4. Target architecture: modern, AI-native, agentic
   Given the domain (field sales force, route-based delivery, van reconciliation, scheme-heavy pricing, multi-party ledger), here's the shape I'd build toward:
   Core platform

Postgres as system of record (your MigrateDest_v1/init_postgres.sql scaffold is a start, but it's a raw 1:1 DBF dump — needs real normalization: ACT/CSR conflate customer + scheme-claim concerns, _NullFlags VFP artifacts should disappear, and IDs like ACT_IDY (char(10)) should become proper UUIDs/sequences)
Domain services (NestJS/FastAPI, pick one per your stack preference) split along the module map above: catalog, pricing-schemes, sales-orders, van-reconciliation, inventory, finance-ledger, route-planning
Event-driven core (Postgres logical replication or Kafka) for van-close, invoice-post, scheme-apply — since van reconciliation and Tally sync are inherently async, this maps far better than the current file-based sync
API layer: GraphQL or REST + tRPC for the web front end; gRPC/REST for the mobile/field app

Field force replaces HHT

Native mobile app (React Native/Flutter) replacing the flat-file HHT sync entirely — real-time order capture, stock visibility, offline-first with conflict resolution instead of nightly TXT exchange

AI-native / agentic layer (this is where it gets interesting for a DSD business)

Scheme/pricing agent: given the slab logic you have today is opaque and error-prone (the SCH_AMT/LSCH_AMT duplication I found is a smell), an LLM-assisted rules engine that can explain why a scheme applied, catch conflicting/stacking schemes, and simulate margin impact before a scheme goes live
Route optimization agent: market/route/salesman data (MAR, SMN, ROU_IDY) feeds a route-planning agent (like existing VRP solvers + LLM for exception handling — "this retailer missed 3 visits, reprioritize")
Demand forecasting agent: PRD already tracks reorder/safety levels — replace static thresholds with a forecasting agent per SKU/market using van sales history
Reconciliation/anomaly agent: van cash-vs-stock reconciliation (OLD_CRD, CUR_CRD, SHT_CSH fields show this was manually tracked) is a natural fit for an agent that flags shortages, breakage patterns, or salesman-level anomalies automatically
Conversational ops assistant: distributor/route-owner asks "why did Route 4 miss target this week" and the agent queries the warehouse + explains, backed by the same event stream

Integrations: keep Tally sync but replace the fragile RTSlink COM/XML string-splicing with a proper Tally XML API client; keep SMS but move to a modern provider API.
5. Realistic migration path (given the scale)

Schema extraction & data audit — finish parsing all 743 DBF headers (I did the root folder; HHT/, Reports/, DASXtract/, CRSEXTRACT/, MMEXTRACT/ subfolders still need the same treatment) → produce a canonical ERD, not a 1:1 dump
Business logic extraction, module by module — starting with Sales/Scheme (highest risk, highest file weight) — this needs Claude Code run locally against the extracted repo, form by form, producing a spec doc per module before any code is written
Strangler-fig cutover — stand up new services behind the same data (dual-write to Postgres + VFP during transition), migrate one workflow at a time starting with master data (lowest risk), then sales entry, then scheme engine, then reporting last
Reporting rebuild last — the ~300 legacy report files are the least valuable to reverse-engineer line-by-line; better to rebuild from the new normalized schema against actual reporting requirements

Want me to (a) finish the DBF schema extraction across the remaining subfolders and produce a clean ERD/migration doc, or (b) start the module-by-module logic extraction on the Sales/Scheme forms specifically, since that's the highest-complexity, highest-value piece? Given the scale, Claude Code running against this extracted folder directly would be a much better fit than continuing purely in this chat.
