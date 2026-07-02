# Graph Report - .  (2026-07-02)

## Corpus Check
- cluster-only mode — file stats not available

## Summary
- 143 nodes · 93 edges · 58 communities (21 shown, 37 thin omitted)
- Extraction: 73% EXTRACTED · 27% INFERRED · 0% AMBIGUOUS · INFERRED: 25 edges (avg confidence: 0.76)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `bb8258f1`
- Run `git rev-parse HEAD` and compare to check if the graph is stale.
- Run `graphify update .` after code changes (no API cost).

## Community Hubs (Navigation)
- [[_COMMUNITY_MM_EX_FILE|MM_EX_FILE]]
- [[_COMMUNITY_DBFReader|DBFReader]]
- [[_COMMUNITY_EFF.DBC|EFF.DBC]]
- [[_COMMUNITY_EFF.DBC (FoxPro Database Container)|EFF.DBC (FoxPro Database Container)]]
- [[_COMMUNITY_MM_EX_FILE|MM_EX_FILE]]
- [[_COMMUNITY_Index.vbp|Index.vbp]]
- [[_COMMUNITY_LOG.TXT|LOG.TXT]]
- [[_COMMUNITY_ACT.DBF|ACT.DBF]]
- [[_COMMUNITY_EFF_LIB.bas|EFF_LIB.bas]]
- [[_COMMUNITY_Table idy|Table: idy]]
- [[_COMMUNITY_Class1.cls|Class1.cls]]
- [[_COMMUNITY_Class1.cls|Class1.cls]]
- [[_COMMUNITY_ModTallyInterface.bas|ModTallyInterface.bas]]
- [[_COMMUNITY_dsync.PRG|dsync.PRG]]
- [[_COMMUNITY_ModTallyInterface.bas|ModTallyInterface.bas]]
- [[_COMMUNITY_ETLInterfaceForm.frm|ETLInterfaceForm.frm]]
- [[_COMMUNITY_FOXUSER.DBF|FOXUSER.DBF]]
- [[_COMMUNITY_Class1.cls|Class1.cls]]
- [[_COMMUNITY_M_FRM_MIBS.frm|M_FRM_MIBS.frm]]
- [[_COMMUNITY_M_FRM_TRS.frm|M_FRM_TRS.frm]]
- [[_COMMUNITY_foxuser (Metadata)|foxuser (Metadata)]]
- [[_COMMUNITY_Table lev|Table: lev]]
- [[_COMMUNITY_STM|STM]]
- [[_COMMUNITY_Direct Store Distribution (DSD) ERP|Direct Store Distribution (DSD) ERP]]
- [[_COMMUNITY_COM Port Settings|COM Port Settings]]
- [[_COMMUNITY_Effmcg4 Path Configuration|Effmcg4 Path Configuration]]
- [[_COMMUNITY_CSPL Logo|CSPL Logo]]
- [[_COMMUNITY_Consoft Logo|Consoft Logo]]
- [[_COMMUNITY_Consoft Logo|Consoft Logo]]
- [[_COMMUNITY_Consoft Logo|Consoft Logo]]
- [[_COMMUNITY_Consoft Logo|Consoft Logo]]
- [[_COMMUNITY_Consoft Logo|Consoft Logo]]
- [[_COMMUNITY_DAS_DBU.TXT (VFP RuntimeUtility)|DAS_DBU.TXT (VFP Runtime/Utility)]]
- [[_COMMUNITY_Help Desk Contacts|Help Desk Contacts]]
- [[_COMMUNITY_Consoft Logo|Consoft Logo]]
- [[_COMMUNITY_Consoft Logo|Consoft Logo]]
- [[_COMMUNITY_Consoft Logo|Consoft Logo]]
- [[_COMMUNITY_Accounting Primary Groups|Accounting Primary Groups]]
- [[_COMMUNITY_M_FRM_SCH.frm|M_FRM_SCH.frm]]
- [[_COMMUNITY_Table fun|Table: fun]]
- [[_COMMUNITY_Table gcc|Table: gcc]]
- [[_COMMUNITY_Table grp|Table: grp]]
- [[_COMMUNITY_Table hht|Table: hht]]
- [[_COMMUNITY_Table inc|Table: inc]]
- [[_COMMUNITY_Table lds|Table: lds]]
- [[_COMMUNITY_Table loc|Table: loc]]
- [[_COMMUNITY_MMExtract.EXE|MMExtract.EXE]]

## God Nodes (most connected - your core abstractions)
1. `MM_EX_FILE` - 16 edges
2. `DBFReader` - 7 edges
3. `generate_postgres_dump()` - 5 edges
4. `EFF.DBC` - 5 edges
5. `Index.vbp` - 4 edges
6. `MM_EX_FILE` - 4 edges
7. `EFF.DBC (FoxPro Database Container)` - 4 edges
8. `clean_name()` - 3 edges
9. `parse_source_files()` - 3 edges
10. `Table: idy` - 3 edges

## Surprising Connections (you probably didn't know these)
- `INL.DBF Structure` --references--> `EFF.DBC (FoxPro Database Container)`  [INFERRED]
  LegacyCodebase/INLSTRU.TXT → claudeguidance.md
- `Cola Connect Smarter Decisions Logo` --conceptually_related_to--> `DAS Mobile GSM Interface`  [INFERRED]
  LegacyCodebase/NCC.JPG → LegacyCodebase/Welcome Message.txt
- `MM_EX_FILE` --conceptually_related_to--> `DAS Mobile GSM Interface`  [INFERRED]
  LegacyCodebase/CRSEXTRACT/MMExtract.txt → LegacyCodebase/Welcome Message.txt
- `M_FRM_TRS.frm` --calls--> `SNV.DBF`  [EXTRACTED]
  MigrateDest_v1/initdb.md → LegacyCodebase/MMEXTRACT/VER_APP.TXT
- `M_FRM_PCS_PCS.frm` --calls--> `INV.DBF`  [EXTRACTED]
  MigrateDest_v1/initdb.md → LegacyCodebase/MMEXTRACT/VER_APP.TXT

## Import Cycles
- None detected.

## Hyperedges (group relationships)
- **DAS Data Extraction and Transfer Flow** — legacycodebase_crsextract_mmextract_mm_ex_file, legacycodebase_dbf_act, legacycodebase_dbf_inv, legacycodebase_dbf_inl, legacycodebase_dbf_prd, legacycodebase_dbf_brd [EXTRACTED 0.90]
- **Visual FoxPro Database Core** — eff_dbc, eff_dct, eff_dcx [EXTRACTED 1.00]
- **DAS Extraction Utility Project** — index_vbp, indexdata_exe, fbkp_bas, cmp_sel_frm, m_frm_str_frm [EXTRACTED 1.00]
- **HHT Data Exchange Set** — hht_party_txt, hht_location_txt, hht_market_txt, hht_salesman_txt, legacycodebase_hht_inv_str_txt, legacycodebase_hht_inl_str_txt [INFERRED 0.90]
- **Hand Held Terminal Data Files** — legacycodebase_hht_party, legacycodebase_hht_salesman, legacycodebase_hht_version [EXTRACTED 0.90]
- **DAS Data Extraction and Export Flow** — legacycodebase_mmextract_mmextract_mm_ex_file, legacycodebase_mmextract_dir_inf, concept_das_data_transfer [EXTRACTED 0.95]
- **EFFMCG Application Suite** — effmcg_exe, eff_dbc, mmextract_exe, m_frm_trs_frm, m_frm_pcs_pcs_frm [INFERRED 0.85]
- **SMS Notification Subsystem** — legacycodebase_smsmessagetemplate_txt, legacycodebase_comsettings_txt, claudeguidance_effmcg_vbp [INFERRED 0.85]
- **Legacy Data Extraction Flow** — mmextract_exe, eff_dbc, foxpro_driver, etlinterfaceform_frm [INFERRED 0.80]
- **Accounting Schema Core** — migratedest_v1_initdb_grp, migratedest_v1_initdb_led, migratedest_v1_initdb_idy [INFERRED 0.90]
- **Inventory Transaction Flow** — migratedest_v1_initdb_ine, migratedest_v1_initdb_lnv, migratedest_v1_initdb_loc [INFERRED 0.85]
- **DAS Core ERP Modules** — claudeguidance_m_frm_sal, claudeguidance_modtallyinterface, claudeguidance_dsync_prg, claudeguidance_eff_dbc [EXTRACTED 1.00]
- **Consoft Systems Branding Assets** — legacycodebase_cspl_png, legacycodebase_csplnew_gif, legacycodebase_csplnew_jpg, legacycodebase_csplnew_png, legacycodebase_cspl_tiny_gif, legacycodebase_logo_gif, legacycodebase_logo1_gif, legacycodebase_logo2_gif [EXTRACTED 1.00]
- **Core Business Forms** — legacycodebase_m_frm_sal_frm [EXTRACTED 0.90]
- **Stock Update Logic** — ai_migration_analysis_05_business_logic_extraction_n_upd_stk, ai_migration_analysis_05_business_logic_extraction_stl_led, ai_migration_analysis_05_business_logic_extraction_eff_lib [EXTRACTED 0.90]

## Communities (58 total, 37 thin omitted)

### Community 0 - "MM_EX_FILE"
Cohesion: 0.11
Nodes (18): DAS Mobile GSM Interface, GetValue, MailSend, MM_EX_FILE, STKTXTDEX, ACT.DBF, BRD.DBF, CAT.DBF (+10 more)

### Community 1 - "DBFReader"
Cohesion: 0.24
Nodes (8): clean_name(), DBFReader, generate_postgres_dump(), map_type_to_postgres(), parse_source_files(), Scan source files (.frm, .bas, .cls, .prg) for SQL queries, table references, an, A simple, dependency-free reader for dBase / Visual FoxPro .dbf and .fpt files., Clean table or column names to be PostgreSQL compatible.

### Community 3 - "EFF.DBC"
Cohesion: 0.29
Nodes (7): DASXtract.exe, EFF.DBC, EFF.DCT, EFF.DCX, EFFMCG.exe, FastBkp.exe, Prince Chopra Volume

### Community 4 - "EFF.DBC (FoxPro Database Container)"
Cohesion: 0.33
Nodes (6): EFF.DBC (FoxPro Database Container), Effmcg.vbp (VB6 Project), M_FRM_SAL.frm (Sales Invoice Form), INL.DBF Structure, SMS Message Templates, File Inventory (VER_APP.TXT)

### Community 5 - "MM_EX_FILE"
Cohesion: 0.33
Nodes (5): ColaZip.GSColaZip32, DAS Data Transfer Protocol, MailSend, MM_EX_FILE, STKTXTDEX

### Community 6 - "Index.vbp"
Cohesion: 0.40
Nodes (5): CMP_SEL.frm, fbkp.bas, Index.vbp, IndexData.exe, M_FRM_STR.frm

### Community 7 - "LOG.TXT"
Cohesion: 0.40
Nodes (4): LOCATION.TXT, MARKET.TXT, PARTY.TXT, SALESMAN.TXT

### Community 8 - "ACT.DBF"
Cohesion: 0.50
Nodes (4): ACT.DBF, INV.DBF, M_FRM_ORD.frm, M_FRM_PCS_PCS.frm

### Community 9 - "EFF_LIB.bas"
Cohesion: 0.50
Nodes (4): EFF_LIB.bas, MTS, N_UPD_STK, STL_LED

### Community 10 - "Table: idy"
Cohesion: 0.50
Nodes (4): Table: idy, Table: ine, Table: irt, Table: led

## Knowledge Gaps
- **91 isolated node(s):** `delete`, `update`, `delete`, `delete`, `update` (+86 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **37 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Are the 3 inferred relationships involving `EFF.DBC` (e.g. with `DASXtract.exe` and `EFFMCG.exe`) actually correct?**
  _`EFF.DBC` has 3 INFERRED edges - model-reasoned connections that need verification._
- **What connects `delete`, `update`, `delete` to the rest of the system?**
  _94 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `MM_EX_FILE` be split into smaller, more focused modules?**
  _Cohesion score 0.1111111111111111 - nodes in this community are weakly interconnected._