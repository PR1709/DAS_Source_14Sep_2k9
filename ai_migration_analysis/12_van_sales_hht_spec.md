# 12. Van Sales, Field Force, and HHT (Hand Held Terminal) Subsystem

This document outlines the business logic for the Van Sales and Field Force components in the legacy VB6 system. It covers how Van Load Sheets are created, how mobile numbers (SIM cards) are tracked and billed, and how data is synchronized with Hand Held Terminals (HHT/PDA devices).

## 1. Van Sales & Load Sheet (`M_VAN_SAL.frm`)

The Van Sales module allows users to assign inventory to a specific "Van" (Challan) and salesman. 

### Key Discoveries
- **Not a separate inventory warehouse**: Unlike standard ERPs that move stock to a Transit/Van warehouse, `M_VAN_SAL.frm` operates by directly creating a Sales Invoice (`TXN_TYP = 'SAL'`).
- **Load Sheet Identifier**: The Van/Challan ID is mapped to the `ORD_IDY` column in the `INV` (Invoice) and `INL` (Invoice Lines) tables. A valid Van Invoice must have an `ORD_IDY` of exactly 10 characters (`LEN(TRIM(ord_idy)) = 10`).
- **Reconciliation**: Because Van Sales directly decrement main inventory (`PRD.CUR_LEV`), any unsold goods must be returned via standard Stock Returns (`RSR` - Return Sales Receipt), which is handled by querying `INL` for `TXN_TYP LIKE 'RSR'`.

### Core Subroutines

#### `saverecord`
- **Source**: `M_VAN_SAL.frm:2130-2354`
- **Tables Read**: `INV`, `INL`, `PRD`, `IRT`, `ACT`, `SMN`
- **Tables Written**: `INV`, `INL`, `PRD`, `IRT`
- **Logic**: 
  - Validates required fields, Dealer (`ACT`), Salesman (`SMN`), and valid Invoice ID.
  - If modifying (`MOD_REC = True`), it identifies previously saved stock quantities. It queries `RSR` (Returns) and recalculates the difference. It restores the original stock back to `PRD.CUR_LEV` by calling `ADDQTY()`.
  - It then deletes the existing `INV` records.
  - Finally, it iterates through the `MS` Grid and inserts new `INV` rows (`TXN_TYP = 'SAL'`), recalculates Trade Discount and Taxes, and inserts Ledger entries into `IRT`. 

#### `dispdetails`
- **Source**: `M_VAN_SAL.frm:3613-3710`
- **Tables Read**: `INL`, `ACT`, `PRD`, `SMN`, `INV`
- **Logic**: 
  - Fetches a Van Sale by querying `INL WHERE inv_idy = M_INV_IDY AND txn_typ = 'SAL' AND LEN(TRIM(ord_idy)) = 10`. 
  - Populates the grid (`MS`) by joining with `PRD` to get Product Names, Conversion Factors (`CON_FAC`), and prices.

---

## 2. Mobile Invoice Billing / Activations (`M_FRM_MBL.frm`)

Despite its name ("Mobile Billing"), this massive 11k+ line file is **not** a PDA/Van sales app. It is a specialized module for **Tracking and Selling Serialized Mobile Numbers (e.g., SIM Cards)**. It interacts heavily with the `AVT` (Activations) table.

### Key Discoveries
- **`AVT` Table**: Tracks the lifecycle of individual 10-digit mobile numbers/SIM cards from Purchase (`PUR`) to Sale (`SAL`) to Return (`SRN`).
- **File Imports**: The module can import Excel/CSV files containing Mobile Numbers (using `MS Excel Object Library` and `CommonDialog1`).

### Core Subroutines

#### `updateavt`
- **Source**: `M_FRM_MBL.frm:10761-10860`
- **Tables Read**: `DIR`, `AVT`
- **Tables Written**: `AVT`
- **Logic**: 
  - Iterates through stock grids (Purchase, Stock In, Opening Stock).
  - Validates that each entered number is exactly 10 digits long.
  - Deletes any empty records in `AVT` for the current invoice.
  - Inserts new Mobile Numbers into `AVT` with their corresponding Product ID (`PRD_IDY`) and Transaction Type (`PUR`, `OPS`, `STI`).

#### `cmdsUpdate_Click` (Command Sales Update)
- **Source**: `M_FRM_MBL.frm:4121-4221`
- **Tables Read**: `INV`, `AVT`, `PRD`
- **Tables Written**: `AVT`
- **Logic**: 
  - Links sold Mobile Numbers to specific Sales Invoices. 
  - It iterates over selected items in a ListView (`lv`) containing Mobile Numbers.
  - Updates the `AVT` table to set `sal_idy` (Sales Invoice ID), `sal_dat` (Sale Date), and `cst_idy` (Customer ID) for each mobile number, transitioning its status to Sold.

#### `fileimpcmd_Click` & `Act_imp_Click`
- **Source**: `M_FRM_MBL.frm:2620-2791` (and others)
- **Tables Read**: `RP2` (Temporary Report Table)
- **Tables Written**: `RP2`
- **Logic**: 
  - Uses `Excel.Application` to open `.xls` files provided by Telecom Operators.
  - Imports Activation details, matching `CHR_ONE` to the Mobile Number. 
  - Temporarily stores imported data in `RP2` before processing it against the `AVT` database to mark SIMs as activated (`AVT_STS = 'Yes'`).

#### Grid Loading / Checkbox Filters (Lines 1601 - 3500)
- **Source**: `M_FRM_MBL.frm:2793-3500`
- **Tables Read**: `AVT`, `PRD`, `ACT`
- **Logic**: 
  - This segment contains extensive UI event handlers (`KeyDown`, `LostFocus`, `Click`) for search filters and "Select All" checkboxes (`chkAAll`, `chkCAll`, `chkRAll`).
  - When a filter changes, inline SQL directly queries the `AVT` table to find pending records (e.g., pending activations, pending CAFs where `caf_sts is null`, pending churns where `chr_sts is null`).
  - Fetched records are manually iterated and bound row-by-row into `MSFlexGrid` controls (`MSA`, `MSC`, `MSR`).
  - Noticeable performance bottleneck: Inside the grid population loop, it performs individual `SELECT` queries against `PRD` (Products) and `ACT` (Accounts) for every single row to resolve IDs to names (N+1 query problem).

---

## 3. Data Synchronization & PDA Interfacing

The legacy system uses a serial interface (COM Ports) via `MSCOMM32.OCX` to communicate with physical Hand Held Terminals (HHTs).

### `M_FRM_MIBS.frm` (Mobile Invoice Billing System Sync)
- **Purpose**: "DAS Mobile Data XChange". Acts as the data pump between the DAS Database and the mobile devices.
- **Components**: Contains `MSComm1` (COM Port controller), `Timer1` (Poller), and Progress Bars (`PBar1`).
- **Logic**: 
  - Exports Product Master Data (`Command4`), Retailer Master Data (`Command1`), and syncs Invoices (`Command3`).
  - Serializes data into specific ASCII layouts to transmit over RS-232/USB-to-Serial.

### `M_FRM_HHT.frm` (Hand Held Terminal Details)
- **Purpose**: Displays and configures HHT synchronization status.
- **Logic**: 
  - Manages the UI for synchronization sessions.
  - Shows connected devices and their sync status using a Tabbed Dialog (`SSTab1`) and ImageLists for status icons (Pending, Success, Error).

---

## 4. Unresolved Items & Assumptions
- **ASSUMPTION**: `M_FRM_MBL.frm` is strictly for SIM card / Mobile number tracking (common in FMCG distributors that also distribute prepaid SIMs) and has no relation to Van/Truck routing.
- **UNRESOLVED**: The exact text/binary format expected by the HHT devices (`dsync.PRG` and `HHT/*.TXT`) relies on older DOS-based parsing which would need to be replaced with standard REST APIs (JSON/XML) in modern implementations.

#### Bulk Updates & Batch Processing (Lines 3501 - 5500)
- **Source**: `M_FRM_MBL.frm:3501-5500`
- **Tables Read**: `AVT`, `INV`, `INL`, `PRD`, `RP2`
- **Tables Written**: `AVT`
- **Logic**:
  - **`CMDAUPDATE_Click`**: Bulk updates SIM activation records in the `AVT` table. It iterates through the activation grid (`MSA`), sets `AVT_STS = 'Yes'`, `AVT_DAT`, and `AVT_TIM`. It also allows swapping the Mobile ID if an alternate number was activated (`MBL_ODY`).
  - **`cmdCUpdate_Click`**: Processes Customer Agreement Forms (CAF). Updates `caf_STS` in `AVT` where a mobile number has been activated (`avt_sts='Yes'`) and not churned.
  - **`cmdModUpdate_Click`**: Handles Retail and Purchase returns. It sums `PRD_QTY` from `INV` to determine expected returns and validates this against selected grid rows. For retail returns, it nullifies the sale (`smn_idy=""`, `sal_idy=""`, `sal_typ=""`); for purchase returns, it physically deletes the record (`DELETE FROM AVT`).
  - **`cmdRUpdate_Click`**: Updates Churn status (`chr_STS`, `chr_dat`) for numbers that are activated.
  - **`cmdsUpdate_Click`**: Allocates mobile numbers to sales invoices. Enforces quantity limits based on `INV` records (`stk_typ IN ('Salable','Free')`). Updates `AVT` to link the SIM to `sal_idy`, `sal_dat`, `cst_idy`, and `smn_idy`.
  - **`Command1_Click`**: Generates serialized records in `AVT` based on Purchase transactions (`PUR`, `OPS`, `STI`, `SRN`). Creates N records for a given `prd_qty` matching `INV`, populating `inv_idy`, `prd_idy`, and vendor info.
  - **Excel Imports**: Various routines (`Command4_Click`, `Command9_Click`, `fileimpcmd_Click`) use `Excel.Application` to parse telecom spreadsheets for Churn, Activation, and Purchase files.

#### UI Grid and Selection Logic (Lines 5501 - 7500)
- **Source**: `M_FRM_MBL.frm:5501-7500`
- **Tables Read**: `AVT`, `INV`, `PRD`, `ACT`, `SMN`, `DIR`, `INL`
- **Logic**:
  - **Grid Navigation & Filtering (`LVP_KeyDown`)**: Extensive logic that intercepts the Enter key (`KeyCode = 13`) on the primary list view (`LVP`) to load secondary grid views based on the context of what the user is selecting (indicated by `Label49.Caption`).
  - **Adjust Returns Context**: Queries the `AVT` table grouping by `prd_idy` to compare the count of returned SIMs against the `prd_qty` from the `INV` (Invoice) table. If there is a mismatch, the system flags that products are not fully returned.
  - **Invoice Selection Context (Purchase/Sale)**: For incoming stock (`PUR`, `STI`, `SRN`, `OPS`), it verifies that the `mbl_idy` is pending (null or empty) for the invoice date, checking the supplier's product mapping (`ALT_IDY` in `PRD`). For outgoing sales (`SAL`), it checks `AVT` for mobiles mapped to the invoice and calculates remaining quantities (`prd_qty` in `INV` vs count in `AVT`).
  - **Activation, CAF, and Churn Lists**: Provides dynamic loading of mobile numbers into various `MSFlexGrid` controls (`MSA`, `MSC`, `MSR`) based on the current state. For example, for CAF processing, it filters `AVT` for `avt_sts='Yes'` and `caf_sts` is null/empty. This logic repeats for filtering by Invoice, Mobile Number, Salesman (`smn_idy`), and Party Account (`cst_idy`).
  - **N+1 Query Issue**: Similar to earlier sections, when populating these grids, the code performs individual `SELECT` queries for `prd_nme` and `act_nme` in a `While Not rsAvt.EOF` loop, which is a significant performance bottleneck in the legacy application.
  - **Mobile Import Tools**: Contains `salimpcmd_Click` and `sal_imp_Click`, which utilize `CommonDialog1` and `Excel.Application` to let the user import `.xls` or text files containing mobile numbers for batch operations.

#### User Input & Validation Logic (Lines 7501 - 11550)
- **Source**: `M_FRM_MBL.frm:7501-11550`
- **Tables Read**: `AVT`, `INV`, `INL`, `PRD`, `ACT`, `SMN`, `DIR`
- **Tables Written**: `AVT`
- **Logic**:
  - **Grid Population**: Repeated `While Not RS.EOF` loops fill `MSFlexGrid` and `ListView` controls with mobile numbers based on complex SQL queries checking `txn_typ` (e.g. `'SAL'`, `'STO'`, `'PRT'`, `'SRN'`), date ranges, and status fields (`avt_sts`, `caf_sts`, `chr_sts`).
  - **Mobile Number Allocation to Invoices (`txtSalInv_KeyDown`)**: Validates the invoice ID entered (`SAL...`) against the `INV` table. Checks that `inv.stk_typ IN ('Salable','Free')` and `prd.exc_inf='Y'`. It verifies the required quantity (`inv.prd_qty`) against the number of SIMs scanned/selected from the `AVT` table.
  - **Duplicate & Validation Checks (`txtSMbl_KeyDown`, `txtTemp_KeyDown`, `txtStart_KeyDown`)**: Enforces that scanned/typed mobile numbers are exactly 10 digits (`Len(Trim(txtTemp)) = 10`). Checks the `AVT` table to prevent adding duplicate mobile numbers (`chk_sts<>'Yes'`). Uses `list_chk` to prevent selecting more SIMs than invoiced for a specific product.
  - **Save Mobile Numbers (`updateavt`)**: When updating opening stock or purchases (`PUR`, `OPS`, `STI`, `SRN`), it loops through the `MS` grid, validates numbers are 10 digits, checks against `AVT` for duplicates, deletes any empty records for the current invoice (`inv_idy`), and inserts new rows into `AVT` using `.AddNew`.
  - **Sales Returns (`updateavt` - SRN section)**: Compares the total checked items in the returns list (`LVSRN`) against the `prd_qty` from the `INV` table. If the quantities match, it updates the `sal_typ` in `AVT` to `'SRN'` for the selected mobile numbers.
  - **Product Verification (`CHK_SAL_IMP`)**: A helper function to verify if the imported/selected mobile numbers match the expected product mapping based on the supplier's product ID (`ALT_IDY` in `PRD` table).
  - **Invoice Selection Logic (`invlist`, `invlist_RETURNS`)**: Loads a list view (`LVP`) with invoices that have missing or pending mobile number allocations (`mbl_idy is null or mbl_idy=''`). Resolves product and supplier names using sequential queries to `PRD` and `ACT`.
