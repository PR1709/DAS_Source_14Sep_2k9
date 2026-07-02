* :: Effmcg Ver.4.0
* :: Date: 09.01.2003
* :: P N Ramaneshwar
* :: Module to index all tables (CDX) - Free Tables - Effmcg Ver.4.0
* :: Ind.prg
* ::
set escape on                  		&& Environment Settings
set ansi on
set autosave on
set bell off
set border to
set carry off
set century on
set color to
set collate to "MACHINE"
set compatible on
set compatible FOXPLUS
set cpcompile to 1252
set cpdialog off
set confirm off
set currency to "Rs."
set currency right
set cursor on
set date british
set decimals to 2
set debug off
set deleted on
set delimiters off
set development on
set device to screen
set display to VGA50
set echo off
set escape off
set exact off
set exclusive on
set fixed off
set hours to 12
set intensity on
set keycomp to DOS
set notify off
set optimize on
set point to "."
set precision to
set reprocess to automatic
set safety off
set scoreboard off
set separator to ","
set status on
set status bar on
set step off
set title on
set talk off
set udfparms to reference
set unique off
set view off
* ::
     use ACT exclusive					&& Accounts DBF
     index on ACT_IDY tag ACT_IDY of ACT
     index on ACT_NME tag ACT_NME of ACT
     pack
     use
* :
     use AGD exclusive					&& Ageing Temporary Database
     index on DUE_DAT tag DUE_DAT of AGD
     pack
     use
* :
     use BRD exclusive                                  && Batches DBF
     index on PRD_IDY tag PRD_IDY of BRD
     index on BAT_IDY tag BAT_IDY of BRD
     pack
     use
* :
     use CAT exclusive					&& Categories DBF
     index on CAT_IDY tag CAT_IDY of CAT
     index on CAT_NME tag CAT_NME of CAT
     pack
     use
* :
     use CLD exclusive					&& Stocks Consolidated Output Database
     index on PRD_IDY tag PRD_IDY of CLD
     pack
     use
* :
     use DIR exclusive					&& Directory DBF
     index on DIR_IDY+DIR_IDX tag DIR_IDY of DIR
     use
* : 
     use GRP exclusive					&& Groups DBF
     index on GRP_IDY tag GRP_IDY of GRP
     index on GRP_NME tag GRP_NME of GRP
     pack
     use
* :
     use INL exclusive
     index on INV_IDY tag INV_IDY of INL
     index on dtos(INV_DAT) tag INV_DAT of INL
     index on CST_IDY tag CST_IDY of INL
     index on ORD_IDY tag ORD_IDY of INL
     index on VCH_IDY tag VCH_IDY of INL
     index on SCH_IDY tag SCH_IDY of INL
     index on DCH_IDY tag DCH_IDY of INL
     pack
     use
* :
     use INV exclusive
     index on INV_IDY tag INV_IDY of INV
     index on dtos(INV_DAT) tag INV_DAT of INV
     index on PRD_IDY tag PRD_IDY of INV
     index on BAT_IDY tag BAT_IDY of INV
     index on SCH_IDY tag SCH_IDY of INV
     index on ORD_IDY tag ORD_IDY of INV
     pack
     use
* :
     use IRT exclusive					&& Invoices Vs.Receipts Matching DBF
     index on INV_IDY tag INV_IDY of IRT
     index on ACT_IDY tag ACT_IDY of IRT
     index on PTY_IDY tag PTY_IDY of IRT
     index on SMN_IDY tag SMN_IDY of IRT
     pack
     use
* :
     use LDS exclusive					&& Loading Sheet DBF
     index on LDS_IDY tag LDS_IDY of LDS
     index on dtos(LDS_DAT) tag LDS_DAT of LDS
     pack
     use
* :
     use LED exclusive					&& Ledgers Output Database
     index on dtos(INV_DAT) tag INV_DAT of LED
     pack
     use
* :
     use LEV exclusive					&& Discounts & Taxes DBF
     index on LEV_IDY tag LEV_IDY of LEV
     index on BAS_IDY tag BAS_IDY of LEV
     pack
     use
* :
     use LNV exclusive					&& Discounts & Taxes Detail DBF for INL
     index on INV_IDY tag INV_IDY of LNV
     index on LEV_IDY tag LEV_IDY of LNV
     pack
     use
* :
     use LOC exclusive					&& Locations DBF
     index on LOC_IDY tag LOC_IDY of LOC
     index on LOC_NME tag LOC_NME of LOC
     pack
     use
* :
     use MAR exclusive					&& Markets DBF
     index on MAR_IDY tag MAR_IDY of MAR
     index on MAR_NME tag MAR_NME of MAR
     pack
     use
* :
     use MBR exclusive					&& Monthly Business Record Database
     index on str(MTH_IDY,2) tag MTH_IDY of MBR
     pack
     use
* :
     use MCL exclusive
     index on INV_IDY tag INV_IDY of MCL
     index on dtos(INV_DAT) tag INV_DAT of MCL
     index on CST_IDY tag CST_IDY of MCL
     pack
     use
* :
     use OPA exclusive					&& Opening Accounts DBF
     index on ACT_IDY tag ACT_IDY of OPA
     pack
     use
* :
     use OPS exclusive					&& Opening Stocks DBF
     index on PRD_IDY tag PRD_IDY of OPS
     pack
     use
* :
     use ORD exclusive					&& Orders DBF
     index on ORD_IDY tag ORD_IDY of ORD
     index on dtos(ORD_DAT) tag ORD_DAT of ORD
     index on PRD_IDY tag PRD_IDY of ORD
     pack
     use
* :
     use PGR exclusive					&& Product Group DBF
     index on GRP_IDY tag GRP_IDY of PGR
     index on GRP_CAT tag GRP_CAT of PGR
     pack
     use
* :
     use PNL exclusive
     index on INV_IDY tag INV_IDY of PNL
     index on dtos(INV_DAT) tag INV_DAT of PNL
     index on CST_IDY tag CST_IDY of PNL
     index on ORD_IDY tag ORD_IDY of PNL
     index on VCH_IDY tag VCH_IDY of PNL
     index on SCH_IDY tag SCH_IDY of PNL
     index on DCH_IDY tag DCH_IDY of PNL
     pack
     use
* :
     use PNV exclusive
     index on INV_IDY tag INV_IDY of PNV
     index on dtos(INV_DAT) tag INV_DAT of PNV
     pack
     use
* :
     use PRD exclusive					&& Products DBF
     index on PRD_IDY tag PRD_IDY of PRD
     index on PRD_NME tag PRD_NME of PRD
     index on PRD_PRN tag PRD_PRN of PRD
     pack
     use
* :
     use SCH exclusive 					&& Schemes DBF
     index on SCH_IDY tag SCH_IDY of SCH
     index on BAT_IDY tag BAT_IDY of SCH
     pack
     use
* :
     use SMN exclusive					&& Salesman DBF
     index on SMN_IDY tag SMN_IDY of SMN
     index on SMN_NME tag SMN_NME of SMN
     pack
     use
* :
     use STA exclusive					&& Stocks Ageing Temporary Database
     index on dtos(TXN_DAT) tag TXN_DAT of STA
     pack
     use
* :
     use STL exclusive					&& Stocks Ledger Output Database
     index on dtos(INV_DAT) tag INV_DAT of STL
     index on PRD_IDY tag PRD_IDY of STL
     pack
     use
* :
     use TPD exclusive					&& Temporary Product Database (for Stock Movement Calculation)
     index on PRD_IDY tag PRD_IDY of TPD
     pack
     use
* :
     use TYP exclusive					&& Account Types DBF
     index on TYP_DEF tag TYP_DEF of TYP
     use
* :
     use VAN exclusive
     index on INV_IDY tag INV_IDY of VAN
     index on dtos(INV_DAT) tag INV_DAT of VAN
     index on PRD_IDY tag PRD_IDY of VAN
     index on BAT_IDY tag BAT_IDY of VAN
     use
* :
     use VCH exclusive					&& Vouchers Database
     index on TXN_IDY tag TXN_IDY of VCH
     index on dtos(TXN_DAT) tag TXN_DAT of VCH
     index on INV_IDY tag INV_IDY of VCH
     pack
     use
* :
     use YER exclusive					&& Years DBF
     index on YER_IDY tag YER_IDY of YER
     use
* ::
close all
clear all
set exclusive on
set unique on
set exact on
clear
return
* :
* : End of ind.prg
