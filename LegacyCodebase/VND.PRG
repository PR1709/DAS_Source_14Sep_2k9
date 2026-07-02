* :: Effmcg Ver.4.0
* :: Date: 09.01.2003
* :: P N Ramaneshwar
* :: Module to index all tables (CDX) - Free Tables - Effmcg Ver.4.0
* :: vnd.prg
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
*set deleted on
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
close all
clear all
select 0
	 use VAN exclusive
     index on INV_IDY tag INV_IDY of VAN
     index on dtos(INV_DAT) tag INV_DAT of VAN
     index on PRD_IDY tag PRD_IDY of VAN
     index on BAT_IDY tag BAT_IDY of VAN
     index on ACT_IDY tag ACT_IDY of VAN
     index on CST_IDY tag CST_IDY of VAN
     index on SMN_IDY tag SMN_IDY of VAN
     index on BIL_IDY tag BIL_IDY of VAN
     pack
     use
* ::
* ::
close all
clear all
return
set exclusive on
set unique on
set exact on
clear
return
* :
* : End of vnd.prg
