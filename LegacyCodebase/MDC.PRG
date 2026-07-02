* do IND
close all
clear all
* Master Detail Check and Write Utility
select 0
use INV index INV alias INV
set order to 1
select 0
use PNL index PNL alias PNL
set order to 1
select 0
use INL index INL alias INL
set order to 1
* : First Part 
select INV
go top 
do while .not. eof("INV")
   select INL
   seek INV.INV_IDY
   if .not. found()
      select INL
      append blank
      replace INL.INV_IDY with INV.INV_IDY, INL.TXN_TYP with INV.TXN_TYP, INL.INV_DAT with INV.INV_DAT,;
      		  INL.DUE_DAT with INV.INV_DAT, INL.WBL_DAT with INV.INV_DAT, INL.LRN_DAT with INV.INV_DAT,;
      		  INL.DCH_DAT with INV.INV_DAT, INL.VCH_DAT with INV.INV_DAT, INL.SYS_DAT with INV.INV_DAT
   endif
   select INV
   skip 1
enddo
* : Second Part
select INL
go top
do while .not. eof("INL")
   select PNL
   seek INL.INV_IDY
   if found()
      if len(alltrim(INL.CST_IDY)) = 0
          replace INL.CST_IDY with PNL.CST_IDY
      endif
   else
      if len(alltrim(INL.CST_IDY)) = 0
          replace INL.CST_IDY with "0000000001"
      endif   
   endif
   select INL
   skip 1
enddo
* : Third Part
select INL
go top
do while .not. eof("INL")
   select INV
   seek INL.INV_IDY
   if .not. found()
       select INV
       append blank
       replace INV.INV_IDY with INL.INV_IDY, INV.TXN_TYP with INL.TXN_TYP, INV.INV_DAT with INL.INV_DAT
   else
       if INV.INV_DAT = ctod("  /  /    ")
          replace INV.INV_DAT with INL.INV_DAT
       endif
   endif
   select INL
   skip 1
enddo
* :
close all
clear all
* do IND
return

