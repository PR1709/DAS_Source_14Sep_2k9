select 0
use prd index prd alias prd
set order to 1
select 0
use brd alias BRD
go top
do while .not. eof("BRD")
   select PRD
   seek BRD.PRD_IDY
   if found()
      replace BRD.PRD_PDR WITH PRD.PUR_RAT, BRD.PRD_SDR with PRD.SAL_RAT, BRD.PRD_MRP with PRD.MRP_RAT
   endif
   select BRD
   skip 1
enddo
close all
return
