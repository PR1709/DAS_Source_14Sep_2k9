close all
clear all
select 0
use INV alias INV
go top
do while .not. eof("INV")
	if TXN_TYP = 'SAL'
		M_UNT_CST = (INV->PRD_RAT / INV->SUB_UNT)
		M_UNT_TAX = ((M_UNT_CST * INV->TRD_TAX) / 100)
		M_TAX_AMT = (INV->TOT_PCS * M_UNT_TAX)
		replace INV->TAX_AMT with M_TAX_AMT
	endif
	select INV
	skip 1
enddo
close all
clear all
return
