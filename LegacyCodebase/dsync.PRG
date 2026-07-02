close all
clear all
set date british
clear
wait window "Enter Date to Synchronize Records On or After (>=) this Date..., Press any key"
store date() to CHK_DAT
@05,05 get CHK_DAT
read
wait window "Please Ensure Backup is taken before you run this utility.., Press <Esc> to Quit, Press <Enter> to proceed..."
if lastkey() = 27
	close all
	clear all
	return
endif
close all

select 0
use INL alias INL
INL_COUNT = reccount()
select 0
use PNL alias PNL
PNL_COUNT = reccount()
if INL_COUNT > PNL_COUNT
	select INL
	set filter to INL.INV_DAT >=  CHK_DAT
	go top	
	copy to OINL type fox2x
	select PNL
	zap
	append from OINL
endif
if PNL_COUNT > INL_COUNT
	select PNL
	set filter to PNL.INV_DAT >=  CHK_DAT
	go top	
	copy to OPNL type fox2x
	select INL
	zap
	append from OPNL
endif
close all
select 0
use INV alias INV
INV_COUNT = reccount()
select 0
use PNV alias PNV
PNV_COUNT = reccount()
if INV_COUNT > PNV_COUNT
	select INV
	set filter to INV.INV_DAT >=  CHK_DAT
	go top		
	copy to OINV type fox2x
	select PNV
	zap
	append from OINV
endif
if PNV_COUNT > INV_COUNT
	select PNV
	set filter to PNV.INV_DAT >=  CHK_DAT
	go top		
	copy to OPNV type fox2x
	select INV
	zap
	append from OPNV
	GO TOP
	REPLACE ALL ISS_QTY WITH .NULL.
endif
close all
clear all
wait window "Process Complete, Proceeding to Index Files... Takes Time, Press any key"
do IND
do CND
return

