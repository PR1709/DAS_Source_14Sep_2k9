CLOSE ALL
CLEAR ALL
set safety off
wait window "This Utility Clears any Invalid Records in the DataBase Due to Power Failure or Improper ShutDowns! Press any Key to Proceed"
use PNV
delete all for year(INV_DAT) = 0
use PNL
delete all for year(INV_DAT) = 0
use INV
delete all for year(INV_DAT) = 0
use INL
delete all for year(INV_DAT) = 0
use IRT
delete all for year(INV_DAT) = 0
use VAN
delete all for year(INV_DAT) = 0
use SNV
delete all for year(INV_DAT) = 0
use VCH
delete all for year(TXN_DAT) = 0
close all
clear all
select 0
use INL index INL alias INL
* set filter to TXN_TYP = "SAL"
set order to 1
go top
select 0
use INV index INV alias INV
* set filter to TXN_TYP = "SAL"
set order to 1
go top
do while .not. eof("INV")
   select INL
   seek INV.INV_IDY
   if .not. found()
   		select INV
		delete
   	endif
   	select INV
   	skip 1
enddo
close all
clear all   
return
