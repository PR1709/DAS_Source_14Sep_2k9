close all
clear all
clear
if file("c:\tmp.xls") = .t.
	delete file c:\tmp.xls
endif
		

set oleobject on
oleapp = createobject("Excel.Application")
oleapp.visible = .t.
oleapp.workbooks.add
oleapp.cells(1,1).value = "Account ID"
oleapp.cells(1,2).value = "Name"
oleapp.cells(1,3).value = "Address"
select 0
use ACT index ACT alias ACT
set order to ACT_IDY
select ACT
go top
store 2 to LIN_CTR
do while .not. eof("ACT")
	if ACT.ACT_TYP = "SAL"
		oleapp.cells(LIN_CTR,1).value = ACT.ACT_IDY
		oleapp.cells(LIN_CTR,2).value = ACT.ACT_NME
		oleapp.cells(LIN_CTR,3).value = ACT.ACT_AD1
		oleapp.cells(LIN_CTR,4).value = ACT.ACT_AD2
		oleapp.cells(LIN_CTR,5).value = ACT.ACT_AD3
		oleapp.cells(LIN_CTR,6).value = ACT.ACT_AD4
		LIN_CTR = (LIN_CTR + 1)
	endif
	skip 1
enddo
oleapp.activeworkbook.saveas("c:\tmp.xls")
oleapp.quit
close all
return
