close all
clear all
select 0
if file('ifix.txt')
	delete file ifix.txt
endif
*
IF FILE('ifix.txt')  && Does file exist? 
   gnErrFile = FOPEN('ifix.txt',12)  && If so, open read-write
ELSE
   gnErrFile = FCREATE('ifix.txt')  && If not, create it
ENDIF
IF gnErrFile < 0  && Check for error opening file
   WAIT 'Cannot open or create output file' WINDOW NOWAIT
ELSE  && If no error, write to file
   *=FWRITE(gnErrFile, 'DB Integrity Fix Errors Log')
	=FPUTS(gnErrFile, 'DB Integrity Fix Errors Log')
ENDIF
*=FCLOSE(gnErrFile)  && Close file
*MODIFY FILE ifix.txt NOWAIT  && Open file in edit window 

use PRD alias PRD
set order to PRD_IDY
select 0
use BRD alias BRD
set order to BAT_IDY
select 0
use PGR alias PGR
set order to GRP_IDY
select 0
use INV alias INV
go top
select 0
use SCH alias SCH
set order to SCH_IDY
select 0
use LNV alias LNV
select 0
use PNV alias PNV
select 0
use VAN alias VAN
set order to INV_IDY
select 0
use CUR alias CUR
select 0
use SHL alias SHL
set order to SHL_IDY
select 0
use ACT alias ACT
set order to ACT_IDY
select 0
use INL alias INL
select 0
use PNL alias PNL
select 0
*
use LOC alias LOC
set order to LOC_IDY
select 0
use CAT alias CAT
set order to CAT_IDY
select 0
use MAR alias MAR
set order to MAR_IDY
select 0
use SMN alias SMN
set order to SMN_IDY
*
select 0
use VCH alias VCH
*
select 0
use SNV alias SNV
*
select BRD
go top
do while .not. eof("BRD")
	select PRD
	seek BRD->PRD_IDY
	if .not. found()
		*wait window "Product ID from Batch Table " + BRD->PRD_IDY + " is Not Found in Product Table "
	   =FPUTS(gnErrFile, 'Product ID from Batch Table ' + BRD->PRD_IDY + ' is not found in product table')
	endif
	select BRD
	skip 1
enddo
*
select PRD
go top
do while .not. eof("PRD")
	select PGR
	seek PRD->GRP_IDY
	if .not. found()
		*wait window "Product Group ID from Product Table " + PRD->GRP_IDY + " Not Found in Product Group Table "
		=FPUTS(gnErrFile, 'Product Group ID from Product Table ' + PRD->GRP_IDY + ' is not found in product group table')
	endif
	select PRD
	skip 1
enddo
*
select INV
go top
do while .not. eof("INV")
	select BRD
	seek INV->BAT_IDY
	if .not. found()
		*wait window "Batch ID from Inv Table " +INV->BAT_IDY + " Not found in Batch Table"
		=FPUTS(gnErrFile, 'Batch ID from Inv Table ' + INV->BAT_IDY + ' is not found in batch table')
	endif
	select PRD
	seek INV->PRD_IDY
	if .not. found()
		*wait window "Product ID from Inv Table " + INV->PRD_IDY + " of INV Not found in Product Table"
		=FPUTS(gnErrFile, 'Product ID from Inv Table ' + INV->PRD_IDY + ' is not found in product table')
	endif
	select INV
	skip 1
enddo
*
select LNV
go top
do while .not. eof("LNV")
	select SCH
	seek LNV->LEV_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Scheme ID from Lnv Table ' + LNV->LEV_IDY + ' is not found in scheme table')
	endif
	select LNV
	skip 1
enddo
*
select PNV
go top
do while .not. eof("PNV")
	select BRD
	seek PNV->BAT_IDY
	if .not. found()
		*wait window "Batch ID from PNV Table " +PNV->BAT_IDY + " Not found in Batch Table"
		=FPUTS(gnErrFile, 'Batch ID from PNV Table ' + PNV->BAT_IDY + ' is not found in batch table')
	endif
	select PRD
	seek PNV->PRD_IDY
	if .not. found()
		*wait window "Product ID from PNV Table " + PNV->PRD_IDY + " of PNV Not found in Product Table"
		=FPUTS(gnErrFile, 'Product ID from PNV Table ' + PNV->PRD_IDY + ' is not found in product table')
	endif
	select PNV
	skip 1
enddo
*
select VAN
go top
do while .not. eof("VAN")
	select BRD
	seek VAN->BAT_IDY
	if .not. found()
		*wait window "Batch ID from VAN Table " +VAN->BAT_IDY + " Not found in Batch Table"
		=FPUTS(gnErrFile, 'Batch ID from VAN Table ' + VAN->BAT_IDY + ' is not found in batch table')
	endif
	select PRD
	seek VAN->PRD_IDY
	if .not. found()
		*wait window "Product ID from VAN Table " + VAN->PRD_IDY + " of VAN Not found in Product Table"
		=FPUTS(gnErrFile, 'Product ID from VAN Table ' + VAN->PRD_IDY + ' is not found in product table')
	endif
	select VAN
	skip 1
enddo
*
select CUR
go top
do while .not. eof("CUR")
	select VAN
	seek CUR->INV_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Reference ID from CUR Table ' + CUR->INV_IDY + ' is not found in van table')
	endif
	select CUR
	skip 1
enddo
*
select SHL
go top
do while .not. eof("SHL")
	select PRD
	seek SHL->PRD_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Product ID from SHL (Shell) Table ' + SHL->PRD_IDY + ' is not found in product table')
	endif
	select SHL
	skip 1
enddo
*
select INL
go top
do while .not. eof("INL")
	select ACT
	seek INL->ACT_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Account ID from INL Table ' + INL->ACT_IDY + ' is not found in account table')	
	endif
	select ACT
	seek INL->CST_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Customer ID from INL Table ' + INL->CST_IDY + ' is not found in account table')
	endif
	select INL
	skip 1
enddo
*
select PNL
go top
do while .not. eof("PNL")
	select ACT
	seek PNL->ACT_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Account ID from PNL Table ' + PNL->ACT_IDY + ' is not found in account table')	
	endif
	select ACT
	seek PNL->CST_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Customer ID from PNL Table ' + PNL->CST_IDY + ' is not found in account table')
	endif
	select PNL
	skip 1
enddo
*
select PNL
go top
do while .not. eof("PNL")
	select LOC
	seek PNL->LOC_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Location ID from PNL Table ' + PNL->LOC_IDY + ' is not found in locations table')	
	endif
	select PNL
	skip 1
enddo
select PNL
go top
do while .not. eof ("PNL")
	select CAT
	seek PNL->CAT_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Category/Channel ID from PNL Table ' + PNL->CAT_IDY + ' is not found in category/channel table')	
	endif
	select PNL
	skip 1
enddo
select PNL
go top
do while .not. eof("PNL")
	select MAR
	seek PNL->MAR_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Market ID from PNL Table ' + PNL->MAR_IDY + ' is not found in market table')	
	endif
	select PNL
	skip 1
enddo
select PNL
go top
do while .not. eof("PNL")
	select SMN
	seek PNL->SMN_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Salesman ID from PNL Table ' + PNL->ACT_IDY + ' is not found in salesman table')	
	endif
	select PNL
	skip 1
enddo
*
select VCH
go top
do while .not. eof("VCH")
	select ACT
	seek VCH->DBT_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Debit ID from VCH Table ' + VCH->DBT_IDY + ' is not found in voucher table')
	endif
	select ACT
	seek VCH->CRT_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Credit ID from VCH Table ' + VCH->CRT_IDY + ' is not found in voucher table')
	endif
	select VCH
	skip 1
enddo
*
select SNV
go top
do while .not. eof("SNV")
	select SHL
	seek SNV->SHL_IDY
	if .not. found()
		=FPUTS(gnErrFile, 'Shell ID from SNV Table ' + SNV->SHL_IDY + ' is not found in Shell table')
	endif
	if SNV->TXN_TYP = "PUR" .or. SNV->TXN_TYP = "SAL"
		select PRD
		seek SNV->PRD_IDY
		if .not. found()
			=FPUTS(gnErrFile, 'Product ID from SNV Table ' + SNV->PRD_IDY + ' is not found in Product table')
		endif
	endif
	select SNV
	skip 1
enddo
*
=FCLOSE(gnErrFile)  && Close file
close all
clear all
MODIFY FILE ifix.txt NOWAIT  && Open file in edit window
