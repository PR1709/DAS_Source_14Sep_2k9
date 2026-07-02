
clear
close all
clear all
set talk off
set echo off
set status off
set safety off

use PNV
zap
use PNL
zap
close all
clear all

select 0
use INV alias INV
select 0
use INL alias INL
select 0
use PNV alias PNV
select 0
use PNL alias PNL
select 0
use DIR index DIR alias DIR
set order to 1
select 0
use ACT index ACT alias ACT
set order to 1
select 0
use PGR index PGR alias PGR
set order to 1
select 0
use PRD index PRD alias PRD
set order to 1
select 0 
use BRD index BRD alias BRD
set order to 2
select 0
use MAR index MAR alias MAR
set order to 1
select 0
use SMN index SMN alias SMN
set order to 1
select 0
use CAT index CAT alias CAT
set order to 1
select 0
use LOC index LOC alias LOC
set order to 1

select INV
GO TOP
DO WHILE .NOT. EOF("INV")
if len(trim(INV.PRD_IDY)) > 0 .and. len(trim(INV.BAT_IDY)) > 0
	select PRD
	seek INV.PRD_IDY
	select PGR
	seek PRD.GRP_IDY
	select BRD
	seek INV.BAT_IDY
	
	select PNV
	append blank
	if isnull(INV.LST_RAT)
		replace PNV.LST_RAT with 0
	else
		replace PNV.LST_RAT with INV.LST_RAT
	endif
	if isnull(INV.LST_DIS)
		replace PNV.LST_DIS with 0
	else
		replace PNV.LST_DIS with INV.LST_DIS
	endif

	* PNV.LST_RAT with INV.LST_RAT, PNV.LST_DIS with INV.LST_DIS, 
	replace PNV.INV_IDY with INV.INV_IDY, PNV.INV_DAT with INV.INV_DAT, PNV.PRD_IDY with INV.PRD_IDY,;
			PNV.PRD_QTY with INV.PRD_QTY
	replace PNV.PRD_NME with PRD.PRD_NME, PNV.GRP_IDY with PRD.GRP_IDY, PNV.GRP_CLS with PGR.GRP_CLS,;
			PNV.GRP_CAT with PGR.GRP_CAT
	if isnull(INV.RET_QTY)
		replace PNV.RET_QTY with 0
	else
		replace PNV.RET_QTY with INV.RET_QTY
	endif
	if isnull(INV.RAT_APP)
		replace PNV.RAT_APP with space(01)
	else
		replace PNV.RAT_APP with INV.RAT_APP
	endif
	if isnull(INV.PRD_RAT)
		replace PNV.PRD_RAT with 0
	else
		replace PNV.PRD_RAT with INV.PRD_RAT
	endif 
	if isnull(INV.CPN_RAT)
		replace PNV.CPN_RAT with 0
	else
		replace PNV.CPN_RAT with INV.CPN_RAT
	endif
	if isnull(INV.CPN_AMT)
		replace PNV.CPN_AMT with 0
	else
		replace PNV.CPN_AMT with INV.CPN_AMT
	endif
	if isnull(INV.ADD_DIS)
		replace PNV.ADD_DIS with 0
	else
		replace PNV.ADD_DIS with INV.ADD_DIS
	endif
	if isnull(INV.ADD_AMT)
		replace PNV.ADD_AMT with 0
	else
		replace PNV.ADD_AMT with INV.ADD_AMT
	endif
	if isnull(INV.TRD_DIS)
		replace PNV.TRD_DIS with 0
	else
		replace PNV.TRD_DIS with INV.TRD_DIS
	endif
	if isnull(INV.DIS_AMT)
		replace PNV.DIS_AMT with 0
	else
		replace PNV.DIS_AMT with INV.DIS_AMT
	endif
	if isnull(INV.TRD_TAX)
		replace PNV.TRD_TAX with 0
	else
		replace PNV.TRD_TAX with INV.TRD_TAX
	endif
	if isnull(INV.TAX_AMT)
		replace PNV.TAX_AMT with 0
	else
		replace PNV.TAX_AMT with INV.TAX_AMT
	endif		
	replace PNV.COL_AMT with INV.COL_AMT, PNV.STK_UPD with INV.STK_UPD, PNV.TXN_TYP with INV.TXN_TYP
	if isnull(INV.FST_FLG) 
		replace PNV.FST_FLG with space(01)
	else
		replace PNV.FST_FLG with INV.FST_FLG
	endif
    If IsNull(INV.OPN_QTY)
        replace PNV.OPN_QTY with 0
    Else
        replace PNV.OPN_QTY with INV.OPN_QTY
    endif
    If IsNull(INV.ISS_QTY)
        replace PNV.ISS_QTY with 0
    Else
        replace PNV.ISS_QTY with INV.ISS_QTY
    endif
    if IsNull(INV.TOT_QTY)
        replace PNV.TOT_QTY with 0
    Else
        replace PNV.TOT_QTY with INV.TOT_QTY
    endif
    If IsNull(INV.SCH_AMT)
        replace PNV.SCH_AMT with 0 
    Else
        replace PNV.SCH_AMT with INV.SCH_AMT
    endif
    replace PNV.ORD_IDY with INV.ORD_IDY, PNV.BAT_IDY with INV.BAT_IDY, PNV.BAT_NME with BRD.BAT_NME,;
    		PNV.PRD_UOM with INV.PRD_UOM, PNV.SUB_UNT with INV.SUB_UNT, PNV.SUB_UNT with INV.SUB_UNT,;
    		PNV.PRD_UNT with INV.PRD_UNT, PNV.PRD_PCS with INV.PRD_PCS, PNV.TOT_PCS with INV.TOT_PCS,;
    		PNV.STK_TYP with INV.STK_TYP
    If IsNull(INV.SCH_IDY) 
        replace PNV.SCH_IDY with space(01)
    else
        replace PNV.SCH_IDY with INV.SCH_IDY
    endif
    replace PNV.MRP_RAT with BRD.PRD_MRP 
    If IsNull(INV.INC_TTL)
        replace PNV.INC_TTL with 0
    Else
        replace PNV.INC_TTL with INV.INC_TTL
    endif
    replace PNV.MAR_PGR with space(01)
endif
select INV
skip 1
enddo    

select INL
go top
do while .not. eof("INL")
if len(Trim(INL.ACT_IDY)) > 0 and len(trim(INL.CST_IDY)) > 0
	select PNL
	append blank

	select ACT
	seek INL.CST_IDY

	replace PNL.INV_IDY with INL.INV_IDY, PNL.INV_DAT with INL.INV_DAT, PNL.ORD_IDY with INL.ORD_IDY,;
			PNL.CST_IDY with INL.CST_IDY, PNL.CST_NME with ACT.ACT_NME, PNL.CST_AD1 with ACT.ACT_AD1,;
			PNL.CST_AD2 with ACT.ACT_AD2, PNL.CST_AD3 with ACT.ACT_AD3, PNL.CST_AD4 with ACT.ACT_AD4,;
			PNL.CST_CST with ACT.ACT_CST, PNL.CST_GST with ACT.ACT_GST, PNL.CST_ZIP with ACT.ACT_ZIP
			if len(trim(ACT.MAR_IDY)) > 0
				select MAR
				seek ACT.MAR_IDY
				replace PNL.MAR_IDY with ACT.MAR_IDY, PNL.MAR_NME with MAR.MAR_NME
			else
				replace PNL.MAR_IDY with space(10), PNL.MAR_NME with space(10)
			endif
			if len(trim(ACT.SMN_IDY)) > 0
				select SMN
				seek ACT.SMN_IDY
				replace PNL.SMN_IDY with ACT.SMN_IDY, PNL.SMN_NME with SMN.SMN_NME
			else
				replace PNL.SMN_IDY with space(10), PNL.SMN_NME with space(10)
			endif
			if len(trim(ACT.CAT_IDY)) > 0
				select CAT
				seek ACT.CAT_IDY
				replace PNL.CAT_IDY with ACT.CAT_IDY, PNL.CAT_NME WITH CAT.CAT_NME
			else
				replace PNL.CAT_IDY with space(10), PNL.CAT_NME with space(10)
			endif
			if len(trim(LOC.LOC_IDY)) > 0
				select LOC
				seek ACT.LOC_IDY
				replace PNL.LOC_IDY with ACT.LOC_IDY, PNL.LOC_NME with LOC.LOC_NME
			else
				replace PNL.LOC_IDY with space(10), PNL.LOC_NME with space(10)
			endif
			if isnull(INL.ROU_IDY)
				replace PNL.ROU_IDY with space(01)
			else
				replace PNL.ROU_IDY with INL.ROU_IDY
			endif
			if isnull(INL.GRS_AMT)
				replace PNL.GRS_AMT with 0
			else
				replace PNL.GRS_AMT with INL.GRS_AMT
			endif
			* PNL.GRS_AMT with INL.GRS_AMT, 
			if isnull(INL.SPD_PCG)
				replace PNL.SPD_PCG with 0
			else
				replace PNL.SPD_PCG with INL.SPD_PCG
			endif
			if isnull(INL.CSD_PCG)
				replace PNL.CSD_PCG with 0
			else
				replace PNL.CSD_PCG with INL.CSD_PCG
			endif
			if isnull(INL.CSD_AMT)
				replace PNL.CSD_AMT with 0
			else
				replace PNL.CSD_AMT with INL.CSD_AMT
			endif
			if isnull(INL.EXD_PCG)
				replace PNL.EXD_PCG with 0
			else
				replace PNL.EXD_PCG with INL.EXD_PCG
			endif		    
	        If IsNull(INL.EXD_CAL) 
		        replace PNL.EXD_CAL with space(01)
		    else
		    	replace PNL.EXD_CAL with INL.EXD_CAL
		    endif 
	    	replace PNL.EXD_AMT with INL.EXD_AMT, PNL.TAX_PCG with INL.TAX_PCG, PNL.TAX_AMT with INL.TAX_AMT,;
	    			PNL.NET_AMT with INL.NET_AMT, PNL.TXN_TYP with INL.TXN_TYP
	 	    If IsNull(INL.FRT_TYP) 
	        	replace PNL.FRT_TYP with space(01)
	        else
	        	replace PNL.FRT_TYP with INL.FRT_TYP
	        endif
	     	replace PNL.FRT_AMT with INL.FRT_AMT, PNL.DCH_IDY with INL.DCH_IDY
	 	    If IsNull(INL.DCH_DAT) 
	 	       replace PNL.DCH_DAT with INL.INV_DAT  
		    else
	        	replace PNL.DCH_DAT with INL.DCH_DAT
	        endif
	    	replace PNL.LRN_IDY with INL.LRN_IDY, PNL.LRN_DAT with INL.LRN_DAT, PNL.LRN_AMT with INL.LRN_AMT,;
	    			PNL.WBL_IDY with INL.WBL_IDY, PNL.WBL_DAT with INL.WBL_DAT, PNL.PKG_CST with INL.PKG_CST,;
	    			PNL.TTL_CAS with INL.TTL_CAS, PNL.TXN_RMK with INL.TXN_RMK, PNL.TXN_TPT with INL.TXN_TYP,;
	    			PNL.ROU_OFF with INL.ROU_OFF
	     	If IsNull(INL.TAX_ACT) 
	     	    replace PNL.TAX_ACT with space(01)
	     	else
	     		replace PNL.TAX_ACT with INL.TAX_ACT
	     	endif
	        If IsNull(INL.ROU_ACT)
	        	replace PNL.ROU_ACT with space(01)
	        else
	        	replace PNL.ROU_ACT with INL.ROU_ACT
	        endif
	        if IsNull(INL.FRT_ACT)
	        	replace PNL.FRT_ACT with space(01)
	        else
	        	replace PNL.FRT_ACT with INL.FRT_ACT
	        endif
	        replace PNL.CRT_DAY with INL.CRT_DAY, PNL.DUE_DAT with INL.DUE_DAT
	        
	    	If IsNull(INL.DES_MOD)
	        	replace PNL.DES_MOD with space(01)
	        else
	        	replace PNL.DES_MOD with INL.DES_MOD
	        endif
	 	    If IsNull(INL.STO_RMK)
	        	replace PNL.STO_RMK with space(01)
	        else
	        	replace PNL.STO_RMK with INL.STO_RMK
	        endif
	        replace PNL.ROU_IDY with INL.ROU_IDY, PNL.TXN_RM1 with trim(INL.TXN_RM1), PNL.TXN_RM2 with trim(INL.TXN_RM2),;
	        		PNL.VCH_IDY with INL.VCH_IDY
		    If IsNull(INL.VCH_DAT)
				replace PNL.VCH_DAT with INL.INV_DAT
			else
				replace PNL.VCH_DAT with INL.VCH_DAT
			endif
			replace PNL.SCH_CHK with INL.SCH_CHK, PNL.SRN_AMT with INL.SRN_AMT, PNL.SCH_IDY with INL.SCH_IDY,;
					PNL.USR_NME with INL.USR_IDY, PNL.SYS_DAT with INL.SYS_DAT, PNL.SYS_TIM with INL.SYS_TIM
	 		replace PNL.FRM_DAT with INL.INV_DAT, PNL.TO_DAT with INL.INV_DAT, PNL.ACT_IDY with INL.ACT_IDY,;
	 				PNL.MAR_DOP with 0
	 		
		    If IsNull(INL.DIS_CHK)
		        replace PNL.DIS_CHK with space(01) 
		    else 
	        	replace PNL.DIS_CHK with INL.DIS_CHK 
		    endif
		
	    	select DIR
			seek "SYS"+"TT1"
			replace PNL.INS_NME with DIR.DIR_MSG
			seek "SYS"+"TT2"
			replace PNL.INS_ADR with DIR.DIR_MSG
			seek "SYS"+"RM1"
			replace PNL.SYS_RM1 with DIR.DIR_MSG
			seek "SYS"+"RM2"
			replace PNL.SYS_RM2 with DIR.DIR_MSG
			seek "SYS"+"RM3"
			replace PNL.SYS_RM3 with DIR.DIR_MSG
			seek "SYS"+"RM4"
			replace PNL.SYS_RM4 with DIR.DIR_MSG
			seek "SYS"+"GST"
			replace PNL.INS_GST with DIR.DIR_MSG
			seek "SYS"+"CST"
			replace PNL.INS_CST with DIR.DIR_MSG
else
	select PNL
	append blank
	replace PNL.INV_IDY with INL.INV_IDY, PNL.INV_DAT with INL.INV_DAT, PNL.TXN_TYP with INL.TXN_TYP
endif
 	select INL
 	skip 1
enddo
close all
clear all
return
