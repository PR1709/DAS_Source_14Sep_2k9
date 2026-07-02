close all
clear all
clear
store 0 to R_INL, R_PNL
select 0
use INL alias INL
select INL
R_INL = reccount()
select 0
use PNL alias PNL
select PNL
R_PNL = reccount()
*
? r_inl
? r_pnl
if R_INL > R_PNL
	select PNL
	delete all
	select INL
	go top
	do while .NOT. eof("INL")
		select PNL
		append blank
		replace PNL.INV_IDY with INL.INV_IDY,;
				PNL.INV_DAT with INL.INV_DAT,;
				PNL.ORD_IDY with INL.ORD_IDY,;
				PNL.CST_IDY with INL.CST_IDY,;
				PNL.ACT_IDY with INL.ACT_IDY,;
				PNL.GRS_AMT with INL.GRS_AMT,;
				PNL.SPD_PCG with INL.SPD_PCG,;
				PNL.SPD_AMT with INL.SPD_AMT,;
				PNL.CSD_PCG with INL.CSD_PCG,;
	            PNL.CSD_AMT with INL.CSD_AMT,;
	            PNL.EXD_PCG with INL.EXD_PCG,;
	            PNL.EXD_AMT with INL.EXD_AMT,;
	            PNL.TAX_PCG with INL.TAX_PCG,;
	            PNL.TAX_AMT with INL.TAX_AMT,;
	            PNL.NET_AMT with INL.NET_AMT,;
	            PNL.TXN_TYP with INL.TXN_TYP,;
	            PNL.FRT_AMT with INL.FRT_AMT,;
	            PNL.DCH_IDY with INL.DCH_IDY,;
	            PNL.LRN_IDY with INL.LRN_IDY,;			
	            PNL.LRN_DAT with INL.LRN_DAT,;
	            PNL.LRN_AMT with INL.LRN_AMT,;
	            PNL.WBL_IDY with INL.WBL_IDY,;
	            PNL.WBL_DAT with INL.WBL_DAT,;
	            PNL.PKG_CST with INL.PKG_CST,;
	            PNL.TTL_CAS with INL.TTL_CAS,;
	            PNL.TXN_RMK with INL.TXN_RMK,;
	            PNL.TXN_TPT with INL.TXN_TPT,;
	            PNL.ROU_OFF with INL.ROU_OFF,;
	            PNL.CRT_DAY with INL.CRT_DAY,;
	            PNL.DUE_DAT with INL.DUE_DAT,;
	            PNL.TXN_RM1 with INL.TXN_RM1,;
	            PNL.TXN_RM2 with INL.TXN_RM2,;
	            PNL.VCH_IDY with INL.VCH_IDY,;
	            PNL.SCH_CHK with INL.SCH_CHK,;
	            PNL.SYS_DAT with INL.SYS_DAT,;
	            PNL.SYS_TIM with INL.SYS_TIM,;
	            PNL.SCH_IDY with INL.SCH_IDY,;
	            PNL.SRN_AMT with INL.SRN_AMT

	            
*	            PNL.DCH_DAT with INL.DCH_DAT,;
*	            PNL.TAX_ACT with INL.TAX_ACT,;
*	            PNL.ROU_ACT with INL.ROU_ACT,;
*	            PNL.FRT_ACT with INL.FRT_ACT,;
*	            PNL.DES_MOD with INL.DES_MOD,;
*	            PNL.STO_RMK with INL.STO_RMK,;
*	            PNL.VCH_DAT with INL.VCH_DAT,;
*	            PNL.DIS_CHK with INL.DIS_CHK
	            	            	            	            	            	            	            	          	
	          	if ISNULL(INL.FRT_TYP)
	          		replace PNL.FRT_TYP with space(10)
	          	else
					replace PNL.FRT_TYP with PNL.FRT_TYP
		        endif

	          	if ISNULL(INL.SMN_IDY)
	          		replace PNL.SMN_IDY with space(10)
	          	else
					replace PNL.SMN_IDY with INL.SMN_IDY
		        endif

	    replace PNL.EXD_CAL with space(1)
	    		* INL.EXD_CAL,;	            
	            * PNL.USR_IDY with INL.USR_IDY,;
		select INL
		skip 1
	enddo
endif	
if R_PNL > R_INL
	select INL
	delete all
	select PNL
	go top
	do while eof("PNL")
		select INL
		append blank
		replace INL.INV_IDY with PNL.INV_IDY,;
				INL.INV_DAT with PNL.INV_DAT,;
				INL.ORD_IDY with PNL.ORD_IDY,;
				INL.CST_IDY with PNL.CST_IDY,;
				INL.ACT_IDY with PNL.ACT_IDY,;
				INL.GRS_AMT with PNL.GRS_AMT,;
				INL.SPD_PCG with PNL.SPD_PCG,;
				INL.SPD_AMT with PNL.SPD_AMT,;
				INL.CSD_PCG with PNL.CSD_PCG,;
	            INL.CSD_AMT with PNL.CSD_AMT,;
	            INL.EXD_PCG with PNL.EXD_PCG,;
	            INL.EXD_CAL with PNL.EXD_CAL,;
	            INL.EXD_AMT with PNL.EXD_AMT,;
	            INL.TAX_PCG with PNL.TAX_PCG,;
	            INL.TAX_AMT with PNL.TAX_AMT,;
	            INL.NET_AMT with PNL.NET_AMT,;
	            INL.SMN_IDY with PNL.SMN_IDY,;
	            INL.TXN_TYP with PNL.TXN_TYP,;
	            INL.FRT_TYP with PNL.FRT_TYP,;
	            INL.FRT_AMT with PNL.FRT_AMT,;
	            INL.DCH_IDY with PNL.DCH_IDY,;
	            INL.DCH_DAT with PNL.DCH_DAT,;
	            INL.LRN_IDY with PNL.LRN_IDY,;			
	            INL.LRN_DAT with PNL.LRN_DAT,;
	            INL.LRN_AMT with PNL.LRN_AMT,;
	            INL.WBL_IDY with PNL.WBL_IDY,;
	            INL.WBL_DAT with PNL.WBL_DAT,;
	            INL.PKG_CST with PNL.PKG_CST,;
	            INL.TTL_CAS with PNL.TTL_CAS,;
	            INL.TXN_RMK with PNL.TXN_RMK,;
	            INL.TXN_TPT with PNL.TXN_TPT,;
	            INL.TAX_ACT with PNL.TAX_ACT,;
	            INL.ROU_ACT with PNL.ROU_ACT,;
	            INL.ROU_OFF with PNL.ROU_OFF,;
	            INL.FRT_ACT with PNL.FRT_ACT,;
	            INL.CRT_DAY with PNL.CRT_DAY,;
	            INL.DUE_DAT with PNL.DUE_DAT,;
	            INL.DES_MOD with PNL.DES_MOD,;
	            INL.STO_RMK with PNL.STO_RMK,;
	            INL.TXN_RM1 with PNL.TXN_RM1,;
	            INL.TXN_RM2 with PNL.TXN_RM2,;
	            INL.VCH_IDY with PNL.VCH_IDY,;
	            INL.VCH_DAT with PNL.VCH_DAT,;
	            INL.SCH_CHK with PNL.SCH_CHK,;
	            INL.USR_IDY with PNL.USR_IDY,;
	            INL.SYS_DAT with PNL.SYS_DAT,;
	            INL.SYS_TIM with PNL.SYS_TIM,;
	            INL.SCH_IDY with PNL.SCH_IDY,;
	            INL.SRN_AMT with PNL.SRN_AMT,;
	            INL.DIS_CHK with PNL.DIS_CHK
		select PNL
		skip 1
	enddo
endif
close all
RETURN
