CLOSE ALL
CLEAR ALL
CLEAR
SELECT 0
USE BRD ALIAS BRD
go top 
do while .not. eof("BRD")
	if isnull(P_INC_DIS)
		replace P_INC_DIS with 0
	endif
	if isnull(S_INC_DIS)
		replace S_INC_DIS with 0
	endif
	if isnull(B_PUR_TAX)
		replace B_PUR_TAX with 0
	endif
	if isnull(B_SAL_TAX)
		replace B_SAL_TAX with 0
	endif
	if isnull(B_CRT_RAT)
		replace B_CRT_RAT with 0
	endif
	if isnull(B_CRT_DIS)
		replace B_CRT_DIS with 0
	endif
	if isnull(SKU_SBU)
		replace SKU_SBU with 0
	endif
	if isnull(P_PUR_VAT)
		replace P_PUR_VAT with 0
	endif
	if isnull(P_CRT_RNT)
		replace P_CRT_RNT with 0
	endif
	if isnull(P_CRT_VAT)
		replace P_CRT_VAT with 0
	endif
	if isnull(P_TOT_VAT)
		replace P_TOT_VAT with 0
	endif
	if isnull(U_PUR_LPP)
		replace U_PUR_LPP with 0
	endif
	if isnull(PUR_LPP)
		replace PUR_LPP with 0
	endif
	if isnull(S_SAL_VAT)
		replace S_SAL_VAT with 0
	endif
	if isnull(S_CRT_RNT)
		replace S_CRT_RNT with 0
	endif
	if isnull(S_CRT_VAT)
		replace S_CRT_VAT with 0
	endif
	if isnull(S_TOT_VAT)
		replace S_TOT_VAT with 0
	endif
	if isnull(U_SAL_LPP)
		replace U_SAL_LPP with 0
	endif
	if isnull(SAL_LPP)
		replace SAL_LPP with 0
	endif
	skip 1
enddo
close all
return
