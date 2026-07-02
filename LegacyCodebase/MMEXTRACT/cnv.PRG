close all
clear all
use inv
delete all for substr(inv_idy, 1, 3) = 'SAL' or substr(inv_idy, 1, 3) = 'TAX'
REPLACE ALL ORD_IDY WITH SPACE(10) FOR TXN_TYP = "SSO"
use inl
delete all for substr(inv_idy, 1, 3) = 'SAL' or substr(inv_idy, 1, 3) = 'TAX'
REPLACE ALL ORD_IDY WITH SPACE(10) FOR TXN_TYP = "SSO"
REPLACE ALL WBL_IDY WITH SPACE(10) FOR TXN_TYP = "SSO"
REPLACE ALL DCH_IDY WITH SPACE(10) FOR TXN_TYP = "SSO"
use pnv
delete all for substr(inv_idy, 1, 3) = 'SAL' or substr(inv_idy, 1, 3) = 'TAX'
REPLACE ALL ORD_IDY WITH SPACE(10) FOR TXN_TYP = "SSO"
use pnl
delete all for substr(inv_idy, 1, 3) = 'SAL' or substr(inv_idy, 1, 3) = 'TAX'
REPLACE ALL ORD_IDY WITH SPACE(10) FOR TXN_TYP = "SSO"
REPLACE ALL WBL_IDY WITH SPACE(10) FOR TXN_TYP = "SSO"
REPLACE ALL DCH_IDY WITH SPACE(10) FOR TXN_TYP = "SSO"
use irt
delete all for substr(inv_idy, 1, 3) = 'SAL' or substr(inv_idy, 1, 3) = 'TAX'
use van
delete all for len(alltrim(lds_idy)) > 0
use lds
zap
close all
do ind
clear all
return

