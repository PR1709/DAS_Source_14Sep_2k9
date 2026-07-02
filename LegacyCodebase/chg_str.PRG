set talk off
close all
CLEAR ALL
OPEN DATABASE EFF
alter table pre add column pre_lev c(5) null
alter table pre add column pre_ain c(5) null
alter table pre add column pre_aou c(5) null
alter table idy add column idy_asm c(10) null
alter table idy add column idy_ain c(10) null
alter table idy add column idy_aou c(10) null
alter table pre add column pre_asm c(5) null
alter table van add column fre_qty n(15,3) null
alter table van add column cas_amt n(15,2) null
alter table van add column chq_amt n(15,2) null
alter table van add column dis_amt n(15,2) null
alter table van add column exp_amt n(15,2) null
alter table van add column add_amt n(15,2) null
alter table AST add column PTY_idy c(10) null
alter table tgt add column tgt_cfc n(15,3) null
alter table wek add column mth_nos n(2) null
CREATE TABLE SCL(INV_DAT DATE,SCH_IDY C(10),PRD_IDY C(10),OPN_STK N(15,3),RCT_STK N(15,3),ISS_STK N(15,3),CLS_STK N(15,3),TOT_SAL N(15,3),SAL_STK N(15,3),FRE_STK N(15,3),SCH_AMT N(15,2),BAS_TYP C(3))
CREATE TABLE TRT(PRD_IDY C(10),PRD_CAT C(25)NULL,PRD_CLS C(25),BRD_NME C(25),T_DAY_STK N(15,3),T_WEK_STK N(15,3),T_MTH_STK N(15,3),A_DAY_STK N(15,3),A_WEK_STK N(15,3),A_MTH_STK N(15,3),MTH_DYS N(5),WEK_DYS N(5),CMP_WEK N(5),CMP_MTH N(5))
close all

