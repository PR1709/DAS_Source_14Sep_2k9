close all
clear all
select 0
use DIR alias DIR
REPLACE ALL DIR_MSG with space(80) FOR DIR_IDY = "SYS" AND DIR_IDX = "ABK"
REPLACE ALL DIR_ALS WITH SPACE(80) FOR DIR_IDY = "SYS" AND DIR_IDX = "IND"
REPLACE ALL DIR_MSG WITH "das@apac.ko.com" FOR DIR_IDY = "EML" AND DIR_IDX = "IDY"
close all
return
