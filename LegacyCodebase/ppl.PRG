select 0
use inl alias INL
M_CTR = 0
do while .t.
   if M_CTR >= 5000
      EXIT
   endif
   select INL
   append from ZNL
   M_CTR = (M_CTR + 1)
enddo
select INL
go top
replace all INV_IDY with alltrim(str(recno()))
close all
return
