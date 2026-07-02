close all
select 0
use INL alias INL
* scatter MEMVAR
copy to array MEMARR
use PNL alias PNL
select PNL
* gather MEMVAR
append from array MEMARR
release MEMVAR
close all
return
