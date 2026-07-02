close all
clear all
clear
store space(10) to M_VAN_CHN
set confirm on
select 0 
use VAN alias VAN
@5,5 say "Enter Van Challan ID" get M_VAN_CHN picture "@!"
read
if M_VAN_CHN = space(10)
    close all
    return
else
	replace all VAN_CLS with "C" for INV_IDY = trim(M_VAN_CHN)
	wait window "Challan Force Closed for Challan Number " + trim(M_VAN_CHN)
	close all
	return
endif
close all
clear all
return
