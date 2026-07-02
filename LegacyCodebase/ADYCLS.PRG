SET TALK OFF
SET ECHO OFF

close all
use VAN
PACK
REPLACE ALL VAN_CLS WITH "C" for alltrim(VAN_CLS) = "S"
close all

use DIR index DIR alias DIR
PACK
set order to 1
SET DATE BRITISH
REPLACE ALL DIR_MSG WITH "Y" FOR DIR_DAT >= CTOD("01/04/2007") .AND. DIR_DAT <= CTOD("31/12/2007")
wait window "Day Closures Over! Press any key!"

close all
return
