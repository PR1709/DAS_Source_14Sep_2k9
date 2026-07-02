store "Mala       Ramaneshwar      Ritvik Sundari Jyotirlingam" to var1
store var1 to var2
store 1 to ctr
do while .t.
	if len(alltrim(var1)) = ctr
	
		exit
	endif
	if substr(var1, ctr, ctr) = chr(32)
		? ctr
		var1 = stuff(var1, ctr, 1, "_")
	endif
	ctr = ctr + 1
enddo
? ctr
? var1

return
