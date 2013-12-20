<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Function RoundTo(intNum, intRn)
	RoundTo= Int((intNum / intRn)+.5) * intRn
End Function 

function money(anumber)
	dim test
	if anumber="" OR anumber=0 then
		anumber=0
		money=0
	end if
	if cdbl(anumber)<0 then
	test=true
	anumber=abs(anumber)
	end if
	' round number
	if anumber<>0 then
		money=RoundTo(anumber,.01)
	end if

	' replace , by .
	money=replace(money,",",".")

	' locate dec division
	dim indexPoint
	indexPoint=instr(money, scDecSign)
	
	' for integer, add .00
	if indexPoint=0 then
			money=Cstr(money)+scDecSign+"00"
	end if

	' calculate if 0 or 00
	dim moneyLarge, decPart
	moneyLarge=len(money)
	decPart=right(money,moneyLarge-indexPoint)
	
	' add to original numbers
	dim g
	for g=0 to (1- (moneyLarge-indexPoint))
			money=Cstr(money)+"0"    
	next

	if test=true then
	money="-" & money
	test=false
	end if
end function
%>