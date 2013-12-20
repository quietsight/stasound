<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Function RoundTo(intNum, intRn)
	RoundTo= Int((intNum / intRn)+.5) * intRn
End Function 

function money(anumber)
	if anumber="" OR anumber=0 then
		anumber=0
		money=0
	end if
	if cdbl(anumber)<0 then
		test=true
		bnumber=abs(anumber)
	else
		bnumber=anumber
	end if
	' round number
	if bnumber<>0 then
		if InStr(Cstr(10/3),",")>0 then		
			if Instr(bnumber,".")>0 then
				money=FormatNumber(bnumber,,,,0)
				money=RoundTo(money,.01)
				money=replace(money,".",",")
			else
				money=RoundTo(bnumber,.01)
			end if
		else
			money=RoundTo(bnumber,.01)
		end if
	end if
	if scDecSign="," then
		' replace . by ,
		money=replace(money,".",",")
	else
		' replace , by .
		money=replace(money,",",".")
	end if

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
	If len(money)>6 then
		money=divideNo(money)
	End If
	dim test
	if test=true then
	money="-" & money
	test=false
	end if
end function

function extmoney(anumber)
	if anumber="" OR anumber=0 then
		anumber=0
		extmoney=0
	end if
	if cdbl(anumber)<0 then
	test=true
	anumber=abs(anumber)
	end if
	' round number
	if anumber<>0 then
		extmoney=round(anumber,4)
	end if
	if scDecSign="," then
		' replace . by ,
		extmoney=replace(extmoney,".",",")
	else
		' replace , by .
		extmoney=replace(extmoney,",",".")
	end if

	' locate dec division
	dim indexPoint
	indexPoint=instr(extmoney, scDecSign)
	
	' for integer, add .00
	if indexPoint=0 then
			extmoney=Cstr(extmoney)+scDecSign+"00"
	end if

	' calculate if 0 or 00
	dim extmoneyLarge, decPart
	extmoneyLarge=len(extmoney)
	decPart=right(extmoney,extmoneyLarge-indexPoint)
	
	' add to original numbers
	dim g
	for g=0 to (1- (extmoneyLarge-indexPoint))
			extmoney=Cstr(extmoney)+"0"    
	next
	If len(extmoney)>6 then
		extmoney=divideNo(extmoney)
	End If
	if test=true then
	extmoney="-" & extmoney
	test=false
	end if
end function

' use this function to display numbers like 125,000,000
function divideNo(anumber)

' locate scDecSign position
scDecSignExist = instr(anumber, scDecSign) - 1
if scDecSignExist <= 0 then 
	scDecSignExist = len(anumber)
end if

' add divideNo to integers
CntNo = 0
if scDecSignExist > 3 then
 for indexPoint = scDecSignExist to 1 step -1
	pNumber = mid(anumber,indexPoint,1)
	if pNumber <> scDecSign then
		CntNo  = CntNo + 1
		pdivideNo  = pNumber & pdivideNo 
 	end if	
	if CntNo = 3 and (indexPoint > 1) then
    pdivideNo  = scDivSign & pdivideNo 
		CntNo  = 0
	end if
	pNumber = ""
 next 

 ' add decimals
 pdivideNo = pdivideNo & mid(anumber, scDecSignExist+1, Len(anumber))
 divideNo  = pdivideNo
else
 divideNo  = anumber
end if

end function

function taxmoney(anumber)
' round number
taxmoney=round(anumber,4)
end function
%>