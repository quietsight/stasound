<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->    
<!--#include file="../includes/stringfunctions.asp"-->
<%
dim rstemp, conntemp, query, AddrList
Dim purchaseType

AddrList=""

call opendb()

SOptedIn=getUserInput(request("SOptedIn"),0)

SIDProduct=getUserInput(request("SIDProduct"),0)
SIDCategory=getUserInput(request("SIDCategory"),0)
	if not validNum(SIDProduct) then SIDProduct=0
	if not validNum(SIDCategory) then SIDCategory=0
	if SIDCategory>0 then SIDProduct=0

SCustType=getUserInput(request("SCustType"),0)

SStartDate=getUserInput(request("SStartDate"),0)
SEndDate=getUserInput(request("SEndDate"),0)
	if not IsDate(SStartDate) then SStartDate=""
	if not IsDate(SEndDate) then SEndDate=""

purchaseType=getUserInput(request("purchaseType"),0)
	if purchaseType="" then
		purchaseType=0
	end if
	'// If the customer's purchases are to be ignored, clear product and category IDs
	if purchaseType="0" then
		SIDProduct=0
		SIDCategory=0
	end if


function checkNoOrderCustomer(COptedIn,CCustType)

	dim mrs4,query4, myTemp4

	myTemp4=" WHERE (idcustomer NOT IN (SELECT DISTINCT idcustomer FROM orders)) "

	if (COptedIn<>"") and (COptedIn<>"0") then
		if myTemp4="" then
			myTemp4=" where "
		else
			myTemp4=myTemp4 & " and "
		end if
		myTemp4=myTemp4 & " RecvNews=" & COptedIn
	end if
	
	if (CCustType<>"") and (CCustType<>"0") then
		if myTemp4="" then
			myTemp4=" where "
		else
			myTemp4=myTemp4 & " and "
		end if
		if CCustType="1" then
			myTemp4=myTemp4 & " customers.customerType<>1 AND customers.idCustomerCategory=0"
		end if
		if CCustType="2" then
			myTemp4=myTemp4 & " customers.customerType=1 AND customers.idCustomerCategory=0"
		end if
		if instr(CCustType,"CC_")>0 then
			tmp_Arr=split(CCustType,"CC_")
			if tmp_Arr(1)<>"" then
				myTemp4=myTemp4 & " customers.idCustomerCategory=" & tmp_Arr(1)
			end if
		end if
	end if

	query4="SELECT email FROM Customers" & myTemp4
	set mrs4=server.CreateObject("ADODB.RecordSet")
	set mrs4=connTemp.execute(query4)
	
	do while not mrs4.eof
		if instr(AddrList,mrs4("email") & "**")=0 then
			AddrList=AddrList & mrs4("email") & "**"
		end if
		mrs4.movenext
	loop

	set mrs4=nothing

end function


function checkCustomer(CIDCustomer,COptedIn,CCustType)

	dim mrs4,query4, myTemp4

	myTemp4=""

	if (CIDCustomer<>"") and (CIDCustomer<>"0") then
		if myTemp4="" then
			myTemp4=" where "
		else
			myTemp4=myTemp4 & " and "
		end if
		myTemp4=myTemp4 & " idCustomer=" & CIDCustomer
	end if

	if purchaseType="1" or purchaseType="3" then
		if myTemp4="" then
			myTemp4=" where "
		else
			myTemp4=myTemp4 & " and "
		end if
		myTemp4=myTemp4 & " (idcustomer IN (SELECT DISTINCT idcustomer FROM orders)) "
	end if
	
	if (COptedIn<>"") and (COptedIn<>"0") then
		if myTemp4="" then
			myTemp4=" where "
		else
			myTemp4=myTemp4 & " and "
		end if
		myTemp4=myTemp4 & " RecvNews=" & COptedIn
	end if
	
	if (CCustType<>"") and (CCustType<>"0") then
		if myTemp4="" then
			myTemp4=" where "
		else
			myTemp4=myTemp4 & " and "
		end if
		if CCustType="1" then
			myTemp4=myTemp4 & " customers.customerType<>1 AND customers.idCustomerCategory=0"
		end if
		if CCustType="2" then
			myTemp4=myTemp4 & " customers.customerType=1 AND customers.idCustomerCategory=0"
		end if
		if instr(CCustType,"CC_")>0 then
			tmp_Arr=split(CCustType,"CC_")
			if tmp_Arr(1)<>"" then
				myTemp4=myTemp4 & " customers.idCustomerCategory=" & tmp_Arr(1)
			end if
		end if
	end if

	query4="SELECT email FROM Customers" & myTemp4
	set mrs4=server.CreateObject("ADODB.RecordSet")
	set mrs4=connTemp.execute(query4)

	do while not mrs4.eof
		if instr(AddrList,mrs4("email") & "**")=0 then
			AddrList=AddrList & mrs4("email") & "**"
		end if
		mrs4.movenext
	loop

	set mrs4=nothing

end function

function checkOrder(CIDOrder,CStartDate,CEndDate)

	dim mrs3,query3, myTemp3

	myTemp3=""

	if (CIDOrder<>"") and (CIDOrder<>"0") then
		if myTemp3="" then
			myTemp3=" where "
		else
			myTemp3=myTemp3 & " and "
		end if
		myTemp3=myTemp3 & " idOrder=" & CIDOrder
	end if

	if (CStartDate<>"") and (IsDate(CStartDate)) then
		if myTemp3="" then
			myTemp3=" where "
		else
			myTemp3=myTemp3 & " and "
		end if
		if scDB="Access" then
			myTemp3=myTemp3 & " orderDate>=#" & CStartDate & "#"
		else
			myTemp3=myTemp3 & " orderDate>='" & CStartDate & "'"
		end if
	end if

	if (CEndDate<>"") and (IsDate(CEndDate)) then
		if myTemp3="" then
			myTemp3=" where "
		else
			myTemp3=myTemp3 & " and "
		end if
		if scDB="Access" then
			myTemp3=myTemp3 & " orderDate<=#" & CEndDate & "#"
		else
			myTemp3=myTemp3 & " orderDate<='" & CEndDate & "'"
		end if
	end if

	query3="select idcustomer from Orders " & myTemp3
	set mrs3=server.CreateObject("ADODB.RecordSet")
	set mrs3=connTemp.execute(query3)

	do while not mrs3.eof
		call checkCustomer(mrs3("idcustomer"),SOptedIn,SCustType)
		mrs3.movenext
	loop

	set mrs3=nothing

end function

function checkproduct(CIDProduct)

	dim mrs2,query2
	
	if purchaseType="1" then
		query2="SELECT DISTINCT idorder FROM ProductsOrdered WHERE idproduct=" & CIDProduct
	elseif purchaseType="2" then
		query2="SELECT DISTINCT idorder FROM Orders WHERE (idcustomer NOT IN (SELECT DISTINCT idcustomer FROM Orders INNER JOIN ProductsOrdered ON Orders.idorder=ProductsOrdered.idOrder WHERE ProductsOrdered.idproduct=" & CIDProduct & "))"
	else
		query2=""
	end if
	
	if query2<>"" then
		set mrs2=connTemp.execute(query2)
		do while not mrs2.eof
			call checkOrder(mrs2("idOrder"),SStartDate,SEndDate)
		mrs2.movenext
		loop
		set mrs2=nothing
	end if

end function

function checkcategory(CIDCategory)

	dim mrs1,query1

	IF purchaseType="1" THEN
		query1="select idproduct from categories_products where idcategory=" & CIDCategory
		set mrs1=connTemp.execute(query1)
	
		do while not mrs1.eof
		call checkproduct(mrs1("idproduct"))
		mrs1.movenext
		loop
		set mrs1=nothing
	ELSEIF purchaseType="2" THEN
		query1="SELECT DISTINCT idorder FROM Orders WHERE (idcustomer NOT IN (SELECT DISTINCT idcustomer FROM Orders INNER JOIN ProductsOrdered ON Orders.idorder=ProductsOrdered.idOrder WHERE (ProductsOrdered.idproduct IN (select distinct idproduct from categories_products where idcategory=" & CIDCategory & "))))"
		set mrs1=connTemp.execute(query1)

		do while not mrs1.eof
			call checkOrder(mrs1("idOrder"),SStartDate,SEndDate)
			mrs1.movenext
		loop
		set mrs1=nothing
	END IF

end function


'// Determine query to run
if purchaseType="3" then
	if SEndDate="" and SStartDate="" then
		call checkCustomer("0",SOptedIn,SCustType)
	else
		call checkOrder("0",SStartDate,SEndDate)
	end if
elseif purchaseType="4" then
	call checkNoOrderCustomer(SOptedIn,SCustType)
else
	if (SIDCategory<>"") and (SIDCategory<>"0") then
		call checkcategory(SIDCategory)
	else
		if (SIDProduct<>"") and (SIDProduct<>"0") then
			call checkproduct(SIDProduct)
		else
			if (SStartDate<>"" AND SEndDate<>"") then
				call checkOrder("0",SStartDate,SEndDate)
			else
				if purchaseType="2" then
					call checkNoOrderCustomer(SOptedIn,SCustType)
				else
					call checkCustomer("0",SOptedIn,SCustType)
				end if
			end if
		end if
	end if
	if purchaseType="2" and SEndDate="" and SStartDate="" then
		call checkNoOrderCustomer(SOptedIn,SCustType)
	end if
end if

'Sort e-mail address list
if AddrList<>"" then

	AList=split(AddrList,"**")
	
	For dem=lbound(AList) to ubound(AList)
		For dem1=dem+1 to ubound(AList)
			if AList(dem)>AList(dem1) then
				TempStr=AList(dem1)
				AList(dem1)=AList(dem)
				AList(dem)=TempStr
			end if
		next
	next
	
	session("AddrList")=AList
	session("AddrCount")=ubound(AList)

else

	dim BList(1)
	session("AddrList")=BList
	session("AddrCount")=0

end if

call closedb()
response.redirect "newsWizStep2.asp?from=1"
%>