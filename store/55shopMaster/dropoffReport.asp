<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<html>
<head>
<title>Drop-Off Report - by Date</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:10px;">
<div id="pcCPmain" style="background-image: none;">
<table class="pcCPcontent" style="width: auto;">
        
<%
Dim connTemp
Dim con
Set con=CreateObject("ADODB.Connection")
con.Open scDSN 
	
' Choose the records to display
err.clear
Dim strTDateVar, strTDateVar2, DateVar, DateVar2
strTDateVar=Request.QueryString("FromDate")
DateVar=strTDateVar
if scDateFrmt="DD/MM/YY" then
	DateVarArray=split(strTDateVar,"/")
	if SQL_Format="1" then
	    DateVar=(DateVarArray(0)&"/"&DateVarArray(1)&"/"&DateVarArray(2))
	else
	    DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
	end if
else
    DateVarArray=split(strTDateVar,"/")
	if SQL_Format="1" then
	    DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
	else
	    DateVar=(DateVarArray(0)&"/"&DateVarArray(1)&"/"&DateVarArray(2))
	end if
end if

strTDateVar2=Request.QueryString("ToDate")
DateVar2=strTDateVar2
if scDateFrmt="DD/MM/YY" then
	DateVarArray2=split(strTDateVar2,"/")
	if SQL_Format = "1" then
	    DateVar2=(DateVarArray2(0)&"/"&DateVarArray2(1)&"/"&DateVarArray2(2))
	else
	    DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	end if
else
    DateVarArray2=split(strTDateVar2,"/")
	if SQL_Format = "1" then
	    DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	else
	    DateVar2=(DateVarArray2(0)&"/"&DateVarArray2(1)&"/"&DateVarArray2(2))
	end if
end if

if err.number<>0 then
	DateVar=Request.QueryString("FromDate")
	DateVar2=Request.QueryString("ToDate")
end if

err.clear

tmpD="orders.orderDate"
tmpD1="processDate"
tmpD2="Processed On"

if (DateVar<>"") and IsDate(DateVar) then
	if scDB="Access" then
		TempSQL1=" AND " & tmpD & " >=#" & DateVar & "# "
	else
		TempSQL1=" AND " & tmpD & " >='" & DateVar & "' "
	end if
else
	TempSQL1=""
end if
if (DateVar2<>"") and IsDate(DateVar2) then
	if scDB="Access" then
		TempSQL2=" AND " & tmpD & " <=#" & DateVar2 & "# "
	else
		TempSQL2=" AND " & tmpD & " <='" & DateVar2 & "' "
	end if
else
	TempSQL2=""	
end if

CustCountry=request("CountryCode")
if CustCountry<>"" then
	showDetailedReports = "1"
	SQLCC=" ,customers "
	SQLCC1=" and customers.idcustomer=orders.idcustomer and customers.countryCode='" & CustCountry & "' "
	call opendb()
	mySQL="SELECT countryName FROM countries WHERE CountryCode='" & CustCountry & "'"
	set rs1=conntemp.execute(mySQL)	
	CountryName=rs1("countryName")
	call closedb()
else
	SQLCC=""
	SQLCC1=""
end if

query="SELECT orders.idorder,orders.idcustomer,orders.total,orders.rmaCredit,orders.pcOrd_CatDiscounts,orders.iRewardValue,orders.shipmentDetails,orders.paymentDetails,orders.taxAmount,orders.discountDetails,orders.orderDate,orders.processDate,orders.shipDate FROM Orders" & SQLCC & " WHERE orders.orderStatus=1 " & TempSQL1 & TempSQL2 & SQLCC1 & " ORDER BY " & tmpD & " DESC;"

' Our Recordset Object
Dim rs
Set rs=CreateObject("ADODB.Recordset")
rs.CursorLocation=adUseClient
rs.Open query, scDSN , 3, 3
	
' If the returning recordset is not empty
If rs.EOF Then %>
	<tr> 
		<td colspan="6">No records match your query</td>
	</tr>
<% Else %>
	<tr> 
		<td colspan="6"><h2>Total drop-offs recorded from <%=strTDateVar%> to <%=strTDateVar2%> <%if CountryName<>"" then%>
&nbsp;in <%=CountryName%><%end if%></h2></td>
</tr>
<tr> 
	<td colspan="6"><% Response.Write "Total Records Found : " & rs.RecordCount %></td>
</tr>
<tr> 
	<th nowrap>Order #</th>
	<th nowrap>Date</th>
	<th nowrap>Customer Name</th>
	<th nowrap>Customer E-mail</th>
	<th nowrap>Referrer</th>
	<th nowrap>Order Total</th>
</tr>
<tr>
	<td colspan="6" class="pcCPspacer"></td>
</tr>
<% 
gTotalNumberOrders=rs.RecordCount
gTotalsales=0
gTotaltaxes=0
gTotalcom=0
gTotalshipfees=0
gTotalhandfees=0
gTotalpayfees=0
do until rs.EOF
	pc_idorder=rs("idorder")
	pc_idcustomer=rs("idcustomer")
	pc_total=rs("total")
	pc_rmaCredit=rs("rmaCredit")
	if pc_rmaCredit<>"" then
	else
		pc_rmaCredit=0
	end if
	pc_total=pc_total-pc_rmaCredit
	pc_CatDiscounts=rs("pcOrd_CatDiscounts")
	if pc_CatDiscounts<>"" then
	else
		pc_CatDiscounts=0
	end if
	pc_Rewards=rs("iRewardValue")
	if pc_Rewards<>"" then
	else
	pc_Rewards=0
	end if
	pc_shipmentDetails=rs("shipmentDetails")
	pc_paymentDetails=rs("paymentDetails")
	pc_taxAmount=rs("taxAmount")
	discountDetails=rs("discountDetails")
	pc_orderDate=rs("orderDate")
	if scDateFrmt="DD/MM/YY" then
		pc_orderDate=(day(pc_orderDate)&"/"&month(pc_orderDate)&"/"&year(pc_orderDate))
	end if
	pc_processDate=rs(tmpD1)
	if scDateFrmt="DD/MM/YY" then
		pc_processDate=(day(pc_processDate)&"/"&month(pc_processDate)&"/"&year(pc_processDate))
	end if 

	querySTR="SELECT * FROM ProductsOrdered WHERE idorder="& pc_idorder
	Set rsSTR=CreateObject("ADODB.Recordset")
	rsSTR.CursorLocation=adUseClient
	rsSTR.Open querySTR, scDSN , 3, 3
	bOrderTotal=0
	do until rsSTR.eof
		querySTR="SELECT name,lastname,email,IDRefer FROM customers WHERE idcustomer="& pc_idcustomer
		Set rsCust=CreateObject("ADODB.Recordset")
		rsCust.CursorLocation=adUseClient
		rsCust.Open querySTR, scDSN , 3, 3
		CustName=rsCust("name")& " "&rsCust("lastname")
		CustEmail=rsCust("email")
		CustRefer=rsCust("IDRefer")
		rsCust.Close
		set rsCust=nothing
		
		if CustRefer <> "" Then
			queryStrRef="SELECT Name FROM Referrer WHERE IdRefer="& CustRefer
			Set rsCustRef=CreateObject("ADODB.Recordset")
			rsCustRef.CursorLocation=adUseClient
			rsCustRef.Open queryStrRef, scDSN , 3, 3
			if rsCustRef.EOF Then
				Referrer="N/A"
			else
				Referrer=rsCustRef("Name")
			end if
			rsCustRef.Close
			set rsCustRef=nothing
		else
			Referrer="N/A"
		end if
		
		unitTotal=rsSTR("unitPrice")
		quantity=rsSTR("quantity")
		bOrderTotal=0 + (unitTotal * quantity)
		rsSTR.moveNext
	loop
	rsSTR.Close
	set rsSTR=nothing
	gTotalsales=gTotalsales + pc_total
	
	'Shipping & Handling Fees
	pshipmentDetails=pc_shipmentDetails
	shipping=split(pshipmentDetails,",")
	if ubound(shipping)>1 then
		if NOT isNumeric(trim(shipping(2))) then
			shipfees=0
		else
			shipfees=cdbl(trim(shipping(2)))
		end if	
		if NOT isNumeric(trim(shipping(2))) then
			HandFees=0
		else
			HandFees=cdbl(trim(shipping(3)))
		end if
	else
		shipfees=0
		Handfees=0
	end if
	gTotalshipfees=gTotalshipfees + shipFees
	gTotalhandfees=gTotalhandfees + HandFees
	
	'Payment Fees
	err.clear
	ppaymentDetails=pc_paymentDetails
	payment = split(ppaymentDetails,"||")
	PaymentType=payment(0)
	on error resume next
	If payment(1)="" then
	 if err.number<>0 then
	 	PayCharge=0
	 end if
		PayCharge=0
	else
		PayCharge=cdbl(payment(1))
	end If
	err.number=0
	
	gTotalpayfees=gTotalpayfees+PayCharge
	gTotaltaxes=gTotaltaxes + pc_taxAmount
	pc_Discounts=0
	pc_discountDetails=""
	pc_DCount=0
	if (discountDetails<>"") and (Istr(discountDetails,"||")>0) then
	A=split(discountDetails,",")
	k=lbound(A)-1
	do
		k=k+1
		tmp1=""
		if Istr(discountDetails,"||")>0 then
		tmp1=A(k)
		else
		tmp1=A(k) & "," & A(k+1)
		k=k+1
		end if
		if trim(tmp1)<>"" then
		B=split(tmp1,"|| ")
		pc_Discounts=pc_Discounts+cdbl(B(1))
		tmp1=replace(tmp1,"|| ",scCurSign)
		if (InStr(tmp1,scCurSign & "0")>0) and (InStr(tmp1,scCurSign & "0.")=0) then
		tmp1 = mid(tmp1,1,InStr(tmp1," - " & scCurSign)-1)
		end if
		pc_discountDetails=pc_discountDetails & tmp1 & "<br>"
		pc_DCount=pc_DCount+1
		end if
	loop until k>=ubound(A)
	end if
	if pc_CatDiscounts<>0 then
		pc_discountDetails=pc_discountDetails & "Category-based Quantity Discounts: " & scCurSign & pc_CatDiscounts & "<br>"
		pc_Discounts=pc_Discounts+cdbl(pc_CatDiscounts)
		pc_DCount=pc_DCount+1
	end if
	if pc_Rewards<>0 then
		pc_discountDetails=pc_discountDetails & "Reward Points were used: " & scCurSign & pc_Rewards & "<br>"
		pc_Discounts=pc_Discounts+cdbl(pc_Rewards)
		pc_DCount=pc_DCount+1
	end if
	if pc_DCount>1 and pc_Discounts<>0 then
		pc_discountDetails=pc_discountDetails & "Discount Amount: " & scCurSign & pc_Discounts & "<br>"
	end if
	if pc_rmaCredit<>0 then
		pc_discountDetails=pc_discountDetails & "Credit: " & scCurSign & pc_rmaCredit & "<br>"
	end if
	%>
        
	<tr>  
		<td nowrap valign="top"><a href="orddetails.asp?id=<%=pc_idorder%>" target="_blank" title="View details for this order"><%=(scpre + int(pc_idorder))%></a></td>
		<td nowrap valign="top"><%=pc_orderDate%></td>
		<td nowrap valign="top"><a href="viewCustOrders.asp?idcustomer=<%=pc_idcustomer%>" target="_blank" title="View all orders by this customer"><%=CustName%></a></td>
		<td nowrap valign="top"><a href="mailto:<%=CustEmail%>"><%=CustEmail%></a></td>
		<td nowrap valign="top"><%=Referrer%></td>
		<td nowrap valign="top" align="center"><%=scCurSign&money(pc_total)%></td>
	</tr>
        
	<% gTotalcomm=gTotalcomm + affiliatepay %>
	<% rs.MoveNext
	loop %>
        
<tr>         
	<td colspan="6" class="pcCPspacer"></td>
</tr>
<tr bgcolor="#e1e1e1">
	<td colspan="5" align="right"><strong>Totals</strong></td>
	<td><b><%=scCurSign&money(gTotalsales)%></font></b></td>
</tr>
<tr>
	<td colspan="5" align="right">Product-Only Sales:</td>
	<td>
		<%
		ProductSales=gTotalsales-gTotalshipfees-gTotalhandfees-gTotalpayfees-gTotaltaxes
		response.write scCurSign&money(ProductSales)
		%>
	</td>
</tr>
<tr>         
	<td colspan="6" class="pcSmallText" align="right">Total sales minus shipping, payment and other charges</td>
</tr>
<tr>         
	<td colspan="6" class="pcCPspacer"></td>
</tr>
<% end if%>
</table>
<%	' Done. Now release Objects
	con.Close
	Set con=Nothing
	Set rs=Nothing
%>
</div>
</body>
</html>