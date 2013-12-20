<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<%
Dim SubTitle

if request("onlyShow")="onlyDisc" then
	SubTitle = "(only orders with discount codes)"
end if%>
<%if request("onlyShow")="onlyGC" then
	SubTitle = "(only orders with gift certificate codes)"
end if%>
<html>
<head>
<title>Online Sales Report <% = SubTitle %></title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:10px; background-image: none;">
<div id="pcCPmain">
<% if trim(scCompanyLogo)<>"" then %>
<div id="pcCPcompanyLogo"><img src="../pc/catalog/<%=scCompanyLogo%>" style="margin: 10px 0;"></div>
<% end if %>
<table class="pcCPcontent">
        
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
	DateVar=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
end if
strTDateVar2=Request.QueryString("ToDate")
DateVar2=strTDateVar2
if scDateFrmt="DD/MM/YY" then
	DateVarArray2=split(strTDateVar2,"/")
	DateVar2=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	if err.number<>0 then
		DateVar=Request.QueryString("FromDate")
		DateVar2=Request.QueryString("ToDate")
	end if
end if
err.clear

tmpDate=request("basedon")
tmpD=""
tmpD1=""
tmpD2=""
Select case tmpDate
Case "2": tmpD="orders.processDate"
tmpD1="processDate"
tmpD2="Processed On"
Case "3": tmpD="pcPackageInfo.pcPackageInfo_ShippedDate"
tmpD1="pcPackageInfo_ShippedDate"
tmpD2="Shipped On"
Case Else: tmpD="orders.orderDate"
tmpD1="processDate"
tmpD2="Processed On"
End Select

if SQL_Format="1" then
	DateVar=Day(DateVar)&"/"&Month(DateVar)&"/"&Year(DateVar)
	DateVar2=Day(DateVar2)&"/"&Month(DateVar2)&"/"&Year(DateVar2)
else
	DateVar=Month(DateVar)&"/"&Day(DateVar)&"/"&Year(DateVar)
	DateVar2=Month(DateVar2)&"/"&Day(DateVar2)&"/"&Year(DateVar2)
end if

if (DateVar<>"") and IsDate(DateVar) then
	if scDB="Access" then
		TempSQL1=" AND " & tmpD & " >=#" & DateVar & "# "
	else
		TempSQL1=" AND " & tmpD & " >='" & DateVar & "' "
	end if
else
	TempSQL1=""
end if

if request("onlyShow")="onlyDisc" then
	TempSQL1=TempSQL1 & " AND orders.discountDetails<>'' AND orders.discountDetails<>'No discounts applied.' "
end if

if request("onlyShow")="onlyGC" then
	TempSQL1=TempSQL1 & " AND orders.pcOrd_GCDetails<>'' "
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

TempSpecial=""
if tmpDate="3" then
	tmpStr1=""
	if TempSQL1<>"" then
		tmpStr1=replace(TempSQL1,tmpD,"orders.shipDate")
		tmpStr1=replace(tmpStr1," AND ","")
	end if
	tmpStr2=""
	if TempSQL2<>"" then
		tmpStr2=replace(TempSQL2,tmpD,"orders.shipDate")
		tmpStr2=replace(tmpStr2," AND ","")
	end if
	tmpD="orders.processDate"
	
	TempSpecial=" AND "
	if tmpStr1 & tmpStr2 <> "" then
		TempSpecial=TempSpecial & " ((" & tmpStr1
		if tmpStr2<>"" then
			if tmpStr1<>"" then
				TempSpecial=TempSpecial & " AND "
			end if
			TempSpecial=TempSpecial & tmpStr2 & ") OR "
		end if
	end if
	
	TempSpecial=TempSpecial & " (orders.idorder IN (SELECT DISTINCT idorder FROM pcPackageInfo"
	if TempSQL1<>"" or TempSQL2<>"" then
		TempSpecial=TempSpecial & " WHERE pcPackageInfo_ID>0 " & TempSQL1 & TempSQL2
	end if
	TempSQL1=""
	TempSQL2=""
	TempSpecial=TempSpecial & "))"
	if tmpStr1 & tmpStr2 <> "" then
		TempSpecial=TempSpecial & ")"
	end if
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
CustomerTypeName=""
HaveCCType=0
CustomerType=request("customerType")
if CustomerType<>"" then
	if CustomerType="0" then
		CustomerTypeName="Retail Customer"
	end if
	if CustomerType="1" then
		CustomerTypeName="Wholesale Customer"
	end if
	if Instr(CustomerType,"CC_") then
		tmpA=split(CustomerType,"CC_")
		call opendb()
		query="SELECT pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & tmpA(1) & ";"
		SET rs=Server.CreateObject("ADODB.RecordSet")
		SET rs1=conntemp.execute(query)
		if not rs1.eof then
			CustomerTypeName=rs1("pcCC_Name")
			HaveCCType=1
		end if
		call closedb()
	end if
	if CustomerTypeName<>"" then
		if SQLCC="" then
			SQLCC=" ,customers "
		end if
		if SQLCC1="" then
			SQLCC1=" AND customers.idcustomer=orders.idcustomer "
		end if
		if HaveCCType=1 then
			SQLCC1=SQLCC1 & " AND customers.idCustomerCategory=" & tmpA(1) & " "
		else
			SQLCC1=SQLCC1 & " AND customers.customerType=" & CustomerType & " AND customers.idCustomerCategory=0 "
		end if
	end if
end if

'SM-S
tmpCond=""
pcSaleID=""
pcSaleName=""
IF Ucase(scDB)="SQL" THEN
	IF request("saleID")<>"" AND request("saleID")<>"0" then
		pcSaleID=request("saleID")
		tmpCond=" AND (Orders.idOrder IN (SELECT DISTINCT ProductsOrdered.idOrder FROM ProductsOrdered INNER JOIN pcSales_Completed ON ProductsOrdered.pcSC_ID=pcSales_Completed.pcSC_ID WHERE pcSales_Completed.pcSales_ID=" & pcSaleID & ")) "
		call opendb()
		queryS="SELECT pcSales_Name FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
		set rsS=connTemp.execute(queryS)
		if not rsS.eof then
			pcSaleName=rsS("pcSales_Name")
		end if
		set rsS=nothing
		call closedb()
	End if
END IF
'SM-E

query="SELECT DISTINCT orders.idorder,orders.idcustomer,orders.total,orders.rmaCredit,orders.pcOrd_CatDiscounts,orders.iRewardValue,orders.shipmentDetails,orders.paymentDetails,orders.taxAmount,orders.discountDetails,orders.orderDate,orders.processDate,orders.shipDate,orders.ord_VAT, orders.pcOrd_GCDetails FROM Orders" & SQLCC & " WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & SQLCC1 & tmpCond & " ORDER BY " & tmpD & " DESC;"

' Our Recordset Object
Dim rs
Set rs=CreateObject("ADODB.Recordset")
rs.CursorLocation=adUseClient
rs.Open query, scDSN , 3, 3
	
' If the returning recordset is not empty
If rs.EOF Then %>
	<tr> 
		<td colspan="8">No records match your query</td>
	</tr>
<% Else %>
	<tr> 
		<td colspan="8"><h2>Total sales recorded from <%=strTDateVar%> to <%=strTDateVar2%> <%if CountryName<>"" then%>
&nbsp;in <%=CountryName%><%end if%><%if request("onlyShow")&""<>"" then%>&nbsp;<%=SubTitle%><%end if%><%if CustomerTypeName<>"" then%><br>Customer Type: <%=CustomerTypeName%><%end if%>
<%if pcSaleName<>"" then%><br>Sale Name: <%=pcSaleName%><%end if%></h2></td>
</tr>
<tr> 
	<td colspan="8"><% Response.Write "Total Records Found : " & rs.RecordCount %></td>
</tr>
<tr> 
	<th nowrap><b>Order #</b></th>
	<th nowrap><b>Date</b></th>
	<th nowrap><b>Customer Name</b></th>
	<th nowrap><b>Referrer</b></th>
	<th nowrap><b>Order Total</b></th>
	<th nowrap><b><%if ptaxsetup=1 AND ptaxVAT="1" then%>VAT<%else%>Tax<%end if%></b></th>
    <% if request("onlyShow")="onlyGC" then %>
        <th nowrap><b>Gift Certificate(s) Used</b></th>
    <% else %>
        <th nowrap><b>Discounts Applied</b></th>
    <% end if %>
	<th nowrap><b><%=tmpD2%></b></th>
</tr>
<tr>
	<td colspan="8" class="pcCPspacer"></td>
</tr>
<% 
gTotalNumberOrders=rs.RecordCount
gTotalsales=0
gTotaltaxes=0
gTotalVAT=0
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
	if tmpDate<>"3" then
	pc_processDate=rs(tmpD1)
	if scDateFrmt="DD/MM/YY" then
		pc_processDate=(day(pc_processDate)&"/"&month(pc_processDate)&"/"&year(pc_processDate))
	end if
	else
		call opendb()
		query="SELECT pcPackageInfo_ShippedDate FROM pcPackageInfo WHERE idorder=" & pc_idorder
		set rsStr=connTemp.execute(query)
		pc_processDate=""
		if not rsStr.eof then
			do while not rsStr.eof
				tmp_processDate=rsStr("pcPackageInfo_ShippedDate")
				if scDateFrmt="DD/MM/YY" then
					tmp_processDate=(day(tmp_processDate)&"/"&month(tmp_processDate)&"/"&year(tmp_processDate))
				end if
				if  pc_processDate<>"" then
					pc_processDate=pc_processDate & "<br>"
				end if
				pc_processDate=pc_processDate & tmp_processDate
				rsStr.MoveNext
			loop
		else
			query="SELECT shipDate FROM orders WHERE idorder=" & pc_idorder
			set rsStr=connTemp.execute(query)
			if not rsStr.eof then
				pc_processDate=rsStr("shipDate")
				if scDateFrmt="DD/MM/YY" then
					pc_processDate=(day(pc_processDate)&"/"&month(pc_processDate)&"/"&year(pc_processDate))
				end if
			end if
		end if
		set rsStr=nothing
		
		call closedb()
	end if
	
	pc_VATAmount=rs("ord_VAT")
	pc_GCDetails =rs("pcOrd_GCDetails")

	querySTR="SELECT * FROM ProductsOrdered WHERE idorder="& pc_idorder
	Set rsSTR=CreateObject("ADODB.Recordset")
	rsSTR.CursorLocation=adUseClient
	rsSTR.Open querySTR, scDSN , 3, 3
	bOrderTotal=0
	do until rsSTR.eof
		querySTR="SELECT name,lastname,IDRefer FROM customers WHERE idcustomer="& pc_idcustomer
		Set rsCust=CreateObject("ADODB.Recordset")
		rsCust.CursorLocation=adUseClient
		rsCust.Open querySTR, scDSN , 3, 3
		CustName=rsCust("name")& " "&rsCust("lastname")
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
	gTotalVAT=gTotalVAT + pc_VATAmount
	pc_Discounts=0
	pc_discountDetails=""
	pc_DCount=0
	if (discountDetails<>"") and (Instr(discountDetails,"||")>0) then
	A=split(discountDetails,",")
	k=lbound(A)-1
	do
		k=k+1
		tmp1=""
		if Instr(discountDetails,"||")>0 then
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
    
	<%'start of gift certificates
	pc_GCString = ""
    if pc_GCDetails<>"" then
        GCArry=split(pc_GCDetails,"|g|")
        intArryCnt=ubound(GCArry)
    
        for k=0 to intArryCnt
			if GCArry(k)<>"" then
				GCInfo = split(GCArry(k),"|s|")
				if GCInfo(2)="" OR IsNull(GCInfo(2)) then
					GCInfo(2)=0
				end if
				pc_GCString = pc_GCString &"Gift Code: "&GCInfo(0)&"&nbsp;"
				if Cdbl(GCInfo(2)) <> 0 then
				pc_GCString = pc_GCString &"-"&scCurSign&money(GCInfo(2))
				end if
				pc_GCString = pc_GCString &"<br>"
			end if
        Next
    end if
	%>
        
	<tr>  
		<td nowrap valign="top"><a href="orddetails.asp?id=<%=pc_idorder%>" target="_blank" title="View details for this order"><%=(scpre + int(pc_idorder))%></a></td>
		<td nowrap valign="top"><%=pc_orderDate%></td>
		<td nowrap valign="top"><a href="viewCustOrders.asp?idcustomer=<%=pc_idcustomer%>" target="_blank" title="View all orders by this customer"><%=CustName%></a></td>
		<td nowrap valign="top"><%=Referrer%></td>
		<td nowrap valign="top"><%=scCurSign&money(pc_total)%></td>
		<td nowrap valign="top"><%if ptaxsetup=1 AND ptaxVAT="1" then%><%=scCurSign&money(pc_VATAmount)%><%else%><%=scCurSign&money(pc_taxAmount)%><%end if%></td>
        <% if request("onlyShow")="onlyGC" then %>
       		<td nowrap valign="top"><%=pc_GCString%></td>
        <% else %>
            <td nowrap valign="top"><%=pc_discountDetails%></td>
        <% end if %>
		<td nowrap valign="top"><%=pc_processDate %></td>
	</tr>
        
	<% gTotalcomm=gTotalcomm + affiliatepay %>
	<% rs.MoveNext
	loop %>
        
<tr>         
	<td colspan="8" class="pcCPspacer"></td>
</tr>
<tr bgcolor="#e1e1e1">
	<td colspan="4" align="right"><strong>Totals</strong></td>
	<td><b><%=scCurSign&money(gTotalsales)%></font></b></td>
	<td><b><%if ptaxsetup=1 AND ptaxVAT="1" then%><%=scCurSign&money(gTotalVAT)%><%else%><%=scCurSign&money(gTotaltaxes)%><%end if%></font></b></td>
	<td colspan="2">Shipping: <b><%=scCurSign&money(gTotalshipfees)%></font></b></td>
</tr>
<tr>
	<td colspan="4" align="right">Product-Only Sales:</td>
	<td colspan="4">
		<%
		ProductSales=gTotalsales-gTotalshipfees-gTotalhandfees-gTotalpayfees-gTotaltaxes
		%>
		<strong><%=scCurSign&money(ProductSales)%></strong>
		&nbsp;(Total sales minus shipping, payment and other charges)
	</td>
</tr>
<%if ptaxsetup=1 AND ptaxVAT="1" then%>
<tr>
	<td colspan="4" align="right">&nbsp;</td>
	<td colspan="4">
		<i>VAT is included in the order total</i>
	</td>
</tr>
<%end if%>
<tr>         
	<td colspan="8" class="pcCPspacer"></td>
</tr>

<% if showDetailedReports <> "1" Then %>

<tr bgcolor="#e1e1e1">
	<td colspan="4" align="right"><strong>Other Sales Statistics</strong></td>
	<td colspan="4">&nbsp;</td>
</tr>

<tr>
	<td colspan="4" align="right">Average Order Amount:</td>
	<td><%=scCurSign&money(gTotalsales/gTotalNumberOrders)%></td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="4" align="right">Average Products Ordered Amount:</td>
	<td><%=scCurSign&money(ProductSales/gTotalNumberOrders)%></td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="4" align="right">% of Sales Attributable to Shipping Charges and Handling Fees:</td>
	<td><%=int(((gTotalshipfees + gTotalhandfees)/gTotalsales)*100)%>%</td>
	<td colspan="3">&nbsp;</td>
</tr>

<%
queryDis="SELECT DISTINCT orders.idorder,idcustomer,total,shipmentDetails,paymentDetails,taxAmount,discountDetails,orderDate,processDate FROM Orders WHERE orders.discountDetails <> 'No discounts applied.' AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & tmpCond & " ORDER BY " & tmpD & " DESC;"

' Our Recordset Object
Dim rsDis
Set rsDis=CreateObject("ADODB.Recordset")
rsDis.CursorLocation=adUseClient
rsDis.Open queryDis, scDSN , 3, 3

gTotalOrdersDiscounts=rsDis.RecordCount
gTotalsalesDiscounts=0
do until rsDis.EOF
	pc_total=rsDis("total")
	gTotalsalesDiscounts=gTotalsalesDiscounts + pc_total

rsDis.MoveNext
loop
set rsDis = nothing

if gTotalOrdersDiscounts <> "0" Then
%>

<tr>
	<td colspan="4" align="right">Number of Orders Where a Coupon/Discount was used:</td>
	<td><%=gTotalOrdersDiscounts%></td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="4" align="right">Percentage of Total Orders:</td>
	<td><%=int((gTotalOrdersDiscounts/gTotalNumberOrders)*100)%>%</td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="4" align="right">Total Sales from Orders Using a Coupon/Discount:</td>
	<td><%=scCurSign&money(gTotalsalesDiscounts)%></td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="4" align="right">Percentage of Total Sales:</td>
	<td><%=int((gTotalsalesDiscounts/gTotalsales)*100)%>%</td>
	<td colspan="3">&nbsp;</td>
</tr>
<% end if 
end if ' ShowDetailedReports
End If %>

      
</table>
<%	' Done. Now release Objects
	con.Close
	Set con=Nothing
	Set rs=Nothing
%>
</div>
</body>
</html>