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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image: none;">

<TABLE width="94%" border="0" align="center" cellpadding="4" cellspacing="0">
        
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

IdRefer=request("IdRefer")

call opendb()

mySQL="SELECT Name,IdRefer FROM Referrer WHERE IdRefer=" & IdRefer

set rs1=conntemp.execute(mySQL)	

do while not rs1.eof
	Referrer=rs1("Name")

query="SELECT orders.idorder,orders.idcustomer,orders.total,orders.shipmentDetails,orders.paymentDetails,orders.taxAmount,orders.discountDetails,orders.orderDate,orders.processDate,orders.shipDate,orders.ord_VAT FROM Orders INNER JOIN customers ON (orders.idcustomer=customers.idcustomer) WHERE (orders.IDRefer =" & IDRefer & " OR customers.IDRefer =" & IDRefer & ") AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & " ORDER BY " & tmpD & " DESC;"

' Our Recordset Object
Dim rs
Set rs=CreateObject("ADODB.Recordset")
rs.CursorLocation=adUseClient
rs.Open query, scDSN , 3, 3
	
' If the returning recordset is not empty
If rs.EOF Then %>
	<tr class="normal"> 
		<td colspan="8"><font size="2">No records match your query</font></td>
	</tr>
<% Else %>
	<tr class="normal"> 
		<td colspan="8"><b><font size="2">Total Sales recorded from: <%=strTDateVar%> to: <%=strTDateVar2%> <%if Referrer<>"" then%><br>Referrer: <%=Referrer%></a><%end if%></font></b></td>
</tr>
<tr class="normal"> 
	<td colspan="8"><font size="2"><% Response.Write "Total Records Found : " & rs.RecordCount & "<br><br>" %></font></td>
</tr>
<tr class="normal"> 
	<td bgcolor="#e1e1e1" nowrap><b>Order #</b></td>
	<td bgcolor="#e1e1e1" nowrap><b>Date</b></td>
	<td bgcolor="#e1e1e1" nowrap><b>Customer Name</b></td>
	<td bgcolor="#e1e1e1" nowrap><b>Referrer</b></td>
	<td bgcolor="#e1e1e1" nowrap><b>Order Total</b></td>
	<td bgcolor="#e1e1e1" nowrap><b><%if ptaxsetup=1 AND ptaxVAT="1" then%>VAT<%else%>Tax<%end if%></b></td>
	<td bgcolor="#e1e1e1" nowrap><b>Discounts Applied</b></td>
	<td bgcolor="#e1e1e1" nowrap><b><%=tmpD2%></b></td>
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
gSubDiscTotal=0
do until rs.EOF
	pc_idorder=rs("idorder")
	pc_idcustomer=rs("idcustomer")
	pc_total=rs("total")
	pc_shipmentDetails=rs("shipmentDetails")
	pc_paymentDetails=rs("paymentDetails")
	pc_taxAmount=rs("taxAmount")
	pc_discountDetails=rs("discountDetails")
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
	end if
	
	pc_VATAmount=rs("ord_VAT")

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
	discountDetails=replace(pc_discountDetails,"|| ",scCurSign)
	%>
        
	<tr class="normal" valign="top">  
		<td nowrap><a href="orddetails.asp?id=<%=pc_idorder%>" target="_blank" title="View details for this order"><%=(scpre + int(pc_idorder))%></a></td>
		<td nowrap><%=pc_orderDate%></td>
		<td nowrap><%=CustName%></td>
		<td nowrap><%=Referrer%></td>
		<td nowrap><%=scCurSign&money(pc_total)%></td>
		<td nowrap><%if ptaxsetup=1 AND ptaxVAT="1" then%><%=scCurSign&money(pc_VATAmount)%><%else%><%=scCurSign&money(pc_taxAmount)%><%end if%></td>
		<td nowrap><%=discountDetails%></td>
		<td nowrap><%=pc_processDate %></td>
	</tr>
        
	<% gTotalcomm=gTotalcomm + affiliatepay %>
	<% rs.MoveNext
	loop %>
        
<tr class="normal">         
	<td colspan="8">&nbsp;</td>
</tr>
   
<tr class="normal">
	<td colspan="2" bgcolor="#e1e1e1">&nbsp;</td>
	<td bgcolor="#e1e1e1" nowrap><b>Total Orders</b></td>
	<td bgcolor="#e1e1e1" nowrap><b>Total Sales</b></td>
	<td bgcolor="#e1e1e1"><b>Total Shipping Charges</b></td>
	<td bgcolor="#e1e1e1" nowrap><b>Total Taxes</b></td>
	<td bgcolor="#e1e1e1"><b>Total Discounts</b></td>
	<td bgcolor="#e1e1e1"><b>Product Sales</b></td>
</tr>
        
<tr class="normal">
	<td colspan="2">&nbsp;</td>
	<td><b><font color="#FF0000"><%=gTotalNumberOrders%></font></b></td>
	<td><b><font color="#FF0000"><%=scCurSign&money(gTotalsales)%></font></b></td>
	<td><b><font color="#FF0000"><%=scCurSign&money(gTotalshipfees)%></font></b></td>
	<td><b><font color="#FF0000"><%if ptaxsetup=1 AND ptaxVAT="1" then%><%=scCurSign&money(gTotalVAT)%><%else%><%=scCurSign&money(gTotaltaxes)%><%end if%></font></b></td>
	<td><b><font color="#FF0000"><%="-" & scCurSign&money(gSubDiscTotal)%></font></b></td>
	<td><b><font color="#FF0000"><%
	ProductSales=gTotalsales-gTotalshipfees-gTotalhandfees-gTotalpayfees-gTotaltaxes
	%>
	<%=scCurSign&money(ProductSales)%></font></b></td>
</tr>
<%if ptaxsetup=1 AND ptaxVAT="1" then%>
<tr class="normal">         
	<td colspan="8" align="right"><i>Note: VAT is included in the order total</i></td>
</tr>
<%end if%>
<tr class="normal">         
	<td colspan="8">&nbsp;</td>
</tr>
<%end if
rs1.MoveNext
loop%>
</TABLE>
</body>
</html>