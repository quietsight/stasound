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
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<html>
<head>
	<title>Online Sales Report - Sales by Brand</title>
	<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin: 10px;">
<div id="pcCPmain" style="background-image: none;">
<table class="pcCPcontent">
        
<%
' Our Connection Object
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
TempSQL3=""
Select case tmpDate
Case "2": tmpD="orders.processDate"
tmpD1="processDate"
tmpD2="Processed On"
Case "3": tmpD="pcPackageInfo.pcPackageInfo_ShippedDate"
tmpD1="pcPackageInfo_ShippedDate"
tmpD2="Shipped On"
TempSQL3=",pcPackageInfo"
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

call opendb()

pIDBrand=request("IDBrand")
if pIDBrand="" then
	response.redirect "menu.asp"
end if

query="SELECT BrandName FROM Brands WHERE idbrand=" & pIDBrand
set rs1=conntemp.execute(query)
BrandName=rs1("BrandName")
set rs1=nothing

query="SELECT DISTINCT orders.idorder,orders.idcustomer,orders.total,orders.shipmentDetails,orders.paymentDetails,orders.taxAmount,orders.discountDetails,orders.orderDate,orders.processDate,orders.shipDate FROM (ProductsOrdered INNER JOIN Orders ON ProductsOrdered.idOrder=Orders.idOrder) INNER JOIN Products ON ProductsOrdered.idProduct=Products.idProduct WHERE Products.IDBrand=" & pIDBrand & " AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & " ORDER BY " & tmpD & " DESC;"

' Our Recordset Object
Set rs=CreateObject("ADODB.Recordset")
rs.CursorLocation=adUseClient
rs.Open query, scDSN , 3, 3
			
' If the returning recordset is not empty
If rs.EOF Then %>
	<tr> 
		<td colspan="10">No records match your query</td>
	</tr>
<% Else %>
	<tr> 
		<td colspan="10"><h2>Total Sales recorded from: <%=strTDateVar%> to: <%=strTDateVar2%> <%if BrandName<>"" then%>&nbsp;for <%=BrandName%><%end if%></h2></td>
	</tr>
	<tr> 
		<td colspan="10"><% Response.Write "Total Records Found : " & rs.RecordCount %></td>
	</tr>
	<tr> 
		<td colspan="10" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th nowrap>Order #</th>
		<th nowrap>Date</th>
		<th nowrap>Customer Name</th>
		<th nowrap>Order Details</th>
		<th nowrap>Qty. Ordered</th>
		<th nowrap>Amount Ordered</th>
		<th nowrap><%=tmpD2%></th>
	</tr>
	<tr>
		<td colspan="10" class="pcCPspacer"></td>
	</tr>
<% 
gTotalNumberOrders=rs.RecordCount
gTotalsales=0
gTotaltaxes=0
gTotalcom=0
gTotalshipfees=0
gTotalhandfees=0
gTotalpayfees=0
gTotalQty=0
gTotalPrdAmount=0
do until rs.EOF
	pc_idorder=rs("idorder")
	pc_idcustomer=rs("idcustomer")
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
	
	tmpOrdDetails=""
	tmpQty=0
	tmpAmount=0
	
	query="SELECT Products.idproduct,ProductsOrdered.quantity,(ProductsOrdered.unitPrice*ProductsOrdered.quantity) AS SubPTotal,Products.Description  FROM (ProductsOrdered INNER JOIN Orders ON ProductsOrdered.idOrder=Orders.idOrder) INNER JOIN Products ON ProductsOrdered.idProduct=Products.idProduct WHERE Products.IDBrand=" & pIDBrand & " and Orders.idOrder=" &  pc_idorder & " AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & " ORDER BY Products.Description ASC;"
	Set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		pcArr=rsQ.getRows()
		set rsQ=nothing
		intCount1=ubound(pcArr,2)
		for i=0 to intCount1
			if tmpOrdDetails<>"" then
				tmpOrdDetails=tmpOrdDetails & "<br>"
			end if
			tmpOrdDetails=tmpOrdDetails & pcArr(3,i) & " - Qty: " & pcArr(1,i) & " - Amount: " & scCurSign&money(pcArr(2,i))
			gTotalQty=gTotalQty+clng(pcArr(1,i))
			gTotalPrdAmount=gTotalPrdAmount+cdbl(pcArr(2,i))
			tmpQty=tmpQty+clng(pcArr(1,i))
			tmpAmount=tmpAmount+cdbl(pcArr(2,i))
		Next
	end if
	set rsQ=nothing
	%>
        
	<tr valign="top">  
		<td nowrap><a href="orddetails.asp?id=<%=pc_idorder%>" target="_blank" title="View details for this order"><%=(scpre + int(pc_idorder))%></a></td>
		<td nowrap><%=pc_orderDate%></td>
		<td nowrap><a href="viewCustOrders.asp?idcustomer=<%=pc_idcustomer%>" target="_blank" title="View all orders by this customer"><%=CustName%></a></td>
		<td><%=tmpOrdDetails%></td>
		<td nowrap align="right"><%=tmpQty%></td>
		<td nowrap align="right"><%=scCurSign&money(tmpAmount)%></td>
		<td nowrap align="right"><%=pc_processDate%></td>
	</tr>
        
	<% rs.MoveNext
	loop %>
        
	<tr>         
		<td colspan="10">&nbsp;</td>
	</tr>
	<tr bgcolor="#e1e1e1">
		<td colspan="3"><strong>Totals</strong></td>
		<td align="right">Orders: <b><%=gTotalNumberOrders%></b></td>
		<td align="right"><b><%=gTotalQty%></b></td>
		<td align="right"><b><%=scCurSign&money(gTotalPrdAmount)%></b></td>
		<td></td>
	</tr>

<%End If %>

</table>
<%	' Done. Now release Objects
	con.Close
	Set con=Nothing
	Set rs=Nothing
%>
</div>
</body>
</html>