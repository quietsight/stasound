<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/opendb.asp"-->
<html>
<head>
	<title>Drop-Off Report - by Product</title>
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
	call opendb()

pIDProduct=request("IDProduct")
if pIDProduct<>"" then
	SQLCC=" ProductsOrdered.IDproduct=" & pidproduct & " and "
	mySQL="SELECT description,configOnly,serviceSpec,cost FROM products WHERE idproduct=" & pIDproduct
	set rs1=conntemp.execute(mySQL)
	PrdName=rs1("description")
	configOnly=rs1("configOnly")
	serviceSpec=rs1("serviceSpec")
	'Start SDBA
	pcv_Cost=rs1("cost")
	if IsNull(pcv_Cost) or pcv_Cost="" then
		pcv_Cost=0
	end if
	'End SDBA	
else
	SQLCC=""
end if

IF pIDProduct="" THEN

query="SELECT ProductsOrdered.IDproduct,SUM(ProductsOrdered.quantity) as TotalQty FROM ProductsOrdered,Orders WHERE " &SQLCC & " orders.idorder=ProductsOrdered.idorder and orders.orderStatus=1 " & TempSQL1 & TempSQL2 & " GROUP BY ProductsOrdered.IDProduct ORDER BY sum(ProductsOrdered.quantity) DESC;"

' Our Recordset Object
Dim rs
Set rs=CreateObject("ADODB.Recordset")
rs.CursorLocation=adUseClient
rs.Open query, scDSN , 3, 3
	
' If the returning recordset is not empty
If rs.EOF Then %>
	<tr> 
		<td colspan="4">No records match your query</td>
	</tr>
<% Else
 %>
		<tr> 
			<td colspan="4"><h2>Total Drop-Offs by Product recorded from: <%=strTDateVar%> to: <%=strTDateVar2%></h2></td>
		</tr>
		<tr>
			<th nowrap>SKU (ID)</th> 
			<th nowrap>Product Name</th>
			<th align="center" nowrap>Units Ordered</th>
			<th align="right" nowrap>Amount Ordered</th>
		</tr>
		<tr>
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
			<% 
			do until rs.EOF
				aIDProduct=rs("idproduct")
				
				query="SELECT ProductsOrdered.IDproduct,ProductsOrdered.quantity,ProductsOrdered.unitPrice,Products.Description,Products.cost,Products.sku FROM ProductsOrdered,Orders,products WHERE ProductsOrdered.idproduct=" & aIDProduct & " and products.idproduct=ProductsOrdered.idproduct and orders.idorder=ProductsOrdered.idorder and orders.orderStatus=1 " & TempSQL1 & TempSQL2 & " ORDER BY ProductsOrdered.idproduct asc;"
				Dim rs1
				Set rs1=CreateObject("ADODB.Recordset")
				rs1.CursorLocation=adUseClient
				rs1.Open query, scDSN , 3, 3
				
				gTotalNumberOrders=rs1.RecordCount
				tmpID=rs1("idproduct")
				gTotalUnit=cdbl(rs1("quantity"))
				gTotalAmount=cdbl(rs1("unitPrice"))*gTotalUnit
				gName=rs1("description")
				gSku=rs1("sku")
				
				'Start SDBA
				pcv_Cost=rs1("cost")
				if IsNull(pcv_Cost) or pcv_Cost="" then
					pcv_Cost=0
				end if
				'End SDBA
				
				do until rs1.EOF
					rs1.MoveNext
					if not rs1.eof then
						gTotalUnit=gTotalUnit+cdbl(rs1("quantity"))
						gTotalAmount=gTotalAmount+(cdbl(rs1("unitPrice"))*cdbl(rs1("quantity")))
					end if
				loop
					%>
					<tr>  
						<td nowrap><%=gSku%> (<%=tmpID%>)</td>
						<td nowrap><a href="FindProductType.asp?id=<%=tmpID%>" target="_blank"><%=gName%></a></td>
						<td align="center" nowrap><%=gTotalUnit%></td>
						<td align="right" nowrap><%=scCurSign&money(gTotalAmount)%></td>
					</tr>
					<% 
					rs.MoveNext
					gTotalQty=gTotalQty+gTotalUnit
					gTotalPrdAmount=gTotalPrdAmount+gTotalAmount
			loop
			%>
		
			<tr>         
				<td colspan="4" class="pcCPspacer"></td>
			</tr>
			<tr bgcolor="#e1e1e1">
				<td colspan="2"><strong>Totals</strong></td>
				<td align="center"><b><%=gTotalQty%></b></td>
				<td align="right"><b><%=scCurSign&money(gTotalPrdAmount)%></b></td>
			</tr>
		
		</table>
<%' Done. Now release Objects
		con.Close
		Set con=Nothing
		Set rs=Nothing
		End if
		
	ELSE 'View Report for only 1 product%>
	
		<table class="pcCPcontent">
	
		<%
		query="SELECT orders.idorder,orders.idcustomer,orders.total,orders.shipmentDetails,orders.paymentDetails,orders.taxAmount,orders.discountDetails,orders.orderDate,orders.processDate,orders.shipDate,ProductsOrdered.quantity,ProductsOrdered.unitPrice FROM Orders,ProductsOrdered WHERE ProductsOrdered.IDProduct=" & pIDProduct & " and orders.idorder=ProductsOrdered.idorder and orders.orderStatus=1 " & TempSQL1 & TempSQL2 & SQLCC1 & " ORDER BY " & tmpD & " DESC;"
	
		' Our Recordset Object
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
		<td colspan="6"><h2>Total Drop-Offs recorded from: <%=strTDateVar%> to: <%=strTDateVar2%> <%if PrdName<>"" then%>
&nbsp;for <%=PrdName%><%end if%></h2></td>
	</tr>
	<tr> 
		<td colspan="6"><% Response.Write "Total Records Found : " & rs.RecordCount %></td>
	</tr>
	<tr> 
		<td colspan="6" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th nowrap>Order #</th>
		<th nowrap>Date</th>
		<th nowrap>Customer Name</th>
		<th nowrap>Referrer</th>
		<th align="center" nowrap>Qty. Ordered</th>
		<th align="right" nowrap>Amount Ordered</th>
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
gTotalQty=0
gTotalPrdAmount=0
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
	pc_processDate=rs(tmpD1)
	if scDateFrmt="DD/MM/YY" then
		pc_processDate=(day(pc_processDate)&"/"&month(pc_processDate)&"/"&year(pc_processDate))
	end if
	PrdQty=rs("quantity")
	PrdAmount=rs("quantity")*rs("unitPrice")
	gTotalQty=gTotalQty+PrdQty
	gTotalPrdAmount=gTotalPrdAmount+PrdAmount

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
	discountDetails=replace(pc_discountDetails,"|| ",scCurSign)
	%>
        
	<tr>  
		<td nowrap><a href="orddetails.asp?id=<%=pc_idorder%>" target="_blank" title="View details for this order"><%=(scpre + int(pc_idorder))%></a></td>
		<td nowrap><%=pc_orderDate%></td>
		<td nowrap><a href="viewCustOrders.asp?idcustomer=<%=pc_idcustomer%>" target="_blank" title="View all orders by this customer"><%=CustName%></a></td>
		<td nowrap><%=Referrer%></td>
		<td align="center" nowrap><%=PrdQty%></td>
		<td align="right" nowrap><%=scCurSign&money(PrdAmount)%></td>
	</tr>
        
	<% gTotalcomm=gTotalcomm + affiliatepay %>
	<% rs.MoveNext
	loop %>
        
	<tr>         
		<td colspan="6" class="pcCPspacer"></td>
	</tr>
	<tr bgcolor="#e1e1e1">
		<td colspan="4"><strong>Totals</strong></td>
		<td align="center"><b><%=gTotalQty%></b></td>
		<td align="right"><b><%=scCurSign&money(gTotalPrdAmount)%></b></td>
	</tr>

<%End If %>

</table>
<%	' Done. Now release Objects
	con.Close
	Set con=Nothing
	Set rs=Nothing

	END IF
	'End View All Products
%>
</div>
</body>
</html>