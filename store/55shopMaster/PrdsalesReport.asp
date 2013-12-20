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
	<title>Online Sales Report - Sales by Product</title>
	<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:10px">
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

'SM-S
tmpCond=""
tmpCond1=""
pcSaleID=""
pcSaleName=""
IF Ucase(scDB)="SQL" THEN
	IF request("saleID")<>"" AND request("saleID")<>"0" then
		pcSaleID=request("saleID")
		tmpCond=" AND (Orders.idOrder IN (SELECT DISTINCT ProductsOrdered.idOrder FROM ProductsOrdered INNER JOIN pcSales_Completed ON ProductsOrdered.pcSC_ID=pcSales_Completed.pcSC_ID WHERE pcSales_Completed.pcSales_ID=" & pcSaleID & ")) "
		tmpCond1=" AND (ProductsOrdered.pcSC_ID IN (SELECT DISTINCT pcSales_Completed.pcSC_ID FROM pcSales_Completed WHERE pcSales_Completed.pcSales_ID=" & pcSaleID & ")) "
		call opendb()
		queryS="SELECT pcSales_Name FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
		set rsS=connTemp.execute(queryS)
		if not rsS.eof then
			pcSaleName=rsS("pcSales_Name")
		end if
		set rsS=nothing
	End if
END IF
'SM-E

IF pIDProduct="" THEN

query="SELECT ProductsOrdered.IDproduct,SUM(ProductsOrdered.quantity) as TotalQty FROM ProductsOrdered,Orders WHERE " &SQLCC & " orders.idorder=ProductsOrdered.idorder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & tmpCond1 & " GROUP BY ProductsOrdered.IDProduct ORDER BY sum(ProductsOrdered.quantity) DESC;"

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
<% Else
 %>
		<tr> 
			<td colspan="6"><h2>Total Products Sales recorded from: <%=strTDateVar%> to: <%=strTDateVar2%><%if pcSaleName<>"" then%><br>The Sale Name: <%=pcSaleName%><%end if%></h2></td>
		</tr>
		<tr>
			<th nowrap>SKU (ID)</th> 
			<th nowrap>Product Name</th>
			<th nowrap>Units Sold</th>
			<th nowrap>Amount Sold</th>
			<th nowrap>Cost of Goods</th>
			<th nowrap>Margin</th>
		</tr>
		<tr>
			<td colspan="6" class="pcCPspacer"></td>
		</tr>
			<% 
			do until rs.EOF
			aIDProduct=rs("idproduct")
			
			query="SELECT ProductsOrdered.IDproduct,ProductsOrdered.quantity,ProductsOrdered.unitPrice,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts,Products.Description,Products.cost,Products.sku FROM ProductsOrdered,Orders,products WHERE ProductsOrdered.idproduct=" & aIDProduct & " and products.idproduct=ProductsOrdered.idproduct and orders.idorder=ProductsOrdered.idorder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & tmpCond1 & " ORDER BY ProductsOrdered.idproduct asc;"
			Dim rs1
			Set rs1=CreateObject("ADODB.Recordset")
			rs1.CursorLocation=adUseClient
			rs1.Open query, scDSN , 3, 3
			
			gTotalNumberOrders=rs1.RecordCount
			tmpID=rs1("idproduct")
			gTotalUnit=cdbl(rs1("quantity"))
			gTotalAmount=cdbl(rs1("unitPrice"))*gTotalUnit-cdbl(rs1("QDiscounts"))-cdbl(rs1("ItemsDiscounts"))
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
					gTotalAmount=gTotalAmount+(cdbl(rs1("unitPrice"))*cdbl(rs1("quantity"))-cdbl(rs1("QDiscounts"))-cdbl(rs1("ItemsDiscounts")))
				end if
			loop
				%>
				<tr>  
					<td nowrap><%=gSku%> (<%=tmpID%>)</td>
					<td nowrap><a href="FindProductType.asp?id=<%=tmpID%>" target="_blank"><%=gName%></a></td>
					<td nowrap><%=gTotalUnit%></td>
					<td nowrap><%=scCurSign&money(gTotalAmount)%></td>
					<td nowrap><%if pcv_Cost>0 then%><%=scCurSign&money(pcv_Cost*gTotalUnit)%><%end if%></td>
					<td nowrap><%if pcv_Cost>0 then%><%=scCurSign&money(gTotalAmount-(pcv_Cost*gTotalUnit)) & " (" & money(100*((gTotalAmount-(pcv_Cost*gTotalUnit))/gTotalAmount)) & "%)" %><%end if%></td>
				</tr>
				<% 
				rs.MoveNext
			loop
			%>
		</table>
<%' Done. Now release Objects
		con.Close
		Set con=Nothing
		Set rs=Nothing
		End if
		
	ELSE 'View Report for only 1 product%>
	
		<table class="pcCPcontent">
	
		<%
		query="SELECT orders.idorder,orders.idcustomer,orders.total,orders.shipmentDetails,orders.paymentDetails,orders.taxAmount,orders.discountDetails,orders.orderDate,orders.processDate,orders.shipDate,ProductsOrdered.quantity,ProductsOrdered.unitPrice,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts FROM Orders,ProductsOrdered WHERE ProductsOrdered.IDProduct=" & pIDProduct & " and orders.idorder=ProductsOrdered.idorder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & SQLCC1 & tmpCond1 & " ORDER BY " & tmpD & " DESC;"

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
		<td colspan="10"><h2>Total Sales recorded from: <%=strTDateVar%> to: <%=strTDateVar2%> <%if PrdName<>"" then%>
&nbsp;for <%=PrdName%><%end if%><%if pcSaleName<>"" then%><br>The Sale Name: <%=pcSaleName%><%end if%></h2></td>
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
		<th nowrap>Payment Details</th>
		<th nowrap>Referrer</th>
		<th nowrap>Qty. Ordered</th>
		<th nowrap>Amount Ordered</th>
		<th nowrap>Cost</th>
		<th nowrap>Margin</th>
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
	
	PrdQty=rs("quantity")
	PrdAmount=cdbl(rs("quantity"))*cdbl(rs("unitPrice"))-rs("QDiscounts")-cdbl(rs("ItemsDiscounts"))
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
		qdiscounts=rsSTR("QDiscounts")
		itemsdiscounts=rsSTR("ItemsDiscounts")
		bOrderTotal=0 + (unitTotal * quantity - qdiscounts - itemsdiscounts)
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
	on error goto 0
	
	gTotalpayfees=gTotalpayfees+PayCharge
	gTotaltaxes=gTotaltaxes + pc_taxAmount
	discountDetails=replace(pc_discountDetails,"|| ",scCurSign)
	%>
        
	<tr valign="top">  
		<td nowrap><a href="orddetails.asp?id=<%=pc_idorder%>" target="_blank" title="View details for this order"><%=(scpre + int(pc_idorder))%></a></td>
		<td nowrap><%=pc_orderDate%></td>
		<td nowrap><a href="viewCustOrders.asp?idcustomer=<%=pc_idcustomer%>" target="_blank" title="View all orders by this customer"><%=CustName%></a></td>
		<td nowrap>
		<%query="SELECT paymentDetails,gwAuthCode,gwTransId,paymentCode FROM Orders WHERE idorder=" & pc_idorder & ";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			pcv_paymentDetails=rsQ("paymentDetails")
			pcv_gwAuthCode=rsQ("gwAuthCode")
			pcv_gwTransId=rsQ("gwTransId")
			pcv_paymentCode=rsQ("paymentCode")
			if instr(pcv_paymentDetails,"||")>0 then
				pcv_PayArray=split(pcv_paymentDetails,"||")
				Response.Write("Payment Method: " & trim(pcv_PayArray(0)) & "<br>")
				if trim(pcv_PayArray(1))<>"" then
				if IsNumeric(trim(pcv_PayArray(1))) then
					PayFees=cdbl(trim(pcv_PayArray(1)))
					if PayFees>0 then
						Response.Write("Fees: " & scCurSign & money(PayFees) & "<br>")
					end if
				end if
				end if
			else
				Response.Write("Payment Details: " & pcv_paymentDetails & "<br>")
			end if
			if pcv_paymentCode<>"" then
			Response.Write("Payment Gateway: " & pcv_paymentCode & "<br>")
			end if
			if pcv_gwTransId<>"" then
			Response.Write("Transaction ID: " & pcv_gwTransId & "<br>")
			end if
			if pcv_gwAuthCode<>"" then
			Response.Write("Authorization Code: " & pcv_gwAuthCode & "<br>")
			end if
		end if
		set rsQ=nothing
		%>
		<%query="SELECT strRuleName,strFormValue FROM customCardOrders WHERE idorder=" & pc_idorder & ";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			do while not rsQ.eof
				response.write rsQ("strRuleName") & ": " & rsQ("strFormValue") & "<br>"
				rsQ.MoveNext
			loop
		end if
		set rsQ=nothing%>
		</td>
		<td nowrap><%=Referrer%></td>
		<td nowrap><%=PrdQty%></td>
		<td nowrap><%=scCurSign&money(PrdAmount)%></td>
		<td nowrap><%if pcv_Cost>0 then%><%=scCurSign&money(pcv_Cost*PrdQty)%><%end if%></td>
		<td nowrap><%if (pcv_Cost>0) AND (Cdbl(PrdAmount)>0) AND (Cdbl(PrdAmount)>Cdbl(pcv_Cost)) then%><%=scCurSign&money(PrdAmount-(pcv_Cost*PrdQty)) & " (" & money(100*((PrdAmount-(pcv_Cost*PrdQty))/PrdAmount)) & "%)" %><%end if%></td>
		<td nowrap><%=pc_processDate%></td>
	</tr>
        
	<% gTotalcomm=gTotalcomm + affiliatepay %>
	<% rs.MoveNext
	loop %>
        
	<tr>         
		<td colspan="10">&nbsp;</td>
	</tr>
	<tr bgcolor="#e1e1e1">
		<td colspan="4"><strong>Totals</strong></td>
		<td><b><%=gTotalNumberOrders%></b></td>
		<td><b><%=gTotalQty%></b></td>
		<td><b><%=scCurSign&money(gTotalPrdAmount)%></b></td>
		<td><b><%if pcv_Cost>0 then%><%=scCurSign&money(pcv_Cost*gTotalQty)%><%end if%></b></td>
		<td><b><%if (pcv_Cost>0) AND (Cdbl(gTotalPrdAmount)>0) AND (Cdbl(gTotalPrdAmount)>Cdbl(pcv_Cost)) then%><%=scCurSign&money(gTotalPrdAmount-(pcv_Cost*gTotalQty)) & " (" & money(100*((gTotalPrdAmount-(pcv_Cost*gTotalQty))/gTotalPrdAmount)) & "%)" %><%end if%></b></td>
		<td></td>
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