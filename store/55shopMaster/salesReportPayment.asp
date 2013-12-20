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
<html>
<head>
<title>Sales Report by Payment Option</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:10px;">
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

if Request("PayType")<>"" then
	TempSQL1=TempSQL1 & " AND paymentDetails LIKE '%" & Request.QueryString("PayType") & "%'"
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

query="SELECT orders.idorder,orders.idcustomer,orders.total,orders.rmaCredit,orders.pcOrd_CatDiscounts,orders.iRewardValue,orders.shipmentDetails,orders.paymentDetails,orders.taxAmount,orders.discountDetails,orders.orderDate,orders.processDate,orders.shipDate,orders.ord_VAT FROM Orders" & SQLCC & " WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & SQLCC1 & TempSpecial & " ORDER BY " & tmpD & " DESC;"

' Our Recordset Object
Dim rs
Set rs=CreateObject("ADODB.Recordset")
rs.CursorLocation=adUseClient
rs.Open query, scDSN , 3, 3
	
' If the returning recordset is not empty
If rs.EOF Then %>
	<tr> 
		<td colspan="9">No records match your query</td>
	</tr>
<% Else %>
	<tr> 
		<td colspan="9">
			<h2>Total Sales from <%=strTDateVar%> to <%=strTDateVar2%> <%if Request("PayType")<>"" then%>&nbsp;using the payment type <%=Request("PayType")%><%end if%></h2>
		</td>
	</tr>
	<tr> 
		<td colspan="9"><% Response.Write "Total Records Found : " & rs.RecordCount %></td>
	</tr>
	<tr> 
		<th nowrap>Order #</th>
		<th nowrap>Date</th>
		<th nowrap>Customer Name</th>
		<th nowrap>Payment Details</th>
		<th nowrap>Referrer</th>
		<th nowrap>Order Total</th>
		<th nowrap><%if ptaxsetup=1 AND ptaxVAT="1" then%>VAT<%else%>Tax<%end if%></th>
		<th nowrap>Discounts Applied</th>
		<th nowrap><%=tmpD2%></th>
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
        
	<tr>  
		<td nowrap valign="top"><a href="orddetails.asp?id=<%=pc_idorder%>" target="_blank" title="View details for this order"><%=(scpre + int(pc_idorder))%></a></td>
		<td nowrap valign="top"><%=pc_orderDate%></td>
		<td nowrap valign="top"><a href="viewCustOrders.asp?idcustomer=<%=pc_idcustomer%>" target="_blank" title="View all orders by this customer"><%=CustName%></a></td>
		<td nowrap valign="top">
		<%call opendb()
		query="SELECT paymentDetails,gwAuthCode,gwTransId,paymentCode FROM Orders WHERE idorder=" & pc_idorder & ";"
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
		<td nowrap valign="top"><%=Referrer%></td>
		<td nowrap valign="top"><%=scCurSign&money(pc_total)%></td>
		<td nowrap valign="top"><%if ptaxsetup=1 AND ptaxVAT="1" then%><%=scCurSign&money(pc_VATAmount)%><%else%><%=scCurSign&money(pc_taxAmount)%><%end if%></td>
		<td nowrap valign="top"><%=pc_discountDetails%></td>
		<td nowrap valign="top"><%=pc_processDate %></td>
	</tr>
        
	<% gTotalcomm=gTotalcomm + affiliatepay %>
	<% rs.MoveNext
	loop %>
        
<tr>         
	<td colspan="9">&nbsp;</td>
</tr>
   
<tr bgcolor="#e1e1e1">
	<td colspan="5" align="right"><strong>Totals</strong></td>
	<td nowrap><b><%=scCurSign&money(gTotalsales)%></b></td>
	<td nowrap><b><%if ptaxsetup=1 AND ptaxVAT="1" then%><%=scCurSign&money(gTotalVAT)%><%else%><%=scCurSign&money(gTotaltaxes)%><%end if%></b></td>
	<td colspan="2">Shipping: <b><%=scCurSign&money(gTotalshipfees)%></b></td>
</tr>
<tr>
	<td colspan="5" align="right">Product-Only Sales</td>
	<td colspan="5"><%
	ProductSales=gTotalsales-gTotalshipfees-gTotalhandfees-gTotalpayfees-gTotaltaxes
	%>
	<strong><%=scCurSign&money(ProductSales)%></strong>
	(Sales minus shipping, payment, and other charges)
	</td>
</tr>
<%if ptaxsetup=1 AND ptaxVAT="1" then%>
<tr>
	<td colspan="5" align="right">&nbsp;</td>
	<td colspan="4">
		<i>VAT is included in the order total</i>
	</td>
</tr>
<%end if%>
<tr>         
	<td colspan="9" class="pcCPspacer"></td>
</tr>

<% if showDetailedReports <> "1" Then %>
<tr>
	<td colspan="9" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="5" align="right" bgcolor="#e1e1e1"><strong>Other Sales Statistics</strong></td>
	<td colspan="4" bgcolor="#e1e1e1">&nbsp;</td>
</tr>

<tr>
	<td colspan="5" align="right">Average Order Amount:</td>
	<td><%=scCurSign&money(gTotalsales/gTotalNumberOrders)%></td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="5" align="right">Average Products Ordered Amount:</td>
	<td><%=scCurSign&money(ProductSales/gTotalNumberOrders)%></td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="5" align="right">% of Sales Attributable to Shipping Charges and Handling Fees:</td>
	<td><%=int(((gTotalshipfees + gTotalhandfees)/gTotalsales)*100)%>%</td>
	<td colspan="3">&nbsp;</td>
</tr>

<%
queryDis="SELECT idorder,idcustomer,total,shipmentDetails,paymentDetails,taxAmount,discountDetails,orderDate,processDate,shipDate FROM Orders WHERE orders.discountDetails <> 'No discounts applied.' AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & " ORDER BY " & tmpD & " DESC;"

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
	<td colspan="5" align="right">Number of Orders Where a Coupon/Discount was used:</td>
	<td><%=gTotalOrdersDiscounts%></td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="5" align="right">Percentage of Total Orders:</td>
	<td><%=int((gTotalOrdersDiscounts/gTotalNumberOrders)*100)%>%</td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="5" align="right">Total Sales from Orders Using a Coupon/Discount:</td>
	<td><%=scCurSign&money(gTotalsalesDiscounts)%></td>
	<td colspan="3">&nbsp;</td>
</tr>

<tr>
	<td colspan="5" align="right">Percentage of Total Sales:</td>
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