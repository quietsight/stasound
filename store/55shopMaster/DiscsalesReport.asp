<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image: none;">
<div id="pcCPmain">

<table class="pcCPcontent">
        
<%Dim connTemp,con

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

IDDiscount=request("IDDiscount")
call opendb()
	if validNum(IDDiscount) then
		mySQL="SELECT iddiscount, DiscountDesc, DiscountCode FROM Discounts WHERE IDDiscount=" & IDDiscount
	else
		mySQL="SELECT iddiscount, DiscountDesc, DiscountCode FROM Discounts order by IDDiscount desc"
	end if	
	set rs1=Server.CreateObject("ADODB.Recordset")
	set rs1=conntemp.execute(mySQL)

pcv_NoDisc=1

do while not rs1.eof
	IDDiscountTemp = rs1("iddiscount")
	DiscountName=rs1("DiscountDesc")
	DiscountCode=rs1("DiscountCode")

	query="SELECT orders.idorder,orders.idcustomer,orders.total,orders.rmaCredit,orders.pcOrd_CatDiscounts,orders.iRewardValue,orders.shipmentDetails,orders.paymentDetails,orders.taxAmount,orders.discountDetails,orders.orderDate,orders.processDate,orders.shipDate,orders.ord_VAT FROM Orders WHERE orders.DiscountDetails like '%" & replace(DiscountName, "'", "''") & "%' and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & TempSpecial & " ORDER BY " & tmpD & " DESC;"

	' Our Recordset Object
	Dim rs
	Set rs=CreateObject("ADODB.Recordset")
	rs.CursorLocation=adUseClient
	rs.Open query, scDSN , 3, 3
	
	' If the returning recordset is not empty
	If (rs.EOF) then
		If (IDDiscount<>"") Then 
		pcv_NoDisc=0%>
			<tr> 
				<td colspan="8">No records match your query</td>
			</tr>
		<%
		End if
	Else
	pcv_NoDisc=0
		if strTDateVar="" then strTDateVar="N/A"
		if strTDateVar2="" then strTDateVar2="N/A"
	%>
	<tr> 
		<td colspan="8">Total Sales recorded from: <%=strTDateVar%> to: <%=strTDateVar2%> <%if DiscountName<>"" then%><br>Discount code: <a href="modDiscounts.asp?mode=Edit&iddiscount=<%=IDDiscountTemp%>" target="_blank" title="View discout details"><strong><%=DiscountName%></strong></a> - <%=DiscountCode%><%end if%></td>
</tr>
<tr> 
	<td colspan="8"><% Response.Write "Total Records Found : " & rs.RecordCount & "<br><br>" %></td>
</tr>
<tr> 
	<th nowrap>Order #</th>
	<th nowrap>Order Date</th>
	<th nowrap>Customer Name</th>
	<th nowrap>Referrer</th>
	<th nowrap>Order Total</th>
	<th nowrap><%if ptaxsetup=1 AND ptaxVAT="1" then%>VAT<%else%>Tax<%end if%></b></th>
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
gSubDiscTotal=0
pc_TotalDisc=0
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
		if Instr(tmp1,"|| ")>0 then
			B=split(tmp1,"|| ")
		else
			B=split(tmp1,"||")
		end if
		pc_Discounts=pc_Discounts+cdbl(B(1))
		if Instr(tmp1,DiscountName)>0 then
		pc_TotalDisc=pc_TotalDisc+cdbl(B(1))
		end if
		if Instr(tmp1,"|| ")>0 then
			tmp1=replace(tmp1,"|| ",scCurSign)
		else
			tmp1=replace(tmp1,"||",scCurSign)
		end if
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

	gSubDiscTotal=gSubDiscTotal+pc_Discounts
	%>
        
	<tr style="font-size: 12px;">  
		<td nowrap valign="top"><a href="orddetails.asp?id=<%=pc_idorder%>" target="_blank" title="View details for this order"><%=(scpre + int(pc_idorder))%></a></td>
		<td nowrap valign="top"><%=pc_orderDate%></td>
		<td nowrap valign="top"><%=CustName%></td>
		<td nowrap valign="top"><%=Referrer%></td>
		<td nowrap valign="top"><%=scCurSign&money(pc_total)%></td>
		<td nowrap valign="top"><%if ptaxsetup=1 AND ptaxVAT="1" then%><%=scCurSign&money(pc_VATAmount)%><%else%><%=scCurSign&money(pc_taxAmount)%><%end if%></td>
		<td valign="top"><%=pc_discountDetails%></td>
		<td nowrap valign="top"><%=pc_processDate %></td>
	</tr>
        
	<% gTotalcomm=gTotalcomm + affiliatepay %>
	<% rs.MoveNext
	loop %>
        
<tr>         
	<td colspan="8">&nbsp;</td>
</tr>
</table>

<table class="pcCPcontent">
<tr>
	<th nowrap>Total Orders</th>
	<th nowrap align="right">Total Sales</th>
	<th align="right">Total Shipping Charges</th>
	<th nowrap align="right">Total Taxes</th>
	<th align="right" nowrap>Total Discounts</th>
	<th align="right" nowrap>Total Discounts by this code</th>
	<th align="right">Product Sales</th>
</tr>  
<tr style="font-size: 12px;">
	<td><%=gTotalNumberOrders%></td>
	<td align="right"><%=scCurSign&money(gTotalsales)%></td>
	<td align="right"><%=scCurSign&money(gTotalshipfees)%></td>
	<td align="right"><%if ptaxsetup=1 AND ptaxVAT="1" then%><%=scCurSign&money(pc_VATAmount)%><%else%><%=scCurSign&money(gTotaltaxes)%><%end if%></td>
	<td align="right"><%="-" & scCurSign&money(gSubDiscTotal)%></td>
	<td align="right"><%="-" & scCurSign&money(pc_TotalDisc)%></td>
	<td align="right"><%ProductSales=gTotalsales-gTotalshipfees-gTotalhandfees-gTotalpayfees-gTotaltaxes%><%=scCurSign&money(ProductSales)%></td>
</tr>
<%if ptaxsetup=1 AND ptaxVAT="1" then%>
<tr>         
	<td colspan="7" align="right" style="font-size: 12px;"><i>Note: VAT is included in the order total</i></td>
</tr>
<%end if%>
<tr>         
	<td colspan="7">&nbsp;</td>
</tr>
<%end if
rs1.MoveNext
loop%>
<%if pcv_NoDisc=1 then%>
<tr> 
	<td colspan="7">No records match your query</td>
</tr>
<%end if%>
</TABLE>
</div>
</body>
</html>