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
<title>Sales Report by Pricing Category</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:10px;">
<div id="pcCPmain" style="background-image: none;">
<table class="pcCPcontent">
        
<%
Dim connTemp,rs,rs1
	
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

CustomerTypeName=""
HaveCCType=0

if SQLCC="" then
	SQLCC=" ,customers "
end if
if SQLCC1="" then
	SQLCC1=" AND customers.idcustomer=orders.idcustomer "
end if

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
		if HaveCCType=1 then
			SQLCC1=SQLCC1 & " AND customers.idCustomerCategory=" & tmpA(1) & " "
		else
			SQLCC1=SQLCC1 & " AND customers.customerType=" & CustomerType & " AND customers.idCustomerCategory=0 "
		end if
	end if
end if

call opendb()
query="SELECT count(*) as OrderCount,sum(total) as GrandTotal,customers.idCustomerCategory,customers.customerType FROM Orders" & SQLCC & " WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & TempSQL1 & TempSQL2 & SQLCC1 & TempSpecial & " GROUP BY customers.idCustomerCategory,customers.customerType ORDER BY customers.idCustomerCategory DESC;"
set rs=connTemp.execute(query)
	
' If the returning recordset is not empty
If rs.EOF Then %>
	<tr> 
		<td colspan="4">No records match your query</td>
	</tr>
<% Else %>
	<tr> 
		<td colspan="4"><h2>Sales report by Pricing Category from <%=strTDateVar%> to <%=strTDateVar2%>
		<%if CustomerTypeName<>"" then%><br>Customer Type: <%=CustomerTypeName%><%end if%></h2></td>
</tr>
<tr> 
	<th nowrap><b>Category</b></th>
	<th nowrap><b>Number of Orders</b></th>
	<th nowrap><b>Total Sales</b></th>
	<th nowrap><b>Percentage of Total Sales</b></th>
</tr>
<tr>
	<td colspan="4" class="pcCPspacer"></td>
</tr>
<%TotalOrders=0
TotalSales=0
do while not rs.eof
	TotalOrders=TotalOrders+clng(rs("OrderCount"))
	TotalSales=TotalSales+cdbl(rs("GrandTotal"))
	rs.MoveNext
loop

rs.MoveFirst

do while not rs.eof
	pcv_Orders=clng(rs("OrderCount"))
	pcv_Sales=cdbl(rs("GrandTotal"))
	pcv_PricingCat=rs("idCustomerCategory")
	pcv_CustType=rs("customerType")
	pcv_PricingName=""
	if pcv_PricingCat="0" then
		if pcv_CustType="0" then
			pcv_PricingName="Retail"
		else
			pcv_PricingName="Wholesale"
        end if
	else
		query="SELECT pcCC_Name FROM pcCustomerCategories WHERE idCustomerCategory=" & pcv_PricingCat & ";"
		set rs1=connTemp.execute(query)
		if not rs1.eof then
			pcv_PricingName=rs1("pcCC_Name")
			set rs1=nothing
		end if
		set rs1=nothing
	end if%>
	<tr>  
		<td nowrap><%=pcv_PricingName%></td>
		<td nowrap align="right"><%=pcv_Orders%></td>
		<td nowrap align="right"><%=scCurSign&money(pcv_Sales)%></td>
		<td nowrap align="right"><%=Round((pcv_Sales/TotalSales)*100,1)%>%</td>
	</tr>
	<% rs.MoveNext
	loop %>
        
<tr>         
	<td colspan="4" class="pcCPspacer"></td>
</tr>
<tr bgcolor="#e1e1e1">
	<td align="right"><strong>Totals</strong></td>
	<td align="right"><b><%=TotalOrders%></b></td>
	<td align="right"><b><%=scCurSign&money(TotalSales)%></font></b></td>
	<td>&nbsp;</td>
</tr>
<%END IF%>

      
</table>
<%	' Done. Now release Objects
	Set rs=Nothing
	call closedb()
%>
</div>
</body>
</html>