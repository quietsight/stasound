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
<!--#include file="../includes/SQLFormat.txt" -->
<html>
<head>
<title>RMA Report</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:10px;">
<div id="pcCPmain" style="background-image: none;">
<table class="pcCPcontent">
        
<%
Dim connTemp,rs, RMAStatusCond

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

Select case tmpDate
	Case "2"
		tmpD="orders.processDate"
	Case "3"
		tmpD="pcPackageInfo.pcPackageInfo_ShippedDate"
	Case Else
		tmpD="orders.orderDate"
End Select

if (DateVar<>"") and IsDate(DateVar) then
    if SQL_Format="1" then DateVar=day(DateVar) & "/" & month(DateVar) & "/" & year(DateVar)
	if scDB="Access" then
		TempSQL1=" AND " & tmpD & " >=#" & DateVar & "# "
	else
		TempSQL1=" AND " & tmpD & " >='" & DateVar & "' "
	end if
else
	TempSQL1=""
end if

if (DateVar2<>"") and IsDate(DateVar2) then
    if SQL_Format="1" then DateVar2=day(DateVar2) & "/" & month(DateVar2) & "/" & year(DateVar2)
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

RMAStatusCond=""
if Request("RMAStatus")<>"" then
	if Request("RMAStatus")<>"-1" then
		RMAStatusCond = " AND PCReturns.rmaApproved = " & Request("RMAStatus")
	end if
end if

call opendb()
query="SELECT Orders.idorder,Orders.orderDate,Customers.name,Customers.lastname,PCReturns.idRMA,PCReturns.rmaApproved,PCReturns.rmaNumber FROM PCReturns, Orders, Customers WHERE PCReturns.idOrder = orders.idOrder AND orders.idCustomer = Customers.idCustomer " & TempSQL1 & TempSQL2 & TempSpecial & RMAStatusCond & " Order by " & tmpD & " DESC"
set rs=connTemp.execute(query)

' if there are no records in recordset
If rs.EOF Then %>
	<tr> 
		<td colspan="6">No records match your query</td>
	</tr>
<% Else %>
	<tr> 
		<td colspan="6"><h2>RMA report from <%=strTDateVar%> to <%=strTDateVar2%></h2>
	</tr>
	<tr> 
		<th nowrap>Order #</th>
		<th nowrap>Order Date</th>
		<th nowrap>Customer Name</th>
		<th nowrap>RMA Status</th>
		<th nowrap>RMA Number</th>
		<th nowrap>Details</th>
	</tr>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
	<%
	do while not rs.eof
		pc_idorder=rs("idorder")
		pc_orderDate=rs("orderDate")
		if scDateFrmt="DD/MM/YY" then
			pc_orderDate=(day(pc_orderDate)&"/"&month(pc_orderDate)&"/"&year(pc_orderDate))
		end if
		pc_CustName=rs("name")& " " & rs("lastname")
		pc_IdRMA= rs("idRMA")
		if IsNull(rs("rmaApproved")) = False then
			if rs("rmaApproved")=1 then
				pc_RMAStatus="Approved"
			elseif rs("rmaApproved")=2 then
				pc_RMAStatus="Denied"
			else
				pc_RMAStatus="Requested"
			end if
		else
			pc_RMAStatus="Requested"	
		end if
		if IsNull(rs("rmaNumber")) = False then
			pc_RMANumber = rs("rmaNumber")
		else
			pc_RMANumber=""
		end if
	%>
	<tr>  
		<td nowrap valign="top"><a href="orddetails.asp?id=<%=pc_idorder%>" target="_blank" title="View details for this order"><%=pc_idorder%></a></td>
		<td nowrap valign="top"><%=pc_orderDate%></td>
		<td nowrap valign="top"><%=pc_CustName%></td>
		<td nowrap valign="top"><%=pc_RMAStatus%></td>
		<td nowrap valign="top"><%=pc_RMANumber%></td>
		<td nowrap valign="top"><a href=modRmaa.asp?idOrder=<%=pc_idorder%>&idRMA=<%=pc_IdRMA%> target="_blank" title="View RMA Information">Details</a></td>
	</tr>
	<% 
		rs.MoveNext
	loop
END IF%>     
</table>
<%	' Done. Now release Objects
	Set rs=Nothing
	call closedb()
%>
</div>
</body>
</html>