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
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/languagesCP.asp" -->
<html>
<head>
	<title>Drop-Off Report - by Customer Type</title>
	<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="margin: 10px;">
<div id="pcCPmain" style="background-image: none;">
<table class="pcCPcontent" style="width: auto;">
        
<% 'on error resume next
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

custStr="Retail|WholeSale"
custIdStr="0|1"
mySQL="SELECT idcustomerCategory,pcCC_Name FROM pcCustomerCategories order by idcustomerCategory;"
set rs1=conntemp.execute(mySQL)
do while not rs1.eof
	custStr=custStr&"|"&rs1("pcCC_Name")
	custIdStr=custIdStr&"|cc_"&rs1("idcustomerCategory")
	rs1.movenext
loop
custArr=split(custStr,"|")
custArrType=split(custIdStr,"|")
%>

	<tr> 
		<td colspan="6"><h2>Total Drop-Offs by Customer Type recorded from: <%=strTDateVar%> to: <%=strTDateVar2%></h2></td>
	</tr>
	<tr>
		<th nowrap>Customer Type</th> 
		<th align="center" nowrap>Total Orders</th> 
		<th align="center" nowrap>Total Drop-Off</th>
		<th align="center" nowrap>Drop-Off Percent</th>
		<th align="center" nowrap>Total Drop-Off Units</th> 
		<th align="right" nowrap>Total Drop-Off Amount</th>
	</tr>
	<tr>
		<td colspan="6" class="pcCPspacer"></td>
	</tr>
	<% 
				
		for lCnt=0 to ubound(custArr)
			pcv_Total=0
			query="Select count(*) as Total from orders "
			query=query&"inner join customers on orders.idcustomer=customers.idcustomer " 
			query=query&"WHERE orders.orderStatus>0 " & TempSQL1 & TempSQL2 
			IF custArrType(lcnt)="0" THEN
				query=query&"AND customers.customertype=0 AND customers.idcustomercategory=0 "
			elseif custArrType(lcnt)="1" THEN
				query=query&"AND customers.customertype=1 AND customers.idcustomercategory=0 "
			else
				pCustType=right(custArrType(lcnt),len(custArrType(lcnt))-3)
				query=query&"AND customers.idcustomercategory="&pCustType&" "
			end if
			set rstemp=conntemp.execute(query)
			if not rstemp.eof then
				pcv_Total=rstemp("Total")
				gTotal=gTotal+pcv_Total
			end if 
			set rstemp = nothing
			
			pcv_Incomp=0
			query="Select count(*) as Incomplete from orders "
			query=query&"inner join customers on orders.idcustomer=customers.idcustomer " 
			query=query&"WHERE orders.orderStatus=1 " & TempSQL1 & TempSQL2 
			IF custArrType(lcnt)="0" THEN
				query=query&"AND customers.customertype=0 AND customers.idcustomercategory=0 "
			elseif custArrType(lcnt)="1" THEN
				query=query&"AND customers.customertype=1 AND customers.idcustomercategory=0 "
			else
				pCustType=right(custArrType(lcnt),len(custArrType(lcnt))-3)
				query=query&"AND customers.idcustomercategory="&pCustType&" "
			end if
			set rstemp=conntemp.execute(query)
			if not rstemp.eof then
				pcv_Incomp=rstemp("Incomplete")
				gIncomp=gIncomp+pcv_Incomp
			end if 
			set rstemp = nothing


			gTotalUnit=0
			gTotalAmount=0
			query="select sum(a.Total) as Total,sum(b.TotalQty) as TotalQty from  "
			query=query&"( "
			query=query&"SELECT idorder,sum(orders.Total) as Total from orders "
			query=query&"inner join customers on orders.idcustomer=customers.idcustomer "
			query=query&"WHERE orderStatus=1 " & TempSQL1 & TempSQL2 
			IF custArrType(lcnt)="0" THEN
				query=query&"AND customers.customertype=0 AND customers.idcustomercategory=0 "
			elseif custArrType(lcnt)="1" THEN
				query=query&"AND customers.customertype=1 AND customers.idcustomercategory=0 "
			else
				pCustType=right(custArrType(lcnt),len(custArrType(lcnt))-3)
				query=query&"AND customers.idcustomercategory="&pCustType&" "
			end if
			query=query&"group by idorder "
			query=query&") a "
			query=query&"inner join "
			query=query&"( "
			query=query&"select idorder,SUM(ProductsOrdered.quantity) as TotalQty "
			query=query&"FROM  ProductsOrdered group by idorder "
			query=query&") b "
			query=query&"on a.idorder=b.idorder "
			set rs1=conntemp.execute(query)
			if not rs1.eof then				
				if rs1("TotalQty") <> "" then
					gTotalUnit=cdbl(rs1("TotalQty"))
				end if
				if rs1("Total") <> "" then
					gTotalAmount=cdbl(rs1("Total"))
				end if
				gTotalQty=gTotalQty+gTotalUnit
				gTotalsales=gTotalsales+gTotalAmount
			end if

			%>
			<tr>
				<td nowrap><%=custArr(lcnt)%></td> 
				<td align="center" nowrap><%=pcv_Total%></td> 
				<td align="center" nowrap><%=pcv_Incomp%></td> 
				<%
					dropOffPercent=0
					if pcv_Incomp>0 and pcv_Total>0 then
						dropOffPercent=clng((pcv_Incomp/pcv_Total)*100)
					end if
					if dropOffPercent < 1 then
						dropOffPercent = 0
					end if
				%>
				<td align="center" nowrap><%=dropOffPercent%>%</td> 
				<td align="center" nowrap><%=gTotalUnit%></td>
				<td align="right" nowrap><%=scCurSign&money(gTotalAmount)%></td>
			</tr>
		<% 
			set rs1=nothing
	
		next
		%>
			<tr>         
				<td colspan="6" class="pcCPspacer"></td>
			</tr>
			<tr bgcolor="#e1e1e1">
				<td nowrap><strong>Totals:</strong></td> 
				<td align="center" nowrap><%=gTotal%></td> 
				<td align="center" nowrap><%=gIncomp%></td> 
				<%
					dropOffPercent=0
					if gIncomp>0 and gTotal>0 then
						dropOffPercent=clng((gIncomp/gTotal)*100)
					end if
					if dropOffPercent < 1 then
						dropOffPercent = 0
					end if
				%>
				<td align="center" nowrap><%=dropOffPercent%>%</td> 
				<td align="center" nowrap><%=gTotalQty%></td>
				<td align="right" nowrap><%=scCurSign&money(gTotalsales)%></td>
			</tr>
			<tr>         
				<td colspan="6" class="pcCPspacer"></td>
			</tr>

</table>
<%	' Done. Now release Objects
con.Close
Set con=Nothing
Set rs=Nothing
%>
</div>
</body>
</html>