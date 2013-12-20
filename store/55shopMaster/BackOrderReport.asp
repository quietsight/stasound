<%'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Back Order Report"
Response.Buffer = False
Server.ScriptTimeout = 8000 %>

<% Section="genRpts" %>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp" --> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim connTemp, rs
call openDb()

FromDate=Request("FromDate")
ToDate=Request("ToDate")

if (FromDate<>"") and (not (IsDate(FromDate))) then
	FromDate=Date()
end if
if (ToDate<>"") and (not (IsDate(ToDate))) then
	ToDate=Date()
end if

Dim strTDateVar, strTDateVar2, DateVar, DateVar2
strTDateVar=FromDate

if scDateFrmt="DD/MM/YY" then
	DateVarArray=split(strTDateVar,"/")
	FromDate=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
end if

strTDateVar2=ToDate
if scDateFrmt="DD/MM/YY" then
	DateVarArray2=split(strTDateVar2,"/")
	ToDate=(DateVarArray2(1)&"/"&DateVarArray2(0)&"/"&DateVarArray2(2))
	if err.number<>0 then
		FromDate=(day(FromDate)&"/"&month(FromDate)&"/"&year(FromDate))
		ToDate=(day(ToDate)&"/"&month(ToDate)&"/"&year(ToDate))
	end if
end if

query1=""
query2=""

if SQL_Format="1" then
	FromDate=Day(FromDate)&"/"&Month(FromDate)&"/"&Year(FromDate)
	ToDate=Day(ToDate)&"/"&Month(ToDate)&"/"&Year(ToDate)
else
	FromDate=Month(FromDate)&"/"&Day(FromDate)&"/"&Year(FromDate)
	ToDate=Month(ToDate)&"/"&Day(ToDate)&"/"&Year(ToDate)
end if

if (FromDate<>"") and (IsDate(FromDate)) then
	if scDB="Access" then
		query1 = " AND orders.orderDate >=#" & FromDate & "# "
	else
		query1 = " AND orders.orderDate >='" & FromDate & "' "
	end if
end if

if (ToDate<>"") and (IsDate(ToDate)) then
	if scDB="Access" then
		query2 = " AND orders.orderDate <=#" & ToDate & "# "
	else
		query2 = " AND orders.orderDate <='" & ToDate & "' "
	end if
end if

query="SELECT products.idproduct,products.sku,products.description,ProductsOrdered.quantity,orders.idOrder FROM Products INNER JOIN (ProductsOrdered INNER JOIN Orders ON ProductsOrdered.idOrder=Orders.idOrder) ON Products.idproduct=ProductsOrdered.idproduct WHERE ((Orders.orderStatus>=2 AND Orders.orderStatus<=3) OR (Orders.orderStatus>=7 AND Orders.orderStatus<=9)) AND (ProductsOrdered.pcPrdOrd_BackOrder=1) AND (ProductsOrdered.pcPrdOrd_Shipped=0) AND (Products.active<>0) AND (Products.removed=0) " & query1 & query2 & " ORDER BY Products.Description ASC, Orders.IdOrder ASC;"
set rs=connTemp.execute(query)

pcv_havelist=0

if not rs.eof then
	pcv_havelist=1
	pcArr=rs.getRows()
	set rs=nothing
	intCount=ubound(pcArr,2)
end if
set rs=nothing
call closedb()

%>
<table class="pcCPcontent">
	
		<%if pcv_havelist<>"1" then %>
			<tr> 	
				<td colspan="3"> 
					<div class="pcCPmessage"><img src="images/pcadmin_note.gif" width="20" height="20">No Results Found</div>
				</td>
			</tr>
		<% else%>
		<tr>
			<td colspan="3">
				<%if FromDate<>"" then%>
					From:&nbsp;<%=FromDate%>
				<%end if%>
				<%if ToDate<>"" then%>
					&nbsp;To:&nbsp;<%=ToDate%>
				<%end if%>
			</td>
		<tr>
			<th nowrap>Product Name</th>
			<th nowrap>Ordered Quantity</th>
			<th nowrap>Order #IDs</th>
		</tr>
		<%
		SaveID=""
		PrdName=""
		PrdSku=""
		OrderList=""
		PrdQty=0
		For i=0 to intCount
			if SaveID & ""<>pcArr(0,i) & "" then
				if SaveID<>"" then%>
				<tr bgcolor="<%= strCol %>">
					<td nowrap valign="top"><a href="FindProductType.asp?id=<%=SaveID%>"><%=PrdName%> (<%=PrdSku%>)</a></td>
					<td nowrap valign="top"><a href="viewStock.asp"><b><%=PrdQty%></b></a></td>
					<td  valign="top"><%=OrderList%></td>
				</tr>
				<%end if
				SaveID=pcArr(0,i)
				PrdName=pcArr(1,i)
				PrdSku=pcArr(2,i)
				OrderList="<a href='Orddetails.asp?id=" & pcArr(4,i) & "'>#" & (int(pcArr(4,i))+scpre) & "</a>"
				PrdQty=Clng(pcArr(3,i))
			else
				if OrderList<>"" then
					OrderList=OrderList & ", "
				end if
				OrderList=OrderList & "<a href='Orddetails.asp?id=" & pcArr(4,i) & "'>#" & (int(pcArr(4,i))+scpre) & "</a>"
				PrdQty=Clng(PrdQty)+Clng(pcArr(3,i))
			end if
		Next
		if SaveID<>"" then%>
		<tr bgcolor="<%= strCol %>">
			<td nowrap valign="top"><a href="FindProductType.asp?id=<%=SaveID%>"><%=PrdName%> (<%=PrdSku%>)</a></td>
			<td nowrap valign="top"><a href="viewStock.asp"><b><%=PrdQty%></b></a></td>
			<td  valign="top"><%=OrderList%></td>
		</tr>
		<%end if%>
	<%end if 'Have Product Records%>
    <tr>
        <td colspan="3">&nbsp;</td>
    </tr>
</table>
<!--#include file="AdminFooter.asp"-->