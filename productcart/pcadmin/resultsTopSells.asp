<%'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Sales Reports"
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
Dim connTemp, rstemp
call openDb()
%>
<!--#include file="pcCharts.asp"-->
<%
recordsToShow=Request.QueryString("resultCnt")
srcVar=Request.QueryString("src")
FromDate=Request.QueryString("FromDate")
ToDate=Request.QueryString("ToDate")

query1=""

if (FromDate<>"") and (not (IsDate(FromDate))) then
	FromDate=Date()
end if
if (ToDate<>"") and (not (IsDate(ToDate))) then
	ToDate=Date()
end if

err.clear

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

TempSQL1=""
TempSQL2=""

if SQL_Format="1" then
	FromDate=Day(FromDate)&"/"&Month(FromDate)&"/"&Year(FromDate)
	ToDate=Day(ToDate)&"/"&Month(ToDate)&"/"&Year(ToDate)
else
	FromDate=Month(FromDate)&"/"&Day(FromDate)&"/"&Year(FromDate)
	ToDate=Month(ToDate)&"/"&Day(ToDate)&"/"&Year(ToDate)
end if

if (FromDate<>"") and (IsDate(FromDate)) then
	if scDB="Access" then
		TempSQL1 = " AND " & tmpD & " >=#" & FromDate & "#"
	else
		TempSQL1 = " AND " & tmpD & " >='" & FromDate & "'"
	end if
end if

if (ToDate<>"") and (IsDate(ToDate)) then
	if scDB="Access" then
		TempSQL2 = " AND " & tmpD & " <=#" & ToDate & "#"
	else
		TempSQL2 = " AND " & tmpD & " <='" & ToDate & "'"
	end if
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

query1=query1 & TempSQL1 & TempSQL2 '& TempSpecial

'// Top Viewed Products
if srcVar="2" then
	query="SELECT description, visits, idproduct FROM products WHERE products.visits >0 ORDER BY products.visits DESC;"
end if

'// Top 'Wish List' Products
if srcVar="4" then
	query="SELECT IDProduct, COUNT(*) AS TotalCount FROM WishList GROUP BY IDProduct;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	pcv_havelist=0
	tmp1=""
	tmp2=""
	do while not rs.eof
		if tmp1<>"" then
			tmp1=tmp1 & "*"
			tmp2=tmp2 & "*"
		end if
		tmp1=tmp1&rs("IDProduct")
		tmp2=tmp2&rs("TotalCount")
		pcv_havelist=1
		rs.MoveNext
	loop
	
	set rs=nothing
	
	if pcv_havelist="1" then
		IDList=split(tmp1,"*")
		PCount=split(tmp2,"*")
		For i=lbound(IDList) to ubound(IDList)
			For j=i+1 to ubound(IDList)
				if clng(pCount(i))<clng(pCount(j)) then
				m1=IDList(i)
				m2=PCount(i)
				IDList(i)=IDList(j)
				PCount(i)=PCount(j)
				IDList(j)=m1
				PCount(j)=m2
				end if
			Next
		Next
	end if
	
end if

'// Top Selling Products
if srcVar="1" then
	query="update products set Sales=0" 
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)

	query="SELECT ProductsOrdered.IDProduct, SUM(ProductsOrdered.Quantity) As PrdSales FROM ProductsOrdered, Orders where Orders.IDOrder=ProductsOrdered.IDOrder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) " & query1 & " GROUP BY ProductsOrdered.IDProduct"
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)

	if not rstemp.eof then
		pcArr=rstemp.getRows()
		set rstemp=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
			'insert values into products table
			query="update products set Sales=" & pcArr(1,i) & " where idproduct="& pcArr(0,i)
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
		Next
	end if
	set rstemp=nothing
	
	query="SELECT description, sales, idproduct FROM products WHERE products.sales >0 ORDER BY products.sales DESC;"
end if 

'//Top Customers
if srcVar="3" then
	query="update customers set TotalOrders=0, TotalSales=0"
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)

	query="select idcustomer, sum(total), count(*) from orders where ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12))" & TempSpecial & query1 &" GROUP BY idcustomer"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
		
	if not rs.eof then
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount
			'insert values into customers table
			query="update customers set TotalOrders=" & pcArr(2,i) & ", TotalSales=" & pcArr(1,i) & " where idcustomer="& pcArr(0,i)
			Set rs=CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			set rs=nothing
		Next
	end if
	set rs=nothing
	
	query="Select name, lastname, customerCompany, Totalorders, Totalsales, idcustomer from customers where TotalOrders>0 ORDER BY	TotalSales DESC"
	
end if

' Our Recordset Object
if srcVar<>"4" then
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
end if

Dim rcount, i, x
Dim tmpChartName,tmpLine1
tmpChartName=""
tmpLine1=""
%>
<table class="pcCPcontent" width="100%">
<tr valign="top">
<td width="50%">
<table class="pcCPcontent" style="width:auto;">
	<tr>
		<td colspan="5" nowrap>
		<% if srcVar<>"2" then %>
			<%if FromDate<>"" then%>
				From:&nbsp;<%=FromDate%>
			<%end if%>
			<%if ToDate<>"" then%>
				&nbsp;To:&nbsp;<%=ToDate%>
			<%end if%>
		<%end if%>
		</td>
	</tr>
	<% if srcVar="1" then
	tmpChartName="Top Selling Products" %>
		<tr> 
			<th colspan="2" nowrap>Top Selling Products</th>
			<th nowrap colspan="2">Amount Sold</th>
		</tr>
	<% end if %>

	<% if srcVar="2" then
	tmpChartName="Most Viewed Products" %>
		<tr> 
			<th colspan="2" nowrap>Most Viewed Products</th>
			<th nowrap colspan="2">Total Views</th>
        </tr>
	<% end if %>

	<% if srcVar="4" then
	tmpChartName="Top ""Wish List"" Products" %>
		<tr> 
			<th colspan="2" nowrap>Top "Wish List" Products</th>
			<th nowrap colspan="2">Number of Wish Lists</th>
		</tr>
	<% end if %>

	<% if srcVar="3" then
	tmpChartName="Best Customers" %>
		<tr> 
			<th colspan="2" nowrap>Best Customers</th>
			<th nowrap>Number of Orders</th>
			<th colspan="2" nowrap>Orders Total</th>
		</tr>
	<% end if %>
	
	<% if srcVar="3" then %>
		<tr> 
			<td colspan="5" class="pcCPspacer"></td>
		</tr>
	<%else%>
		<tr> 
			<td colspan="4" class="pcCPspacer"></td>
		</tr>
	<%end if%>

	<% IF srcVar="4" THEN
		if pcv_havelist<>"1" then %>
			<tr> 	
				<td colspan="5"> 
					<div class="pcCPmessage"><img src="images/pcadmin_note.gif" width="20" height="20">No Results Found</div>
				</td>
			</tr>
		<% else
			rCount=0
			do while (clng(rCount-1)<clng(ubound(IDList))) and (clng(rCount)<clng(recordsToShow))
				pIDProduct=IDList(rCount)
				TotalCount=PCount(rCount)
				rCount=rCount+1
				query="SELECT Description FROM Products WHERE IDProduct=" & pIDProduct
				set rs=server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
				PDesc=rs("Description")
				if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				end if
				tmpline1=tmpline1 & "['" & PDesc & "'," & Clng(TotalCount) & "]"
				%>
				<tr bgcolor="<%= strCol %>"> 		
					<td width="6%"><%Response.Write rcount %></td>
					<td width="44%"><a href="FindProductType.asp?id=<%=pIDProduct%>" target="_blank"><%=PDesc%></a></td>
					<td colspan="2"><%=TotalCount%></td>
				</tr>
		
			<% loop		
		end if
	ELSE
		If rs.EOF Then %>
			<tr> 	
                <td colspan="5" height="22"> 
                        <div class="pcCPmessage"><img src="images/pcadmin_note.gif" width="20" height="20">No Results Found</div>
                </td>
			</tr>
		<% Else 
			' Showing relevant records
			rcount=0
			rs.MoveFirst
		
			rcount=0
			do while (not rs.EOF) and (int(rcount)<int(recordsToShow))
				rcount=rcount+1
				IF srcVar="1" then %>
						
                    <tr bgcolor="<%= strCol %>"> 
                        <td width="6%"><%Response.Write rcount %></td>
                        <td width="44%"><a href="FindProductType.asp?id=<%=rs("idProduct")%>" target="_blank"><%=rs("Description")%></a></td>
                        <td colspan="2"><%=rs("sales")%></td>
                    </tr>
					

				<%
				if tmpline1<>"" then
					tmpline1=tmpline1 & ","
				end if
				tmpline1=tmpline1 & "['" & rs("Description") & "'," & Clng(rs("sales")) & "]"
				END IF

				IF srcVar="2" then %>
		
					<tr bgcolor="<%= strCol %>"> 		
						<td width="6%"><%Response.Write rcount %></td>
						<td width="44%"><a href="FindProductType.asp?id=<%=rs("idProduct")%>" target="_blank"><%=rs("Description")%></a></td>
						<td colspan="2"><%=rs("visits")%></td>
					</tr>
	
				<%
				if tmpline1<>"" then
					tmpline1=tmpline1 & ","
				end if
				tmpline1=tmpline1 & "['" & rs("Description") & "'," & Clng(rs("visits")) & "]"
				END IF

				IF srcVar="3" THEN %>
							
					<tr bgcolor="<%= strCol %>"> 			
						<td width="6%"><%Response.Write rcount %></td>
						<td width="44%" nowrap><a href="modCusta.asp?idcustomer=<%=rs("idcustomer")%>" target="_blank"><%=rs("name")%>&nbsp;<%=rs("lastname")%></a>
                        <%
						pcvCustomerCompany = rs("customerCompany")
						if pcvCustomerCompany<>"" and not isNull(pcvCustomerCompany) then
						%>
						&nbsp;(<%=pcvCustomerCompany%>)
						<%
						end if
						%>
						</td>
						<td align="center"><%=rs("Totalorders")%></td>
						<td colspan="2"><%=scCurSign%> <%=money(rs("TotalSales"))%></td>
					</tr>
						
				<%
				if tmpline1<>"" then
					tmpline1=tmpline1 & ","
				end if
				tmpline1=tmpline1 & "['" & rs("name") & " " & rs("lastname") & "'," & Clng(rs("TotalSales")) & "]"
				END IF
					
				rs.MoveNext
			loop
			
			If srcVar="1" then
			
				'reset sales
				query="update products set Sales=0" 
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)

				query="SELECT ProductsOrdered.IDProduct, SUM(ProductsOrdered.Quantity) As PrdSales FROM ProductsOrdered, Orders where Orders.IDOrder=ProductsOrdered.IDOrder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) GROUP BY ProductsOrdered.IDProduct"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
	
				do while not rstemp.eof
					IDOProduct=rstemp("IDProduct")
					Totalsales=rstemp("PrdSales")
			
					'insert values into products table
					query="update products set Sales=" & Totalsales & " where idproduct="& IDOProduct
					set rstemp2=server.CreateObject("ADODB.RecordSet")
					set rstemp2=conntemp.execute(query)
					set rstemp2=nothing
					
					rstemp.movenext
				loop
				
			End If
			
			
		End If 'Have Records
	END IF 'Top Wish List Products
	%>
    <tr>
        <td colspan="4">&nbsp;</td>
    </tr>
</table>
</td>
<td width="50%">
<%if tmpline1<>"" then%>
	<div id="chartTop" style="height:330px; "></div>
	<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.pieRenderer.min.js"></script>
		
	<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		plot2 = $.jqplot('chartTop', [line1], {
    	title: '<%=tmpChartName%>',
    	seriesDefaults:{renderer:$.jqplot.PieRenderer, rendererOptions:{showDataLabels: true,sliceMargin:0}},
    	legend:{show:true}
		});});
	</script>
<%end if%>
</td>
</tr>
</table>
<%  ' Done. Now release Objects
set rs=nothing
call closedb()
%>
<!--#include file="AdminFooter.asp"-->