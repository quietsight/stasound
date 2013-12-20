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
<body style="margin:10px; background-image: none;">
<div id="pcCPmain">
<table class="pcCPcontent">
        
<%
Dim connTemp, query,rs,rstemp

function ShowDateTimeFrmt(datestring)
Dim tmp1,tmp2
	tmp1=split(datestring," ")
	if scDateFrmt="DD/MM/YY" then
		tmp2=day(tmp1(0))&"/"&month(tmp1(0))&"/"&year(tmp1(0))
	else
		tmp2=month(tmp1(0))&"/"&day(tmp1(0))&"/"&year(tmp1(0))
	end if
	if instr(datestring," ") then
		tmp2=tmp2 & " " & tmp1(1) & tmp1(2)
	end if
	ShowDateTimeFrmt=tmp2
end function

	
call opendb()

pcSaleID=request("id")
tmpQ1=""
pcSaleName=""
pcSCID=request("sub")

IF pcSaleID<>"" AND pcSaleID<>"0" then
	queryS="SELECT pcSales_Name FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
	set rsS=connTemp.execute(queryS)
	if not rsS.eof then
		pcSaleName=rsS("pcSales_Name")
	end if
	set rsS=nothing
	tmpQ1=tmpQ1 & " pcSales_Completed.pcSales_ID=" & pcSaleID & " AND "
End if

IF pcSCID<>"" AND pcSCID<>"0" then
	tmpQ1=tmpQ1 & " pcSales_Completed.pcSC_ID=" & pcSCID & " AND "
End if


query="SELECT ProductsOrdered.pcSC_ID,SUM(ProductsOrdered.unitPrice*ProductsOrdered.quantity) FROM pcSales_Completed,orders,ProductsOrdered WHERE " & tmpQ1 & " ProductsOrdered.pcSC_ID=pcSales_Completed.pcSC_ID AND Orders.idOrder=ProductsOrdered.idOrder AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) GROUP BY ProductsOrdered.pcSC_ID,pcSales_Completed.pcSales_ID ORDER BY pcSales_Completed.pcSales_ID ASC;"
set rs=connTemp.execute(query)
	
' If the returning recordset is not empty
If rs.EOF Then %>
	<tr> 
		<td colspan="6">No records match your query</td>
	</tr>
<% Else%>
	<tr> 
		<td colspan="6"><h2>SALE SUMMARY REPORT<%if pcSaleName<>"" then%> for: <strong><%=pcSaleName%></strong><%end if%></h2></td>
	</tr>
	<tr>
		<th nowrap>Sale Name</th> 
		<th nowrap>Started On</th>
		<th nowrap>Completed On</th>
		<th nowrap>Total Orders</th>
		<th nowrap>Product Amount</th>
		<th nowrap>Order Amount</th>
	</tr>
	<tr>
		<td colspan="6" class="pcCPspacer"></td>
	</tr>
	<%
	pcArr=rs.getRows()
	set rs=nothing
	intCount=ubound(pcArr,2)
	For i=0 to intCount
		tmpSCID=pcArr(0,i)
		tmpPrdAmount=pcArr(1,i)
		
		query="SELECT TotalOrders=COUNT(*),TotalOrderAmount=SUM(Orders.total) FROM Orders WHERE Orders.IDOrder IN (SELECT DISTINCT ProductsOrdered.IDOrder FROM ProductsOrdered WHERE ProductsOrdered.pcSC_ID=" & tmpSCID & ");"
		set rstemp=connTemp.execute(query)
		if not rstemp.eof then
			tmpTotalOrders=rstemp("TotalOrders")
			tmpOrdAmount=rstemp("TotalOrderAmount")
		end if
		set rstemp=nothing
		
		query="SELECT pcSC_SaveName,pcSC_StartedDate,pcSC_ComDate FROM pcSales_Completed WHERE pcSC_ID=" & tmpSCID & ";"
		set rstemp=connTemp.execute(query)
		
		tmpSaveName=""
		tmpSDate="N/A"
		tmpCDate="N/A"
		
		if not rstemp.eof then
			tmpSaveName=rstemp("pcSC_SaveName")
			tmpSDate=rstemp("pcSC_StartedDate")
			if Not IsNull(tmpSDate) then
				tmpSDate=ShowDateTimeFrmt(tmpSDate)
			else
				tmpSDate="N/A"
			end if
			tmpCDate=rstemp("pcSC_ComDate")
			if Not IsNull(tmpCDate) then
				tmpCDate=ShowDateTimeFrmt(tmpCDate)
			else
				tmpCDate="N/A"
			end if
		end if
		set rstemp=nothing
			
		%>
		<tr>  
			<td nowrap><%=tmpSaveName%></td>
			<td nowrap><%=tmpSDate%></td>
			<td nowrap><%=tmpCDate%></td>
			<td nowrap><%=tmpTotalOrders%></td>
			<td nowrap><%=scCurSign&money(tmpPrdAmount)%></td>
			<td nowrap><%=scCurSign&money(tmpOrdAmount)%></td>
		</tr>
	<%Next%>

	<%IF pcSCID<>"" AND pcSCID<>"0" then%>
	
		<tr>
		<td colspan="6" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td colspan="6" class="pcCPspacer"></td>
		</tr>
		<%
		query="SELECT orders.idorder,orders.total,orders.orderDate,Customers.idCustomer,Customers.Name,Customers.Lastname,Products.Description,ProductsOrdered.quantity,ProductsOrdered.unitPrice FROM Orders,ProductsOrdered,Products,Customers WHERE ProductsOrdered.pcSC_ID=" & pcSCID & " AND Products.IDProduct=ProductsOrdered.IDProduct AND orders.idorder=ProductsOrdered.idorder AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND Customers.IDCustomer=Orders.IDCustomer ORDER BY Orders.IDOrder DESC;"
		set rs=connTemp.execute(query)			
		' If the returning recordset is not empty
		If not rs.EOF Then
			pcArr=rs.getRows()
			intCount=ubound(pcArr,2)
			set rs=nothing
			%>
			<tr> 
				<th nowrap>Order #</th>
				<th nowrap>Date</th>
				<th nowrap>Customer Name</th>
				<th nowrap>Products</th>
				<th nowrap>Product Amount</th>
				<th nowrap>Order Amount</th>
			</tr>
			<tr>
				<td colspan="6" class="pcCPspacer"></td>
			</tr>
			<% 
			gTotalNumberOrders=0
			gTotalQty=0
			gTotalPrdAmount=0
			gTotalOrdAmount=0
			
			tmpOrdDetails=""
			tmpOrdID=0
			tmpPrdAmount=0
			tmpOrdQty=0
			For i=0 to intCount
				if clng(tmpOrdID)<>clng(pcArr(0,i)) then
					if tmpOrdDetails<>"" then%>
						<tr valign="top">
							<td><a href="orddetails.asp?id=<%=tmpOrdID%>" target="_blank"><%=(scpre+int(tmpOrdID))%></a></td>
							<td><%=ShowDateTimeFrmt(pcArr(2,i-1))%></td>
							<td><a href="viewCustOrders.asp?idcustomer=<%=pcArr(3,i-1)%>" target="_blank"><%=pcArr(4,i-1) & " " & pcArr(5,i-1)%></a></td>
							<td><%=tmpOrdDetails%></td>
							<td><%=scCurSign & money(tmpPrdAmount)%></td>
							<td><%=scCurSign & money(pcArr(1,i-1))%></td>
						</tr>
						<%
						gTotalNumberOrders=gTotalNumberOrders+1
						gTotalQty=gTotalQty+tmpOrdQty
						gTotalPrdAmount=gTotalPrdAmount+tmpPrdAmount
						gTotalOrdAmount=gTotalOrdAmount+Cdbl(pcArr(1,i-1))
						
						tmpOrdDetails=""
						tmpPrdAmount=0
						tmpOrdQty=0
					end if
					tmpOrdID=pcArr(0,i)
				end if
				tmpOrdDetails=tmpOrdDetails & pcArr(6,i) & " - Qty: " & pcArr(7,i) & " - Amount: " & scCurSign & money(pcArr(7,i)*pcArr(8,i)) & "<br>"
				tmpOrdQty=tmpOrdQty+Clng(pcArr(7,i))
				tmpPrdAmount=tmpPrdAmount+Cdbl(pcArr(7,i)*pcArr(8,i))
			Next
			
			if tmpOrdDetails<>"" then%>
				<tr valign="top">
					<td><a href="orddetails.asp?id=<%=tmpOrdID%>" target="_blank"><%=(scpre+int(tmpOrdID))%></a></td>
					<td><%=ShowDateTimeFrmt(pcArr(2,intCount))%></td>
					<td><a href="viewCustOrders.asp?idcustomer=<%=pcArr(3,intCount)%>" target="_blank"><%=pcArr(4,intCount) & " " & pcArr(5,intCount)%></a></td>
					<td><%=tmpOrdDetails%></td>
					<td><%=scCurSign & money(tmpPrdAmount)%></td>
					<td><%=scCurSign & money(pcArr(1,intCount))%></td>
				</tr>
				<%
				gTotalNumberOrders=gTotalNumberOrders+1
				gTotalQty=gTotalQty+tmpOrdQty
				gTotalPrdAmount=gTotalPrdAmount+tmpPrdAmount
				gTotalOrdAmount=gTotalOrdAmount+Cdbl(pcArr(1,intCount))
			end if%>
			<tr>         
				<td colspan="6" class="pcCPspacer">&nbsp;</td>
			</tr>
			<tr bgcolor="#e1e1e1">
				<td><strong>Totals</strong></td>
				<td align="right">Orders:</td>
				<td><b><%=gTotalNumberOrders%></b></td>
				<td>Qty: <b><%=gTotalQty%></b></td>
				<td><b><%=scCurSign&money(gTotalPrdAmount)%></b></td>
				<td><b><%=scCurSign&money(gTotalOrdAmount)%></b></td>
			</tr>

		<%End If
		set rs=nothing %>
	<%END IF%>
</table>
<%
END IF
call closedb()
%>
</div>
</body>
</html>