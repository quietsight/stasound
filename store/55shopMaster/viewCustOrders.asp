<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="List Orders for Selected Customer" 
pageIcon="pcv4_icon_people.png"
section= "mngAcc"
pcInt_ShowOrderLegend = 1
%>
<%PmAdmin="7*9*"%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"--> 
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/encrypt.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 

Const iPageSize=15
Dim iPageCurrent
if request.querystring("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request.QueryString("iPageCurrent")
end if

'sorting order
Dim strORD
strORD=request("order")
if strORD="" then
	strORD="idorder"
End If

strSort=request("sort")
if strSort="" Then
	strSort="DESC"
End If 

dim query, conntemp, rstemp

call openDb()

	' Get customer information
		pIdCustomer=Request.QueryString("idcustomer")
		if not validNum(pIdCustomer) then
			response.redirect "menu.asp"
		end if
		query="select name,lastName,customerCompany,email,password from customers where idcustomer=" & pIdCustomer
		set rstemp=connTemp.execute(query)

		pemail=rstemp("email")
		pname=rstemp("name") & " " & rstemp("lastName")
		pcompany=rstemp("customerCompany")
			if isNull(pcompany) or trim(pcompany)="" then
				pcompany="no"
			end if
		ppassword=enDeCrypt(rstemp("password"), scCrypPass)
		ppassword=encrypt(ppassword, 9286803311968)
		ppassword=Server.URLEncode(ppassword)


		' Get order totals
			TotalOrdered=0
			query="SELECT sum(total-rmaCredit) As TotalAmount, Sum(Total) AS TotalLessRMA FROM orders WHERE ((orderStatus>2 AND orderStatus<5) OR (orderStatus>6 AND orderStatus<=9) OR (orderStatus=10 OR orderStatus=12)) AND idcustomer=" & pIdCustomer &";"
			
			set rstemp=Server.CreateObject("ADODB.Recordset") 
			rstemp.Open query, conntemp
			if err.number <> 0 then
				response.redirect "techErr.asp?error="& Server.Urlencode("Error calculating order total: "&Err.Description) 
			end If
			if not rstemp.eof then
				TotalOrdered=rstemp("TotalAmount")
				TotalLessRMA=rstemp("TotalLessRMA")
			end if
			if isNull(TotalOrdered) then
				TotalOrdered = TotalLessRMA
				if isNull(TotalOrdered) then
					TotalOrdered=0
				end if
			end if
			set rstemp=nothing
			
		' Get order totals for this year
			Dim pcvThisYear, pcvLastYear, pcvYearAgo, TotalOrderedYear, TotalOrdered12
			pcvThisYear = year(date)
			TotalOrderedYear=0
			query="SELECT sum(total-rmaCredit) As TotalAmount, Sum(Total) AS TotalLessRMA FROM orders WHERE ((orderStatus>2 AND orderStatus<5) OR (orderStatus>6 AND orderStatus<=9) OR (orderStatus=10 OR orderStatus=12)) AND year(orderDate)="& pcvThisYear & " AND idcustomer=" & pIdCustomer &";"
			set rstemp=Server.CreateObject("ADODB.Recordset") 
			rstemp.Open query, conntemp
			if err.number <> 0 then
				response.redirect "techErr.asp?error="& Server.Urlencode("Error calculating order total: "&Err.Description) 
			end If
			if not rstemp.eof then
				TotalOrderedYear=rstemp("TotalAmount")
				TotalLessRMA=rstemp("TotalLessRMA")
			end if
			if isNull(TotalOrderedYear) then
				TotalOrderedYear = TotalLessRMA
				if isNull(TotalOrderedYear) then
					TotalOrderedYear=0
				end if
			end if
			set rstemp=nothing
			

		' Get order totals for last 12 months
			pcvThisYear = year(date)
			pcvLastYear = (int(pcvThisYear)-1)
			' leap year check and fix
			if month(date)  = 2 and day(date) = 29 then 			 
			        pcvYearAgo = month(date) & "/28/" & pcvLastYear
			 else
			 		pcvYearAgo = month(date) & "/" & day(date) & "/" & pcvLastYear			
			End if 
			if SQL_Format="1" then
				DateVarArray=split(pcvYearAgo,"/")
				pcvYearAgo=(DateVarArray(1)&"/"&DateVarArray(0)&"/"&DateVarArray(2))
			end if
			TotalOrdered12=0
			if scDB="SQL" then
				query="SELECT sum(total-rmaCredit) As TotalAmount, Sum(Total) AS TotalLessRMA FROM orders WHERE ((orderStatus>2 AND orderStatus<5) OR (orderStatus>6 AND orderStatus<=9) OR (orderStatus=10 OR orderStatus=12)) AND orderDate >='"& pcvYearAgo & "' AND idcustomer=" & pIdCustomer &";"
			else
				query="SELECT sum(total-rmaCredit) As TotalAmount, Sum(Total) AS TotalLessRMA FROM orders WHERE ((orderStatus>2 AND orderStatus<5) OR (orderStatus>6 AND orderStatus<=9) OR (orderStatus=10 OR orderStatus=12)) AND orderDate >=#"& pcvYearAgo & "# AND idcustomer=" & pIdCustomer &";"
			end if
			set rstemp=Server.CreateObject("ADODB.Recordset") 
			rstemp.Open query, conntemp
			if err.number <> 0 then
				response.redirect "techErr.asp?error="& Server.Urlencode("Error calculating order total for last 12 months: "&Err.Description) 
			end If
			if not rstemp.eof then
				TotalOrdered12=rstemp("TotalAmount")
				TotalLessRMA=rstemp("TotalLessRMA")
			end if
			if isNull(TotalOrdered12) then
				TotalOrdered12 = TotalLessRMA
				if isNull(TotalOrdered12) then
					TotalOrdered12=0
				end if
			end if
			set rstemp=nothing

	
		' Count total orders except for incomplete
			Dim pcvIntTotalOrders
			query = "SELECT Count(*) AS intTotal FROM orders WHERE orderStatus>2 AND idcustomer=" & pIdCustomer
			Set rsTemp = Server.CreateObject("ADODB.Recordset")
			rsTemp.Open query, conntemp
			pcvIntTotalOrders = rsTemp("intTotal")
			set rsTemp = nothing
		
		' Count incomplete orders
			Dim pcvIntIncOrders
			query = "SELECT Count(*) AS intTotal FROM orders WHERE orderStatus=1 AND idcustomer=" & pIdCustomer
			Set rsTemp = Server.CreateObject("ADODB.Recordset")
			rsTemp.Open query, conntemp
			pcvIntIncOrders = rsTemp("intTotal")
			set rsTemp = nothing


		' Get orders placed by this customer
		
			query="SELECT orders.idorder, orderDate, total, orderstatus,orders.pcOrd_PaymentStatus,orders.comments,orders.admincomments,orders.details,orders.rmaCredit FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND customers.idcustomer="& pIdCustomer &" ORDER BY "& strORD &" "& strSort
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			rstemp.CursorLocation=adUseClient
			rstemp.CacheSize=iPageSize
			rstemp.PageSize=iPageSize
			rstemp.Open query, conntemp
			
			if err.number <> 0 then
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
			end If
			
	if rstemp.eof then
		dim showLinks
		showLinks = 1
	else
	
		rstemp.MoveFirst
		' get the max number of pages
		Dim iPageCount
			iPageCount=rstemp.PageCount
			If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
		If iPageCurrent < 1 Then iPageCurrent=1
		
		' set the absolute page
		rstemp.AbsolutePage=iPageCurrent
		Dim count
		Count=0
	end if

%>
<!--#include file="AdminHeader.asp"-->
<script>
	function openwin(file)
	{
		msgWindow=open(file,'win1','scrollbars=yes,resizable=yes,width=500,height=400');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
</script>
<% 
IF showlinks = 1 THEN '// NO ORDERS
%>
	<table class="pcCPcontent">
		<tr>
			<td>
			<div class="pcCPmessage">There are no orders associated with this customer account.</div>
			<ul>
			<li><a href="modCusta.asp?idcustomer=<%=pIdCustomer%>">View customer details</a></li>
			<li><a href="adminPlaceOrder.asp?idcustomer=<%=pidcustomer%>" target="_blank">Place an order on behalf of this customer</a><br><br></li>
			<li><a href="viewCusta.asp">Look for another customer</a></li>
			<li><a href="viewCustb.asp?mode=ALL">View all customers</a></li>
			<li><a href="javascript: history.go(-1)">Back</a></li>
			</ul>
			</td>
		</tr>
	</table>
<% 
ELSE
%>
	<table class="pcCPcontent">
		<tr>
			<td colspan="2">
            <div style="float: right; padding-top: 5px;"><a href="modCusta.asp?idcustomer=<%=pIdCustomer%>">Edit Customer</a></b><% if pcf_GetCustType(pidcustomer)=0 then%>&nbsp;|&nbsp;<a href="adminPlaceOrder.asp?idcustomer=<%=pidcustomer%>" target="_blank">Place Order</a><%end if%></div>
            <h2>Customer: <strong><%=pname%></strong><%if pcompany<>"no" then response.write " (" & pcompany & ")" end if%></h2></td>
		</tr>
		<tr>
			<td width="25%" nowrap>Number of orders:</td>
			<td width="75%"><%=pcvIntTotalOrders%> (<em>excluding incomplete orders</em>) &nbsp;|&nbsp;Incomplete orders: <%=pcvIntIncOrders%></td>
		</tr>
		<tr>
			<td nowrap>Total amount ordered to date:</td>
            <td><strong><%=scCurSign & money(TotalOrdered)%></strong>&nbsp;&nbsp;|&nbsp;This year: <strong><%=scCurSign & money(TotalOrderedYear)%></strong>&nbsp;&nbsp;|&nbsp;Last 12 months: <strong><%=scCurSign & money(TotalOrdered12)%></strong></td>
		</tr>
	</table>
	
	<table class="pcCPcontent" style="margin-top: 10px;">
	<tr> 
        <th nowrap align="center">Status</th>
        <th width="2%" nowrap><a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&iPageCurrent=<%=iPageCurrent%>&order=idorder&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&iPageCurrent=<%=iPageCurrent%>&order=idorder&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Order#</th>
        <th width="2%" nowrap><a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&iPageCurrent=<%=iPageCurrent%>&order=orderDate&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending"></a><a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&iPageCurrent=<%=iPageCurrent%>&order=orderDate&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending"></a>&nbsp;Date</th>
        <th nowrap>Total</th>
        <th colspan="2" nowrap>Products Ordered</th>
	</tr>
    <tr>
        <td colspan="6" class="pcCPspacer"></td>
    </tr>

	<%
	do while not rstemp.eof And Count < rstemp.PageSize
	
		pidorder=rstemp("idorder")
		porderDate=rstemp("orderDate")
		porderDate=ShowDateFrmt(porderDate)
		ptotal=rstemp("total")
		prmaCredit=rstemp("rmaCredit")
			'// Calculate total adjusted for credits
			if trim(prmaCredit)="" or IsNull(prmaCredit) then
				prmaCredit=0
			end if
			pTotalAdj=pTotal-prmaCredit
		porderstatus=rstemp("orderStatus")
		'Start SDBA
		pcv_PaymentStatus=rstemp("pcOrd_PaymentStatus")
		if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
			pcv_PaymentStatus=0
		end if
		'End SDBA
		pcv_custcomments=trim(rstemp("comments"))
		pcv_admcomments=trim(rstemp("admincomments"))
		pcv_details=trim(rstemp("details"))
			if len(pcv_details)>180 then
				pcv_details=left(pcv_details,180) & "..."
			end if
		pcv_details=replace(pcv_details," ||",""&scCurSign&"")
	%>
	<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
		<td align="center" valign="top" width="5%"><!--#include file="inc_orderStatusIcons.asp"--></td>
    	<td align="center" valign="top" width="5%"><% if porderstatus="1" then %><a href="OrdDetailsIncomplete.asp?id=<%=pidOrder%>"><% else %><a href="Orddetails.asp?id=<%=pidOrder%>"><% end if %><strong><%response.write (scpre+int(pIdOrder))%></strong></a></td>
		<td valign="top" width="5%"><%response.write pOrderDate%></td>
		<td valign="top" width="5%"><%response.write(scCurSign & money(ptotal))%></td>
    	<td valign="top" width="70%">
    
			<%
                query="SELECT ProductsOrdered.idProduct, ProductsOrdered.idOrder, products.description, products.sku, products.idProduct, orders.idOrder FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idOrder=" &pIdOrder
                set rs=Server.CreateObject("ADODB.Recordset")
                set rs=conntemp.execute(query)
            
                While Not rs.EOF
                pIdProduct=rs("idProduct") 
                pSku=rs("sku")
                pDescription=rs("description")
                %>
                <div style="margin-bottom: 3px;"><%=psku%> - <%=pDescription %></div>
                <%
                rs.MoveNext
                Wend
                set rs = nothing
            %>

        </td>
        <td align="right" nowrap valign="top" width="10%">
            <% if porderstatus="1" then %>
             <a href="OrdDetailsIncomplete.asp?id=<%=pidOrder%>">Review</a>
            <% else %>
             <a href="Orddetails.asp?id=<%=pidOrder%>"><img src="images/pcIconNext.jpg" width="12" height="12" alt="View and Process"></a>&nbsp;<a href="OrdInvoice.asp?id=<%response.write pIdOrder%>" target="_blank"><img src="images/print_xsmall.gif" alt="Printer Friendly Version" border="0"></a>
            <% end if %>
            <%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=pidOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%>
        </td>
	</tr>
														
	<%
	 rstemp.movenext
	 Count=Count + 1
	Loop
	set rstemp=nothing
	%>
    <tr>
        <td colspan="6" class="pcCPspacer"></td>
    </tr>
	<tr>
	<td colspan="6">
	<form method="post" action="" name="" class="pcForms">
	<%=("Page "& iPageCurrent & " of "& iPageCount)%>
    <br />
	<%
	'Display Next / Prev buttons
	if iPageCurrent > 1 then
	'We are not at the beginning, show the prev button %>
		<a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&status=<%=pOrderStatus%>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a> 
	<% end If
	If iPageCount <> 1 then
		For I=1 To iPageCount
			If I=iPageCurrent Then %>
				<%=I%> 
			<% Else %>
				<a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&status=<%=pOrderStatus%>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"><%=I%></a> 
			<% End If %>
		<% Next %>
	<% end if %>
	<% if CInt(iPageCurrent) <> CInt(iPageCount) then
	'We are not at the end, show a next link %>
		<a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>&status=<%=pOrderStatus%>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
	<% end If %>
	</form>
	</td>
	</tr>
</table>
<% 
END IF
call closeDb()
%><!--#include file="AdminFooter.asp"-->