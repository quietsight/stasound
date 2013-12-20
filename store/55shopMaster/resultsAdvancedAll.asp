<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
response.Buffer=true
Section="orders"
%>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/GoogleCheckoutConstants.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->

<%
'// Page Options - START
'//
'// ORDERED ITEMS: set the number to show
'// If = 0, then only the number of items ordered will be shown
Dim pcIntOrderedItems, pcIntOrderedItemsOverride
pcIntOrderedItems=3 '// Change this variable to change the number of items shown
pcIntOrderedItemsOverride=getUserInput(request.querystring("hideItemsOrdered"),1)
if not validNum(pcIntOrderedItemsOverride) then 
	if not validNum(session("pcIntOrderedItemsOverride")) then
		session("pcIntOrderedItemsOverride")=0
	end if
else
	session("pcIntOrderedItemsOverride")=pcIntOrderedItemsOverride
end if
if session("pcIntOrderedItemsOverride")=1 then pcIntOrderedItems=0
if session("pcIntOrderedItemsOverride")=0 then pcIntOrderedItems=pcIntOrderedItems

'// Page Options - END

'// Check for archived orders
pcInt_OrdArchived=request("pcIntArchived")
if not validNum(pcInt_OrdArchived) then pcInt_OrdArchived=0

if pcInt_OrdArchived=0 then
	pageTitle="View Orders"
	else
	pageTitle="View Archived Orders"
end if
pageIcon="pcv4_icon_orders.gif"
pcInt_ShowOrderLegend=1
%>

<!--#include file="AdminHeader.asp"-->
<script>
	function openwin(file)
	{
		msgWindow=open(file,'win1','scrollbars=yes,resizable=yes,width=500,height=400');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
	
	function CalPop(sInputName)
	{
		window.open('../Calendar/Calendar.asp?N=' + escape(sInputName) + '&DT=' + escape(window.eval(sInputName).value), 'CalPop','toolbar=0,width=378,height=225' );
	}
</script>
<% Session.LCID = 1033 %>
<%
dim query, conntemp, rstemp, strORD

Const iPageSize=50

Dim iPageCurrent

if request.querystring("iPageCurrent")="" or request.querystring("iPageCurrent")="0" then
	iPageCurrent=1
else
	iPageCurrent=Request.QueryString("iPageCurrent")
end if

if request("B1")<>"" OR request.querystring("iPageCurrent")="" or request.querystring("iPageCurrent")="0" then
	if session("CP_OrdSrcPages")>"0" then
		For i=1 to Clng(session("CP_OrdSrcPages"))
			session("CP_OrdSrcPage"&i)=""
		Next
	end if
	session("CP_OrdSrcPages")=0
else
	if request("curpage")<>"" then
		tmpPre=request("curpage")
		session("CP_OrdSrcPage"&tmpPre)=request("pageval")
	end if
end if
pcChkArr=split(session("CP_OrdSrcPage"&iPageCurrent),",")

'sorting order
strORD=request("order")
if strORD="" then
	strORD="orderDate DESC, idOrder"
End If

strSort=request("sort")
if strSort="" Then
	strSort="DESC"
End If

query1=""

OType=request("OType")
if OType="" then
	OType="0"
end if
if (OType<>"0") then
	query1= " orderstatus=" & OType
	else
	query1= " orderstatus>1"
end if
Dim pcShowIncompleteOnly
pcShowIncompleteOnly=0
if OType="1" then
	pcShowIncompleteOnly = 1
end if

'// Ajust for payment status
pcv_PayType=request("PayType")
if (pcv_PayType<>"") then
	query1= query1 & " AND pcOrd_PaymentStatus=" & pcv_PayType
end if

'// Ajust for archived orders
query1= query1 & " AND pcOrd_Archived=" & pcInt_OrdArchived

err.number=0
FromDate=request("fromdate")
PassFromDate=FromDate
ToDate=request("todate")
PassToDate=ToDate

if FromDate<>"" then
	FromDate=GetDateGUIDatabase(FromDate, 1)
else
	call opendb()
	query="SELECT TOP 1 orders.orderDate FROM orders WHERE orders.orderStatus>1 ORDER BY orderDate ASC;"
	set rstemp=Server.CreateObject("ADODB.Recordset") 
	set rstemp=conntemp.execute(query)
	if NOT rstemp.eof then
		FromDate=rstemp("orderDate")
		if scDateFrmt="DD/MM/YY" then
			FromDate=day(FromDate)&"/"&month(FromDate)&"/"&Year(FromDate)
			PassFromDate=FromDate
		else
			if SQL_Format="1" then
				FromDate=day(FromDate) & "/" & month(FromDate) & "/" & year(FromDate)
			else
				FromDate=month(FromDate) & "/" & day(FromDate) & "/" & year(FromDate)
			end if
		end if
	end if
	call closedb()
end if

if ToDate<>"" then
	ToDate=GetDateGUIDatabase(ToDate, 1)
else
	if SQL_Format="1" then
		ToDate=day(date()) & "/" & month(date()) & "/" & year(date())
	else
		ToDate=month(date()) & "/" & day(date()) & "/" & year(date())
	end if
end if

if request("dd")="1" then
	dtFromDate=Date()-13
	dtFromDateStr=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
	if SQL_Format="1" then
		FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
	else
		FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
	end if
	dtToDate=Date()
	dtToDateStr=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
	if SQL_Format="1" then
		ToDate=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
	else
		ToDate=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
	end if
	if scDateFrmt="DD/MM/YY" then
		dtFromDateStr=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
		dtToDateStr=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
	end if
	PassFromDate=dtFromDateStr
	PassToDate=dtToDateStr
end if

if (FromDate<>"") and (IsDate(FromDate)) then
	if scDB="Access" then
		query1= query1 & " AND orderDate>=#" & FromDate & "#"
	else
		query1= query1 & " AND orderDate>='" & FromDate & "'"
	end if
end if

if (ToDate<>"") and (IsDate(ToDate)) then
	if scDB="Access" then
		query1= query1 & " AND orderDate<=#" & ToDate & "#"
	else
		query1= query1 & " AND orderDate<='" & ToDate & "'"
	end if
end if



call openDb()

' Choose the records to display
Dim srcVar
	SqlVar="SELECT orders.idOrder, orders.idCustomer, orders.paymentDetails, orders.paymentCode, orders.orderstatus, orders.orderDate, orders.total, orders.rmaCredit, orders.pcOrd_PaymentStatus, customers.name, customers.lastName, customers.customerCompany, orders.comments, orders.admincomments FROM orders, customers WHERE orders.idCustomer=customers.idCustomer AND " & query1 & " ORDER BY "& strORD &" "& strSort
%>

<% 
set rstemp=Server.CreateObject("ADODB.Recordset")     

rstemp.CursorLocation=adUseClient
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open SqlVar, conntemp

if err.number <> 0 then
  response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
end If
if NOT rstemp.eof then 
	rstemp.MoveFirst
	' set the absolute page
	rstemp.AbsolutePage=iPageCurrent
end if
' get the max number of pages
Dim iPageCount
iPageCount=rstemp.PageCount
session("CP_OrdSrcPages")= iPageCount
If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
If iPageCurrent < 1 Then iPageCurrent=1

if rstemp.eof then 
	presults="0"
else
	%>
	<table class="pcCPcontent">
		<tr> 
			<td>
				<%
				if PassFromDate<>"" then %>
                From: <strong><%=PassFromDate%></strong>
				<%end if%>
				&nbsp;
				<%if PassToDate="" then
					PassToDate=date()
					if scDateFrmt="DD/MM/YY" then
						PassToDate=day(PassToDate)&"/"&month(PassToDate)&"/"&Year(PassToDate)
					end if %>
				<% end if%>
				To: <strong><%=PassToDate%></strong>
                &nbsp;|&nbsp;
				<% ' Showing total number of pages found and the current page number %>
				Displaying Page <b><%=iPageCurrent%></b> of <b><%=iPageCount%></b>
                &nbsp;|&nbsp;
				Total Records Found: <b><%=rstemp.RecordCount%></b>
                &nbsp;|&nbsp;
                <% if session("pcIntOrderedItemsOverride")=0 then %>
                <a href="resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=iPageCurrent%>&order=<%=strORD%>&sort=<%=strSort%>&OType=<%=OType%>&paytype=<%=pcv_PayType%>&hideItemsOrdered=1">Hide Ordered Items Details</a>
                <% else %>
                <a href="resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=iPageCurrent%>&order=<%=strORD%>&sort=<%=strSort%>&OType=<%=OType%>&paytype=<%=pcv_PayType%>&hideItemsOrdered=0">Show Ordered Items Details</a>                
                <% end if %>
			</td>
		</tr>
	</table>
<% end if %>
<form name="checkboxform" method="post" target="_blank" class="pcForms">
<table class="pcCPcontent">
	<tr>
        <td colspan="11" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>

			<% if pcShowIncompleteOnly = 1 then %>
                <h2>You are viewing Incomplete Orders</h2>
                <div style="margin-bottom: 15px;"><a href="http://wiki.earlyimpact.com/productcart/orders_status#incomplete_orders" target="_blank">Learn about <strong>incomplete orders</strong></a>.</div>
            <% end if %>
		</td>
	</tr>
	<tr> 
		<th align="center" nowrap><a href="resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=iPageCurrent%>&order=orderstatus&sort=ASC"><img src="images/sortasc.gif" alt="Sort Ascending"></a><a href="resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=iPageCurrent%>&order=orderstatus&sort=DESC"><img src="images/sortdesc.gif" alt="Sort Descending"></a>
        <a href="resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=iPageCurrent%>&order=pcOrd_PaymentStatus&sort=ASC"><img src="images/sortasc.gif" alt="Sort Ascending"></a><a href="resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=iPageCurrent%>&order=pcOrd_PaymentStatus&sort=DESC"><img src="images/sortdesc.gif" alt="Sort Descending"></a></th>
		<th align="center" nowrap><a href="resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate %>&iPageCurrent=<%=iPageCurrent%>&order=orderDate&sort=ASC"><img src="images/sortasc.gif" alt="Sort Ascending"></a><a href="resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate%>&iPageCurrent=<%=iPageCurrent%>&order=orderDate&sort=DESC"><img src="images/sortdesc.gif" alt="Sort Descending"></a> Date</th>
		<th nowrap>ID</th>
		<th nowrap>Total</th>
		<th nowrap>Customer</th>
        <th nowrap>Items Ordered</th>
  		<th nowrap>Paid By</th>
		<th colspan="4" nowrap>
			<% if pcShowIncompleteOnly <> 1 then %>
				<div style="text-align: right;" class="pcSmallText"><a href="batchprocessorders.asp">Batch Process</a></div>
			<% end if %>
		</th>
	</tr>
	<tr>
		<td colspan="11" class="pcCPspacer"></td>
	</tr>
	<% 
	Dim mcount
	mcount=0
	If rstemp.EOF Then %>
	<tr>
		<td colspan="11" align="center">
			<div class="pcCPmessage">No Results Found. <% if pcInt_OrdArchived=0 then %>If you have archived orders, check "Archived Orders Only" in the Advanced Filters below to locate them.<% end if %></div>
		</td>
	</tr>
	<% Else
	' Showing relevant records
	Dim rcount, i, x
	
	Do while (not rstemp.eof) and (mcount<rstemp.PageSize)
		pidOrder=rstemp("idOrder")
		pidCustomer=rstemp("idCustomer")
		ppaymentDetails=trim(rstemp("paymentDetails"))
			pcArrayPayment = split(ppaymentDetails,"||")
			pcPaymentType=pcArrayPayment(0)
		ppaymentCode=rstemp("paymentCode")
		porderstatus=rstemp("orderstatus")
		porderDate=rstemp("orderDate")
		ptotal=rstemp("total")
			pc_rmaCredit=rstemp("rmaCredit")
			if trim(pc_rmaCredit)="" or IsNull(pc_rmaCredit) then
				pc_rmaCredit=0
			end if
			pTotal=pTotal-pc_rmaCredit
		pcv_PaymentStatus=rstemp("pcOrd_PaymentStatus")
		if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
			pcv_PaymentStatus=0
		end if
		pfName=rstemp("name")
		plName=rstemp("lastName")
		pcv_CustomerCompany=rstemp("customerCompany")
			if trim(pcv_CustomerCompany)<>"" and not IsNull(pcv_CustomerCompany) then
				pcv_customerName = pfName & " " & plName & "<br />(" & pcv_CustomerCompany & ")"
				else
				pcv_customerName = pfName & " " & plName				
			end if
		pcv_custcomments=trim(rstemp("comments"))
		pcv_admcomments=trim(rstemp("admincomments"))
		rcount=i
		If currentPage > 1 Then
			For x=1 To (currentPage - 1)
				rcount=10 + rcount
			Next
		End If
		
		'// DeActivate Sections for Google Checkout
		pcv_strDeactivateStatus=0						
		if ppaymentCode="Google" then
			pcv_strDeactivateStatus=1
		end if	
		
		If Not rstemp.EOF Then
			mcount=mcount+1 %>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td align="center" valign="top" nowrap> 
					<!--#include file="inc_orderStatusIcons.asp"--> 
				</td>
				<td valign="top"><%=ShowDateFrmt(porderDate)%></td>
				<td valign="top">
				<% if porderstatus="1" then %>
					 <a href="OrdDetailsIncomplete.asp?id=<%=pidOrder%>"><%=(scpre+int(pidOrder))%></a>
					<% else %>
					 <a href="Orddetails.asp?id=<%=pidOrder%>"><%=(scpre+int(pidOrder))%></a>
				<% end if %>
                </td>
				<td valign="top">
				<% if porderstatus="1" then %>
					 <a href="OrdDetailsIncomplete.asp?id=<%=pidOrder%>"><%=scCurSign&money(ptotal)%></a>
					<% else %>
					 <a href="Orddetails.asp?id=<%=pidOrder%>"><%=scCurSign&money(ptotal)%></a>
				<% end if %>
				</td>
				<td valign="top"><a href="modcusta.asp?idcustomer=<%=pidCustomer%>"><%=pcv_customerName%></a></td>
                <td valign="top" style="color: #333; font-size: 11px;">
                <%
                query="SELECT products.description, products.sku, orders.idorder, ProductsOrdered.idOrder, ProductsOrdered.quantity FROM products, orders, ProductsOrdered WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idOrder=" &pIdOrder
                set rs=Server.CreateObject("ADODB.Recordset")
                set rs=conntemp.execute(query)
				if rs.eof then ' Empty recordset. There must be a problem in the database with this order.
					response.write "Not Available"
				else
	            	Dim pcvProductsArray, pcIntCount, pcIntTotalQuantity
					pcvProductsArray=rs.getRows()
					pcvProductsArrayCount=ubound(pcvProductsArray,2)
					set rs=nothing
					if not validNum(pcIntOrderedItems) then pcIntOrderedItems=0
					if pcIntOrderedItems=0 then
						pcIntCount=0
						pcIntTotalQuantity=0
						While pcIntCount<=pcvProductsArrayCount
							pcIntTotalQuantity=pcIntTotalQuantity+pcvProductsArray(4,pcIntCount)
							pcIntCount=pcIntCount+1
						Wend					
						response.write pcvProductsArrayCount+1 & " item(s) [" & pcIntTotalQuantity & " qty]"
					else
						pcIntCount=0
						While pcIntCount<=pcvProductsArrayCount and pcIntCount<pcIntOrderedItems
							%>
							<div style="margin-bottom: 3px;"><% if pcvProductsArray(4,pcIntCount)>1 then %>(<%=pcvProductsArray(4,pcIntCount)%>) <% end if %><%=pcvProductsArray(0,pcIntCount)%> <span style="color:#999;">(<%=pcvProductsArray(1,pcIntCount)%>)</span></div>
							<%
							pcIntCount=pcIntCount+1
						Wend
						if pcvProductsArrayCount=>pcIntCount then
						%>
							<div style="margin-bottom: 3px;"><a href="Orddetails.asp?id=<%=pidOrder%>&activetab=2">More...</a></div>
						<%
						end if
					end if
				end if ' Empty recordset
            	%>
                </td>
				<td valign="top" style="color: #333; font-size: 11px;"><%=pcPaymentType%></td>
				<td valign="top" align="center" nowrap class="cpLinksList">
					<% if porderstatus="1" then %>
					 <a href="OrdDetailsIncomplete.asp?id=<%=pidOrder%>">Review/Reset Status</a>
					<% else %>
					 <a href="Orddetails.asp?id=<%=pidOrder%>">View &amp; Process</a>
					<% end if %>
					<%if pcv_custcomments<>"" or pcv_admcomments<>"" then%>&nbsp;<a href="javascript:openwin('popup_viewOrdCustComments.asp?idorder=<%=pidOrder%>');"><img src="images/pcv3_infoIcon.gif" border="0" alt="Click here to view order comments"></a><%end if%>
				</td>
				<td valign="top" align="center" class="cpLinksList">
				<% if porderstatus>1 AND pcv_strDeactivateStatus<>1 then %>
					<a href="AdminEditOrder.asp?ido=<%=pidOrder%>">Edit</a>
				<% else %>
					&nbsp;
				<% end if %>
				</td>
				<td valign="top" align="center" nowrap class="cpLinksList">
					<% if porderstatus>1 then %>
					<a href="OrdInvoice.asp?id=<%=pidOrder%>" target="_blank"><img src="images/print_xsmall.gif" width="12" height="11" alt="Print" title="Printer Friendly (Invoice style)"></a>
					<% else %>
					&nbsp;
					<% end if %>
				</td>
				<td valign="top" align="center" nowrap>
				<input type="hidden" name="idord<%=mcount%>" value="<%=pidOrder%>">
				<% if porderstatus>1 then %>
				<input type="checkbox" name="check<%=mcount%>" value="1" class="clearBorder" 
				<%For m=0 to ubound(pcChkArr)
					if pcChkArr(m)<>"" then
						if Clng(pcChkArr(m))=Clng(pidOrder) then%>
						checked
						<%end if
					end if
				Next%> onclick="javascript:TestChecked();">
				<% else %>
				&nbsp;
				<% end if %>
				</td>
			</tr>
		<% rstemp.MoveNext
		End If
	Loop
	%>
	<input type="hidden" name="count" value="<%=mcount%>">
<%End If %>
<tr> 
<td colspan="11" align="right">
<input type="hidden" name="curpage" value="<%=iPageCurrent%>">
<input type="hidden" name="pageval" value="<%=session("CP_OrdSrcPage"&iPageCurrent)%>">
<%if mcount>0 and pcShowIncompleteOnly<>1 then%>
<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a><br>
<br>
<%
if pcInt_OrdArchived=0 then
%>
<INPUT type="button" value="Archive" name="button3" onclick="return OnButton3();">&nbsp;
<%
else
%>
<INPUT type="button" value="Unarchive" name="button4" onclick="return OnButton4();">&nbsp;
<%
end if
%>
<INPUT type="button" value="Print Invoices/Packing Slips" name="button1" onclick="return OnButton1();">&nbsp;
<INPUT type="button" value="Print Pick List" name="button2" onclick="return OnButton2();">
&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=443')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
<script language="JavaScript">
<!--
function OnButton1()
{
	var countChecked=0;
	for (var j = 1; j <= <%=mcount%>; j++) {
	box = eval("document.checkboxform.check" + j); 
		if (box.checked == true) countChecked =countChecked+1;
   }
   if(countChecked<1)
   {
	alert("Please select at least one order from the list.");
	return false;
   }
   else
   {
	document.checkboxform.action = "batchprint.asp"
	document.checkboxform.target = "_blank";	// Open in a new window
	document.checkboxform.submit();			// Submit the page
	return true;
	}
}

function OnButton2()
{
	var countChecked=0;
	for (var j = 1; j <= <%=mcount%>; j++) {
	box = eval("document.checkboxform.check" + j); 
		if (box.checked == true) countChecked =countChecked+1;
   }
   if(countChecked<1)
   {
	alert("Please select at least one order from the list.");
	return false;
   }
   else
   {
	document.checkboxform.action = "batchPrintPickList.asp"
	document.checkboxform.target = "_blank";	// Open in a new window
	document.checkboxform.submit();			// Submit the page
	return true;
	}
}

function OnButton3()
{
	var countChecked=0;
	for (var j = 1; j <= <%=mcount%>; j++) {
	box = eval("document.checkboxform.check" + j); 
		if (box.checked == true) countChecked =countChecked+1;
   }
   if(countChecked<1)
   {
	alert("Please select at least one order to archive.");
	return false;
   }
   else
   {
	document.checkboxform.action = "BatchArchiveOrders.asp?action=archive"
	document.checkboxform.target = "_self";
	document.checkboxform.submit();			// Submit the page
	return true;
	}
}

function OnButton4()
{
	var countChecked=0;
	for (var j = 1; j <= <%=mcount%>; j++) {
	box = eval("document.checkboxform.check" + j); 
		if (box.checked == true) countChecked =countChecked+1;
   }
   if(countChecked<1)
   {
	alert("Please select at least one order to unarchive.");
	return false;
   }
   else
   {
	document.checkboxform.action = "BatchArchiveOrders.asp?action=unarchive"
	document.checkboxform.target = "_self";
	document.checkboxform.submit();			// Submit the page
	return true;
	}
}

function checkAll() {
for (var j = 1; j <= <%=mcount%>; j++) {
box = eval("document.checkboxform.check" + j); 
if (box.checked == false) box.checked = true;
   }
TestChecked();
}

function uncheckAll() {
for (var j = 1; j <= <%=mcount%>; j++) {
box = eval("document.checkboxform.check" + j); 
if (box.checked == true) box.checked = false;
   }
TestChecked();
}

function TestChecked() {
var tmpStr="";
for (var j = 1; j <= <%=mcount%>; j++) {
	box = eval("document.checkboxform.check" + j); 
	if (box.checked == true) tmpStr=tmpStr + eval("document.checkboxform.idord" + j).value + ",";
}
document.checkboxform.pageval.value=tmpStr;
}
//-->
</script>
<%end if%>
</td>
</tr>
</table>
</form>          

<%
if pResults<>"0" and iPageCount>1 Then
%>
<table class="pcCPcontent">
	<tr> 
		<td> 
			<%Response.Write("Page "& iPageCurrent & " of "& iPageCount & "<br />")%>
			<%'Display Next / Prev buttons
			if iPageCurrent > 1 then
					'We are not at the beginning, show the prev button %>
					 <a href="javascript:location='resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate%>&ToDate=<%=PassToDate%>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>&OType=<%=OType%>&paytype=<%=pcv_PayType%>&curpage=<%=iPageCurrent%>&pageval='+eval('document.checkboxform.pageval').value;"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a>
			<% end If
			If iPageCount <> 1 then
				For I=1 To iPageCount
					If int(I)=int(iPageCurrent) Then %>
						<%=I%> 
					<% Else %>
					<a href="javascript:location='resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate %>&ToDate=<%=PassToDate %>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>&OType=<%=OType%>&paytype=<%=pcv_PayType%>&curpage=<%=iPageCurrent%>&pageval='+eval('document.checkboxform.pageval').value;" style="text-decoration:underline;"><%=I%></a>
					<% End If %>
				<% Next %>
			<% end if %>
			<% if CInt(iPageCurrent) <> CInt(iPageCount) then
			'We are not at the end, show a next link %>
			<a href="javascript:location='resultsadvancedall.asp?pcIntArchived=<%=pcInt_OrdArchived%>&FromDate=<%=PassFromDate %>&ToDate=<%=PassToDate %>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>&OType=<%=OType%>&paytype=<%=pcv_PayType%>&curpage=<%=iPageCurrent%>&pageval='+eval('document.checkboxform.pageval').value;"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a>
			<% end If 
			call closeDb()
			%>
		</td>
	</tr>
	<tr>
		<td><hr></td>
	</tr>          
</table>
<% end if %> 


<table class="pcCPcontent" style="width:auto;">
	<tr> 
		<td><h2>Advanced Filters</h2></td>
	</tr>
	<tr> 
		<td>Filter orders by date and status.</td>
	</tr>
	<tr> 
		<td>    
		<form action="resultsAdvancedAll.asp" name="advsearch" class="pcForms">
			<table class="pcCPcontent">
				<tr>
					<td align="right">Date From:</td>
					<td nowrap="nowrap">
                    <input type="text" name="fromdate" value="<%=PassFromDate%>" size="10">&nbsp;<a href="javascript:CalPop('document.advsearch.fromdate');"><img SRC="../Calendar/icon_Cal.gif" border="0" align="absbottom"></a>
                    &nbsp;To: <input type="text" name="todate" value="<%=PassToDate%>" size="10">&nbsp;<a href="javascript:CalPop('document.advsearch.todate');"><img SRC="../Calendar/icon_Cal.gif" border="0" align="absbottom"></a>
					</td>
				</tr>
				<tr>
					<td align="right" valign="top" nowrap="nowrap">Order Status:</td>
					<td>
						<select name="otype">
							<option value="0" <%if OType="0" then%>selected<%end if%>>All</option>
							<option value="1" <%if OType="1" then%>selected<%end if%>>Incomplete</option>
							<option value="2" <%if OType="2" then%>selected<%end if%>>Pending</option>
							<option value="3" <%if OType="3" then%>selected<%end if%>>Processed</option>
							<option value="7" <%if OType="7" then%>selected<%end if%>>Partially Shipped</option>
							<option value="8" <%if OType="8" then%>selected<%end if%>>Shipping</option>
							<option value="4" <%if OType="4" then%>selected<%end if%>>Shipped</option>
							<option value="5" <%if OType="5" then%>selected<%end if%>>Canceled</option>
							<option value="9" <%if OType="9" then%>selected<%end if%>>Partially Return</option>
							<option value="6" <%if OType="6" then%>selected<%end if%>>Return</option>							
							<% if GOOGLEACTIVE<>0 then %>
							<option value="10" <%if OType="10" then%>selected<%end if%>>Delivered</option>
							<option value="11" <%if OType="11" then%>selected<%end if%>>Will Not Deliver</option>
							<option value="12" <%if OType="12" then%>selected<%end if%>>Archived</option>
							<% end if %>						
						</select>
					</td>
				</tr>
				<tr>
					<td align="right" valign="top" nowrap="nowrap">Payment Status:</td>
					<td>
						<select name="PayType">
							<option value="" <%if pcv_PayType="" then%>selected<%end if%>>All</option>
							<option value="0" <%if pcv_PayType="0" then%>selected<%end if%>>Pending</option>
							<option value="1" <%if pcv_PayType="1" then%>selected<%end if%>>Authorized</option>
							<option value="2" <%if pcv_PayType="2" then%>selected<%end if%>>Paid</option>
							<% if GOOGLEACTIVE<>0 then %>
							<option value="3" <%if pcv_PayType="3" then%>selected<%end if%>>Declined</option>
							<option value="4" <%if pcv_PayType="4" then%>selected<%end if%>>Cancelled</option>
							<option value="5" <%if pcv_PayType="5" then%>selected<%end if%>>Cancelled By Google</option>							
							<option value="7" <%if pcv_PayType="7" then%>selected<%end if%>>Charging</option>
							<% end if %>
							<option value="6" <%if pcv_PayType="6" then%>selected<%end if%>>Refunded</option>
							<option value="8" <%if pcv_PayType="8" then%>selected<%end if%>>Voided</option>
						</select>
					</td>
				</tr>
                <tr>
                	<td align="right" valign="top" nowrap>Archived Orders Only:</td>
                    <td>
                    	<input type="checkbox" name="pcIntArchived" value="1" <%if pcInt_OrdArchived="1" then%>checked<%end if%> class="clearBorder">
                    </td>
                </tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<td colspan="2" align="center">
						<input type="submit" name="B1" value="Search Orders" class="submit2">
						&nbsp;
						<input type="button" name="Button" value="Back" onClick="location='invoicing.asp'">
					</td>
				</tr>
			</table>
		</form>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->