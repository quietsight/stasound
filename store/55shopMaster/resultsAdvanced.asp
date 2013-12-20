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

pageTitle="View Orders"
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

'sorting order
strORD=request("order")
if strORD="" then
	strORD="orderDate DESC, idOrder"
End If

strSort=request("sort")
if strSort="" Then
	strSort="DESC"
End If 

call openDb()

	query="SELECT orders.idOrder, orders.idCustomer, orders.paymentDetails, orders.paymentCode, orders.orderstatus, orders.orderDate, orders.total, orders.rmaCredit, orders.pcOrd_paymentStatus, customers.name, customers.lastName, customers.customerCompany, orders.comments, orders.admincomments FROM orders, customers"

If Request.QueryString("TypeSearch")="idOrder" Then	
	tempqryStr=Request.QueryString("advquery")
		if not validNum(tempqryStr) then
			call closeDb()
			response.redirect "msg.asp?message=14"
			response.End()
		end if
	if tempqryStr="" then
		query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 ORDER BY "& strORD &" "& strSort
	else
		tempqryStr=(int(tempqryStr) - scpre)
		query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 AND idOrder LIKE '%" & _
		tempqryStr & "%' ORDER BY "& strORD &" "& strSort
	end if
End If

If Request.QueryString("TypeSearch")="orderCode" Then	
	tempqryStr=trim(getUserInput(Request.QueryString("advquery"),0))
	if tempqryStr="" then
		query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 ORDER BY "& strORD &" "& strSort
	else
		query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 AND pcOrd_OrderKey LIKE '%" & _
		tempqryStr & "%' ORDER BY "& strORD &" "& strSort
	end if
End If

If Request.QueryString("TypeSearch")="GoogleOrderID" Then
	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 AND pcOrd_GoogleIDOrder LIKE '%" & _
	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
End If

If Request.QueryString("TypeSearch")="details" Then
	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 AND details LIKE '%" & _
	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
End If

If Request.QueryString("TypeSearch")="stateCode" Then
	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 AND orders.stateCode LIKE '%" & _
	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
End If

If Request.QueryString("TypeSearch")="CountryCode" Then
	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 AND orders.CountryCode LIKE '%" & _
	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
End If

If Request.QueryString("TypeSearch")="shipmentDetails" Then
	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 AND shipmentDetails LIKE '%" & _
	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
End If

If Request.QueryString("TypeSearch")="discountDetails" Then
	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 AND discountDetails LIKE '%" & _
	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
End If

If Request.QueryString("TypeSearch")="orderstatus" Then
	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderstatus LIKE '%" & _
	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
End If

If Request.QueryString("TypeSearch")="registry" Then
	pcIntRegistryID=Request.QueryString("pcIntRegistryID") 
	if validNum(pcIntRegistryID) then
		query=query & " WHERE orders.idCustomer=customers.idCustomer AND orders.pcOrd_IDEvent=" & pcIntRegistryID & " ORDER BY "& strORD &" "& strSort
	else
		query=query & " WHERE orders.idCustomer=customers.idCustomer AND orders.pcOrd_IDEvent=0 ORDER BY "& strORD &" "& strSort		
	end if
End If

If Request.QueryString("TypeSearch")="payment" Then
	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderStatus>1 AND orders.paymentDetails LIKE '%" & _
	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
End If

set rstemp=Server.CreateObject("ADODB.Recordset")     

rstemp.CursorLocation=adUseClient
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, conntemp

if err.number <> 0 then
	pcErrDescription = err.description
	call rstemp.Close
	set rstemp=nothing
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error: " & pcErrDescription) 
end If

if rstemp.eof then 
	presults="0"
else
	rstemp.MoveFirst
	' get the max number of pages
	Dim iPageCount
	iPageCount=rstemp.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
	' set the absolute page
	rstemp.AbsolutePage=iPageCurrent

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
                <a href="resultsAdvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent%>&order=<%=strORD%>&sort=<%=strSort%>&hideItemsOrdered=1">Hide Ordered Items Details</a>
                <% else %>
                <a href="resultsAdvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent%>&order=<%=strORD%>&sort=<%=strSort%>&hideItemsOrdered=0">Show Ordered Items Details</a>                
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
	
			</td>
		</tr>
		<tr> 
			<th align="center" nowrap><a href="resultsadvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent%>&order=orderstatus&sort=ASC"><img src="images/sortasc.gif" border="0" alt="Sort Ascending"></a><a href="resultsadvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent%>&order=orderstatus&sort=DESC"><img src="images/sortdesc.gif" border="0" alt="Sort Descending"></a>
			<a href="resultsadvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent%>&order=pcOrd_PaymentStatus&sort=ASC"><img src="images/sortasc.gif" border="0" alt="Sort Ascending"></a><a href="resultsadvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent%>&order=pcOrd_PaymentStatus&sort=DESC"><img src="images/sortdesc.gif" border="0" alt="Sort Descending"></a></th>
            <th align="center" nowrap><a href="resultsadvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent%>&order=orderDate&sort=ASC"><img src="images/sortasc.gif" border="0" alt="Sort Ascending"></a><a href="resultsadvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent%>&order=orderDate&sort=DESC"><img src="images/sortdesc.gif" border="0" alt="Sort Descending"></a> Date</th>
			<th nowrap>ID</th>
			<th nowrap>Total</th>
			<th nowrap>Customer</th>
	        <th nowrap>Items Ordered</th>
  			<th nowrap>Paid By</th>
			<th colspan="4" nowrap><div style="text-align: right;" class="pcSmallText"><a href="batchprocessorders.asp">Batch Process</a></div></th>
		</tr>
        <tr>
        	<td colspan="11" class="pcCPspacer"></td>
        </tr>
		<% 
		Dim mcount
		mcount=0
		If rstemp.EOF Then %>
			<tr>
			<td colspan="11">
				<p>&nbsp;</p>
				<p>No Results Found - <a href="javascript:history.back();">Back</a></p>
			</td>
			</tr>
		<% Else
			' Showing relevant records

			Dim i, x
			
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
							 <a href="Orddetails.asp?id=<%=pidOrder%>"><%=scCurSign&money(ptotal)%></a>
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
						<input type="checkbox" name="check<%=mcount%>" value="1" class="clearBorder">
						<% else %>
						&nbsp;
						<% end if %>
						</td>
					</tr>
					<% rstemp.MoveNext
				End If
			Loop
			set rstemp=nothing
			call closeDb()
			%>
	<input type="hidden" name="count" value="<%=mcount%>">
<%End If %>
		<tr> 
<td colspan="11" align="right">
				<%if mcount>0 then%>
<span class="cpLinksList"><a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a></span>
<br>
<br>
<INPUT type="button" value="Print Invoices/Packing Slips" name="button1" onclick="return OnButton1();">
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

					function checkAll() {
					for (var j = 1; j <= <%=mcount%>; j++) {
					box = eval("document.checkboxform.check" + j); 
					if (box.checked == false) box.checked = true;
						 }
					}
						
					function uncheckAll() {
					for (var j = 1; j <= <%=mcount%>; j++) {
					box = eval("document.checkboxform.check" + j); 
					if (box.checked == true) box.checked = false;
						 }
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
					<a href="resultsAdvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a> 
				<% end If
				If iPageCount <> 1 then
					For I=1 To iPageCount
						If int(I)=int(iPageCurrent) Then %>
							<%=I%> 
						<% Else %>
							<a href="resultsAdvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>" style="text-decoration: underline;"><%=I%></a> 
						<% End If %>
					<% Next %>
				<% end if %>
				<% if CInt(iPageCurrent) <> CInt(iPageCount) then
					'We are not at the end, show a next link %>
					<a href="resultsAdvanced.asp?TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
				<% end If %>
			</td>
		</tr>
	<tr>
		<td><hr></td>
	</tr>          
	</table>
<% end if %>
<!--#include file="AdminFooter.asp"-->