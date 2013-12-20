<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Dim pageTitle, Section
pageTitle="Batch Ship Orders"
Section="orders" %>
<% Server.ScriptTimeout = 3600 %>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim query, rs, conntemp

iPageSize=50
iPageCurrent=request("iPageCurrent")
if iPageCurrent="" then
	iPageCurrent="1"
end if	

AuthOrder=request.QueryString("AuthOrder")
AuthSort=request.QueryString("AuthSort")
if AuthOrder="" then
	AuthOrder="orders.orderdate"
	AuthSort="DESC"
end if
GenOrder=request("GenOrder")
GenSort=request("GenSort")
if (GenOrder="") and AuthOrder<>"orders.orderdate" then
	GenOrder="orders.orderdate"
	GenSort="DESC"
else
	GenOrder="orders.idorder"
	GenSort="DESC"
end if

call opendb() 

Dim iCnt, gwa, gwvpfp, gwpp, gwpsi, gwit, gwlp, gwvpfl, gwwp, gwmoneris, gwbofa, gw2Checkout, gwAIM, varTemp, varActive, actGW


query="SELECT DISTINCT orders.idOrder, orders.orderDate, orders.orderstatus, orders.total, orders.idCustomer, orders.paymentCode, orders.shipmentdetails FROM orders WHERE orderstatus=3 AND idorder IN (SELECT DISTINCT ProductsOrdered.idorder FROM Products INNER JOIN ProductsOrdered ON Products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.pcPrdOrd_Shipped<>1 AND Products.pcProd_IsDropShipped<>1) ORDER BY " & AuthOrder & " " & AuthSort
if GenOrder<>"" then
	query=query & ","&GenOrder&" "&GenSort&";"
end if
set rs=server.CreateObject("ADODB.RecordSet")
rs.CacheSize=iPageSize
rs.PageSize=iPageSize
rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

%>
<form name="form1" method="post" action="batchshiporders_submit.asp" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="7">
		The following orders have been <strong>processed</strong>.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=316')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a><br>
		You can also <a href="ship-index_import_help.asp">import 'Order Shipped' information</a> and automatically ship orders during the import process (see the User Guide for details).
		</td>
	</tr>
	<tr>
		<td colspan="7" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th nowrap><a href="batchshiporders.asp?AuthOrder=orders.idOrder&AuthSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending" align="absmiddle"></a><a href="batchshiporders.asp?AuthOrder=orders.idOrder&AuthSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending" align="absmiddle"></a>&nbsp;Order</th>
		<th width="5%" nowrap><a href="batchshiporders.asp?AuthOrder=orders.orderdate&AuthSort=ASC&GenOrder=orders.idOrder&GenSort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" border="0" alt="Sort Ascending" align="absmiddle"></a><a href="batchshiporders.asp?AuthOrder=orders.orderdate&AuthSort=Desc&GenOrder=orders.idOrder&GenSort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" border="0" alt="Sort Descending" align="absmiddle"></a>&nbsp;Date</th>
		<th width="5%">Ship</th>
		<th width="5%">Email</th>
		<th nowrap="nowrap">Ship Method</th>
		<th nowrap="nowrap">Ship Date</th>
		<th nowrap="nowrap">Tracking #</th>
	</tr>
	<tr>
		<td colspan="7" class="pcCPspacer"></td>
	</tr>
	<% dim noORDRec
	noORDRec=0
	checkboxCnt=0
	iPageCount=1
	if rs.eof then
		noORDRec=1 
		%>
		<tr> 
			<td colspan="7"><div class="pcCPmessage">There are no <strong>processed</strong> orders that have not yet been shipped.</div></td>
		</tr>
		<tr>
			<td colspan="7" class="pcCPspacer"></td>
		</tr>
	<%set rs=nothing
	ELSE
		iPageCount=rs.PageCount
		If Cint(iPageCurrent) > Cint(iPageCount) Then Cint(iPageCurrent)=Cint(iPageCount)
		rs.AbsolutePage=iPageCurrent
		pcv_Arr=rs.getRows(iPageSize)
		set rs=nothing
		intCount=ubound(pcv_Arr,2)
		For i=0 to intCount
			idOrder=pcv_Arr(0,i)
			orderDate=pcv_Arr(1,i)
			orderStatus=pcv_Arr(2,i)
			total=pcv_Arr(3,i)
			idCustomer=pcv_Arr(4,i)
			paymentCode=pcv_Arr(5,i)
			if isNULL(paymentCode)=True then
				paymentCode=""
			end if
			shipmentdetails=pcv_Arr(6,i)
	
			If instr(shipmentdetails,",") > 0 Then
				shipArray=split(shipmentdetails,",")
				shipmethod = shipArray(1)
			Else
				shipmethod = ""
			End If
			
			'// Get customer information
			if validNum(idCustomer) then
				query="SELECT name,lastname FROM customers WHERE idCustomer = " & idCustomer
				set rsCust=server.CreateObject("ADODB.RecordSet")
				set rsCust=conntemp.execute(query)
				pcv_custName = rsCust("name") & " " & rsCust("lastname")
				set rsCust=nothing
			else
				pcv_custName = "NA"
			end if
			
			'// Select all the products in this order
			pcv_PrdList=""
			query="SELECT Products.idproduct,Products.Description,Products.Stock,Products.noStock,Products.pcProd_IsDropShipped,Products.pcDropShipper_ID, ProductsOrdered.idProductOrdered, ProductsOrdered.quantity,ProductsOrdered.pcPrdOrd_BackOrder,ProductsOrdered.pcPrdOrd_Shipped FROM Products INNER JOIN ProductsOrdered ON Products.idproduct=ProductsOrdered.idproduct WHERE ProductsOrdered.idorder=" & idOrder
			set rsPrdList=server.CreateObject("ADODB.RecordSet")
			set rsPrdList=conntemp.execute(query)
			IF NOT rsPrdList.eof THEN
				'// Make a product list with the ones that are available					
				pcv_BatchOverRide=0
				Do While NOT rsPrdList.eof 
					pcv_cancheck=0
					pcv_count=pcv_count+1
					pcv_IDProduct=rsPrdList("idproduct")
					pcv_IDProductOrdered=rsPrdList("idProductOrdered")
					pcv_Description=rsPrdList("description")
					pcv_Stock=rsPrdList("stock")
					pcv_DisregardStock=rsPrdList("noStock")
					pcv_IsDropShipped=rsPrdList("pcProd_IsDropShipped")
					if IsNull(pcv_IsDropShipped) or pcv_IsDropShipped="" then
						pcv_IsDropShipped=0
					end if
					pcv_IDDropShipper=rsPrdList("pcDropShipper_ID")
					if IsNull(pcv_IDDropShipper) or pcv_IDDropShipper="" then
						pcv_IDDropShipper=0
					end if
					pcv_Qty=rsPrdList("quantity")
					if IsNull(pcv_Qty) or pcv_Qty="" then
						pcv_Qty=0
					end if
					pcv_BackOrder=rsPrdList("pcPrdOrd_BackOrder")
					if IsNull(pcv_BackOrder) or pcv_BackOrder="" then
						pcv_BackOrder=0
					end if						
					pcv_Shipped=rsPrdList("pcPrdOrd_Shipped")
					if IsNull(pcv_Shipped) or pcv_Shipped="" then
						pcv_Shipped=0
					end if	
					if pcv_Shipped=0 AND pcv_IsDropShipped=0 then
						pcv_available=1
					else
						pcv_available=0
					end if
					if (pcv_Shipped=0) AND (((pcv_BackOrder="1") and (clng(pcv_Stock)>=clng(pcv_Qty))) or ((pcv_BackOrder="0") and (clng(pcv_Stock)>=0)) or ((pcv_DisregardStock<>0) and (pcv_IDDropShipper=0)) OR (scOutOfStockPurchase=0)) then
						pcv_cancheck=1
					else
						'// one item did not meet the "checkable criteria" so we must over-ride the entire order
						pcv_BatchOverRide=1
					end if	
					if pcv_cancheck=1 then
						pcv_PrdList=pcv_PrdList & pcv_IDProductOrdered & ","
					end if								
				rsPrdList.movenext
				Loop
				Set rsPrdList = nothing			
			END IF	
			'// Batch Shipping Over-ride
			if pcv_BatchOverRide=1 then
				pcv_cancheck=0
				pcv_PrdList=""
			end if
			'// Hide package already shipped
			If pcv_available=1 AND NOT (trim(paymentCode)="Google Checkout") Then	
				'// Count
				checkboxCnt=checkboxCnt+1
				%>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td>
						<a href="ordDetails.asp?id=<%=int(idOrder)%>" target="_blank" style="text-decoration: none;" title="View Order Details"><%=int(idOrder)+scpre%></a>&nbsp;|&nbsp;<a href="ordDetails.asp?id=<%=int(idOrder)%>&activetab=2" target="_blank" style="text-decoration: none;" title="View Order Details"><%=scCurSign&money(total)%></a>&nbsp;|&nbsp;<a href="modcusta.asp?idcustomer=<%=idCustomer%>" target="_blank" style="text-decoration:none;" title="View Customer Details"><%=pcv_custName%></a>
					</td>
					<td><%=ShowDateFrmt(orderDate)%></td>
					<td align="center">
						<% if pcv_cancheck=1 then %>
							<input name="checkOrd<%=checkboxCnt%>" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" checked class="clearBorder">
						<% else %>
							<input name="checkOrd<%=checkboxCnt%>_disabled" type="checkbox" id="checkOrd<%=checkboxCnt%>" value="YES" disabled class="clearBorder">
						<% end if %>
					</td>
					<td align="center">
						<% if pcv_cancheck=1 then %>
							<input name="checkEmail<%=checkboxCnt%>" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" checked class="clearBorder">
						<% else %>
							<input name="checkEmail<%=checkboxCnt%>_disabled" type="checkbox" id="checkEmail<%=checkboxCnt%>" value="YES" disabled class="clearBorder">
						<% end if %>					
					</td>
					<td><input type="text" name="shipmethod<%=checkboxCnt%>" size="15" value="<%= shipmethod %>"></td>
					<td><input type="text" name="shipdate<%=checkboxCnt%>" size="8" value="<%=ShowDateFrmt(Date())%>"></td>
					<td>
						<input type="text" name="tracking<%=checkboxCnt%>" size="25">
						<input type="hidden" name="orderstatus<%=checkboxCnt%>" value="<%=orderStatus%>">
						<input type="hidden" name="idOrder<%=checkboxCnt%>" value="<%=idOrder%>">
						<input type="hidden" name="amt<%=checkboxCnt%>" value="<%=total%>">
						<input type="hidden" name="PrdList<%=checkboxCnt%>" value="<%=pcv_PrdList%>">	
					</td>
				</tr>
				<% if pcv_cancheck=0 then %>
				<tr>
					<td colspan="7" style="padding-bottom:8px"><span class="pcCPnotes">The order listed above cannot be batch-shipped because it contains one or more products that are currently out-of-stock or back-ordered. Even if you click on <em>Select All</em>, the order will not be shipped (although the checkbox may appear checked).</span></td>
				</tr>
				<% end if %>
				<%
			End If '// If pcv_available=1 Then
			
		Next
	end if
	if noORDRec=0 AND checkboxCnt=0 then
		noORDRec=1 
		%>
		<tr>
			<td colspan="7" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="7"><div class="pcCPmessage">There are no <strong>processed</strong> orders that have not yet been shipped.</div></td>
		</tr>
		<tr>
			<td colspan="7" class="pcCPspacer"></td>
		</tr>
	<% end if %>
	<tr>
		<td colspan="7" class="pcCPspacer"></td>
	</tr>
	<tr>
        <td colspan="7" class="cpLinksList">
			<input type="hidden" name="checkboxCnt" value="<%=checkboxCnt%>">
			<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
		</td>
	</tr>
		<tr>
			<td colspan="7" class="pcCPspacer"></td>
		</tr>
	<% if noORDRec=0  then %>
	<tr>
		<td colspan="9"><input type="submit" name="Submit" value="Ship Selected Orders" class="submit2"></td>
	</tr>
	<% end if %>
</table>
<table class="pcCPcontent">
<tr>
	<td>
		<% If iPageCount>1 Then %>
            <%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount & "<br>")%>
			<p class="pcPageNav">
				<%if iPageCurrent > 1 then %>
					<a href="batchshiporders.asp?iPageCurrent=<%=iPageCurrent - 1%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&GenOrder=<%=GenOrder%>&GenSort=<%=GenSort%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a>
				<% end If
				For I = 1 To iPageCount
					If Cint(I) = Cint(iPageCurrent) Then %>
						<b><%=I%></b>
					<% Else %>
						<a href="batchshiporders.asp?iPageCurrent=<%=I%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&GenOrder=<%=GenOrder%>&GenSort=<%=GenSort%>"><%=I%></a>
					<% End If %>
				<%Next %>
				<%if CInt(iPageCurrent) < CInt(iPageCount) then %>
					<a href="batchshiporders.asp?iPageCurrent=<%=iPageCurrent + 1%>&AuthOrder=<%=AuthOrder%>&AuthSort=<%=AuthSort%>&GenOrder=<%=GenOrder%>&GenSort=<%=GenSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
				<%end If %>
			</p>
		<% End If %>
	</td>
</tr>
</table>
</form>
<%if checkboxCnt>0 then%>
<script language="JavaScript">
<!--
function checkAll() {
for (var j = 1; j <= <%=checkboxCnt%>; j++) {
box = eval("document.form1.checkOrd" + j); 
if (box.checked == false) box.checked = true;
box1 = eval("document.form1.checkEmail" + j); 
if (box1.checked == false) box1.checked = true;
   }
}

function uncheckAll() {
for (var j = 1; j <= <%=checkboxCnt%>; j++) {
box = eval("document.form1.checkOrd" + j); 
if (box.checked == true) box.checked = false;
box1 = eval("document.form1.checkEmail" + j); 
if (box1.checked == true) box1.checked = false;
   }
}

//-->
</script>
<%
end if
call closeDb()
%>
<!--#include file="adminfooter.asp"-->