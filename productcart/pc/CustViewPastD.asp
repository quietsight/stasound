<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%'Allow Guest Account
AllowGuestAccess=1
%>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/shipFromsettings.asp"--> 
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<%
err.number=0
dim query, conntemp, rs, rstemp, pIdOrder
call openDb()
%>
<!--#include file="prv_getsettings.asp"-->
<%
pcv_RWActive=pcv_Active
pIdOrder=getUserInput(request("idOrder"),10)
if not validNum(pIdOrder) then response.Redirect "custPref.asp"

' extract real idorder (without prefix)
pIdOrder=(int(pIdOrder)-scpre)

Dim pord_DeliveryDate, pord_OrderName

if request("action")="rename" then
	pord_OrderName=getUserInput(request("ord_OrderName"),0)
	If pord_OrderName = "" Then
		pord_OrderName = "No Name"
	End If
	query="update orders set ord_OrderName='" & pord_OrderName & "' where idOrder=" & pidOrder
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
end if

tmpReSent=0

IF request("action")="resend" THEN
	GC_ReName=getUserInput(request("GC_RecName"),0)
	GC_ReEmail=getUserInput(request("GC_RecEmail"),0)
	GC_ReMsg=getUserInput(request("GC_RecMsg"),0)
	
	query="UPDATE orders SET pcOrd_GcReName='" & GC_ReName & "',pcOrd_GcReEmail='" & GC_ReEmail & "',pcOrd_GcReMsg='" & GC_ReMsg & "' WHERE idOrder="& pIdOrder
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	Set rs=nothing
	
	ReciEmail=""

	query="select idproduct from ProductsOrdered WHERE idOrder="& pIdOrder
	pidorder=pIdOrder
	set rs11=connTemp.execute(query)
	do while not rs11.eof
		query="select products.Description,pcGCOrdered.pcGO_GcCode,pcGc.pcGc_EOnly from Products,pcGc,pcGCOrdered where products.idproduct=" & rs11("idproduct") & " and pcGC.pcGc_IDProduct=products.idproduct and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& pIdOrder
		set rs=connTemp.execute(query)
	
		if not rs.eof then
			pIdproduct=rs11("idproduct")
			pName=rs("Description")
			pCode=rs("pcGO_GcCode")
			pEOnly=rs("pcGc_EOnly")
			
				query="select pcGO_Amount,pcGO_GcCode,pcGO_ExpDate from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & pidorder
				set rs19=connTemp.execute(query)
				
				do while not rs19.eof
				pAmount=rs19("pcGO_Amount")
				if pAmount<>"" then
				else
				pAmount="0"
				end if
				
				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_68") & scCurSign & money(pAmount) & vbcrlf
				
				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_69") & rs19("pcGO_GcCode") & vbcrlf
				pExpDate=rs19("pcGO_ExpDate")
				 
				if year(pExpDate)="1900" then
				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_45b") & vbcrlf
				else
				if scDateFrmt="DD/MM/YY" then
				pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
				else
				pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
				end if
				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_70") & pExpDate & vbcrlf
				end if
				if pEOnly="1" then
				ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_71") & vbcrlf
				end if
				ReciEmail=ReciEmail & vbcrlf
				rs19.movenext
				loop

		end if
	rs11.MoveNext
	loop
	set rs11=nothing
	
	query="SELECT customers.name,customers.lastname,customers.email,orders.pcOrd_GcReName,orders.pcOrd_GcReEmail,orders.pcOrd_GcReMsg FROM customers INNER JOIN Orders ON customers.idcustomer=orders.idcustomer WHERE idOrder="& pIdOrder
	set rs11=connTemp.execute(query)

	if not rs11.eof then
		pCustomerFullName=rs11("name") & " " & rs11("lastname")
		pCustomerFullNamePlusEmail=pCustomerFullName & " (" & rs11("email") & ")"
		GcReName=rs11("pcOrd_GcReName")
		GcReEmail=rs11("pcOrd_GcReEmail")
		GcReMsg=rs11("pcOrd_GcReMsg")
	
		if GcReEmail<>"" then
			if GcReName<>"" then
			else
				GcReName=GcReEmail
			end if
			ReciEmail1=replace(dictLanguage.Item(Session("language")&"_sendMail_66"),"<recipient name>",GcReName)
			ReciEmail2=replace(dictLanguage.Item(Session("language")&"_sendMail_67"),"<customer name>",pCustomerFullNamePlusEmail)
			if GcReMsg<>"" then
				ReciEmail3=replace(dictLanguage.Item(Session("language")&"_sendMail_72"),"<customer name>",pCustomerFullNamePlusEmail) & vbcrlf & GcReMsg & vbcrlf
			else
				ReciEmail3=""
			end if
			ReciEmail=ReciEmail1 & vbcrlf & vbcrlf & ReciEmail2 & vbcrlf & vbcrlf & ReciEmail & ReciEmail3
			ReciEmail=ReciEmail & vbcrlf & scCompanyName & vbCrLf & scStoreURL & vbcrlf & vbCrLf
			call sendmail (scCompanyName, scEmail, GcReEmail,pCustomerFullName & dictLanguage.Item(Session("language")&"_sendMail_73"), replace(ReciEmail, "&quot;", chr(34)))
			tmpReSent=1
		end if
	end if
	set rs11=nothing
END IF


query="SELECT orders.pcOrd_OrderKey,customers.email,customers.fax,orders.pcOrd_ShippingEmail,orders.pcOrd_ShippingFax,orders.pcOrd_ShowShipAddr,orders.idCustomer, orders.pcOrd_PaymentStatus,orders.orderDate, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.customerType, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.comments, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.pcOrd_shippingPhone, orders.shippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.idOrder, orders.rmaCredit, orders.ordPackageNum, orders.ord_DeliveryDate, orders.ord_OrderName, orders.ord_VAT,orders.pcOrd_CatDiscounts, orders.paymentDetails, orders.gwAuthCode, orders.gwTransId, orders.paymentCode FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pIdOrder&"));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rs.eof then
	set rs=nothing
	call closeDb()
 	response.redirect "msg.asp?message=35"     
end if 

dim pidCustomer, porderDate, pfirstname, plastname,pcustomerCompany, pphone, paddress, pzip, pstate, pcity, pcountryCode, pcomments, pshippingAddress, pshippingState, pshippingCity, pshippingCountryCode, pshippingZip, paddress2, pshippingFullName, pshippingCompany, pshippingAddress2, pshippingPhone

pidCustomer=rs("idCustomer")
if int(Session("idcustomer"))<=0 then
	if session("REGidCustomer")>"0" then
		testidCustomer=int(session("REGidCustomer"))
	end if
else
	testidCustomer=int(Session("idcustomer"))
end if
if testidCustomer<>int(pidCustomer) then
	set rs=nothing
	call closeDb()
	session("REGidCustomer")=""
	response.redirect "msg.asp?message=11"    
end if

'Start SDBA
pcv_PaymentStatus=rs("pcOrd_PaymentStatus")
if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
	pcv_PaymentStatus=0
end if
'End SDBA

pcOrderKey=rs("pcOrd_OrderKey")
pEmail=rs("email")
pFax=rs("fax")
pshippingEmail=rs("pcOrd_ShippingEmail")
pshippingFax=rs("pcOrd_ShippingFax")
pcShowShipAddr=rs("pcOrd_ShowShipAddr")
porderDate=rs("orderDate")
porderDate=showdateFrmt(porderDate)
pfirstname=rs("name")
plastName=rs("lastName")
pcustomerCompany=rs("customerCompany")
pphone=rs("phone")
pcustomerType=rs("customerType")
paddress=rs("address")
pzip=rs("zip")
pstate=rs("stateCode")
if pstate="" then
	pstate=rs("state")
end if
pcity=rs("city")
pcountryCode=rs("countryCode")
pcomments=rs("comments")
pshippingAddress=rs("shippingAddress")

	'// START - Test for existence of separate shipping address
	if IsNull(pcShowShipAddr) OR (pcShowShipAddr="") OR (pcShowShipAddr="0") then
		'This might be a v3 store, check another field
		if trim(pshippingAddress)="" then
			pcShowShipAddr=0
			else
			pcShowShipAddr=1
		end if
	end if
	'// END

pshippingState=rs("shippingStateCode")
if pshippingState="" then
	pshippingState=rs("shippingState")
end if
pshippingCity=rs("shippingCity")
pshippingCountryCode=rs("shippingCountryCode")
pshippingZip=rs("shippingZip")
pshippingPhone=rs("pcOrd_shippingPhone")
pshippingFullName=rs("shippingFullName")
paddress2=rs("address2")
pshippingCompany=rs("shippingCompany")
pshippingAddress2=rs("shippingAddress2")
pidOrder=rs("idOrder")
pRmaCredit=rs("rmaCredit")
pOrdPackageNum=rs("ordPackageNum")
pord_DeliveryDate=rs("ord_DeliveryDate")
pord_OrderName=rs("ord_OrderName")
pord_VAT=rs("ord_VAT")
pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
if isNULL(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
	pcv_CatDiscounts="0"
end if
pcpaymentDetails=trim(rs("paymentDetails"))
pcgwAuthCode=rs("gwAuthCode")
pcgwTransId=rs("gwTransId")
pcpaymentCode=rs("paymentCode")

query="SELECT Orders.pcOrd_GWTotal,Orders.pcOrd_IDEvent,ProductsOrdered.pcPO_GWOpt,ProductsOrdered.pcPO_GWNote,ProductsOrdered.pcPO_GWPrice,orders.pcOrd_GCs,orders.pcOrd_GcCode,orders.pcOrd_GcUsed,ProductsOrdered.idProduct, ProductsOrdered.pcPrdOrd_Shipped, ProductsOrdered.quantity, ProductsOrdered.unitPrice,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.xfdetails  "
'BTO ADDON-S
If scBTO=1 then
	query=query&", ProductsOrdered.idconfigSession"
End If
'BTO ADDON-E
query=query&", products.description, products.sku, orders.total, orders.paymentDetails, orders.taxamount, orders.shipmentDetails, orders.discountDetails, orders.pcOrd_GCDetails, orders.orderstatus,orders.processDate, orders.shipdate, orders.shipvia, orders.trackingNum, orders.returnDate, orders.returnReason, orders.iRewardPoints, orders.iRewardValue, orders.iRewardPointsCustAccrued, orders.taxdetails,orders.dps,pcPrdOrd_BundledDisc FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idCustomer=" &pidCustomer& " AND orders.idOrder=" &pIdOrder
set rsOrdObj=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsOrdObj=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rsOrdObj.eof then
	set rsOrdObj=nothing
	call closeDb()
 	response.redirect "msg.asp?message=35"
end if

'GGG Add-on start
pGWTotal=rsOrdObj("pcOrd_GWTotal")
if pGWTotal<>"" then
else
pGWTotal="0"
end if
gIDEvent=rsOrdObj("pcOrd_IDEvent")
if gIDEvent<>"" then
else
gIDEvent="0"
end if
'GGG Add-on end

query="SELECT pcPrdOrd_Shipped FROM ProductsOrdered WHERE idOrder=" & pIdOrder & " AND pcPrdOrd_Shipped=1;"
set rsQ=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
pcv_HaveShipped=0
if not rsQ.eof then
	pcv_HaveShipped=1
end if
set rsQ=nothing

%>
<!--#include file="header.asp"-->
<script>
	function openbrowser(url) {
			self.name = "productPageWin";
			popUpWin = window.open(url,'rating','toolbar=0,location=0,directories=0,status=0,top=0,scrollbars=yes,resizable=1,width=705,height=535');
			if (navigator.appName == 'Netscape') {
			popUpWin.focus();
		}
	}
</script>
<div id="pcMain">
	<div id="GlobalAjaxErrorDialog" title="Communication Error" style="display:none">
		<div class="pcErrorMessage">
			Can not connect to server to exchange information. Please contact store owner or try again later
		</div>
	</div>
	<table class="pcMainTable">   
		<tr>
			<td>
				<h1>
					<%response.write dictLanguage.Item(Session("language")&"_CustviewPast_4")%>
				</h1>
				<h2>
					<%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_9")&(int(pIdOrder)+scpre) & " - " & dictLanguage.Item(Session("language")&"_CustviewPastD_14") & porderDate%>
                    <%if pcOrderKey<>"" then%> - <%=dictLanguage.Item(Session("language")&"_opc_common_1")%>&nbsp;<%=pcOrderKey%><%end if%>
				</h2>
			</td>
		</tr>
		<%if tmpReSent=1 then%>
		<tr>
			<td>
				<div class="pcSuccessMessage">
					<%response.write dictLanguage.Item(Session("language")&"_GCRecipient_3")%>
				</div>
			</td>
		</tr>
		<%end if%>
		<%if session("REGidCustomer")>"0" then %>
		<tr>
			<td>
				<div class="pcInfoMessage">
					<%response.write dictLanguage.Item(Session("language")&"_opc_85")%>
				</div>
			</td>
		</tr>
		<%end if%>
		<tr>
			<td>
				<table class="pcShowContent">
					<tr>
						<td>
							<p>
							<a href="custOrdInvoice.asp?id=<%=pIdOrder%>" target="_blank"><img src="images/document.gif" border="0" align="middle" style="margin-right: 2px;"></a> <a href="custOrdInvoice.asp?id=<%=pIdOrder%>" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_33")%></a>
							<a href="custOrdInvoicePDF.asp?id=<%=pIdOrder%>" target="_blank"><img src="images/document.gif" border="0" align="middle" style="margin: 0 2px 0 6px;"></a> <a href="custOrdInvoicePDF.asp?id=<%=pIdOrder%>" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_75")%></a>
							<%IF (Session("CustomerGuest")="0") AND (Session("idCustomer")>"0") THEN%> - <a href="RepeatOrder.asp?idOrder=<%=pIdOrder%>"><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_32")%></a>
							<% 'Hide/show link to Help Desk
								If scShowHD <> 0 then %>
								&nbsp;-&nbsp;<a href="userviewallposts.asp?idOrder=<%=clng(scpre)+clng(pIdOrder)%>"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_3")%></a>
							<% end if %>
							<%END IF%>
							</p>
						</td>
						<td>
						<%IF (Session("CustomerGuest")="0") AND (Session("idCustomer")>"0") THEN%>
						<p align="right">
						<a href="custViewPast.asp"><img src="<%=rslayout("back")%>"></a>
						</p>
						<%END IF%>
						</td>
					</tr>
					<% 'SB S %>
                    <% Dim pcv_strCustEmail, pcv_strGUID
					pcv_strCustEmail=""
					pcv_strGUID=""
					query = "SELECT customers.email, SB_Orders.SB_GUID FROM (orders "
					query = query & "Inner Join customers on orders.idCustomer = customers.idCustomer) "
					query = query & "Inner Join SB_Orders on orders.idorder = SB_Orders.idorder "
					query = query & "WHERE orders.idorder = " & pIdOrder
					set rsSB=Server.CreateObject("ADODB.Recordset")
					set rsSB=conntemp.execute(query)
					if NOT rsSB.eof then
						   pcv_strCustEmail = rsSb("email")
						   pcv_strGUID = rsSb("SB_GUID")
					end if  
					set rsSB=nothing 
					%>
                    <% If len(pcv_strGUID)>0 Then %>
					<tr>
						<td colspan="2">
                        	<%response.write dictLanguage.Item(Session("language")&"_SB_1")%> <a href="<%=gv_RootURL%>/CustomerCenter/AutoLogin.asp?ID=<%=pcv_strGUID%>&Email=<%=pcv_strCustEmail%>&mode=details" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_SB_2")%></a>
						</td>
					</tr>  
                    <% End If %>
					<% 'SB E %>
					<tr>
					<td colspan="2">
						<%'GGG Add-on start
						if gIDEvent<>"0" then

							query="select pcEvents.pcEv_name,pcEvents.pcEv_Date, pcEv_HideAddress,customers.name,customers.lastname from pcEvents,Customers where Customers.idcustomer=pcEvents.pcEv_idcustomer and pcEvents.pcEv_IDEvent=" & gIDEvent
							set rs1=connTemp.execute(query)

							if err.number<>0 then
								call LogErrorToDatabase()
								set rs1=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if

							geName=rs1("pcEv_name")
							geDate=rs1("pcEv_Date")

							if year(geDate)="1900" then
								geDate=""
							end if
							if gedate<>"" then
								if scDateFrmt="DD/MM/YY" then
									gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
								else
									gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
								end if
							end if
							geHideAddress=rs1("pcEv_HideAddress")
							if geHideAddress="" then
								geHideAddress=0
							end if
							gReg=rs1("name") & " " & rs1("lastname")
							
							set rs1=nothing
							%>
							<table class="pcCPcontent">
							<tr>
								<td nowrap><b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_39")%></b></td>
								<td width="100%"><%=gename%></td>
							</tr>
							<tr>
								<td nowrap><b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_40")%></b></td>
								<td width="100%"><%=geDate%></td>
							</tr>
							<tr>
								<td nowrap><b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_41")%></b></td>
								<td width="100%"><%=gReg%></td>
							</tr>
							</table>
						<% else
							geHideAddress=0
						end if
						'GGG Add-on end%>
					</td>
					</tr>
				</table>
			</td>
		</tr>
		
		<% if scOrderName="1" Then 'Allow customer to nickname this order %>
		<tr>
			<td><hr></td>
		</tr>
		<tr>
			<td>
					<% if pord_OrderName="" then
						pord_OrderName="No Name"
					end if%>
					<form method="post" name="form1" id="form1" action="CustViewPastD.asp" class="pcForms">
					<input type=hidden name="action" value="rename">
					<input type=hidden name="IDOrder" value="<%=int(pIdOrder)+scpre%>">
					<%=dictLanguage.Item(Session("language")&"_CustviewOrd_40")%> <input type="text" size="30" maxsize="50" name="ord_OrderName" value="<%=pord_OrderName%>">&nbsp;<input type="submit" name="Submit" value="Update" class="submit2">
					</form>
					</font>
			</td>
		</tr>
		<tr>
			<td><hr></td>
		</tr>
	<% end if 'End allow customer to nickname this order %>
	
	<% 
		' START order delivery date, if any
		if (pord_DeliveryDate<>"") then
			if scDateFrmt="DD/MM/YY" then
				pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 4)
			else
				pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 3)
			end if
		pord_DeliveryDate = showdateFrmt(pord_DeliveryDate)
		
			if not scOrderName="1" Then 'Add <hr> only if the Order Name section is not shown %>
			<tr>
				<td><hr></td>
			</tr>
		<% end if %>
		
			<tr>
				<td valign="top">
				<%=dictLanguage.Item(Session("language")&"_CustviewOrd_39")%><%=pord_DeliveryDate%> <% If pord_DeliveryTime <> "00:00" Then %><%=", " & pord_DeliveryTime%><% End If %>
				</td>
			</tr>
			<tr>
				<td><hr></td>
			</tr>
		<%
		end if
		' END order delivery date
		'
		' START Billing and Shipping Addresses
		%>

		<tr>
			<td>
				<table class="pcShowContent">
					<tr>
						<th colspan="2">
							<strong><%response.write dictLanguage.Item(Session("language")&"_orderverify_23")%></strong>
						</th>
						<th>&nbsp;</th>
						<th>
							<%if pcShowShipAddr="1" AND geHideAddress=0 then %>
							<strong><%response.write dictLanguage.Item(Session("language")&"_orderverify_24")%></strong>
							<%end if%>
						</th>
					</tr>
	
					<tr>
						<td width="20%">
							<p><% response.write replace(dictLanguage.Item(Session("language")&"_orderverify_7"),"''","'")%></p>
						</td>
						<td width="30%">
							<p><% response.write pFirstName&" "&plastname %></p>
						</td>
						<td width="20%">&nbsp;</td>
						<td width="30%">
							<%if pcShowShipAddr="1" AND geHideAddress=0 then%>
								<p><% response.write pshippingFullName %></p>
							<% end if%>
						</td>
					</tr>
	
					<tr>
						<td>
							<p>
							<% response.write dictLanguage.Item(Session("language")&"_orderverify_8")%>
							</p>
						</td>
						<td>
							<p><%=pcustomerCompany%></p>
						</td>
						<td>&nbsp;</td>
						<td>
						<p>
							<% if pcShowShipAddr="1" AND geHideAddress=0 then
								if pshippingCompany<>"" then
									response.write pshippingCompany
								end if
							end if %>
						</p>
						</td>
					</tr>
					
					<%if pEmail<>pshippingEmail AND pshippingEmail<>"" then%>
					<tr>
						<td>
							<p>
							<%=dictLanguage.Item(Session("language")&"_opc_5")%>
							</p>
						</td>
						<td valign="top">
							<p><%=pEmail%></p>
						</td>
						<td>&nbsp;</td>
						<td>
						<p>
							<%if pcShowShipAddr="1" AND geHideAddress=0 then
								response.write pshippingEmail
							end if %>
						</p>
						</td>
					</tr>
					<%end if%>
	
					<tr>
						<td>
							<p>
							<% response.write dictLanguage.Item(Session("language")&"_orderverify_9")%>
							</p>
						</td>
						<td valign="top">
							<p><%=pPhone%></p>
						</td>
						<td>&nbsp;</td>
						<td>
						<p>
							<%if pcShowShipAddr="1" AND geHideAddress=0 then
								response.write pshippingPhone
							end if %>
						</p>
						</td>
					</tr>
					
					<%if pFax<>"" OR pshippingFax<>"" then%>
					<tr>
						<td>
							<p>
							<%=dictLanguage.Item(Session("language")&"_opc_18")%>
							</p>
						</td>
						<td valign="top">
							<p><%=pFax%></p>
						</td>
						<td>&nbsp;</td>
						<td>
						<p>
							<%if pcShowShipAddr="1" AND geHideAddress=0 then
								response.write pshippingFax
							end if %>
						</p>
						</td>
					</tr>
					<%end if%>
	
					<tr>
						<td>
						<p>
							<% response.write dictLanguage.Item(Session("language")&"_orderverify_10")%>
						</p>
						</td>
						<td valign="top">
						<p>
							<%=paddress%>
						</p>
						</td>
						<td>&nbsp;</td>
						<td valign="top">
						<p>
							<% if pcShowShipAddr="1" AND geHideAddress=0 then
								if pshippingAddress="" then
									response.write "Same as Billing Address"
								else
									response.write pshippingAddress
								end if
							else
								if pcShowShipAddr="0" AND geHideAddress=0 then
									response.write "Same as Billing Address"
								end if
							end if %>
						</p>
						</td>
					</tr>
	
					<tr>
						<td>&nbsp;</td>
						<td valign="top">
							<p>
							<%=paddress2%>
							</p>
						</td>
						<td>&nbsp;</td>
						<td>
						<p>
							<% if pcShowShipAddr="1" AND geHideAddress=0 then
								if pshippingAddress2<>"" then
									response.write pshippingAddress2
								end if
							end if %>
						</p>
						</td>
					</tr>
	
					<tr>
						<td>&nbsp;</td>
						<td>
							<p>
							<%=pCity&", "&pState&" "&pzip%>
							</p>
						</td>
						<td>&nbsp;</td>
						<td>
						<p>
							<% if pcShowShipAddr="1" AND geHideAddress=0 then
								if pshippingAddress<>"" then
									response.write pShippingCity&", "&pshippingState
									If pshippingState="" then
										response.write pshippingStateCode
									End If
									response.write " "&pshippingZip
								end if
							end if %>
						</p>
						</td>
					</tr>
	
					<tr>
						<td>&nbsp;</td>
						<td>
						<p>
						<%=pCountryCode%>
						</p>
						</td>
						<td>&nbsp;</td>
						<td>
						<p>
							<%if pcShowShipAddr="1" AND geHideAddress=0 then
								response.write pshippingCountryCode
								strFedExCountryCode=pshippingCountryCode
							else
								strFedExCountryCode=pCountryCode
							end if %>
						</p>
						</td>
					</tr>
				</table>
			</td>
		</tr>

		<% 
		' END Billing and Shipping Addresses
		'
		' START of payment details
		payment = split(pcpaymentDetails,"||")
		PaymentType=trim(payment(0))
		
		'Get payment nickname
		query="SELECT paymentDesc,paymentNickName FROM paytypes WHERE paymentDesc = '" & replace(PaymentType,"'","''") & "';"
		Set rsTemp=Server.CreateObject("ADODB.Recordset")
		Set rsTemp=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsTemp=nothing
				call closedb()

				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if not rsTemp.EOF then
				PaymentName=trim(rsTemp("paymentNickName"))
				else
				PaymentName=""
			end if
		Set rsTemp = nothing
		'End get payment nickname
	
		'Get authorization and transaction IDs, if any
		varTransID=""
		varTransName= dictLanguage.Item(Session("language")&"_CustviewPastD_102")
		varAuthCode=""
		varAuthName= dictLanguage.Item(Session("language")&"_CustviewPastD_103")
	
		if NOT isNull(pcpaymentCode) AND pcpaymentCode<>"" then 
			varShowCCInfo=0
			select case pcpaymentCode
			case "LinkPoint"
				varAry=split(pcgwAuthCode,":")
				varTransName="Approval Number"
				varAuthName="Reference Number"
				varTransID=left(varAry(1),6)
				varAuthCode=right(varAry(1),10)
			case "PFLink", "PFPro", "PFPRO", "PFLINK"
				varTransID=pcgwTransId
				varAuthCode=pcgwAuthCode
				varShowCCInfo=1
				varGWInfo="P"
			case "Authorize"
				varTransID=pcgwTransId
				varAuthCode=pcgwAuthCode
				varShowCCInfo=1
				if instr(ucase(PaymentType),"CHECK") then
					varShowCCInfo=0
				end if
				varGWInfo="A"
			case "twoCheckout"
				varTransName="2Checkout Order No"
				varTransID=pcgwTransId
			case "BOFA"
				varTransName="Order No"
				varAuthName="Authorization Code"
				varTransID=pcgwTransId
				varAuthCode=pcgwAuthCode
			case "WorldPay"
				varTransID=""
				varAuthCode=""
			case "iTransact"
				varTransName="Transaction ID"
				varAuthName="Authorization Code"
				varTransID=pcgwTransId
				varAuthCode=pcgwAuthCode
			case "PSI", "PSIGate"
				varTransName="Transaction ID"
				varAuthName="Authorization Code"
				varTransID=pcgwTransId
				varAuthCode=pcgwAuthCode
			case "fasttransact", "FastTransact", "FAST","CyberSource"
				varTransName="Transaction ID"
				varAuthName="Authorization Code"
				varTransID=pcgwTransId
				varAuthCode=pcgwAuthCode
			case "USAePay","FastCharge"
				varTransName="Transaction reference code"
				varAuthName="Authorization code"
				varTransID=pcgwTransId
				varAuthCode=pcgwAuthCode
			case "PxPay"
				varTransName="DPS Transaction Reference Number"
				varAuthName="Authorization code"
				varTransID=pcgwTransId
				varAuthCode=pcgwAuthCode
			end select
		end if
		
		'End get authorization and transaction IDs
	
		If payment(1)="" then
		 if err.number<>0 then
			PayCharge=0
		 end if
			PayCharge=0
		else
			PayCharge=payment(1)
		end If
		err.number=0
		if instr(PaymentType,"FREE") AND len(PaymentType)<6 then
		else %>
		<tr>
			<td><hr></td>
		</tr>
		<tr>
			<td>
			<p>
			<%=dictLanguage.Item(Session("language")&"_CustviewPastD_101")%>
			<%
				if PaymentName <> "" and PaymentName <> PaymentType then
					Dim pcv_strPaymentType
					Select Case PaymentType
						Case "PayPal Website Payments Pro": pcv_strPaymentType=PaymentName
						Case Else: pcv_strPaymentType=PaymentName & " (" & PaymentType & ")"
					End Select
					Response.Write pcv_strPaymentType
					else
					Response.Write PaymentType
				end if
			%>
			<% if PayCharge>0 then %>
				<br><%=dictLanguage.Item(Session("language")&"_CustviewOrd_14b")%><%= " " & scCurSign&money(PayCharge)%>
			<% end if %>
			<% if varTransID<>"" then %>
			<br><%=varTransName%>: <%=varTransID%>
			<% end if %>
			<% if varAuthCode<>"" then %>
			<br><%=varAuthName%>: <%=varAuthCode%>
			<% end if %>
			</p>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	<% 
	end if
	' END Payment details
	' 
	' START Order Comments
	if len(pcomments)>3 then %>
		<tr>
			<td>
				<div class="pcSectionTitle">
				<% response.write dictLanguage.Item(Session("language")&"_orderverify_11")%>
				</b> <%=pcomments%>
			</div>
			</td>
		</tr>
	<% 
	end if 
	' END Order Comments
	'
	' START Order Details
	%>
	<tr>
		<td>
			<table class="pcShowContent">
				<tr>
					<th width="10%">
						<% response.write dictLanguage.Item(Session("language")&"_orderverify_25")%>
					</th>
					<th width="15%">
						<% response.write dictLanguage.Item(Session("language")&"_orderverify_26")%>
					</th>
					<th width="50%">
						<% response.write dictLanguage.Item(Session("language")&"_orderverify_27")%>
					</th>
					<th width="15%">
						<div align="right">
						<% response.write dictLanguage.Item(Session("language")&"_orderverify_32")%>
						</div>
					</th>
					<th width="10%">
						<div align="right">
							<% response.write dictLanguage.Item(Session("language")&"_orderverify_28")%>
						</div>
					</th>
					<%'Start SDBA%>
					<th><%if pcv_HaveShipped=1 then
					response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_1")
					end if%></th>
					<%'End SDBA%>
				</tr>
	
				<% dim pidProduct, pquantity, punitPrice, pxfdetails, pidconfigSession, pdescription, pSku, pcDPs, ptotal, ppaymentDetails,ptaxamount,pshipmentDetails, pdiscountDetails, porderstatus
				dim pprocessDate, pshipdate, pshipvia, ptrackingNum, preturnDate, preturnReason, piRewardPoints, piRewardValue, piRewardPointsCustAccrued,ptaxdetails, pOpPrices, rsObjOptions, pRowPrice, count, rsConfigObj,stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory,i, s,OptPrice,xfdetails, xfarray, q
				
				Dim pcv_strSelectedOptions, pcv_strOptionsPriceArray, pcv_strOptionsArray
				Dim pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice
				Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions
				Dim subTotal
				subTotal=0
				
				do while not rsOrdObj.eof
				
					'GGG Add-on start
					pGWOpt=rsOrdObj("pcPO_GWOpt")
					if pGWOpt<>"" then
					else
						pGWOpt="0"
					end if
					pGWText=rsOrdObj("pcPO_GWNote")
					pGWPrice=rsOrdObj("pcPO_GWPrice")
					if pGWPrice<>"" then
					else
						pGWPrice="0"
					end if
					pGCs=rsOrdObj("pcOrd_GCs")
					pGiftCode=rsOrdObj("pcOrd_GcCode")
					pGiftUsed=rsOrdObj("pcOrd_GcUsed")
					'GGG Add-on end
				
					pidProduct=rsOrdObj("idProduct")
					pcv_Shipped=rsOrdObj("pcPrdOrd_Shipped")
					if IsNull(pcv_Shipped) or pcv_Shipped="" then
						pcv_Shipped=1
					end if
					pquantity=rsOrdObj("quantity")
					punitPrice=rsOrdObj("unitPrice")
					QDiscounts=rsOrdObj("QDiscounts")
					ItemsDiscounts=rsOrdObj("ItemsDiscounts")
					
					'// Product Options Arrays
					pcv_strSelectedOptions = rsOrdObj("pcPrdOrd_SelectedOptions") ' Column 11
					pcv_strOptionsPriceArray = rsOrdObj("pcPrdOrd_OptionsPriceArray") ' Column 25
					pcv_strOptionsArray = rsOrdObj("pcPrdOrd_OptionsArray") ' Column 4
					
					pxdetails=rsOrdObj("xfdetails")
					pxdetails=replace(pxdetails,"|","<br>")
					pxdetails=replace(pxdetails,"::",":")
					'BTO ADDON-S
					if scBTO=1 then
						pidconfigSession=rsOrdObj("idconfigSession")
						if pidconfigSession="" then
							pidconfigSession="0"
						end if
					End If
					'BTO ADDON-E
					pdescription=rsOrdObj("description")
					pSku=rsOrdObj("sku")
					ptotal=rsOrdObj("total")
					ppaymentDetails=trim(rsOrdObj("paymentDetails"))
					ptaxamount=rsOrdObj("taxamount")
					pshipmentDetails=rsOrdObj("shipmentDetails")
					pdiscountDetails=rsOrdObj("discountDetails")
					GCDetails=rsOrdObj("pcOrd_GCDetails")
					porderstatus=rsOrdObj("orderstatus")
					pprocessDate=rsOrdObj("processDate")
					pprocessDate=ShowDateFrmt(pprocessDate)
					pshipdate=rsOrdObj("shipdate")
					pshipdate=ShowDateFrmt(pshipdate)
					pshipvia=rsOrdObj("shipvia")
					ptrackingNum=rsOrdObj("trackingNum")
					preturnDate=rsOrdObj("returnDate")
					preturnDate=ShowDateFrmt(preturnDate)
					preturnReason=rsOrdObj("returnReason")
					piRewardPoints=rsOrdObj("iRewardPoints")
					piRewardValue=rsOrdObj("iRewardValue")
					piRewardPointsCustAccrued=rsOrdObj("iRewardPointsCustAccrued")
					ptaxdetails=rsOrdObj("taxdetails")
					pcDPs=rsOrdObj("DPs")
					pcPrdOrd_BundledDisc=rsOrdObj("pcPrdOrd_BundledDisc")
					pIdConfigSession=trim(pidconfigSession)
					
					pOpPrices=0
					
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: Get the total Price of all options
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					pOpPrices=0
					dim pcv_tmpOptionLoopCounter, pcArray_TmpCounter
					
					If len(pcv_strOptionsPriceArray)>0 then
					
						pcArray_TmpCounter = split(pcv_strOptionsPriceArray,chr(124))
						For pcv_tmpOptionLoopCounter = 0 to ubound(pcArray_TmpCounter)
							pOpPrices = pOpPrices + pcArray_TmpCounter(pcv_tmpOptionLoopCounter)
						Next
						
					end if				

					if NOT isNumeric(pOpPrices) then
						pOpPrices=0
					end if	
					
					'// Apply Discounts to Options Total
					'   >>> call function "pcf_DiscountedOptions(OriginalOptionsTotal, ProductID, Quantity, CustomerType)" from stringfunctions.asp
					Dim pcv_intDiscountPerUnit
					pOpPrices = pcf_DiscountedOptions(pOpPrices, pidProduct, pquantity, pcustomerType)
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END: Get the total Price of all options
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
					%>
					
					<% 'BTO ADDON-S
					err.number=0
					TotalUnit=0
					If scBTO=1 then
						pIdConfigSession=trim(pidconfigSession)
						if pIdConfigSession<>"0" then 
							query="SELECT stringProducts, stringValues, stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
							set rsConfigObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsConfigObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							stringProducts=rsConfigObj("stringProducts")
							stringValues=rsConfigObj("stringValues")
							stringCategories=rsConfigObj("stringCategories")
							stringQuantity=rsConfigObj("stringQuantity")
							stringPrice=rsConfigObj("stringPrice")
							ArrProduct=Split(stringProducts, ",")
							ArrValue=Split(stringValues, ",")
							ArrCategory=Split(stringCategories, ",")
							ArrQuantity=Split(stringQuantity, ",")
							ArrPrice=Split(stringPrice, ",")
							set rsConfigObj=nothing
							for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
							
							query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
							set rsQ=connTemp.execute(query)
							tmpMinQty=1
							if not rsQ.eof then
								tmpMinQty=rsQ("pcprod_minimumqty")
								if IsNull(tmpMinQty) or tmpMinQty="" then
									tmpMinQty=1
								else
									if tmpMinQty="0" then
										tmpMinQty=1
									end if
								end if
							end if
							set rsQ=nothing
							tmpDefault=0
							query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
							set rsQ=connTemp.execute(query)
							if not rsQ.eof then
								tmpDefault=rsQ("cdefault")
								if IsNull(tmpDefault) or tmpDefault="" then
									tmpDefault=0
								else
									if tmpDefault<>"0" then
									 	tmpDefault=1
									end if
								end if
							end if
							set rsQ=nothing
										
							query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
							set rsConfigObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsConfigObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							if NOT validNum(ArrQuantity(i)) then
								pIntQty=1
							else
								pIntQty=ArrQuantity(i)
							end if
							
							if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
								if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
									if tmpDefault=1 then
										UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
									else
										UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
									end if
								else
									UPrice=0
								end if
								TotalUnit=TotalUnit+((ArrValue(i)+UPrice)*pQuantity)
							end if
							set rsConfigObj=nothing
							next
						end if 
					End If 
					'BTO ADDON-E


		
					if TotalUnit>0 then
						punitPrice1=punitPrice
						if pIdConfigSession<>"0" AND pIdConfigSession<>"" then
							pRowPrice1=Cdbl(pquantity * ( punitPrice1 )) - TotalUnit
							punitPrice1=Round(pRowPrice1/pquantity,2)
						else
							pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
						end if
					else
						punitPrice1=punitPrice
						if pIdConfigSession<>"0" AND pIdConfigSession<>"" then
							pRowPrice1=Cdbl(pquantity * ( punitPrice1 ))
						else
							pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
							punitPrice1=Round(pRowPrice1/pquantity,2)
						end if
					end if
					%>
		
					<tr> 
						<td><p><%=pquantity%></p></td>
						<td><p><%=pSku%></p></td>
						<td>
							<%=pdescription%>
							<%IF pcv_RWActive="1" THEN
							query="SELECT pcRE_IDProduct FROM pcRevExc WHERE pcRE_IDProduct=" & pidProduct
							set rsQ=server.CreateObject("ADODB.RecordSet")
							set rsQ=connTemp.execute(query)

							if err.number<>0 then
								call LogErrorToDatabase()
								set rs=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
	
							if rsQ.eof then
								Prv_Accept=1
							else
								Prv_Accept=0
							end if
							set rsQ=nothing
	
							IF Prv_Accept=1 THEN%>
							<br />
							<a href="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pidProduct%>');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a>
							<%END IF
							END IF%>
						</td>
						<td>
							<p align="right">
							<% if punitPrice1 > 0 then response.write(scCurSign&money(punitPrice1)) end if %>
							</p>
						</td>
						<td>
							<p align="right">
							<% if pRowPrice1 > 0 then response.write(scCurSign&money(pRowPrice1)) end if %>
							</p>
						</td>
						<%'Start SDBA%>
						<td>
							<%if pcv_HaveShipped=1 then
								if pcv_Shipped="1" then
									response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_3")
								else
									response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_2")
								end if
							end if%>
						</td>
						<%'End SDBA%>
					</tr>
		
					<% 'BTO ADDON-S
					err.number=0
					If scBTO=1 then
						pIdConfigSession=trim(pidconfigSession)
						if pIdConfigSession<>"0" then 
							query="SELECT stringProducts, stringValues, stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
							set rsConfigObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsConfigObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							stringProducts=rsConfigObj("stringProducts")
							stringValues=rsConfigObj("stringValues")
							stringCategories=rsConfigObj("stringCategories")
							stringQuantity=rsConfigObj("stringQuantity")
							stringPrice=rsConfigObj("stringPrice")
							ArrProduct=Split(stringProducts, ",")
							ArrValue=Split(stringValues, ",")
							ArrCategory=Split(stringCategories, ",")
							ArrQuantity=Split(stringQuantity, ",")
							ArrPrice=Split(stringPrice, ",")
							set rsConfigObj=nothing
							%>
				
							<tr> 
								<td>&nbsp;</td>
								<td colspan="4"> 
									<table class="pcShowBTOconfiguration">
										<tr> 
											<td colspan="2">  
												<p><%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_1")%></p>
											</td>
											<td>&nbsp;</td>
										</tr>
										<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
											query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and specProduct=" & pidProduct 
											set rs=server.CreateObject("ADODB.RecordSet") 
											set rs=conntemp.execute(query)
														
											btDisplayQF=rs("displayQF")
											set rs=nothing
											err.clear
											
											query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
											set rsQ=connTemp.execute(query)
											tmpMinQty=1
											if not rsQ.eof then
												tmpMinQty=rsQ("pcprod_minimumqty")
												if IsNull(tmpMinQty) or tmpMinQty="" then
													tmpMinQty=1
												else
													if tmpMinQty="0" then
														tmpMinQty=1
													end if
												end if
											end if
											set rsQ=nothing
											tmpDefault=0
											query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
											set rsQ=connTemp.execute(query)
											if not rsQ.eof then
												tmpDefault=rsQ("cdefault")
												if IsNull(tmpDefault) or tmpDefault="" then
													tmpDefault=0
												else
													if tmpDefault<>"0" then
													 	tmpDefault=1
													end if
												end if
											end if
											set rsQ=nothing
											
											query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
											set rsConfigObj=conntemp.execute(query)
											if err.number<>0 then
												call LogErrorToDatabase()
												set rsConfigObj=nothing
												call closedb()
												response.redirect "techErr.asp?err="&pcStrCustRefID
											end if
											if NOT validNum(ArrQuantity(i)) then
												pIntQty=1
											else
												pIntQty=ArrQuantity(i)
											end if
											strCategoryDesc=rsConfigObj("categoryDesc")
											strDescription=rsConfigObj("description") %>
											<tr> 
												<td width="85%" colspan="2" valign="top"> 
													<p>
													<%=strCategoryDesc%>:&nbsp;
													<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%>(<%=ArrQuantity(i)%>)&nbsp;<%end if%>
													<%=strDescription%>
													</p>
													</td>
												<%if pnoprices<2 then%>
												<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
												if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
													if tmpDefault=1 then
														UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
													else
														UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
													end if
												else
													UPrice=0
												end if
												'pfPrice=pfPrice+cdbl((ArrValue(i)+UPrice)*pQuantity) %>
												<%end if%> 
												<% end if %>
												<td width="15%" nowrap align="right" valign="top">
												<p>
												<%if pnoprices<2 then%>
													<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then%>
														<%=scCurSign & money((ArrValue(i)+UPrice)*pQuantity)%>
													<%else
														if tmpDefault=1 then%>
															<%=dictLanguage.Item(Session("language")&"_defaultnotice_1")%>
														<%end if
													end if%>
												<%end if%>
												</p>
												</td>
												
											</tr>
									<% set rsConfigObj=nothing
									next %>
									</table>
								</td>
								<td>&nbsp;</td>
							</tr>
						<% end if 
					End If 
					'BTO ADDON-E
					%>
					
					<!-- start options -->
					<%
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: SHOW PRODUCT OPTIONS
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					
					if (len(pcv_strSelectedOptions)>0 AND pcv_strSelectedOptions<>"NULL") then 
						%>
						<tr valign="top">
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td colspan="3">
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
							<%
							'#####################
							' START LOOP
							'#####################	
							
							'// Generate Our Local Arrays from our Stored Arrays  
							
							' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
							pcArray_strSelectedOptions = ""					
							pcArray_strSelectedOptions = Split(pcv_strSelectedOptions,chr(124))
							
							' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
							pcArray_strOptionsPrice = ""
							pcArray_strOptionsPrice = Split(pcv_strOptionsPriceArray,chr(124))
							
							' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
							pcArray_strOptions = ""
							pcArray_strOptions = Split(pcv_strOptionsArray,chr(124))
							
							' Get Our Loop Size
							pcv_intOptionLoopSize = 0
							pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
							
							' Start in Position One
							pcv_intOptionLoopCounter = 0
							
							' Display Our Options
							For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
							%>
							<tr>
							<td width="75%"><p><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></p></td>
							
							<td align="right" width="25%">									
							<% 
							tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
							
							if tempPrice="" or tempPrice=0 then
								response.write "&nbsp;"
							else
								'// Adjust for Quantity Discounts
								tempPrice = tempPrice - ((pcv_intDiscountPerUnit/100) * tempPrice)
								%>
								<table width="100%" cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td align="left" width="50%">
											<%=scCurSign&money(tempPrice)%>
										</td>
										<td align="right" width="50%">
											<%									
											tAprice=(tempPrice*Cdbl(pquantity))
											response.write scCurSign&money(tAprice) 
											%>
										</td>
									</tr>
								</table>
							<% 
							end if 
							%>			
							
							</td>
							</tr>
							<%
							Next
							'#####################
							' END LOOP
							'#####################					
							%>
                            </table>
                                
                            </td>
                        </tr>															
					<%					
                    end if
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ' END: SHOW PRODUCT OPTIONS
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    %>
                    <!-- end options -->
					
					<% 'BTO ADDON-S
    
                    err.number=0
                    pRowPrice=(punitPrice)*(pquantity)
					pExtRowPrice=pRowPrice
					Charges=0
                    If scBTO=1 then
                        pIdConfigSession=trim(pidconfigSession)
                        if pIdConfigSession<>"0" then
                            ItemsDiscounts=trim(ItemsDiscounts)
                            if ItemsDiscounts="" then
                                ItemsDiscounts=0
                            end if
                            if (ItemsDiscounts<>"") and (CDbl(ItemsDiscounts)<>"0") then
                                %>
                                <tr valign="top"> 
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><p><%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_2")%></p></td>
                                    <td>&nbsp;</td>
                                    <td><p align="right"><%=scCurSign&money(-1*ItemsDiscounts)%></p></td>
                                    <td>&nbsp;</td>
                                </tr>
                                <% pRowPrice=pRowPrice-Cdbl(ItemsDiscounts)
                            end if%>
							<% 'BTO Additional Charges
                                pIdConfigSession=trim(pidconfigSession)
                                if pIdConfigSession<>"0" then 
                                    query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
                                    set rsConfigObj=conntemp.execute(query)
                                    if err.number<>0 then
                                        call LogErrorToDatabase()
                                        set rsConfigObj=nothing
                                        call closedb()
                                        response.redirect "techErr.asp?err="&pcStrCustRefID
                                    end if
                                    stringCProducts=rsConfigObj("stringCProducts")
                                    stringCValues=rsConfigObj("stringCValues")
                                    stringCCategories=rsConfigObj("stringCCategories")
                                    ArrCProduct=Split(stringCProducts, ",")
                                    ArrCValue=Split(stringCValues, ",")
                                    ArrCCategory=Split(stringCCategories, ",")
                                    set rsConfigObj=nothing
                                    if ArrCProduct(0)<>"na" then%>
                                        <tr> 
                                            <td>&nbsp;</td>
                                            <td colspan="4"> 
                                                <table class="pcShowBTOconfiguration">
                                                    <tr> 
                                                        <td colspan="2">
                                                        <p><%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_5")%></p> 
                                                        </td>
                                                    </tr>
                                                <% for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
                                                        query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
                                                        set rsConfigObj=conntemp.execute(query)
                                                        if err.number<>0 then
                                                            call LogErrorToDatabase()
                                                            set rsConfigObj=nothing
                                                            call closedb()
                                                            response.redirect "techErr.asp?err="&pcStrCustRefID
                                                        end if
														strCategoryDesc=rsConfigObj("categoryDesc")
														strDescription=rsConfigObj("description")
														if (CDbl(ArrCValue(i))>0)then
															Charges=Charges+cdbl(ArrCValue(i))
														end if %>
														<tr> 
															<td width="85%" valign="top"><p><%=strCategoryDesc%>:	<%=strDescription%></p></td>
															<td width="15%" valign="top" nowrap align="right"><%if pnoprices<2 then%><%if ArrCValue(i)>0 then%><p><%=scCurSign & money(ArrCValue(i))%></p><%end if%><%end if%></td>
														</tr>
														<% set rsConfigObj=nothing
													next %>
												</table>
											</td>
											<td>&nbsp;</td>
										</tr>
						
										<% pRowPrice=pRowPrice+Cdbl(Charges)
                                    end if 'Have Charges
                                end if 
                            'BTO Additional Charges
						end if
					end if 'BTO
						
							QDiscounts=trim(QDiscounts)
							if QDiscounts="" then
								QDiscounts=0
							end if
							if (QDiscounts<>"") and (CDbl(QDiscounts)<>"0") then
								%>
                        
                        <tr>
									<td>&nbsp;</td>
                            <td colspan="4"> 
                                <table class="pcShowBTOconfiguration">
                                    <tr> 
                                    <td width="85%" colspan="2" valign="top"> 
                                    <p><%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_3")%></p></td>
                                    <td width="15%" nowrap align="right" valign="top">
                                    <p><%=scCurSign&money(-1*QDiscounts)%></p></td>
                                    </tr>
                                </table>
                            </td>
									<td>&nbsp;</td>
								</tr>
								<% pRowPrice=pRowPrice-Cdbl(QDiscounts)
							end if%>
						
						<% if pExtRowPrice<>pRowPrice then %>
							<tr valign="top">
								<td colspan="4" align="right"><p><strong><%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_4")%></strong></p></td>
								<td><p align="right"><% if pRowPrice > 0 then response.write(scCurSign&money(pRowPrice)) end if %></p></td>
								<td>&nbsp;</td>
							</tr>
                           <% end if %>
				
					<% if pRowPrice<>pRowPrice1 then
						subTotal=subTotal + cdbl(pRowPrice)
					else
						subTotal=subTotal + cdbl(pRowPrice1)
					end if
					%>
					
					<% 'show xtra options
                    '-----------------				
                    if pxdetails<>"" then
                    %>
                        <tr> 
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td colspan="2" style="padding-left:10px;"><%=pxdetails%></td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                        </tr>
                    <%
                    end if
					'----------------- %>
					<% count=count+1
                    If pshippingAddress="" then
                        'grab shipping address from shipping...
                        pshippingAddress=pAddress
                        pshippingAddress2=pAddress2
                        pshippingCity=pCity
                        pshippingState=pState
                        pshippingZip=pZip
                        pshippingCountryCode=pCountryCode
                    End if
                    'GGG Add-on start
					if pGWOpt<>"0" then
						query="select pcGW_OptName,pcGW_optPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
						set rsG=connTemp.execute(query)
						if not rsG.eof then%>
							<tr valign="top"> 
								<td width="41" >&nbsp;</td>
								<td colspan="3">
									<b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_42")%></b><%=rsG("pcGW_OptName")%><br>
									<%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_43")%><%=scCurSign & money(pGWPrice)%>
									<%if pGWText<>"" then%>
									<br>
									<b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_44")%></b><br><%=pGWText%>
									<%end if%>
									<br><br>
								</td>
								<td>&nbsp;</td>
							</tr>
						<%end if
					end if
					'GGG Add-on end
					if pcPrdOrd_BundledDisc>0 then %>
                        <tr valign="top"> 
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_54")%></p></td>
                            <td>&nbsp;</td>
                            <td><p align="right">-<%=scCurSign&money(pcPrdOrd_BundledDisc)%></p></td>
                            <td>&nbsp;</td>
                        </tr>
                    <%subTotal=subTotal - cdbl(pcPrdOrd_BundledDisc)
					end if

					rsOrdObj.movenext
				loop
				%>
			
			<tr> 
				<td class="pcSpacer" colspan="4"></td>
			</tr>
	
			<!-- start of processing charges -->
			<% dim payment, PaymentType,PayCharge
			payment = split(ppaymentDetails,"||")
			err.clear
			on error resume next
			PaymentType=payment(0)
			If payment(1)="" then
				if err.number<>0 then
					PayCharge=0
				end if
				PayCharge=0
			else
				PayCharge=payment(1)
			end If
			err.number=0
			subTotal=subTotal+PayCharge
			%>
			
			<% if PayCharge>0 then %>
				<tr>
					<td>&nbsp;</td>
					<td colspan="3" valign="top"><p align="left">
					<b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_14b")%></b>
					</p></td>
					<td valign="top"><p align="right"><%=scCurSign&money(PayCharge)%></p></td>
					<td>&nbsp;</td>
				</tr>
			<% end if %>
			<!-- end of processing charges -->
			
			<%if subTotal>0 then%>
			<tr>
					<td>&nbsp;</td>
					<td colspan="3" valign="top"><p align="left">
					<b><%response.write dictLanguage.Item(Session("language")&"_orderverify_15")%></b>
					</p></td>
					<td valign="top"><p align="right"><%=scCurSign&money(subTotal)%></p></td>
					<td>&nbsp;</td>
				</tr>
			<% end if %>
			<!-- start of discount details -->
			<% if pcv_CatDiscounts>"0" then	%>
						<td>&nbsp;</td>
						<td colspan="3" valign="top"><p align="left"><b><%response.write dictLanguage.Item(Session("language")&"_catdisc_2")%></b></p></td>
				<td valign="top"><p align="right">-<%=scCurSign&money(pcv_CatDiscounts)%></p></td>
				<td>&nbsp;</td>
				</tr>
			<% end if %>
			
			<% if instr(pdiscountDetails,",") then
				DiscountDetailsArry=split(pdiscountDetails,",")
				intArryCnt=ubound(DiscountDetailsArry)
				for k=0 to intArryCnt
					if (DiscountDetailsArry(k)<>"") AND (instr(DiscountDetailsArry(k),"- ||")=0) then
						DiscountDetailsArry(k+1)=DiscountDetailsArry(k)+"," + DiscountDetailsArry(k+1)
						DiscountDetailsArry(k)=""
					end if
				next
			else
				intArryCnt=0
			end if
			
			dim discounts, discountType 
			
			for k=0 to intArryCnt
				if intArryCnt=0 then
					pTempDiscountDetails=pdiscountDetails
				else
					pTempDiscountDetails=DiscountDetailsArry(k)
				end if
				if instr(pTempDiscountDetails,"- ||") then
					discounts = split(pTempDiscountDetails,"- ||")
					discountType = discounts(0)
					discount = discounts(1)
					%>
					<tr> 
						<td>&nbsp;</td>
						<td colspan="3" valign="top"><p align="left"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_15")%></b>&nbsp;<% response.write discountType%></p></td>
					<td valign="top">
					<% if discount <> 0 then %>
					<p align="right">-<%=scCurSign&money(discount)%></p>
					<% end if %>
					</td>
					<td>&nbsp;</td>
					</tr>
				<% end if
			Next %>
			<!-- end if discount details -->
			
			<%'start of gift certificates
				if GCDetails<>"" then
					GCArry=split(GCDetails,"|g|")
					intArryCnt=ubound(GCArry)
				
					for k=0 to intArryCnt
					
					if GCArry(k)<>"" then
						GCInfo = split(GCArry(k),"|s|")
						if GCInfo(2)="" OR IsNull(GCInfo(2)) then
							GCInfo(2)=0
						end if
						%>
						<tr> 
							<td>&nbsp;</td>
							<td colspan="3" valign="top"><p align="left"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_15A")%></b>&nbsp;<%=GCInfo(1)%> (<%=GCInfo(0)%>)</p></td>
						<td valign="top" align="right">
						<% if Cdbl(GCInfo(2)) <> 0 then %>
						-<%=scCurSign&money(GCInfo(2))%>
						<% end if %>
						</td>
						</tr>
					<% end if
					Next
				end if
			'end if gift certificates
			%>
			
			<!-- start of rewards used -->
			<% if piRewardPoints>0 then %>
				<tr> 
					<td>&nbsp;</td>
					<td colspan="3" valign="top"><p align="left"><%response.write "<b>"&piRewardPoints&"&nbsp;"&RewardsLabel&" </b>" & dictLanguage.Item(Session("language")&"_orderverify_31")%></p></td>
					<td valign="top"><p align="right">- <% response.write scCurSign& money(piRewardValue) %></p></td>
					<td>&nbsp;</td>
				</tr>
			<% end if %>
			<!-- end if rewards -->
			
			<!-- start of rewards earned -->
			<% if piRewardPointsCustAccrued>0 then %>
				<tr>
					<td>&nbsp;</td>
					<td colspan="4" valign="top">
					<p align="left"><b><% response.write dictRewardsLanguage.Item(Session("rewards_language")&"_orderverify")%></b><%=dictLanguage.Item(Session("language")&"_orderverify_30")%><% response.Write(piRewardPointsCustAccrued) %>
					</p>
					</td>
					<td>&nbsp;</td>
				</tr>
			<% end if %>
			<!-- end if rewards -->
			<%'GGG Add-on start
			if pGWTotal>0 then%>
				<tr> 
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td colspan="2"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_42")%></b></td>
					<td align="right"><%=scCurSign & money(pGWTotal)%></td>
				</tr>
			<%end if
			'GGG Add-on end%>
			<!-- start of shipping -->
			<% dim shipping, varShip, Shipper, Service, Postage, serviceHandlingFee
			shipping=split(pshipmentDetails,",")

			if ubound(shipping)>1 then
				if NOT isNumeric(trim(shipping(2))) then
					varShip="0"
					response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
				else
					Shipper=shipping(0)
					Service=shipping(1)
					Postage=trim(shipping(2))
					if ubound(shipping)=>3 then
						serviceHandlingFee=trim(shipping(3))
						if NOT isNumeric(serviceHandlingFee) then
							serviceHandlingFee=0
						end if
					else
						serviceHandlingFee=0
					end if
				end if
			else
				varShip="0"
			end if 
			
			'// Postage Total
			if varShip<>"0" then
			%>
				<tr> 
					<td>&nbsp;</td>
					<td colspan="3" valign="top"><p align="left"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_13")%></b>&nbsp;<%=Service%></p>
					</td>
					<td valign="top">
					<p align="right"><%=scCurSign&money(Postage)%></p>
					</td>
					<td>&nbsp;</td>
				</tr>
			<% End If %>
			
			<% 
			'// Handling Fee Charge
			if serviceHandlingFee>0 then %>
				<tr> 
					<td valign="top">&nbsp;</td>
					<td colspan="3" valign="top"><p align="right"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_14")%></b></p></td>
					<td valign="top"><p align="right"><%=scCurSign&money(serviceHandlingFee)%></p></td>
					<td>&nbsp;</td>
				</tr>
			<% end if %>
			<!-- end of shipping -->
			
			<!-- start of taxes -->
			<% ' If the store is using VAT and VAT is > 0, don't show any taxes here, but show VAT after the total
			if pord_VAT>0 then
			else
				if isNull(ptaxDetails) or trim(ptaxDetails)="" then %>
					<tr> 
						<td valign="top">&nbsp;</td>
						<td colspan="3" valign="top"><p align="right"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_12")%></b></p></td>
						<td valign="top"><p align="right"><% response.write scCurSign& money(ptaxAmount)%></p></td>
						<td>&nbsp;</td>
					</tr>
				<% else %>
					<% dim taxArray, taxDesc
					taxArray=split(ptaxDetails,",")
					for i=0 to (ubound(taxArray)-1)
						taxDesc=split(taxArray(i),"|")
						if taxDesc(0)<>"" then %>
						<tr> 
							<td valign="top">&nbsp;</td>
							<td colspan="3" valign="top"><p align="right"><b><%=taxDesc(0)%></b></p></td>
							<td valign="top"><p align="right"><% response.write scCurSign& money(taxDesc(1))%></p></td>
							<td>&nbsp;</td>
						</tr>
						<% end if
					next %>
				<% end if 
			end if %>
			<!-- end if taxes -->
			<tr> 
				<td valign="top">&nbsp;</td>
				<td colspan="3" valign="top"><p align="right"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_12")%></b></p></td>
				<td valign="top"><p align="right"><% response.write scCurSign& money(ptotal) %></p></td>
				<td>&nbsp;</td>
			</tr>
			
			<% 
			' START - VAT
			' If the store is using VAT and VAT > 0, show it here
			' Show different message depending on whether VAT is included or excluded
			if pord_VAT>0 then 
				Dim pcv_IsEUMemberState
				pcv_IsEUMemberState = pcf_IsEUMemberState(pshippingCountryCode)
				VATRemovedTotal=0
				if pcv_IsEUMemberState=0 then
					VATRemovedTotal=pord_VAT
				end if	
			%>
				<tr> 
					<td colspan="5" align="right" class="pcSmallText">
						<% if VATRemovedTotal=0 then %>
                            <p><% response.write dictLanguage.Item(Session("language")&"_orderverify_35") & scCurSign & money(pord_VAT) %></p>
                        <% else %>
                            <p><% response.write dictLanguage.Item(Session("language")&"_orderverify_42") & scCurSign & money(pord_VAT) %></p>
                        <% end if %>
					</td>
				</tr>
			<% 
			end if 
			' END - VAT
			%>
			
			<!--RMA CREDIT-->
			<% if NOT isNull(prmaCredit) AND prmaCredit<>"" AND prmaCredit>0 then %>
				<tr> 
					<td valign="top">&nbsp;</td>
					<td colspan="3" valign="top"><p align="right"><b><%response.write  dictLanguage.Item(Session("language")&"_CustviewPastD_31")%></b></p></td>
					<td valign="top" nowrap="nowrap"><p align="right"><% response.write "-"&scCurSign& money(prmaCredit) %></p></td>
					<td>&nbsp;</td>
				</tr>
			<% end if %>
			</table>
			</td>
		</tr>
		<% 
		' END Order Details
		
		' START Other order information
		%>
		<tr>
			<td>
        	<hr>
            <form method="post" name="form2" id="form2" action="#" class="pcForms">
            <table class="pcShowContent">
						<%'Start SDBA
                        'Show PaymentStatus%>
                        <tr>
                            <td>
                            <p><%response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_4")
                            Select Case pcv_PaymentStatus
                                Case 1: response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_5")
                                Case 2: response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_6")
                                Case 3:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_9")
                                Case 4:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_10")
                                Case 5:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_11")
                                Case 6:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_12")
                                Case 7:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_13")
                                Case 8:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_14")
                                Case Else: response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_2")
                            End Select%></p>
                            </td>
                        </tr>
						<%'End SDBA%>
                        <% if piRewardPointsCustAccrued>0 AND int(pOrderStatus)>2 AND int(pOrderStatus)<>5 AND int(pOrderStatus)<>6 then %>
                            <tr> 
                                <td>
                                <p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_16")%><%=piRewardPointsCustAccrued%>&nbsp;<%=RewardsLabel%><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_17")%>
                                </p>
                                </td>
                            </tr>
                        <% end if %>
					
                        <!-- if order was cancelled -->
                        <% if int(pOrderStatus)=5 then %>
                            <tr> 
                                <td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_18")%></p></td>
                            </tr>
                        <% else %>
                            
                            <!-- if order was returned -->
                            <% if int(pOrderStatus)=6 then %>
                                <tr> 
                                    <td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_26")%></p></td>
                                </tr>
                                <tr> 
                                    <td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_37")%></p></td>
                                </tr>
                                <tr> 
                                    <td><hr></td>
                                </tr>
                            <% end if %>
                            <!-- end order returned -->
		
				
                            <!-- order has been processed, show date -->
                            <% if int(pOrderStatus)>2 then %>
                                <tr> 
                                    <td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_22b")%></p></td>
                                </tr>
                                <tr> 
                                    <td>
                                    <p><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_22") & pprocessDate %></p>
                                    </td>
                                </tr>
                            <% else %>
                            <!-- else if order has not been processed, tell customer -->
                                <tr> 
                                    <td> 
                                        <p><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_20")%></p>
                                    </td>
                                </tr>
                            <% end if %>
							<!-- end order processed check -->	
						
							<!-- if order has been shipped, show information -->
							<% 
							if (int(pOrderStatus)=4 OR int(pOrderStatus)>= 6) then %>
								<tr> 
									<td><hr></td>
								</tr>
								<% if int(pOrderStatus)=7 then %>
									<tr> 
										<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_46")%></td>
									</tr>
								<% else %>
									<tr> 
										<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_19")%></td>
									</tr>
								<% end if %>
								<% if pShippingFullName<>"" then %>
									<tr>
										<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_20")%></p></td>
									</tr>
									<tr> 
										<td><p><%=pShippingFullName%></p></td>
									</tr>
									<tr>
										<td><p><%=pShippingCompany%></p></td>
									</tr>
									<tr>
										<td><p><%=pShippingAddress%></p></td>
									</tr>
								<% else %>
									<tr>
										<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_20")%></p></td>
									</tr>
									<tr>
										<td><p><%=pShippingAddress%></p></td>
									</tr>
								<% end if %>
								
								<% if pShippingAddress2<>"" then %>
									<tr>
										<td><p><%=pShippingAddress2 %></p></td>
									</tr>
								<% end if %>
									
								<tr> 
									<td><p><% response.write pShippingCity&", "&pshippingStateCode&" "&pShippingZip %></p></td>
								</tr>
								<tr> 
									<td><p><%=pShippingCountryCode %></p></td>
								</tr>
								<tr class="pcSpacer"> 
									<td></td>
								</tr>
								
								<%
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: Shippment Information
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~				
								%>	
								<% if pshipDate="//" OR isNULL(pshipVia)=True then %>
									<%
									Dim rsShipInfo
									query="SELECT pcPackageInfo_ID, pcPackageInfo_ShipMethod,pcPackageInfo_TrackingNumber,pcPackageInfo_ShippedDate,pcPackageInfo_Comments,pcPackageInfo_UPSPackageType FROM pcPackageInfo WHERE idOrder=" & pidorder & ";"
									set rsPackages=server.CreateObject("ADODB.RecordSet")
									set rsPackages=connTemp.execute(query)
									if not rsPackages.eof then
										pcIdNum = 1								
										do while not rsPackages.eof	
											pcv_PackageID =	rsPackages("pcPackageInfo_ID")
											ptrackingNum = ""
											pshipVia = ""
											tmp_ShipMethod=rsPackages("pcPackageInfo_ShipMethod")
											tmp_TrackingNumber=rsPackages("pcPackageInfo_TrackingNumber")
											tmp_ShippedDate=rsPackages("pcPackageInfo_ShippedDate")
											tmp_UPSPackageType=rsPackages("pcPackageInfo_UPSPackageType")
											ptrackingNum=tmp_TrackingNumber
											pshipVia=tmp_ShipMethod
												
											'// Show the Shipment Info for v3 package
											query="SELECT quantity, Description FROM Products INNER JOIN ProductsOrdered ON (products.idproduct=ProductsOrdered.idproduct) WHERE productsOrdered.pcPackageInfo_ID=" & pcv_PackageID & " ORDER BY ProductsOrdered.pcPackageInfo_ID;"
											set rsShipInfo=server.CreateObject("ADODB.RecordSet")
											set rsShipInfo=conntemp.execute(query)
											if err.number<>0 then
												call LogErrorToDatabase()
												set rsShipInfo=nothing
												call closedb()
												response.redirect "techErr.asp?err="&pcStrCustRefID
											end if
											%>
											<tr style="background-color: #e5e5e5"> 
												<td><p><strong><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_41")%></strong></p></td>
											</tr>
											<tr> 
												<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_21")%>&nbsp;<%=ShowDateFrmt(tmp_ShippedDate) %></p></td>
											</tr>
											<tr> 
												<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_22")%>&nbsp;<%=tmp_ShipMethod %></p></td>
											</tr>
											<% 
											'Show Tracking Number
											if ptrackingNum<>"" then 
												%>
												<tr> 
													<td>
													<p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_25")%>
													<% 
													'//  Start: Tracking Link
													if instr(ucase(tmp_ShipMethod),"UPS:") OR tmp_UPSPackageType<>"" then %>
														<a href="custUPSTracking.asp?itracknumber=<%=ptrackingNum%>"><%=ptrackingNum %></a>
													<% elseif instr(ucase(tmp_ShipMethod),"FEDEX:") then %>
														&nbsp;
                                                        <a href="http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=<%=ptrackingNum%>" target="_blank"><%=ptrackingNum %></a>
													<% else 
															response.write " " & ptrackingNum
													end if 
													'//  End: Tracking Link	
													%>
													</p>
													</td>
												</tr>
												<% 
											end if
											'End if Tracking Number 
											%>
											<tr> 
												<td><p><a href="JavaScript:pcf_ShowContents('PackageTable<%=pcIdNum%>');"><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_42")%></a></p></td>
											</tr>
											<tr> 
												<td>
													<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
													<!--
													function pcf_ShowContents(obj){										
														if(document.getElementById){											
														var tablename = document.getElementById(obj);											
															if(tablename.style.display != ''){
																tablename.style.display='';
															} else {
																tablename.style.display = 'none';
															}
														}
													}
													//-->
													</SCRIPT>
													<p>											
													<table id="PackageTable<%=pcIdNum%>" class="pcShowContent" style="display: none;">
														<tr style="background-color: #ffffcc">
															<td>Product Name</td>
															<td>Qty</td>
														</tr>	
														<%
														if not rsShipInfo.eof then
															do while NOT rsShipInfo.eof
																if tmp_ShipMethod<>"" OR tmp_TrackingNumber<>"" OR tmp_ShippedDate<>"" then				
																pcv_PrdName=rsShipInfo("Description")
																pcv_PrdQty=rsShipInfo("quantity")
																%>												
																<tr>
																	<td><%=pcv_PrdName%></td>
																	<td><%=pcv_PrdQty%></td>
																</tr>							
																<%
																pcIdNum = pcIdNum + 1
																end if
																rsShipInfo.movenext
															loop
															set rsShipInfo = nothing
														end if %>
													</table>											
													</p>
												</td>
											</tr>
											<% rsPackages.movenext
										loop
									end if
									set rsPackages = nothing
									%>
									<tr>
										<td><hr></td>
									</tr>	
									<% if pOrdPackageNum <> "" then %>
										<% if int(pOrderStatus)=7 then %>							
											<tr> 
												<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_43")%>&nbsp;<%=pOrdPackageNum %></p></td>
											</tr>
										<% else %>
											<tr> 
												<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_38")%>&nbsp;<%=pOrdPackageNum %></p></td>
											</tr>
										<% end if %>
									<% end if %>
									<% if pOrdPackageNum <> "" AND int(pOrderStatus)=7 then %>
										<tr> 
											<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_44")%><%=(pcIdNum - 1)%><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_45")%><%=pOrdPackageNum %></p></td>
										</tr>
									<% end if %>
								<% else %>				
									<%
									'// Show the Shipment Info for v2.76 package
									%>
									<tr> 
										<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_21")%><%=pshipDate %></p></td>
									</tr>
									<tr> 
										<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_22")%><%=pshipVia %></p></td>
									</tr>
									<tr> 
										<td><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_38")%><%=pOrdPackageNum %></p></td>
									</tr>
									
									<% 'Show Tracking Number
									if ptrackingNum<>"" then %>
										<tr> 
											<td>
												<p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_25")%>
												<% if instr(ucase(pshipVia),"UPS") then %>
													<a href="custUPSTracking.asp?itracknumber=<%=ptrackingNum%>"><%=ptrackingNum %></a>
												<% else 
													if instr(ucase(pshipVia),"FEDEX") then 
														if ucase(strFedExCountryCode)="US" then %>
															<a href="http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=<%=ptrackingNum%>" target="_blank"><%=ptrackingNum %></a>
														<% else %>
															<a href="http://www.fedex.com/Tracking?cntry_code=<%=strFedExCountryCode%>" target="_blank"><%=ptrackingNum %></a>
														<% end if %>
													<% else 
														response.write ptrackingNum
													end if
												end if %>
												</p>
											</td>
										</tr>
									<% 
									end if
									'End if Tracking Number 
									%>					
								
								<% end if %>
								<%
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: Shippment Information
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~				
								%>
									
						
								<!-- if RMA has not been issued, show link to RMA request form, otherwise show message -->
								<%
								IF scHideRMA <> 1 THEN ' START - Check if the store allows customers to request an RMA
									Dim rsRma, rmaVar, rmaNumber, rmaReturnStatus, queryrma
							
									queryrma="SELECT rmaNumber, rmaReturnStatus, rmaApproved FROM PCReturns WHERE idOrder=" &pIdOrder
									set rsRma=conntemp.execute(queryrma)
									if err.number<>0 then
										call LogErrorToDatabase()
										set rsRma=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									if NOT rsRma.eof then
										rmaNumber = rsRma("rmaNumber")
										rmaReturnStatus = rsRma("rmaReturnStatus")
										rmaApproved = rsRma("rmaApproved")
										'0=pending, 1=approved, 2=denied
										rmaVar = 1
									else
										rmaVar = 0
									end if
							
									Set rsRma = nothing
									%>				
									<tr>
										<td><hr></td>
									</tr>
									<tr style="background-color: #e5e5e5">
										<td>
										<p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_47")%></p>
										</td>
									</tr>	
									<% 
									if rmaVar = 0 then
									' RMA can be requested. The customer has not requested it. Show link.
									%>
										<tr>
											<td>
											<p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_23")%><a href="rmaindex.asp?idorder=<%=pIdOrder%>"><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_24")%></a></p>
											</td>
										</tr>
									<% 
									else ' An RMA has already been requested by the customer or issued by the store manager
										if rmaApproved=0 then ' An RMA request has not yet been approved
										%>
											<tr>
												<td>
												<p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_30")%></p>
												</td>
											</tr>
										<%
										end if 
										if rmaApproved=1 then ' An RMA request has been approved
										%>
											<tr><td colspan="2"><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_34")%></p></td></tr>
											<tr><td colspan="2"><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_31")%> <b><%=rmaNumber%></b></p></td>
											</tr>
										<%
										end if 
										if rmaApproved=2 then ' An RMA request has been denied
										%>
											<tr><td colspan="2"><hr></td></tr>
											<tr><td colspan="2"><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_35")%></p></td></tr>
										<%
										end if %>
										<%
										if trim(RmaReturnStatus) <> "" then ' Admin comments related to the RMA request
										%>
											<tr>
												<td colspan="2"><p><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_32")%>&nbsp;<%=RmaReturnStatus%></p></td>
											</tr>
										<%
										end if ' End RMA Comments
									end if ' End RMA has already been requested
								END IF ' END - Check if the store allows customers to request an RMA
								%>
								<!-- End RMA link -->
						
								<!-- end shipping info -->
							<% end if
						end if	%>
					
                        <!-- START GGG Infor and Downloadable Products Information -->
                        <tr> 
                            <td colspan="2">
                                <%'GGG Add-on start
                                IF (GCDetails<>"") then %>
                                    <hr>
                                    <table class="pcShowContent">
                                    <tr>
                                        <th colspan="2"><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_45")%></th>
                                    </tr>
                                    <%
									GCArry=split(GCDetails,"|g|")
									intArryCnt=ubound(GCArry)
				
									for k=0 to intArryCnt
					
									if GCArry(k)<>"" then
										GCInfo = split(GCArry(k),"|s|")
										if GCInfo(2)="" OR IsNull(GCInfo(2)) then
										GCInfo(2)=0
										end if
										pGiftCode=GCInfo(0)
										pGiftUsed=GCInfo(2)
                                    query="select products.IDProduct,products.Description from pcGCOrdered,Products where products.idproduct=pcGCOrdered.pcGO_idproduct and pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
                                    set rsG=connTemp.execute(query)
        
                                    if not rsG.eof then
                                        pIdproduct=rsG("idproduct")
                                        pName=rsG("Description")
                                        pCode=pGiftCode
                                        %>
                                        <tr> 
                                            <td width="18%" nowrap><b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_46")%></b></font></td>
                                            <td width="82%"><b><%=pName%></b></font></td>
                                        </tr>
                                        <tr> 
                                            <td width="18%" nowrap valign="top">&nbsp;</td>
                                            <td width="82%" valign="top">
                                                <%
                                                query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_GcCode='" & pGiftCode & "'"
                                                set rs19=connTemp.execute(query)
                                
                                                if not rs19.eof then%>
                                                    <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_47")%><b><%=rs19("pcGO_GcCode")%></b><br>
                                                    <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_48")%><%=scCurSign & money(pGiftUsed)%><br><br>
                                                    <%
                                                    pGCAmount=rs19("pcGO_Amount")
                                                    if cdbl(pGCAmount)<=0 then%>
                                                        <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_49")%>
                                                    <%else%>
                                                        <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_50")%><b><%=scCurSign & money(pGCAmount)%></b>
                                                        <br>
                                                        <%pExpDate=rs19("pcGO_ExpDate")
                                                        if year(pExpDate)="1900" then%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_51")%>
                                                        <%else
                                                            if scDateFrmt="DD/MM/YY" then
                                                                pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
                                                            else
                                                                pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
                                                            end if%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_52")%><font color=#ff0000><b><%=pExpDate%></b></font>
                                                        <%end if%>
                                                        <br>
                                                        <%
                                                        pGCStatus=rs19("pcGO_Status")
                                                        if pGCStatus="1" then%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_53")%><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_53a")%>
                                                        <%else%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_53")%><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_53b")%>
                                                        <%end if%>
                                                    <%end if%>
                                                    <br><br>
                                                <%end if
                                                set rs19=nothing%>
                                            </td>
                                        </tr>
                                    <%end if
                                    set rsG=nothing
									end if
									Next%>
                                    </table>
                                <% END IF
                                'GGG Add-on end%>
                        
                                <% 
                                '// we do not hide the download link on partial return
                                If (int(pOrderStatus)>2 AND int(pOrderStatus)<=4) OR (int(pOrderStatus)>=7)  Then
                                    If (pcDPs<>"") and (pcDPs="1") then %>
                                        <hr>
                                        <table class="pcShowContent">
                                            <tr>
                                                <th colspan="2">
                                                <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_23")%>
                                                </th>
                                            </tr>
                                                
                                            <% query="select IdProduct from DPRequests WHERE IdOrder=" & pidorder & ";"
                                            set rsLic=connTemp.execute(query)
                                            if err.number<>0 then
                                                call LogErrorToDatabase()
                                                set rsLic=nothing
                                                call closedb()
                                                response.redirect "techErr.asp?err="&pcStrCustRefID
                                            end if
                                            do while not rsLic.eof
                                                pIdproduct=rsLic("idproduct")
                                                query="select Description, URLExpire, ExpireDays, License, LicenseLabel1, LicenseLabel2, LicenseLabel3, LicenseLabel4, LicenseLabel5 from Products,DProducts where products.idproduct=" & pIdproduct & " and DProducts.idproduct=Products.idproduct and products.downloadable=1"
                                                set rstemp=connTemp.execute(query)
                                                if err.number<>0 then
                                                    call LogErrorToDatabase()
                                                    set rstemp=nothing
                                                    call closedb()
                                                    response.redirect "techErr.asp?err="&pcStrCustRefID
                                                end if
                            
                                                if not rstemp.eof then
                                                    pName=rstemp("Description")
                                                    pURLExpire=rstemp("URLExpire")
                                                    pExpireDays=rstemp("ExpireDays")	
                                                    pLicense=rstemp("License")
                                                    pLL1=rstemp("LicenseLabel1")
                                                    pLL2=rstemp("LicenseLabel2")
                                                    pLL3=rstemp("LicenseLabel3")
                                                    pLL4=rstemp("LicenseLabel4")
                                                    pLL5=rstemp("LicenseLabel5")
                                                    
                                                    set rstemp = nothing
                                    
                                                    query="select RequestSTR,StartDate from DPRequests where idproduct=" & pIdproduct & " and idorder=" & pidorder & " and idcustomer=" & pidcustomer
                                                    set rstemp=connTemp.execute(query)
                                                    if err.number<>0 then
                                                        call LogErrorToDatabase()
                                                        set rstemp=nothing
                                                        call closedb()
                                                        response.redirect "techErr.asp?err="&pcStrCustRefID
                                                    end if
                                                    pdownloadStr=rstemp("RequestSTR")
                                                    pStartDate=rstemp("StartDate")
                                                    SPath1=split(Request.ServerVariables("PATH_INFO"),"/pc/")
                                                    
                                                    if SPath1(0)<>"" then
                                                    else
                                                        SPath1(0)="/"
                                                    end if
                                                    
                                                    if SPath1(0)<>"/" then
                                                        if Left(SPath1(0),1)<>"/" then
                                                            SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & "/" & SPath1(0)
                                                        else
                                                            SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1(0)
                                                        end if
                                                    else
                                                        SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & "/"
                                                    end if
                                
                                                    if Right(SPathInfo,1)="/" then
                                                        pdownloadStr=SPathInfo & "pc/pcdownload.asp?id=" & pdownloadStr					
                                                    else
                                                        pdownloadStr=SPathInfo & "/pc/pcdownload.asp?id=" & pdownloadStr
                                                    end if
                                
                                                    set rstemp=nothing %>
                                                    <tr> 
                                                        <td width="18%" nowrap>
                                                        <p><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_24")%></p>
                                                        </td>
                                                        <td width="82%"><p><b><%=pName%></b></p></td>
                                                    </tr>
                                                    <tr> 
                                                        <td nowrap valign="top">
                                                        <p><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_25")%></p>
                                                        </td>
                                                        <td width="82%">
                                                        <p><a href="<%=pdownloadStr%>" target="_blank"><%=pdownloadStr%></a></p>
                                                        <p>
                                                        <% if (pURLExpire<>"") and (pURLExpire="1") then
                                                                if date()-(CDate(pStartDate)+pExpireDays)<0 then%>
                                                                <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_26")%><%=(CDate(pStartDate)+pExpireDays)-date()%><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_27")%>
                                                            <%else
                                                                if date()-(CDate(pStartDate)+pExpireDays)=0 then%>
                                                                    <p><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_28")%></p>
                                                                <%else%>
                                                                    <p><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_29")%></p>
                                                                <%end if
                                                            end if
                                                        end if%>
                                                        </p>
                                                        </td>
                                                    </tr>
                                                    <%if (pLicense<>"") and (pLicense="1") then %>
                                                        <tr> 
                                                            <td nowrap valign="top">
                                                            <p><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_30")%></p>
                                                            </td>
                                                            <td>
                                                            <% query="select Lic1, Lic2, Lic3, Lic4, Lic5 from DPLicenses where idproduct=" & pIdproduct & " and idorder=" & pidorder
                                                            set rstemp=connTemp.execute(query)
                                                            if err.number<>0 then
                                                                call LogErrorToDatabase()
                                                                set rstemp=nothing
                                                                call closedb()
                                                                response.redirect "techErr.asp?err="&pcStrCustRefID
                                                            end if
                                                            do while not rstemp.eof
                                                                Lic1=rstemp("Lic1")
                                                                Lic2=rstemp("Lic2")
                                                                Lic3=rstemp("Lic3")
                                                                Lic4=rstemp("Lic4")
                                                                Lic5=rstemp("Lic5")
                                                                %>
                                                                <table class="pcShowContent">
                                                                <% if Lic1<>"" then%>
                                                                    <tr><td nowrap><p><%=pLL1%>:</p></td><td><p><%=Lic1%></p></td></tr>
                                                                <%end if
                                                                if Lic2<>"" then%>
                                                                    <tr><td nowrap><p><%=pLL2%>:</p></td><td><p><%=Lic2%></p></td></tr>
                                                                <%end if
                                                                if Lic3<>"" then%>
                                                                    <tr><td nowrap><p><%=pLL3%>:</p></td><td><p><%=Lic3%></p></td></tr>
                                                                <%end if
                                                                if Lic4<>"" then%>
                                                                    <tr><td nowrap><p><%=pLL4%>:</p></td><td><p><%=Lic4%></p></td></tr>
                                                                <%end if
                                                                if Lic5<>"" then%>
                                                                    <tr><td nowrap><p><%=pLL5%>:</p></td><td><p><%=Lic5%></p></td></tr>
                                                                <%end if%>
                                                                </table>
                                                                <%rstemp.movenext
                                                            loop
                                                            set rstemp=nothing
                                                            %>
                                                            </td>
                                                        </tr>
                                                    <%end if
                                                end if
                                                rsLic.MoveNext
                                            loop
                                            set rsLic=nothing
                                            call closedb()
                                            %>
                                        </table>
                                        <!-- END Downloadable products -->
                                        <% 
                                    end if
                                end if 'Order Status 3, 4
                                %>
                                    
                                <%'GGG Add-on start
                                If (int(pOrderStatus)>2 AND int(pOrderStatus)<=4) OR (int(pOrderStatus)>=7) Then
                                    IF (pGCs<>"") and (pGCs="1") then
									call opendb() %>
                                        <hr>
                                        <table class="pcShowContent">
                                        <tr>
                                            <th colspan="2"><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_33")%></th>
                                        </tr>
                                        <%
                                        query="select * from ProductsOrdered WHERE idOrder="& pidorder
                                        set rs11=connTemp.execute(query)
                                        do while not rs11.eof
                                            query="select products.Description,pcGCOrdered.pcGO_GcCode from Products,pcGCOrdered where products.idproduct=" & rs11("idproduct") & " and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& pidorder
                                            set rsG=connTemp.execute(query)
        
                                            if not rsG.eof then
                                                pIdproduct=rs11("idproduct")
                                                pGCName=rsG("Description")
                                                pCode=rsG("pcGO_GcCode")
                                                %>
                                                <tr>
                                                    <td width="18%" nowrap><b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_34")%></b></td>
                                                    <td width="82%"><b><%=pGCName%></b></td>
                                                </tr>
                                                <tr> 
                                                    <td width="18%" nowrap>&nbsp;</td>
                                                    <td width="82%" valign="top">
                                                    <%
                                                    query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & pidorder
                                                    set rs19=connTemp.execute(query)
            
                                                    do while not rs19.eof%>
                                                        <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_35")%>&nbsp;<b><%=rs19("pcGO_GcCode")%></b><br>
                                                        <%pExpDate=rs19("pcGO_ExpDate")
                                                        if year(pExpDate)="1900" then%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_36b")%>
                                                        <%else
                                                            if scDateFrmt="DD/MM/YY" then
                                                                pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
                                                            else
                                                                pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
                                                            end if%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_36")%>&nbsp;<font color=#ff0000><b><%=pExpDate%></b></font>
                                                        <%end if%>
                                                        <br>
                                                        <%
                                                        pGCAmount=rs19("pcGO_Amount")
                                                        if cdbl(pGCAmount)<=0 then%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_37b")%>
                                                        <%else%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_37")%>&nbsp;<b><%=scCurSign & money(pGCAmount)%></b>
                                                        <%end if%><br>
                                                        <%
                                                        pGCStatus=rs19("pcGO_Status")
                                                        if pGCStatus="1" then%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_38")%>&nbsp;<%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_38a")%>
                                                        <%else%>
                                                            <%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_38")%>&nbsp;<%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_38b")%>
                                                        <%end if%>
                                                        <br><br>
                                                        <%	rs19.movenext
                                                    loop
                                                    set rs19=nothing
                                                    %>
                                                    </td>
                                                </tr>
                                            <%end if
                                            set rsG=nothing
                                            rs11.MoveNext
                                        loop
                                        set rs11=nothing
                                        call closedb()%>
                                        </table>
                                    <% END IF
                                end if
                                'GGG Add-on end%>
                                    
                            </td>
                        </tr>
                        <%' ------------------------------------------------------
                        'Start SDBA - Notify Drop-Shipping
                        ' ------------------------------------------------------
                        if scShipNotifySeparate="1" then
                            call opendb()
                            tmp_showmsg=0
                            query="SELECT products.pcProd_IsDropShipped FROM products INNER JOIN productsOrdered ON (products.idproduct=productsOrdered.idproduct AND products.pcProd_IsDropShipped=1) WHERE ProductsOrdered.idOrder=" & pIdOrder & ";"
                            set rs=connTemp.execute(query)
                            if err.number<>0 then
                                call LogErrorToDatabase()
                                set rs=nothing
                                call closedb()
                                response.redirect "techErr.asp?err="&pcStrCustRefID
                            end if
                            if not rs.eof then
                                tmp_showmsg=1
                            end if
                            set rs=nothing
                            if tmp_showmsg=1 then%>
                            <tr> 
                                <td colspan="2" class="pcSpacer">&nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <div class="pcTextMessage"><%response.write ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%></div>
                                </td>
                            </tr>
                        <%end if
                        call closedb()
                    end if
                    ' ------------------------------------------------------
                    'End SDBA - Notify Drop-Shipping
                    ' ------------------------------------------------------%>
          </table>
				</form>

					<%
					'Gift Certificates Recipient Information
					' Show only if the order has been processed
					if int(pOrderStatus)>2 AND int(pOrderStatus)<>5 AND int(pOrderStatus)<>6 then

							call opendb()
							query="SELECT pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg FROM Orders WHERE idOrder="& pidorder &" AND pcOrd_GcReEmail<>'';"
							SET rsGCObj=server.CreateObject("ADODB.RecordSet")
							SET rsGCObj=connTemp.execute(query)
							if not rsGCObj.eof then
								Gc_ReName=rsGCObj("pcOrd_GcReName")
								Gc_ReEmail=rsGCObj("pcOrd_GcReEmail")
								Gc_ReMsg=rsGCObj("pcOrd_GcReMsg")
								%>
								<form method="post" name="form3" action="CustViewPastD.asp?action=resend" class="pcForms">
								<table class="pcShowContent">
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2"><%response.write dictLanguage.Item(Session("language")&"_GCRecipient_1")%></th>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr> 
									<td><b><%response.write dictLanguage.Item(Session("language")&"_NotifyRe_3")%></b></td>
									<td><input type="text" name="GC_RecName" size="30" value="<%=Gc_ReName%>"></b></td>
								</tr>
								<tr> 
									<td><b><%response.write dictLanguage.Item(Session("language")&"_NotifyRe_4")%></b></td>
									<td><input type="text" name="GC_RecEmail" size="30" value="<%=Gc_ReEmail%>"></b></td>
								</tr>
								<tr> 
									<td valign="top"><b><%response.write dictLanguage.Item(Session("language")&"_NotifyRe_5")%></b></td>
									<td valign="top"><textarea name="GC_RecMsg" cols="60" rows="5" wrap="VIRTUAL"><%=GC_ReMsg%></textarea></td>
								</tr>
								<tr> 
									<td valign="top">&nbsp;</td>
									<td valign="top">
										<input type="hidden" name="idOrder" value="<%=int(pIdOrder)+scpre%>">
										<input type="submit" name="submitReSendGCRec" value="<%response.write dictLanguage.Item(Session("language")&"_GCRecipient_2")%>">
									</td>
								</tr>
								</table>
								</form>
							<%
							end if
							set rsGCObj=nothing									
										
					end if	
					%>
        </td>
    </tr>
	<%if Session("CustomerGuest")="1" then%>
	<tr>
		<td>
			<div id="PwdArea">
				<form id="PwdForm" name="PwdForm">
				<table class="pcShowContent">
				<tr>
					<th colspan="4" class="pcSectionTitle"><%=dictLanguage.Item(Session("language")&"_opc_common_2")%></th>
				</tr>
				<tr>
					<td colspan="4"><%=dictLanguage.Item(Session("language")&"_opc_common_3")%></td>
				</tr>
				<tr>
					<td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_6")%></td>
					<td width="30%"><input type="password" name="newPass1" id="newPass1" size="20"></td>
					<td width="20%" nowrap><%=dictLanguage.Item(Session("language")&"_opc_38")%></td>
					<td width="30%"><input type="password" name="newPass2" id="newPass2" size="20"></td>
				</tr>
				<tr>
					<td colspan="4" style="padding-top: 10px;"></td>
				</tr>
				<tr>
					<td colspan="4" style="padding-top: 10px;"><input type="button" name="PwdSubmit" id="PwdSubmit" value="<%=dictLanguage.Item(Session("language")&"_opc_common_4")%>" class="submit2"></td>
				</tr>
				</table>
				</form>
				<div id="PwdLoader" style="display:none"></div>
		</div>
		</td>
	</tr>
	<%end if%>
	<tr>
		<td>
			<% '// Account Consolidation %>
            <% call openDB() %>
			<!--#include file="opc_inc_CustConsolidate.asp"-->
		</td>
	</tr>
	<tr>
		<td>
<script>
$(document).ready(function()
{
	jQuery.validator.setDefaults({
		success: function(element) {
			$(element).parent("td").children("input, textarea").addClass("success")
		}
	});
	
	//*Ajax Global Settings
	$("#GlobalAjaxErrorDialog").ajaxError(function(event, request, settings){
		$(this).dialog('open');
		$("#PwdLoader").hide();
		$("#ConLoader").hide();
	});

	
	//*Dialogs
	$("#GlobalAjaxErrorDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			resizable: false,
			width: 450,
			height: 230,
			modal: true,
			buttons: {
				' OK ': function() {
						$(this).dialog('close');
					}
			}
	});
	
	<%if Session("CustomerGuest")="1" then
	Session("SFStrRedirectUrl")="CustPref.asp"%>
	//*Validate Password Form
	$("#PwdForm").validate({
		rules: {
			newPass1: 
			{
				required: true
			},
			newPass2:
			{
				required: true,
				equalTo: "#newPass1"
			}
		},
		messages: {
			newPass1: {
				required: "<%=dictLanguage.Item(Session("language")&"_opc_js_4")%>",
				minlength: "<%=dictLanguage.Item(Session("language")&"_opc_js_5")%>"
			},
			newPass2: {
				required: "<%=dictLanguage.Item(Session("language")&"_opc_js_47")%>",
				minlength: "<%=dictLanguage.Item(Session("language")&"_opc_js_5")%>",
				equalTo: "<%=dictLanguage.Item(Session("language")&"_opc_js_48")%>"
			}
		}
	})
	
	$('#PwdSubmit').click(function(){
		if ($('#PwdForm').validate().form())
		{
			$("#PwdLoader").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_5")%>');
			$("#PwdLoader").show();	
			$.ajax({
				type: "POST",
				url: "opc_createacc.asp",
				data: $('#PwdForm').formSerialize() + "&action=create",
				timeout: 5000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						$("#PwdArea").html("");
						$("#PwdArea").hide();
						$("#PwdLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_6")%>');
						var callbackPwd=function (){setTimeout(function(){$("#PwdLoader").hide();},1000);}
						$("#PwdLoader").effect('pulsate',{},500,callbackPwd);
					}
					else
					{
					if ((data=="OK") || (data=="REG") || (data=="OKA") || (data=="REGA"))
					{

						if ((data=="OK") || (data=="OKA"))
						{
							$("#PwdLoader").html('<img src="images/pcv4_st_icon_success_small.png" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_7")%>');
						}
						else
						{
							$("#PwdLoader").html('<img src="images/pcv4_st_icon_success_small.png" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_8")%>');
						}
						var callbackPwd=function (){}
						$("#PwdLoader").effect('pulsate',{},500,callbackPwd);
						$("#PwdArea").html("");
						$("#PwdArea").hide();
						if (data=="OKA")
						{
							$("#ConArea").show();
						}
						else
						{
							location="login.asp?lmode=2";
						}
					}
					else
					{
						$("#PwdLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> '+data);
						var callbackPwd=function (){setTimeout(function(){$("#PwdLoader").hide();},1000);}
						$("#PwdLoader").effect('pulsate',{},500,callbackPwd);
					}
					}
				}
	 		});
			return(false);
		}
		return(false);
	});
	<%end if%>


});
</script>
		</td>
	</tr>
	<%IF (Session("CustomerGuest")="0") AND (Session("idCustomer")>"0") THEN%>
    <tr> 
        <td align="right"><a href="custViewPast.asp"><img src="<%=rslayout("back")%>"></a></td>
    </tr>
	<%END IF%>
    <tr>
        <td class="pcSpacer"></td>
    </tr>
</table>
</div>
<% call closedb() %>
<!--#include file="footer.asp"-->