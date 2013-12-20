<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/shipFromsettings.asp"--> 
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/CashbackConstants.asp"-->
<!--#include file="header.asp"-->
<% 
err.number=0
dim query, conntemp, rs, rstemp, pIdOrder, pOID, pnValid, pOrderStatus, pcv_noDoubleTracking
call openDb()
%>
<!--#include file="prv_getsettings.asp"-->
<script>
	function openbrowser(url) {
			self.name = "productPageWin";
			popUpWin = window.open(url,'rating','toolbar=0,location=0,directories=0,status=0,top=0,scrollbars=yes,resizable=1,width=705,height=535');
			if (navigator.appName == 'Netscape') {
			popUpWin.focus();
		}
	}
</script>
<%pcv_RWActive=pcv_Active
pnValid=0
If len(session("idOrder"))>0 Then
	pOID=session("idOrder")
	session("idOrderConfirm")=pOID
Else
	pOID=session("idOrderConfirm")
	pcv_noDoubleTracking=1
End If
if pOID = "" then
	pOID = 0
	pnValid=1
end if
session("idOrder")=""
session("GWOrderId")="" '// PayPal Standard
if NOT validNum(pOID) then
	pnValid=1
end if

' Create "View Previous Order" link
if scSSL="1" AND scIntSSLPage="1" then
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/CustviewPastD.asp"),"//","/")
else
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/CustviewPastD.asp"),"//","/")
end if
tempURL=replace(tempURL,"https:/","https://")
tempURL=replace(tempURL,"http:/","http://")
tempURL=tempURL & "?idOrder=" & (int(pOID)+scpre)
	
' clear cart data
if len(session("pcSFIdDbSession"))>0 then
	on error resume next
	query="DELETE FROM pcCustomerSessions WHERE idDbSession="&session("pcSFIdDbSession")&" AND randomKey="&session("pcSFRandomKey")&" AND idCustomer="&session("idCustomer")&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	err.number=0
	err.clear
end if

dim pcCartArray2(100,45)
Session("pcCartSession")=pcCartArray2
Session("pcCartIndex")=Cint(0)
session("pcSFIdDbSession")=""
session("pcSFRandomKey")=""
session("iOrderTotal")=""
session("pcSFCartRewards")=Cint(0)
session("pcSFUseRewards")=Cint(0)
session("IDRefer")=""
session("specialdiscount")=""
Session("ContinueRef")=""
session("TF1")=""
session("DF1")=""
session("shippingFullName")=""
session("shippingCompany")=""
session("shippingAddress")=""
session("shippingAddress2")=""
session("shippingStateCode")=""
session("shippingState")=""
session("shippingZip")=""
session("shippingPhone")=""
session("shippingCity")=""
session("shippingCountryCode")=""
session("DCODE")=""
session("idOrderSaved")=""
session("ExpressCheckoutPayment")=""
session("GWOrderDone")=""
session("redirectPage")=""
Session("SFStrRedirectUrl")=""
session("idGWSubmit")=""
session("idGWSubmit2")=""
session("idGWSubmit3")=""
session("Gateway")=""
session("SaveOrder")=""
Session("pcPromoSession")=""
Session("pcPromoIndex")=""
session("OPCstep")=""
session("Entered-" & session("GWPaymentId"))=""
Session("CurrentPanel")=""
session("NeedToUpdatePay")=""
session("SF_DiscountTotal")=""
session("SF_RewardPointTotal")=""
IDSC=0
tmpGUID=getUserInput(Request.Cookies("SavedCartGUID"),0)
IF tmpGUID<>"" THEN
	query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "';"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		IDSC=rsQ("SavedCartID")
		HasSavedCart=1
	end if
	set rsQ=nothing
	if HasSavedCart=1 then
		query="DELETE FROM pcSavedCartArray WHERE SavedCartID=" & IDSC & ";"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
		query="DELETE FROM pcSavedCarts WHERE SavedCartID=" & IDSC & ";"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	end if
	Response.Cookies("SavedCartGUID")=""
END IF
%>
<div id="pcMain">
<div id="GlobalAjaxErrorDialog" title="Communication Error" style="display:none">
	<div class="pcErrorMessage">
		Can not connect to server to exchange information. Please contact store owner or try again later
	</div>
</div>
<table class="pcMainTable">	
	<% if pnValid=1 then 'Order number not valid %>
	<tr> 
		<td><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_viewPostings_a")%></div>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		</td>
	</tr>
	<% else 'Order number is valid
	
		' Get order status and customer ID
		query = "SELECT orders.idCustomer,orders.orderStatus,orders.pcOrd_OrderKey FROM orders WHERE orders.idOrder =" & pOID
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
				call closeDb() %>
				<tr> 
					<td><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_viewPostings_a")%></div>
					<p>&nbsp;</p>
					<p>&nbsp;</p>
					<p>&nbsp;</p>
					<p>&nbsp;</p>
					<p>&nbsp;</p>
					</td>
				</tr>
		<% end if 
		
		'Get the customer ID if the session is empty
		if int(Session("idcustomer")) = 0 then
			Session("idcustomer") = rs("idCustomer")
		end if
			
		'Get the order status
		pOrderStatus=rs("orderStatus")
		pcOrderKey=rs("pcOrd_OrderKey")
		set rs=nothing
		
		'If order has already been processed, show corresponding message
		if pOrderStatus="3" then %>
		<tr> 
			<td>
			<h1><%=dictLanguage.Item(Session("language")&"_updOrdStats_2a")%><%=(int(pOID)+scpre)%></h1>
			<p><%=dictLanguage.Item(Session("language")&"_updOrdStats_2b")%> <a href="<%=tempURL%>"><%=dictLanguage.Item(Session("language")&"_orderComplete_1")%></a></p>
			</td>
		</tr>
		<% else %>
		<tr> 
			<td>
			<h1><%=dictLanguage.Item(Session("language")&"_updOrdStats_2a")%><%=(int(pOID)+scpre)%></h1>
			<p><%=dictLanguage.Item(Session("language")&"_updOrdStats_2")%> <a href="<%=tempURL%>"><%=dictLanguage.Item(Session("language")&"_orderComplete_1")%></a></p>
			</td>
		</tr>
	<% 
		end if 'End if order has already been processed
	%>
	<tr>
		<td><p><%=dictLanguage.Item(Session("language")&"_orderComplete_2")%></p></td>
	</tr>
	<tr>
		<td><hr></td>
	</tr>
    <% 
	pcv_intNewAcct=getUserInput(Request("newAcct"),0)
	if pcv_intNewAcct="1" then 'New Account Created %>
	<tr> 
		<td><div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_common_7")%></div></td>
	</tr>
    <% end if %>
	<%if pcOrderKey<>"" then%>
		<tr> 
			<td>
			<div id="OrderCodeArea" class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_common_1")%>&nbsp;<%=pcOrderKey%></div>
			<p><%=dictLanguage.Item(Session("language")&"_opc_common_9")%></p>
			</td>
		</tr>
	<%end if%>
	<tr>
		<td valign="top">
			<% 
			' Start Order Details section
			pIdOrder=pOID
			
			query="SELECT customers.email,customers.fax,orders.pcOrd_ShippingEmail,orders.pcOrd_ShippingFax,orders.pcOrd_ShowShipAddr,orders.orderDate, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.customerType, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.comments, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.pcOrd_shippingPhone, orders.shippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.idOrder, orders.rmaCredit, orders.ordPackageNum, orders.ord_DeliveryDate, orders.ord_OrderName, orders.ord_VAT,orders.pcOrd_CatDiscounts, orders.paymentDetails, orders.gwAuthCode, orders.gwTransId, orders.paymentCode, orders.pcOrd_GWTotal FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pIdOrder&"));"
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
			
			dim pidCustomer, porderDate, pfirstname, plastname,pcustomerCompany, pphone, paddress, pzip, pstate, pcity, pcountryCode, pcomments, pshippingAddress, pshippingState, pshippingCity, pshippingCountryCode, pshippingZip, paddress2, pshippingFullName, pshippingCompany, pshippingAddress2, pshippingPhone, pcustomerType
			
			
			pEmail=rs("email")
			pFax=rs("fax")
			pshippingEmail=rs("pcOrd_ShippingEmail")
			pshippingFax=rs("pcOrd_ShippingFax")
			pcShowShipAddr=rs("pcOrd_ShowShipAddr")
			if IsNull(pcShowShipAddr) OR (pcShowShipAddr="") then
				pcShowShipAddr=0
			end if
			pidCustomer=Session("idcustomer")
			porderDate=rs("orderDate")
			porderDate=showdateFrmt(porderDate)
			pfirstname=rs("name")
			plastName=rs("lastName")
			pCustomerName=pfirstname& " " & plastName
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
			'GGG Add-on start
			pGWTotal=rs("pcOrd_GWTotal")
			if pGWTotal<>"" then
			else
			pGWTotal="0"
			end if
			'GGG Add-on end
			
			'// Check if the Customer is European Union 
			Dim pcv_IsEUMemberState
			pcv_IsEUMemberState = pcf_IsEUMemberState(pshippingCountryCode)
			
			query="SELECT ProductsOrdered.idProduct, ProductsOrdered.pcSubscription_ID, ProductsOrdered.quantity, ProductsOrdered.unitPrice, ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts  "
			'BTO ADDON-S
			If scBTO=1 then
				query=query&", ProductsOrdered.idconfigSession"
			End If
			'BTO ADDON-E
			query=query&", pcPO_GWOpt, pcPO_GWPrice, products.description, products.sku, orders.total, orders.paymentDetails, orders.taxamount, orders.shipmentDetails, orders.discountDetails, orders.pcOrd_GCDetails, orders.orderstatus,orders.processDate, orders.shipdate, orders.shipvia, orders.trackingNum, orders.returnDate, orders.returnReason, orders.iRewardPoints, orders.iRewardValue, orders.iRewardPointsCustAccrued, orders.taxdetails, orders.dps, ProductsOrdered.xfdetails, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, pcPrdOrd_BundledDisc, pcPO_GWNote FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idCustomer=" &Session("idcustomer")& " AND orders.idOrder=" &pIdOrder
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
			%>
			<table class="pcShowContent">
				<tr>
					<td colspan="5">
					<table class="pcShowContent">
						<tr>
							<td>
							
							<%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_14")%>
							<%response.write porderDate%> - 
							<%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_9")&": "&(int(pIdOrder)+scpre)%>
							
							</td>
							<td align="right">
							<div class="pcSmallText"><a href="custOrdInvoice.asp?id=<%=pIdOrder%>" target="_blank"><img src="images/document.gif" width="16" border="0" align="middle" vspace="5" hspace="2"></a> <a href="custOrdInvoice.asp?id=<%=pIdOrder%>" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_33")%></a></div>
							</td>
						</tr>
					</table>
					</td>
				</tr>
				
				<% if (pord_DeliveryDate<>"") then
					if scDateFrmt="DD/MM/YY" then
						pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 4)
					else
						pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 3)
					end if
					pord_DeliveryDate = showdateFrmt(pord_DeliveryDate)
					%>
					<tr>
						<td colspan="5" valign="top">
						
						<%=dictLanguage.Item(Session("language")&"_CustviewOrd_39")%><%=pord_DeliveryDate%> <% If pord_DeliveryTime <> "00:00" Then %><%=", " & pord_DeliveryTime%><% End If %>
							
						</td>
					</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>
				<%end if%>
				
				<tr>
					<th colspan="3"><%response.write dictLanguage.Item(Session("language")&"_orderverify_23")%></th>
					<th colspan="2"><%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
							response.write dictLanguage.Item(Session("language")&"_orderverify_24")
						end if%>
					</th>
				</tr>
				
				<tr>
					<td colspan="2" valign="top"><b>
						<% response.write replace(dictLanguage.Item(Session("language")&"_orderverify_7"),"''","'")%>
					</b></td>
					<td valign="top"> 
						<% response.write pFirstName&" "&plastname %>
					</td>
					<td colspan="2" valign="top">
						<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then%>
							<% response.write pshippingFullName %>
						<% end if%>
					</td>
				</tr>
				
				<tr>
					<td colspan="2" valign="top"><b> 
						<% response.write dictLanguage.Item(Session("language")&"_orderverify_8")%>
					</b></td>
					<td valign="top"><%=pcustomerCompany%></td>
					<td colspan="2" valign="top">
						<% if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
							if pshippingCompany<>"" then
								response.write pshippingCompany
							else
								if (pshippingFullName = "" or pshippingFullName = pCustomerName) and pCustomerCompany <> "" then
									response.write pCustomerCompany
								end if
							end if
						end if %>
					</td>
				</tr>
				
				<%if pEmail<>pshippingEmail AND pshippingEmail<>"" then%>
				<tr>
					<td colspan="2" valign="top"><b> 
						<%=dictLanguage.Item(Session("language")&"_opc_5")%>
					</b></td>
					<td valign="top"><%=pEmail%></td>
					<td colspan="2" valign="top">
						<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
							response.write pshippingEmail
						end if %>
					</td>
				</tr>
				<%end if%>
				
				<tr>
					<td colspan="2" valign="top"><b> 
						<% response.write dictLanguage.Item(Session("language")&"_orderverify_9")%>
					</b></td>
					<td valign="top"><%=pPhone%></td>
					<td colspan="2" valign="top">
						<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
							response.write pshippingPhone
						end if %>
					</td>
				</tr>
				
				<%if pFax<>"" OR pshippingFax<>"" then%>
				<tr>
					<td colspan="2" valign="top"><b> 
						<%=dictLanguage.Item(Session("language")&"_opc_18")%>
					</b></td>
					<td valign="top"><%=pFax%></td>
					<td colspan="2" valign="top">
						<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
							response.write pshippingFax
						end if %>
					</td>
				</tr>
				<%end if%>
				
				<tr>
					<td colspan="2" valign="top"><b> 
						<% response.write dictLanguage.Item(Session("language")&"_orderverify_10")%>
					</b></td>
					<td valign="top"><%=paddress%></td>
					<td colspan="2" valign="top">
						<% if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
							if pshippingAddress="" then
								response.write "Same as Billing Address"
							else
								response.write pshippingAddress
							end if
						else
							if pcShowShipAddr="0" AND session("gHideAddress")<>"1" then
								response.write "Same as Billing Address"
							end if
						end if %>
					 </td>
				</tr>
				
				<tr>
					<td colspan="2" valign="top">&nbsp;</td>
					<td valign="top"><%=paddress2%></td>
					<td colspan="2" valign="top">
						<% if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
							if pshippingAddress2<>"" then
								response.write pshippingAddress2
							end if
						end if %>
					</td>
				</tr>
				
				<tr>
					<td colspan="2" valign="top">&nbsp;</td>
					<td valign="top"><%=pCity&", "&pState&" "&pzip%></td>
					<td colspan="2" valign="top">
						<% if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
							if pshippingAddress<>"" then
								response.write pShippingCity&", "&pshippingState
								If pshippingState="" then
									response.write pshippingStateCode
								End If
								response.write " "&pshippingZip
							end if
						end if %>
					</td>
				</tr>
				
				<tr>
					<td colspan="2" valign="top">&nbsp;</td>
					<td valign="top">  <%=pCountryCode%> </td>
					<td colspan="2" valign="top">
						<%if pcShowShipAddr="1" AND session("gHideAddress")<>"1" then
							response.write pshippingCountryCode
							strFedExCountryCode=pshippingCountryCode
						else
							strFedExCountryCode=pCountryCode
						end if %>
					</td>
				</tr>
			
				<% ' Start of payment details
				payment = split(pcpaymentDetails,"||")
				PaymentType=trim(payment(0))
				
				'Get payment nickname
				query="SELECT paymentDesc, paymentNickName FROM paytypes WHERE paymentDesc = '" & replace(PaymentType,"'","''") & "';"
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
					 case "Moneris2"								     					  
						  Dim varIDEBIT_ISSCONF, varIDEBIT_ISSNAME,varRespName,varResponseCode
						   varTransName="Sequence Number"
						   varAuthName="Approval Code"
						   varRespName="Response / ISO Code"
						   varTransID=pcgwTransId
						   varAuthCode=pcgwAuthCode
						
						   query = "Select pcPay_MOrder_responseCode, pcPay_MOrder_ISOcode, pcPay_MOrder_IDEBIT_ISSCONF, pcPay_MOrder_IDEBIT_ISSNAME from pcPay_OrdersMoneris Where pcPay_MOrder_TransId='"& pcgwTransId &"';" 
						   set rstemp=server.CreateObject("ADODB.RecordSet")
						   set rstemp=conntemp.execute(query)											  
							if err.number<>0 then
								call LogErrorToDatabase()
								set rstemp=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
			
						if not rs.eof then
						   varResponseCode = RStemp("pcPay_MOrder_responseCode")
						   varISO_Code = RStemp("pcPay_MOrder_ISOcode")
						   varIDEBIT_ISSCONF = rstemp("pcPay_MOrder_IDEBIT_ISSCONF")
						   varIDEBIT_ISSNAME = rstemp("pcPay_MOrder_IDEBIT_ISSNAME")								 							
						end if
						set rstemp=nothing
						  
					end select
			
				end if
				
				'End get authorization and transaction IDs
				
				on error resume next
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
						<td colspan="5"><hr></td>
					</tr>
					<tr>
						<td colspan="5">
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
						<%if varResponseCode <> ""  or varISO_Code <> "" Then%>
						<BR><%=varRespName%>&nbsp;<%=varResponseCode%>/<%=varISO_Code%>
						<%end if %>
						<% if varIDEBIT_ISSCONF <> ""  and varIDEBIT_ISSNAME <> "" then %>
						<br><%=dictLanguage.Item(Session("language")&"_CustviewOrd_48")%>
						<BR><%=dictLanguage.Item(Session("language")&"_CustviewOrd_49")%>&nbsp;<%=varIDEBIT_ISSNAME%>
						<BR><%=dictLanguage.Item(Session("language")&"_CustviewOrd_50")%>&nbsp;<%=varIDEBIT_ISSCONF%>						
						<% end if%>
						
						<br><br>
						</td>
					</tr>
				<% end if
					' End of payment details
				%>
			
				<% ' Start of order comments
					if len(pcomments)>3 then %>
					<tr>
						<td colspan="5"> <b>
							<% response.write dictLanguage.Item(Session("language")&"_orderverify_11")%>
							</b> <%=pcomments%> <br>
							<br>
						</td>
					</tr>
				<% end if 
					' End of order comments
				%>
				
				<tr>
					<th width="41"><% response.write dictLanguage.Item(Session("language")&"_orderverify_25")%></th>
					<th width="59"><% response.write dictLanguage.Item(Session("language")&"_orderverify_26")%></th>
					<th><% response.write dictLanguage.Item(Session("language")&"_orderverify_27")%></th>
					<th width="15%"><% response.write dictLanguage.Item(Session("language")&"_orderverify_32")%></th>
					<th width="92" align="right"><% response.write dictLanguage.Item(Session("language")&"_orderverify_28")%></th>
				</tr>
				
				<% 
				dim pidProduct, pquantity, punitPrice, pxfdetails, pidconfigSession, pdescription, pSku, pcDPs, ptotal, ppaymentDetails,ptaxamount,pshipmentDetails, pdiscountDetails
				dim pprocessDate, pshipdate, pshipvia, ptrackingNum, preturnDate, preturnReason, piRewardPoints, piRewardValue, piRewardPointsCustAccrued,ptaxdetails, pOpPrices, rsObjOptions, pRowPrice, count, rsConfigObj,stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory,i, s,OptPrice,xfdetails, xfarray, q
				Dim GCDetails
				Dim pcv_strSelectedOptions, pcv_strOptionsPriceArray, pcv_strOptionsArray
				Dim pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice
				Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions
				
				do while not rsOrdObj.eof
					pidProduct=rsOrdObj("idProduct")
					pcSubscription_ID=rsOrdObj("pcSubscription_ID")
					pquantity=rsOrdObj("quantity")
					punitPrice=rsOrdObj("unitPrice")
					QDiscounts=rsOrdObj("QDiscounts")
					ItemsDiscounts=rsOrdObj("ItemsDiscounts")
					'BTO ADDON-S
					if scBTO=1 then
						pidconfigSession=rsOrdObj("idconfigSession")
						if pidconfigSession="" then
							pidconfigSession="0"
						end if
					End If
					'BTO ADDON-E
					'GGG Add-on start  
					pGWOpt=rsOrdObj("pcPO_GWOpt")
					if pGWOpt<>"" then
					else
						pGWOpt="0"
					end if
					pGWPrice=rsOrdObj("pcPO_GWPrice")
					if pGWPrice<>"" then
					else
						pGWPrice="0"
					end if
					'GGG Add-on end
					
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
					pxfdetails=rsOrdObj("xfdetails")
					'// Product Options Arrays
					pcv_strSelectedOptions = rsOrdObj("pcPrdOrd_SelectedOptions") ' Column 11
					pcv_strOptionsPriceArray = rsOrdObj("pcPrdOrd_OptionsPriceArray") ' Column 25
					pcv_strOptionsArray = rsOrdObj("pcPrdOrd_OptionsArray") ' Column 4
					pcPrdOrd_BundledDisc=rsOrdObj("pcPrdOrd_BundledDisc")
					pGWText=rsOrdObj("pcPO_GWNote")
					pprocessDate=ShowDateFrmt(pprocessDate)
					
					pIdConfigSession=trim(pidconfigSession)
					
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
                        if pIdConfigSession<>"0" then
                            pRowPrice1=Cdbl(pquantity * ( punitPrice1 )) - TotalUnit
                            punitPrice1=Round(pRowPrice1/pquantity,2)
                        else
                            pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
                        end if
                    else
                        punitPrice1=punitPrice
                        if pIdConfigSession<>"0" then
                            pRowPrice1=Cdbl(pquantity * ( punitPrice1 ))
                        else
                            pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
                            punitPrice1=Round(pRowPrice1/pquantity,2)
                        end if
                    end if
                    %>
					
					<tr valign="top"> 
						<td width="41"> <%response.write pquantity%></td>
						<td width="59"> <%response.write pSku%></td>
						<td>
							<%response.write pdescription%>
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
                            <%
							'SB S
							If pcSubscription_ID>0 Then
								query="SELECT SB_Terms FROM SB_Orders WHERE idOrder=" & pIdOrder & ";"
								Set rsSB=Server.CreateObject("ADODB.Recordset")
								Set rsSB=connTemp.execute(query)
								If NOT rsSB.eof Then
									pcv_strTerms = rsSB("SB_Terms")
									if len(pcv_strTerms)>0 then
										response.Write(pcv_strTerms)
									end if
								End If
							End If
							'SB E
							%>     
						</td>
						<td width="15%" align="right"><p align="right"><% response.write(scCurSign&money(punitPrice1)) %></p></td>
						<td width="92" align="right"><p align="right"><% response.write(scCurSign&money(pRowPrice1)) %></p></td>
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
								<td valign="top" width="41">&nbsp;</td>
								<td colspan="4" valign="top"> 
									<div class="pcShowBTOconfiguration">
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr> 
											<td colspan="3">  
												<%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_1")%>
												</td>
										</tr>
										<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
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
											
											query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i) & " and specProduct=" & pidProduct 
												set rs=server.CreateObject("ADODB.RecordSet")
												set rs=conntemp.execute(query)
												if err.number<>0 then
													'//Logs error to the database
													call LogErrorToDatabase()
													'//clear any objects
													set rs=nothing
													'//close any connections
													call closedb()
													'//redirect to error page
													response.redirect "techErr.asp?err="&pcStrCustRefID
												end if
												
												btDisplayQF=rs("displayQF")
												set rs=nothing
												
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
												<td width="85%" valign="top">
													<p>
													<%=strCategoryDesc%>:
													<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%>(<%=ArrQuantity(i)%>)&nbsp;<%end if%>
													<%=strDescription%>
													</p> 
												</td>
												<td width="15%" align="right" valign="top"> 
													<%
														if pnoprices<2 then
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
															end if
														end if
													%>
												<%if pnoprices<2 then%>
													<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then%>
														<%=scCurSign & money((ArrValue(i)+UPrice)*pQuantity)%>
													<%else
														if tmpDefault=1 then%>
															<%=dictLanguage.Item(Session("language")&"_defaultnotice_1")%>
														<%end if
													end if
												end if%>
												</td>
											</tr>
											<% set rsConfigObj=nothing
										next %>
									</table>
									</div>
									</td>
							</tr>
						<% end if 
					End If 
					'BTO ADDON-E
					%>
					
					<% 'Start options
					
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: SHOW PRODUCT OPTIONS
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
						pcv_strSelectedOptions = ""
					end if
					
					
					if len(pcv_strSelectedOptions)>0 then 
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
								<td width="54%"><p><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></p></td>
								
								<td align="right" width="46%">									
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
											<td align="left" width="62%">
												<%=scCurSign&money(tempPrice)%>
											</td>
											<td align="right" width="38%">
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
					<%
					'BTO ADDON-S
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
							if (ItemsDiscounts<>"") and (CDbl(ItemsDiscounts)<>"0") then %>
								<tr valign="top"> 
									<td width="41">&nbsp; </td>
									<td width="59">&nbsp; </td>
									<td><%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_2")%></td>
									<td width="15%">&nbsp;</td>
									<td><div align="right"><%=scCurSign&money(-1*ItemsDiscounts)%></div></td>
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
											<td valign="top" width="41">&nbsp;</td>
											<td colspan="4" valign="top">
												<div class="pcShowBTOconfiguration">
												<table width="100%" border="0" cellspacing="0" cellpadding="0">
												
													<tr> 
														<td colspan="3"> 
															<%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_5")%>
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
															<td width="85%%" valign="top"> 
																<p><%=strCategoryDesc%>:&nbsp;<%=strDescription%></p>
															</td>
															<td width="15%" nowrap align="right" valign="top"><%if pnoprices<2 then%><%if ArrCValue(i)>0 then%><%=scCurSign & money(ArrCValue(i))%><%end if%><%end if%>
															</td>
														</tr>
														<% set rsConfigObj=nothing
													next %>
												</table>
												</div>
											</td>
											
										</tr>
							
										<% pRowPrice=pRowPrice+Cdbl(Charges)
									end if 'Have Charges
								end if 
						end if 'BTO 
							'BTO Additional Charges
                    end if
							
							QDiscounts=trim(QDiscounts)
							if QDiscounts="" then
								QDiscounts=0
							end if
						
							if (QDiscounts<>"") and (CDbl(QDiscounts)<>"0") then
								%>
								<tr valign="top"> 
									<td width="41">&nbsp; </td>
									<td width="59">&nbsp; </td>
									<td><%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_3")%></td>
									<td width="15%">&nbsp;</td>
									<td align="right"><%=scCurSign&money(-1*QDiscounts)%></td>
								</tr>
								<% pRowPrice=pRowPrice-Cdbl(QDiscounts)
							end if%>
							
					<% if pExtRowPrice<>pRowPrice then %>
								<tr valign="top"> 
									<td colspan="4" align="right"><b><%response.write bto_dictLanguage.Item(Session("language")&"_CustviewPastD_4")%></b></td>
									<td><div align="right"><% if pRowPrice> 0 then response.write(scCurSign&money(pRowPrice)) end if %></div></td>
								</tr>
                    <% end if
					
					'show xtra options
					'-----------------
					xfdetails=pxfdetails
					If len(xfdetails)>3 then
						xfarray=split(xfdetails,"|")
						for q=lbound(xfarray) to ubound(xfarray) %>
							<tr> 
								<td valign="top" width="41">&nbsp;</td>
								<td valign="top" width="59">&nbsp;</td>
								<td valign="top" colspan="2"><%=xfarray(q)%></td>
								<td valign="top" width="92">&nbsp;</td>
							</tr>
						<% next
					End If 
					'----------------- 
					%>
					<%
					'SB S
					 if pSubscriptionID > 0  then%>				
						<tr>				
							<td></td>
							<td colspan="4" >
								
								<table cellpadding="1" cellspacing="0" width="100%" bgcolor="#EFEFEF">	
									<% if pSubInstall= 1 Then %>
										<tr>
											<td width="5%">&nbsp;</td>
											<td colspan="2">
												<p align="left">
												<%	
												if pRowPrice > 0 and pcv_intBillingCycles > 0  then								
													pExtRowPrice = ((pRowPrice/pcv_intBillingCycles))
													response.write (pcv_intBillingCycles & " payments of " & scCurSign & money(pExtRowPrice) & " = " & scCurSign & money(pRowPrice) )
												end if 
												%>
												</p>
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td colspan="2">
												<p align="left">
													<%	
													if pRowPrice > 0 and pcv_intBillingCycles <> 0  then								
														pExtRowPrice = ((pRowPrice/pcv_intBillingCycles))
														response.write ("First payment of " & scCurSign & money(pExtRowPrice) &" due on: " )
													End if
													if pcv_intIsTrial = "1" then
														response.write  pFirstPaymentDate
													else
														response.write  pSubStartDate
													end if 
													%>
												</p>
											</td>
										</tr>
									<% end if %>
									<% if pcv_intIsTrial = 1 or pSubInstall= 1 then %>
										<tr>
											<td colspan="2" >
												<p align="right">
												<%
												if pcv_intIsTrial = "1" then 
												  pcv_curTrialAmount = (pcv_curTrialAmount * cdbl(pquantity) )				 
												  if pcv_intIsTrial = "1"  then 
													 response.write pcv_intIsTrialDesc
												  end if
												Else
													 response.write "Payment:"  
												End if 
												%>	
												</p>				
											</td>
											<td width="7%" >
											<p align="right">
												<%	
												if pcv_intIsTrial = "1" then 
												   if pcv_curTrialAmount = 0 and pShowFreeTrial="1" then 
														response.write  pFreeTrialDesc
													else
														response.write  (scCurSign &  money(pcv_curTrialAmount)) 
													end if
												else
												   response.write  (scCurSign &  money(pExtRowPrice))  
												End if 
												%>
											</p>
											</td>
										</tr>
									<% end if %>
								</table>
                            </td>
                        </tr>
                        <%
                        ' if there's a trial or startup fee set the line total to the trial price
                        if pcv_intIsTrial = "1" Then
                            pRowPrice = pcv_curTrialAmount
                        else
                            pRowPrice = pExtRowPrice
                        end if 
					end if 									   
					'SB E
					
					'GGG Add-on start
					if pGWOpt<>"0" then
						mySQL="select pcGW_OptName,pcGW_optPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
						set rsG=connTemp.execute(mySQL)
						if not rsG.eof then
						%>
						<tr> 
							<td valign="top" width="41">&nbsp;</td>
							<td valign="top" colspan="4">
							<div><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_4")%>&nbsp;<%=rsG("pcGW_OptName")%><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_5")%>&nbsp;<%=scCurSign & money(pGWPrice)%></div>
							<%if pGWText<>"" then%>
								<div><%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_6")%><br><%=pGWText%><div>
							<%end if%>
							</td>
						</tr>
						<%
						end if
					end if
					'GGG Add-on end
					If pcPrdOrd_BundledDisc>0 then %>
                        <tr valign="top"> 
                            <td width="41">&nbsp; </td>
                            <td width="59">&nbsp; </td>
                            <td><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_54")%></td>
                            <td width="15%">&nbsp;</td>
                            <td align="right">-<%=scCurSign&money(pcPrdOrd_BundledDisc)%></td>
                        </tr>
					<% end if
					%>
					<% count=count+1
					If pshippingAddress="" then
						'grab shipping address from shipping...
						pshippingAddress=pAddress
						pshippingAddress2=pAddress2
						pshippingCity=pCity
						pshippingState=pState
						pshippingZip=pZip
						pshippingCountryCode=pCountryCode
					End if %>
                    <tr> 
					<td valign="top" colspan=5><hr></td>
					</tr>
                    <%
					rsOrdObj.movenext  
				loop
				%>
				
				<tr> 
					<td valign="top" width="41">&nbsp;</td>
					<td valign="top" width="59">&nbsp;</td>
					<td valign="top" colspan="2">&nbsp;</td>
					<td valign="top" width="92">&nbsp;</td>
				</tr>
				
				<% 'start of processing charges
				dim payment, PaymentType,PayCharge
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
				%>
				
				<% if PayCharge>0 then %>
					<tr>
						<td>&nbsp;</td>
						<td colspan="3" valign="top"><div>
						<b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_14b")%></b>
						</div></td>
						<td valign="top" align="right"><%=scCurSign&money(PayCharge)%></td>
					</tr>
				<% end if %>
				<% 'end of processing charges
				
				'start of discount details
				if pcv_CatDiscounts>"0" then %>
							<td>&nbsp;</td>
							<td colspan="3" valign="top"><div><b><%response.write dictLanguage.Item(Session("language")&"_catdisc_2")%></b></div></td>
					<td valign="top" align="right">-<%=scCurSign&money(pcv_CatDiscounts)%></td>
					</tr>
				<% end if %>
				<% if instr(pdiscountDetails,",") then
					DiscountDetailsArry=split(pdiscountDetails,",")
					intArryCnt=ubound(DiscountDetailsArry)
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
							<td colspan="3" valign="top"><div><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_15")%></b>&nbsp;<% response.write discountType%></div></td>
						<td valign="top" align="right">
						<% if discount <> 0 then %>
						-<%=scCurSign&money(discount)%>
						<% end if %>
						</td>
						</tr>
					<% end if
				Next
				'end if discount details
				
				
				'start of gift certificates
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
							<td colspan="3" valign="top"><div><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_15A")%></b>&nbsp;<%=GCInfo(1)%> (<%=GCInfo(0)%>)</div></td>
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
				
				'start of rewards used
				if piRewardPoints>0 then %>
					<tr> 
							<td>&nbsp;</td>
							<td colspan="3" valign="top"><div><%response.write "<b>"&piRewardPoints&"&nbsp;"&RewardsLabel&" </b>used on this purchase, for a discount of: "%>
						</div></td>
						<td valign="top" align="right">-<% response.write scCurSign& money(piRewardValue) %></td>
					</tr>
				<% end if
				'end if rewards
				
				'start of rewards earned
				if piRewardPointsCustAccrued>0 then %>
					<tr>
						<td>&nbsp;</td>
						<td colspan="4" valign="top"><b>
							<% response.write dictRewardsLanguage.Item(Session("rewards_language")&"_orderverify")%>
						</b><%=dictLanguage.Item(Session("language")&"_orderverify_30")%>
						<% response.Write(piRewardPointsCustAccrued) %>
						</td>
					</tr>
				<% end if
				'end if rewards
				
				'GGG Add-on start
				if pGWTotal>0 then
				%>
					<tr>
						<td>&nbsp;</td>
						<td colspan="3" valign="top">
							<div><b><%=dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_4")%></b></div>
						</td>
						<td valign="top" align="right"><%=scCurSign&money(pGWTotal)%></td>
					</tr>
				<%
				end if
				'GGG Add-on end
				
				'start of shipping
				dim shipping, varShip, Shipper, Service, Postage, serviceHandlingFee
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
				
				if varShip<>"0" then %>
					<tr> 
						<td>&nbsp;</td>
						<td colspan="3" valign="top"><div><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_13")%></b>&nbsp;<%=Service%></div>
						</td>
						<td valign="top" align="right"><%=scCurSign&money(Postage)%></td>
					</tr>
				<% End If %>
				
				<% if serviceHandlingFee>0 then %>
					<tr> 
						<td valign="top">&nbsp;</td>
						<td colspan="3" valign="top" align="right"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_14")%></b>
						</td>
						<td valign="top" align="right"><%=scCurSign&money(serviceHandlingFee)%></td>
				</tr>
				<% end if
				'end of shipping
				
				'start of taxes
				'If the store is using VAT and VAT is> 0, don't show any taxes here, but show VAT after the total
				if pord_VAT>0 then
				else
					if isNull(ptaxDetails) or trim(ptaxDetails)="" then %>
						<tr> 
							<td valign="top">&nbsp;</td>
							<td colspan="3" valign="top"><div align="right"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_12")%></b></div>
							</td>
							<td valign="top" align="right"><% response.write scCurSign& money(ptaxAmount)%></td>
						</tr>
					<% else %>
						<% dim taxArray, taxDesc
						taxArray=split(ptaxDetails,",")
						for i=0 to (ubound(taxArray)-1)
							taxDesc=split(taxArray(i),"|")
							if taxDesc(0)<>"" then %>
							<tr> 
								<td valign="top">&nbsp;</td>
								<td colspan="3" valign="top"><div align="right"><b><%=taxDesc(0)%></b></div>
								</td>
								<td valign="top" align="right"><% response.write scCurSign& money(taxDesc(1))%></td>
							</tr>
							<% end if
						next %>
					<% end if 
				end if
				'end if taxes %>
				
				<tr> 
					<td valign="top">&nbsp;</td>
					<td colspan="3" valign="top" align="right"><b><%response.write dictLanguage.Item(Session("language")&"_CustviewPastD_12")%></b></td>
					<td valign="top" align="right"><% response.write scCurSign& money(ptotal) %></td>
				</tr>
				
				<% ' If the store is using VAT and VAT> 0, show it here
				if pord_VAT>0 then %>
					<tr> 
						<td colspan="5" align="right" class="pcSmallText">							
							<% if pcv_IsEUMemberState = 1 then %>
								<% response.write dictLanguage.Item(Session("language")&"_orderverify_35") & scCurSign & money(pord_VAT)%>
							<% else %>
								<% response.write dictLanguage.Item(Session("language")&"_orderverify_42") & scCurSign & money(pord_VAT)%>
							<% end if %>
						</td>
					</tr>
				<% end if %>
			</table>
		</td>
	</tr>
	
	<%' ------------------------------------------------------
	'Start SDBA - Notify Drop-Shipping
	' ------------------------------------------------------
	if scShipNotifySeparate="1" then
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
			<td>
				<hr>
			</td>
		</tr>
		<tr>
			<td>
				<div class="pcTextMessage"><%response.write ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%></div>
			</td>
		</tr>
		<%end if
	end if
	' ------------------------------------------------------
	'End SDBA - Notify Drop-Shipping
	' ------------------------------------------------------%>
	
	<tr>
		<td>
		<p><a href="<%=tempURL%>"><%=dictLanguage.Item(Session("language")&"_orderComplete_1")%></a></p>
		</td>
	</tr>
<%
	end if 'End if order number is valid
%>
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
	<%if Session("CustomerGuest")="2" then
		Session("JustPurchased")="1"
	end if%>
	<tr>
		<td>
			<% '// Account Consolidation %>
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
	
	<%if Session("CustomerGuest")="1" then%>
	//*Validate Password Form
	$("#PwdForm").validate({
		rules: {
			newPass1: 
			{
				required: true,
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
						location='orderComplete.asp?newAcct=1';
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

	<%if pcOrderKey<>"" then%>
	//var callbackOCA=function(){};
	//$("#OrderCodeArea").effect('pulsate',{},800,callbackOCA);
	<%end if%>
});
</script>
		</td>
	</tr>
	<% ' Continue shopping button %>
	<tr>
		<td align="right">
		<% 
			csimage= RSlayout("continueshop")
			contURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
			contURL=replace(contURL,"https:/","https://")
			contURL=replace(contURL,"http:/","http://")	
		%>
		<a href="<%=contURL%>"><img src="<%=csimage%>" border=0></a>
		</td>
	</tr>
	<% ' End Continue shopping button %>
</table>
</div>
<% 
'// Tell the system that this is the order completed page
Dim pcv_intOrderComplete
pcv_intOrderComplete=1

'// Tell the system that there has been a page refresh
if pcv_noDoubleTracking=1 then
	pcv_intOrderComplete=0
end if

%>
<!--#include file="orderCompleteTracking.asp"-->
<!--#include file="inc-Cashback.asp"-->
<%
session("ExpressCheckoutPayment")=""
session("gHideAddress")=""
on error resume next
If Session("Payer")&""<>"" Then
	Session.Abandon
	
	' clear cart data
	redim pcCartArray2(100,45)
	Session("pcCartSession")=pcCartArray2
	Session("pcCartIndex")=Cint(0)
End If

'// Google Analytics (GA)
'// Inform GA script that this is the Order Completed page
'// If GA is not used, this code does not need to be removed as it is harmless
Dim pcGAorderComplete 
pcGAorderComplete=1
%>
<% call closedb() %>
<!--#include file="footer.asp"-->