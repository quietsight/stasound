<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'file name = gwReturn.asp
'file purpose = gather the response from the appropriate real-time gateway.
%>
<% response.Buffer=true %>
<%
'// Avoid seeing external attacks
Dim SPath,strSiteURL
SPath=Request.ServerVariables("PATH_INFO")
SPath=mid(SPath,1,InStrRev(SPath,"/")-1)
If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
	strSiteURL="http://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
Else
  strSiteURL="https://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
End if
Dim pOrderDate,pProcessDate,pIdCustomer,ppStatus
%>

<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="paypalOrdIPN.asp"-->
<!--#include file="gw2CheckoutResponse.asp"-->
<% 'SB S %>
<!--#include file="inc_sb.asp"-->
<% 'SB E %>
<%'Allow Guest Account
AllowGuestAccess=1
%>
<!--#include file="CustLIv.asp"-->
<% dim rt_gateway,pcv_Action
'-----------------------------------
' Check for post from gateway data
'-----------------------------------
'//WorldPay
if request("status")="Y" then
	session("GWAuthCode")=getUserInput(request("rawAuthCode"),0)
	session("GWTransId")=getUserInput(request("transId"),0)
	if session("GWOrderId")="" then
		session("GWOrderId")=getUserInput(request("idOrder"),0)
	end if
	session("GWSessionID")=Session.SessionID
	rt_gateway="WorldPay"
end if



'//ChronoPay
Transaction_Id = getUserInput(request("transaction_id"),0)
if Transaction_Id <> "" then
	session("GWAuthCode")=""
	session("GWTransId")=getUserInput(request("transaction_id"),0)
	if session("GWOrderId")="" then
		session("GWOrderId")=getUserInput(request("idOrder"),0)
	end if
	session("GWSessionID")=Session.SessionID
	rt_gateway="ChronoPay"
end if
'//PSI
If request("OrdNo")<>""  then
	session("GWAuthCode")=getUserInput(request("Code"),0)
	session("GWTransId")=getUserInput(request("RefNo"),0)
	session("GWTransType")=getUserInput(request("ChargeType"),0)
	session("GWSessionID")=Session.SessionID
	if session("GWOrderId")="" then
		session("GWOrderId")=getUserInput(request("OrdNo"),0)
	end if
	rt_gateway="PSI"
End If

'//PayFlow Link
pcv_PNREF = getUserInput(request("PNREF"),0)
if pcv_PNREF<>"" then
	gwTransID=getUserInput(request("order_number"),0)
	session("GWAuthCode")=getUserInput(request("AUTHCODE"),0)
	session("GWTransId")=pcv_PNREF
	if session("GWOrderId")="" then
		session("GWOrderId")=getUserInput(request("INVOICE"),0)
	end if
	session("GWSessionID")=Session.SessionID
	session("GWTransType")=getUserInput(request("TYPE"),0)
	rt_gateway="PFLink"
end if

'//LinkPoint
If request("approval_code")<>"" OR request("fail_reason")<>""  then
	pcIntOrderFailed=0
	if request("fail_reason")<>"" then
		pcIntOrderFailed=1
		pcStrFailedReason=getUserInput(request("fail_reason"),0)
	end if

	tempTransId=getUserInput(request("oid"),0)
	if instr(tempTransId,",") then
		arryTransId=split(tempTransId,",")
		pTransID=arryTransId(0)
	else
		pTransID=tempTransId
	end if

	tempIdOrder=getUserInput(request("userid"),0)
	if instr(tempIdOrder,",") then
		arryIdOrder=split(tempIdOrder,",")
		pIdOrder=arryIdOrder(0)
	else
		pIdOrder=tempIdOrder
	end if

	if session("GWOrderId")="" then
		session("GWOrderId")=pIdOrder
	end if

	if pcIntOrderFailed=0 then
		session("GWAuthCode")=getUserInput(request("approval_code"),0)
		session("GWTransId")=pTransID
		session("GWSessionID")=Session.SessionID
		rt_gateway="LinkPoint"
	else
		'// Return to payment page
		session("pcStrFailedReason")=pcStrFailedReason
		session("pcIntOrderFailed")=pcIntOrderFailed
		response.Redirect "gwlp.asp"
	end if
end if

'//iTransact
If request("xid")<>"" then
	session("GWAuthCode")=getUserInput(request("authcode"),0)
	session("GWTransId")=getUserInput(request("xid"),0)
	session("GWTransType")=getUserInput(request("xxauth"),0)
	if session("GWOrderId")="" then
		session("GWOrderId")=getUserInput(request("idOrder"),0)
	end if
	session("GWSessionID")=Session.SessionID
	rt_gateway="iTransact"
end if

'//InternetSecure
pcv_Response_IdOrder=getUserInput(request("xxxVar1"),0)
if pcv_Response_IdOrder<>"" then
	pcv_Response_ApprovalCode=getUserInput(request("ApprovalCode"),0)
	pcv_Response_ReceiptNumber=getUserInput(request("receiptnumber"),0)

	'if pcv_Response_StatusCode="F" OR pcv_Response_StatusCode="0" OR pcv_Response_StatusCode="D" then
	'	Msg=pcv_Response_AuthMessage
	'end if
	session("GWAuthCode")=pcv_Response_ApprovalCode
	session("GWTransId")=pcv_Response_ReceiptNumber
	if session("GWOrderId")="" then
		session("GWOrderId")=pcv_Response_IdOrder
	end if
	session("GWSessionID")=Session.SessionID
	rt_gateway="InternetSecure"
end if

'//FastTransact
pcv_Response_StatusCode=getUserInput(request("Ecom_Ezic_Response_StatusCode"),0)
if pcv_Response_StatusCode<>"" then
	pcv_Response_StatusCode=getUserInput(request("Ecom_Ezic_Response_StatusCode"),0)
	pcv_Response_AuthCode=getUserInput(request("Ecom_Ezic_Response_AuthCode"),0)
	pcv_Response_AuthMessage=getUserInput(request("Ecom_Ezic_Response_AuthMessage"),0)
	pcv_Response_TransactionID=getUserInput(request("Ecom_Ezic_Response_TransactionID"),0)
	pcv_Response_Card_AVSCode=getUserInput(request("Ecom_Ezic_Response_Card_AVSCode"),0)
	pcv_Response_Card_VerificationCode=getUserInput(request("Ecom_Ezic_Response_Card_VerificationCode"),0)
	pcv_Response_IssueDate=getUserInput(request("Ecom_Ezic_Response_IssueDate"),0)
	if pcv_Response_StatusCode="F" OR pcv_Response_StatusCode="0" OR pcv_Response_StatusCode="D" then
		Msg=pcv_Response_AuthMessage
		'response.redirect "fasttransact_giveup.asp?msg="&pcv_Response_AuthMessage
	end if
	session("GWAuthCode")=pcv_Response_AuthCode
	session("GWTransId")=pcv_Response_TransactionID
	session("GWSessionID")=Session.SessionID

	rt_gateway="FastTransact"
end if

'//2Checkout
CartOrderID = getUserInput(request("cart_order_id"),0)
if CartOrderID<>"" then
	gwTransID=getUserInput(request("order_number"),0)

	session("GWAuthCode")=""
	session("GWTransId")=gwTransID
	if session("GWOrderId")="" then
		session("GWOrderId")=CartOrderID
	end if
	session("GWSessionID")=Session.SessionID
	rt_gateway="twoCheckout"
end if
'---------------------------------------
' End Check for post from gateway data
'---------------------------------------

call opendb()

gwAuthCode=""
gwTransId=""
gwAVSCode=""
gwCVV2Code=""
paymentCode=""

'Order Status=2 by default (pending)
pOrderStatus=2
'Payment Status=0 by default (pending)
pPaymentStatus=0
'flag to allow override of payment and order status
pOverride=0

'//////////////////////////////////////////////////////////
'// Allow Override of Default Status for Free Mode Only
'//////////////////////////////////////////////////////////
' pFMOrderStatus=2 by default for (pending)
' pFMOrderStatus=3 for (processed)
pFMOrderStatus=2


if request("purchasemode")="FREE" then
	rt_gateway="FREE"
end if

if request("s")="true" then
	Session("PPEI")="1"
	rt_gateway=getUserInput(request("gw"),0)
end if

pIdOrder=session("GWOrderId")
gwAuthCode=session("GWAuthCode")
gwTransID=session("GWTransId")
gwAVSCode=session("AVSCode")
gwCVV2Code=session("CVV2Code")
paymentCode= rt_gateway

dim query, conntemp, rs
dim gwAuthCode,gwTransID,paymentCode,pIdOrder,gwAVSCode,gwCVV2Code
'get variables
select case rt_gateway
	case "OFFLINE_CUSTOM"
		paymentCode="OFFLINE_PAYMENT"
	case "FREE"
		pOrderStatus=pFMOrderStatus
		pIdOrder=getUserInput(request("idOrder"),0)
	case "ParaData"
		if ucase(session("GWTransType"))="AUTH" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "TransFirst"
		If session("GWTransType")="1" then
			pPaymentStatus=1 'Auth only
		else
			pPaymentStatus=2
		end if
	case "PaymentExpress"
		if ucase(session("GWTransType"))="AUTH" then
			pPaymentStatus=0
		end if
	case "PayPalExp"
		if session("GWTransType")="2" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		'SB S
		if NOT request("s")="true" then
			session("GWOrderDone")="YES"
		end if
		'SB E
	case "PayPalWP"
		if session("GWTransType")="2" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		'SB S
		if NOT request("s")="true" then
			session("GWOrderDone")="YES"
		end if
		'SB E
	case "PayPal"
		if session("GWTransType")="P" then 'Pending PayPal order
			pOverride=1
			pPaymentStatus=1
			pOrderStatus=2
		else
			pPaymentStatus=2
		end if
	case "PayPalAdvanced"
		if session("GWTransType")="P" then 'Pending PayPal order
			pOverride=1
			pPaymentStatus=1
			pOrderStatus=2
		else
			pPaymentStatus=2
		end if
	case "AIM"
		if ucase(session("x_type"))="AUTH_ONLY" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		if ucase(session("x_method"))="ECHECK" then
			if session("x_eCheckPending")=1 then
				pOverride=1
				pPaymentStatus=0
			end if
		end if
		paymentCode="Authorize"
		gwAVSCode = session("x_avs_code")
		gwCVV2Code = session("x_cardcode_response") & "||" & session("x_cav_response")
		session("GWOrderTotal")=""
		'SB S
		if NOT request("s")="true" then
			session("GWOrderDone")="YES"
		end if
		'SB E
	case "EIG"
		if ucase(session("GWTransType"))="AUTH_ONLY" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		'SB S
		if NOT request("s")="true" then
			session("GWOrderDone")="YES"
		end if
		'SB E
	case "Moneris2"
		if ucase(session("GWTransType"))="PURCHASE" or ucase(session("GWTransType"))="IDEBIT_PURCHASE" then
			pPaymentStatus=2
		else
			pPaymentStatus=1
		end if
	case "LinkPoint", "LinkPointApi"
		query="SELECT transType FROM LinkPoint WHERE id=1;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		transType=rs("transType")
		set rs=nothing
		if ucase(transType)="SALE" then
			pPaymentStatus=2
		else
			pPaymentStatus=1
		end if
	case "PSI"
		if session("GWTransType")="1" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "PSIGate"
		query="SELECT [mode] FROM PSIGate WHERE id=1;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		ChargeType=rs("mode")
		set rs=nothing
		if ChargeType="1" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "twoCheckout"
		pPaymentStatus=2
	case "iTransact"
		if ucase(session("GWTransType"))="YES" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "WorldPay"
		pPaymentStatus=2
	case "ChronoPay"
		pPaymentStatus=2
	case "PFPRO"
		if ucase(session("GWTransType"))="A" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		session("GWOrderDone")="YES"
	case "PFLink"
		if ucase(session("GWTransType"))="A" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "FastTransact"
		query="SELECT tran_type FROM fasttransact Where id=1;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		transType=rs("tran_type")
		set rs=nothing
		if ucase(transType)="PREAUTH" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="FAST"
	case "ECHO"
		if ucase(session("GWTransType"))="AS" OR ucase(session("GWTransType"))="AV" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "Concord"
		if ucase(session("GWTransType"))="AUTHORIZE" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="CONCORD"
	case "KLIX"
		pPaymentStatus=2
		pOrderStatus=3
	case "TCLink"
		if ucase(session("GWTransType"))="PREAUTH" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "TCLinkCheck"
		if session("GWTransType")=1 then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "SagePay"
		if lcase(session("GWTransType"))="authenticate" OR lcase(session("GWTransType"))="deferred" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "Netbill"
		if ucase(session("GWTransType"))="A" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "NetbillCheck"
		gwAuthCode=session("AuthorizationNumber")
		gwTransID=session("TransactionID")
		pIdOrder=session("NetbillOrdno")
	case "BluePay"
		if ucase(session("GWTransType"))="AUTH" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "SecPay"
		if ucase(session("GWTransType"))="AUTH" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "VM"
		if ucase(session("GWTransType"))="CCAUTHONLY" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="VM"
	case "InternetSecure"
		paymentCode="IntSecure"
	case "eWay"
		pPaymentStatus=2
	case "CBN"
		if session("GWTransType") = 1 then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "Paymentech"
		if session("TransType")="A" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="PAYMENTECH"
		session("GWOrderDone")="YES"
	case "CYS"
		if session("GWTransType")="0" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="CyberSource"
	case "HSBC"
		if lcase(session("GWTransType"))="preauth" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "UEP"
		if request("c")=1 then
			query="SELECT pcPay_Uep_CheckPending FROM pcPay_USAePay WHERE pcPay_Uep_Id=1"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			pcPay_Uep_CheckPending=rs("pcPay_Uep_CheckPending")
			set rs=nothing
			if pcPay_Uep_CheckPending="1" then
				pOverride=1
				pPaymentStatus=0
				pOrderStatus=2
			end if
		else
			if session("GWTransType")="0" then
				pPaymentStatus=1
			else
				pPaymentStatus=2
			end if
		end if
		paymentCode="USAePay"
		session("GWOrderDone")="YES"
	case "FAC"
		if lcase(session("GWTransType"))="yes" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="FastCharge"
	case "ACH"
		if lcase(session("GWTransType"))<>"sale" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="ACHDirect"
	case "NETOne"
		if session("GWTransType")="02" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
	case "GestPay"
		pPaymentStatus=2
	case "EPN"
		paymentCode="eProcessingNetwork"
	case "TD"
		if lcase(session("GWTransType"))="paid" then
			pPaymentStatus=2
		else
			pPaymentStatus=1
		end if
		paymentCode="TripleDeal"
	case "SkipJack"
		pPaymentStatus=2
	case "eMoney"
		pPaymentStatus=1
	case "Ogone"
		if lcase(session("GWTransType"))="sal" then
			pPaymentStatus=2
		else
			pPaymentStatus=1
		end if
		paymentCode="Ogone"
		case "Beanstream"
		if lcase(Beanstreamsession("GWTransType"))="P" then
			pPaymentStatus=2
		else
			pPaymentStatus=1
		end if
		paymentCode="BeanStream"
	case "eMerchant"
		if ucase(session("GWTransType"))="A" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="eMerchant"
	case "TotalWeb"
		 pPaymentStatus=1
		 paymentCode="TotalWeb Solutions"
	case "PayJunction"
		if ucase(session("GWTransType"))="AUTHORIZATION" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="Pay Junction"
	case "TGS"
		if ucase(session("GWTransType"))="0" then
			pPaymentStatus=1
		else
			pPaymentStatus=2
		end if
		paymentCode="Transaction Gateway Systems"
end select

'Start SDBA
if (session("pcSFIdPayment")<>"") AND (session("pcSFIdPayment")<>"0") then
	if session("pcSFIdPayment")=999999 then
		query="SELECT pcPayTypes_processOrder,pcPayTypes_setPayStatus FROM PayTypes WHERE gwcode=999999 OR gwcode=46"
	else
		query="SELECT pcPayTypes_processOrder,pcPayTypes_setPayStatus FROM PayTypes WHERE idPayment=" & session("pcSFIdPayment")
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rs.eof then
		tmp_OrderStatus=rs("pcPayTypes_processOrder")
		if IsNull(tmp_OrderStatus) or tmp_OrderStatus="" then
		 tmp_OrderStatus=0
		end if
		if tmp_OrderStatus=1 AND pOverride=0 then
			pOrderStatus=3 'override Order Status to set payment as processed
		end if
		tmp_PaymentStatus=rs("pcPayTypes_setPayStatus")
		if IsNull(tmp_PaymentStatus) or tmp_PaymentStatus="" then
			tmp_PaymentStatus=0
		end if
		'if flag to override is still set to '0' - proceed to override Payment Status
		if pOverride=0 then
			Select Case tmp_PaymentStatus
				Case 0: pPaymentStatus=0
				Case 1: pPaymentStatus=1
				Case 2: pPaymentStatus=2
			End Select
		end if
	end if
end if
'End SDBA

'SB S
'// Save temp value for subscriptions
dim pSubIdOrder
pSubIdOrder=pIdOrder
'SB E

' extract real idorder (without prefix)
pIdOrder=(int(pIdOrder)-scpre)

'// Check back button from OrderComplete.asp
If len(session("idOrderConfirm"))>0 Then
	If session("idOrderConfirm") = pIdOrder Then
		session("idOrderConfirm") = ""
		response.redirect "msg.asp?message=1"
	End If
End If

If Not validNum(pIdOrder) then
	call closedb()
	response.redirect "msg.asp?message=10"
End If

if trim(pIdOrder)="" then
	call closedb()
	response.redirect "techErr.asp?error=21"&Server.Urlencode(dictLanguage.Item(Session("language")&"_updOrdStats_1") )
end if

' get order details

'SB S
query="SELECT orders.pcOrd_OrderKey, orders.idcustomer, orders.gwAuthCode, orders.address, orders.City, orders.StateCode, orders.State, orders.zip, orders.CountryCode, orders.shippingAddress, orders.shippingCity, orders.shippingStateCode, orders.shippingState, orders.shippingZip,  orders.shippingCountryCode, orders.pcOrd_shippingPhone, orders.ShipmentDetails, orders.PaymentDetails, orders.discountDetails, orders.taxAmount, orders.total, orders.comments, orders.ShippingFullName, orders.address2, orders.ShippingCompany, orders.ShippingAddress2, orders.taxDetails, orders.orderstatus, orders.iRewardPoints, orders.iRewardValue, orders.iRewardRefId, orders.iRewardPointsRef, orders.iRewardPointsCustAccrued, orders.ordPackageNum, customers.phone, orders.ord_DeliveryDate, orders.ord_VAT, orders.pcOrd_DiscountsUsed, orders.pcOrd_Payer, orders.pcOrd_CatDiscounts, orders.pcOrd_SubTax, orders.pcOrd_SubTrialTax, orders.pcOrd_SubShipping, orders.pcOrd_SubTrialShipping FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" &pIdOrder
'SB E

set rsObjOrder=server.CreateObject("ADODB.RecordSet")
set rsObjOrder=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rsObjOrder=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pcOrderKey=rsObjOrder("pcOrd_OrderKey")
pidcustomer=rsObjOrder("idcustomer")
pgwAuthCode=rsObjOrder("gwAuthCode")
paddress=rsObjOrder("address")
pCity=rsObjOrder("City")
pStateCode=rsObjOrder("StateCode")
pState=rsObjOrder("State")
if isNULL(pStateCode) OR pStateCode="" then
	pStateCode=pState
end if
pzip=rsObjOrder("zip")
pCountryCode=rsObjOrder("CountryCode")
pshippingAddress=rsObjOrder("shippingAddress")
pshippingCity=rsObjOrder("shippingCity")
pshippingStateCode=rsObjOrder("shippingStateCode")
pshippingState=rsObjOrder("shippingState")
if isNULL(pshippingStateCode) OR pshippingStateCode="" then
	pshippingStateCode=pshippingState
end if
pshippingZip=rsObjOrder("shippingZip")
pshippingCountryCode=rsObjOrder("shippingCountryCode")
pshippingPhone=rsObjOrder("pcOrd_shippingPhone")
pShipmentDetails=rsObjOrder("ShipmentDetails")
pPaymentDetails=rsObjOrder("PaymentDetails")
pdiscountDetails=rsObjOrder("discountDetails")
ptaxAmount=rsObjOrder("taxAmount")
ptotal=rsObjOrder("total")
pcomments=rsObjOrder("comments")
pShippingFullName=rsObjOrder("ShippingFullName")
paddress2=rsObjOrder("address2")
pShippingCompany=rsObjOrder("ShippingCompany")
pShippingAddress2=rsObjOrder("ShippingAddress2")
ptaxDetails=rsObjOrder("taxDetails")
pCurOrderStatus=rsObjOrder("orderStatus")
piRewardPoints=rsObjOrder("iRewardPoints")
piRewardValue=rsObjOrder("iRewardValue")
piRewardRefId=rsObjOrder("iRewardRefId")
piRewardPointsRef=rsObjOrder("iRewardPointsRef")
piRewardPointsCustAccrued=rsObjOrder("iRewardPointsCustAccrued")
pOrdPackageNum=rsObjOrder("ordPackageNum")
pPhone=rsObjOrder("phone")
pord_DeliveryDate=rsObjOrder("ord_DeliveryDate")
pord_DeliveryDate=showDateFrmt(pord_DeliveryDate)
pord_VAT=rsObjOrder("ord_VAT")
strPcOrd_DiscountsUsed=rsObjOrder("pcOrd_DiscountsUsed")
pcOrd_Payer=rsObjOrder("pcOrd_Payer")
pcOrd_CatDiscounts=rsObjOrder("pcOrd_CatDiscounts")

'SB S
pcSubTax=rsObjOrder("pcOrd_SubTax")
pcSubTrialTax=rsObjOrder("pcOrd_SubTrialTax")
pcSubShipping=rsObjOrder("pcOrd_SubShipping")
pcSubTrialShipping=rsObjOrder("pcOrd_SubTrialShipping")
'SB E

set rsObjOrder=nothing
call closeDB() ' Closing DB connection because it is reopened and closed by pcGateWayData.asp, included below

ppStatus=0 ' This flag will prevent duplicate e-mails when PayPal IPN contacts the store with updates
if paymentCode="PayPal" then
	if (pCurOrderStatus="3") OR (pgwAuthCode<>"" AND isNULL(pgwAuthCode)=False) then
		ppStatus=1
	end if
end if

'SB S
if session("pcIsSubscription") then
	pcGWError = request("GWError")
	pcGatewayDataIdOrder=pSubIdOrder
	%>
	<!--#include file="pcGateWayData.asp"-->
	<%
	Set objSB = NEW pcARBClass

	'// Get Customer CC or Check Values
	objSB.IdOrder =  pcGatewayDataIdOrder
	objSB.IDCustomer =  session("IDCustomer")
	objSB.PaymentCode = PaymentCode

	'// Payment Details
	if Request("c") <> "true" Then

		'// Credit Card OR Token
		If len(Session("PayPalExpressToken"))>0 Then
			objSB.PayInfoType="PP"
			objSB.PayInfoToken = Session("PayPalExpressToken")
			objSB.PayInfoPayerID = Session("PayerID")
		Else
			objSB.PayInfoType="CC"
			objSB.PayInfoExpMonth= session("reqExpMonth")
			If len(session("reqExpYear"))=2 Then
				objSB.PayInfoExpYear = "20" & session("reqExpYear")
			Else
				objSB.PayInfoExpYear = session("reqExpYear")
			End If
			objSB.PayInfoCardNumber = left(session("reqCardNumber"),16)
			objSB.PayInfoAccountNumber = right(PayInfoCardNumber,4)
			objSB.PayInfoCardType = session("reqCardType")
			objSB.PayInfoCVVNumber = session("reqCVV")
		End If

	Else

		'// Check
		objSB.PayInfoType="CHECK"
		objSB.PayInfoDriversLicenseNum =left(trim(session("x_drivers_license_num")),20)
		objSB.PayInfoDriversLicenseState= left(trim(session("x_drivers_license_state")),2)
		objSB.PayInfoDriversLicenseDOB= session("x_drivers_license_dob")
		objSB.PayInfoBankAcctName= left(trim(session("x_bank_acct_name")),22)
		objSB.PayInfoBankABACode= left(trim(session("x_bank_aba_code")),9)
		objSB.PayInfoBankAcctNum =left(trim(session("x_bank_acct_num")),17)
		objSB.PayInfoAccountNumber=right(objSB.PayInfoBankAcctNum,4)
		objSB.PayInfoBankAcctType= session("x_bank_acct_type")
		objSB.PayInfoInfoBankName =session("x_bank_name")
		objSB.PayInfoBankAcctOrgType = session("x_customer_organization_type")
		objSB.PayInfoCustomerTaxId =left(trim(session("x_customer_tax_id")),9)
		objSB.PayInfoExpYear=0
		objSB.PayInfoExpMonth=0

	End if


	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// BILLING ADDRESS
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	objSB.BillingFirstName = pcBillingFirstName
	objSB.BillingLastName = pcBillingLastName
	objSB.BillingCompany = pcBillingCompany
	objSB.BillingAddress = pcBillingAddress
	objSB.BillingAddress2 = pcBillingAddress2
	objSB.BillingCity = pcBillingCity
	objSB.BillingPostalCode = pcBillingPostalCode
	objSB.BillingStateCode = pcBillingStateCode
	objSB.BillingProvince = pcBillingProvince
	objSB.BillingCountryCode = pcBillingCountryCode
	objSB.BillingPhone = pcBillingPhone
	objSB.CustomerEmail = pcCustomerEmail


	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// SHIPPING ADDRESS
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	objSB.ShippingFirstName = pcShippingFirstName
	objSB.ShippingLastName = pcShippingLastName
	objSB.ShippingCompany = pcShippingCompany
	objSB.ShippingAddress = pcShippingAddress
	objSB.ShippingAddress2 = pcShippingAddress2
	objSB.ShippingCity = pcShippingCity
	objSB.ShippingPostalCode = pcShippingPostalCode
	objSB.ShippingStateCode = pcShippingStateCode
	objSB.ShippingProvince = pcShippingProvince
	objSB.ShippingCountryCode = pcShippingCountryCode
	objSB.ShippingPhone = pcShippingPhone
	objSB.ShippingEmail = pcShippingEmail

	call opendb()
	query = "SELECT password FROM customers WHERE idCustomer = " & session("idCustomer")
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs = conntemp.execute(query)
	If NOT rs.eof Then
		objSB.CustomerPassword = enDeCrypt(rs("password"), scCrypPass)
	End If
	set rs = nothing
	call closedb()

	objSB.CustomerAccount = session("idCustomer")


	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// CART
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	'// Tax Name
	dim taxArray, taxDesc
	taxArray=split(ptaxDetails,",")
	for i=0 to (ubound(taxArray)-1)
		taxDesc=split(taxArray(i),"|")
		if taxDesc(0)<>"" then
			pTaxName = taxDesc(0)
		end If
	next
	objSB.CartTaxName = pTaxName


	'// Tax
	objSB.CartRegularTax = pcSubTax


	'// Tax Trial
	objSB.CartTrialTax = pcSubTrialTax


	'// Shipping
	objSB.CartRegularShipping = pcSubShipping


	'// Shipping Trial
	objSB.CartTrialShipping = pcSubTrialShipping


	'// Is Shippable
	If instr(pShipmentDetails,",")>0 Then
		objSB.CartIsShippable = true
	Else
		objSB.CartIsShippable = false
	End If


	'// Ship Name
	shipping=split(pShipmentDetails,",")
	if ubound(shipping)>1 then
		if NOT isNumeric(trim(shipping(2))) then
			pShipName=""
		else
			pShipName = shipping(1)
			pShippingFees = trim(shipping(2))
			if ubound(shipping)=>3 then
				serviceHandlingFee=trim(shipping(3))
				if NOT isNumeric(serviceHandlingFee) then
					serviceHandlingFee=0
				end if
			else
				serviceHandlingFee=0
			end if
			pShippingFees = (cdbl(pShippingFees) + cdbl(serviceHandlingFee))
		end if
	else
		pShipName=""
	end if
	objSB.CartShipName = pShipName

	'// Agree to Terms
	If session("pcAgreeAll") = True Then
		objSB.CartAgreedToTerms = true
	Else
		objSB.CartAgreedToTerms = false
	End If

	'// Language Code
	If scSBLanguageCode<>"" Then
		objSB.CartLanguageCode = scSBLanguageCode
	Else
		objSB.CartLanguageCode = "en-EN"
	End If


	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// SEND REQUEST
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	call opendb()

	query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
	set rsAPI=connTemp.execute(query)
	if not rsAPI.eof then
		Setting_APIUser=rsAPI("Setting_APIUser")
		Setting_APIPassword=enDeCrypt(rsAPI("Setting_APIPassword"), scCrypPass)
		Setting_APIKey=enDeCrypt(rsAPI("Setting_APIKey"), scCrypPass)
	end if
	set rsAPI=nothing

	Dim result
	result = objSB.SubscriptionRequest(Setting_APIUser, Setting_APIPassword, Setting_APIKey)

	If len(SB_ErrMsg)>0 Then
		Dim pcv_strPaymentPage
		Select Case getUserInput(request("gw"),0)
			Case "PayPalWP" :  pcv_strPaymentPage="gwPayPal.asp"
			Case "AIM" :  pcv_strPaymentPage="gwAuthorizeAIM.asp"
			Case "PayPalExp" :  pcv_strPaymentPage="gwPayPal.asp"
			Case "EIG" :  pcv_strPaymentPage="gwEIGateway.asp"
		End Select
		call closeDb()
		response.Redirect(pcv_strPaymentPage & "?message=" & server.URLEncode(SB_ErrMsg))
		response.End()
	Else
		session("GWOrderDone")="YES"
	End If

	call closeDb()
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// CLEAN UP
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	'// CLEAR SESSIONS
	session("reqCardNumber")=""
	session("reqExpMonth")=""
	session("reqExpYear")=""
	session("reqCardType")=""
	session("reqCVV")=""
	session("x_bank_acct_name") = ""
	session("x_bank_aba_code") = ""
	session("x_bank_acct_num") =  ""
	session("x_bank_acct_type") = ""
	session("x_customer_organization_type") = ""
	session("x_bank_name") = ""
	session("x_customer_tax_id") = ""
	session("x_drivers_license_num") = ""
	session("x_drivers_license_state") =  ""
	session("x_drivers_license_dob") = ""
	session("pcIsSubscription") = ""
	session("pcIsSubTrial") = ""
	session("pcAgreeAll") = ""

	'// CLEAR SESSIONS:  Agreement Sessions
	pcCartArray = Session("pcCartSession")
	pcCartIndex =Session("pcCartIndex")
	aCnt = session("pcAgreeCnt")
	for a = 1 to aCnt
		for f=1 to pcCartIndex
			pSubscriptionID = cstr(pcCartArray(f,38))
			if  pSubscriptionID = session("Agree_"&pSubscriptionID&"_"&a) then
				 session("agree_"& pSubscriptionID&"_"&a) = ""
				 exit for
			 end if
		Next
	Next

End if
'SB E

if paymentCode="FREE" then
  if pCurOrderStatus<>"1" then
		response.redirect "msg.asp?message=211"
  end if
end if

' Reopen database connection
call openDb()

' get idCustomer
query="SELECT email,name,lastName,idCustomer,customerCompany FROM customers WHERE idcustomer=" &pidcustomer
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
	response.redirect "techErr.asp?error=49"&Server.Urlencode(dictLanguage.Item(Session("language")&"_updOrdStats_4"))
end if

pEmail=rs("email")
pName=rs("name")
pLName=rs("lastName")
pIdCustomer=rs("idCustomer")
pCustomerCompany=rs("customerCompany")
set rs=nothing

' iterates through order items
query="SELECT idProduct,quantity,idconfigSession FROM ProductsOrdered WHERE idOrder=" &pIdOrder
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

do while not rs.eof
	pIdProduct=rs("idProduct")
	pQuantity=rs("quantity")
	idconfig=rs("idconfigSession")
	'check if stock is ignored or not
	query="SELECT noStock FROM products WHERE idProduct="&pIdProduct
	set rstemp=conntemp.execute(query)
	pNoStock=rstemp("noStock")

	query="SELECT stock, sales, description FROM products WHERE idProduct=" &pIdProduct
	set rstemp=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	pDescription=rstemp("description")

	if (pNoStock=0) OR ((idconfig<>"") AND (idconfig<>"0")) then
		' decrease stock
		if ppStatus=0 then
			if (pNoStock=0) then
				query="UPDATE products SET stock=stock-" &pQuantity&" WHERE idProduct=" &pIdProduct
				set rsTemp=conntemp.execute(query)  
				if err.number<>0 then
					call LogErrorToDatabase()
					set rstemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
			end if
			query="SELECT stock FROM Products WHERE idProduct=" &pIdProduct
			set rsTemp=conntemp.execute(query)
			if not rsTemp.eof then
				tmpStock=rsTemp("stock")
				set rsTemp=nothing
				if clng(tmpStock)<0 then
					query="UPDATE products SET stock=0 WHERE idProduct=" &pIdProduct
					set rsTemp=conntemp.execute(query)
					set rsTemp=nothing
				end if
			end if
			set rsTemp=nothing
			
			'Update BTO Items & Additional Charges stock and sales
			IF (idconfig<>"") and (idconfig<>"0") then
				query="select stringProducts,stringQuantity,stringCProducts from configSessions where idconfigSession=" & idconfig
				set rs1=conntemp.execute(query)
				stringProducts=rs1("stringProducts")
				stringQuantity=rs1("stringQuantity")
				stringCProducts=rs1("stringCProducts")
				if (stringProducts<>"") and (stringProducts<>"na") then
					PrdArr=split(stringProducts,",")
					QtyArr=split(stringQuantity,",")

					for k=lbound(PrdArr) to ubound(PrdArr)
						if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock-" &QtyArr(k)*pQuantity&",sales=sales+" &QtyArr(k)*pQuantity&" WHERE idProduct=" &PrdArr(k)
							set rs1=conntemp.execute(query)
						end if
					next
				end if
				if (stringCProducts<>"") and (stringCProducts<>"na") then
					CPrdArr=split(stringCProducts,",")

					for k=lbound(CPrdArr) to ubound(CPrdArr)
						if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
							query="UPDATE products SET stock=stock-" &pQuantity&",sales=sales+" &pQuantity&" WHERE idProduct=" &CPrdArr(k)
							set rs1=conntemp.execute(query)
						end if
					next
				end if
			END IF
			'End Update BTO Items & Additional Charges

		end if
	end if

	' update sales
	if ppStatus=0 then
		query="UPDATE products SET sales=sales+" &pQuantity&" WHERE idProduct=" &pIdProduct
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query)
		set rstemp=nothing
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	end if
	rs.movenext
loop
set rs=nothing

'// START - Order is processed when placed -> Find out if it contains downloadable products
query="select idproduct,idconfigSession from ProductsOrdered WHERE idOrder="& pIdOrder
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

DPOrder="0"
IF pOrderStatus="3" THEN
	do while not rs.eof
		pIdProduct=rs("idproduct")
		tmpidConfig=rs("idconfigSession")
		query="select downloadable from products where idproduct=" & pIdProduct
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=connTemp.execute(query)
		if not rstemp.eof then
			pdownloadable=rstemp("downloadable")
			if (pdownloadable<>"") and (pdownloadable="1") then
				DPOrder="1"
			end if
		end if
		set rstemp=nothing
		'Find downloadable items in BTO configuration
		if tmpidConfig<>"" AND tmpidConfig>"0" then
			query="SELECT stringProducts,stringQuantity,stringCProducts FROM configSessions WHERE idconfigSession=" & tmpidConfig & ";"
			set rs1=connTemp.execute(query)
			if not rs1.eof then
				stringProducts=rs1("stringProducts")
				stringQuantity=rs1("stringQuantity")
				stringCProducts=rs1("stringCProducts")
				if (stringProducts<>"") and (stringProducts<>"na") then
					PrdArr=split(stringProducts,",")
					QtyArr=split(stringQuantity,",")

					for k=lbound(PrdArr) to ubound(PrdArr)
						if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
							query="SELECT idproduct FROM Products WHERE idProduct=" & PrdArr(k) & " AND Downloadable=1;"
							set rs1=conntemp.execute(query)
							if not rs1.eof then
								DPOrder="1"
							end if
							set rs1=nothing
						end if
					next
				end if
				if (stringCProducts<>"") and (stringCProducts<>"na") then
					CPrdArr=split(stringCProducts,",")
					for k=lbound(CPrdArr) to ubound(CPrdArr)
						if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
							query="SELECT idproduct FROM Products WHERE idProduct=" & CPrdArr(k) & " AND Downloadable=1;"
							set rs1=conntemp.execute(query)
							if not rs1.eof then
								DPOrder="1"
							end if
							set rs1=nothing
						end if
					next
				end if
			end if
			set rs1=nothing
		end if
	rs.moveNext
	loop
END IF
set rs=nothing
'// END - Order is processed when placed -> Find out if it contains downloadable products


'GGG Add-on start
'// START - Order is processed when placed -> Find out if it contains Gift Certificates
query="select idproduct from ProductsOrdered WHERE idOrder="& pIdOrder
set rstemp=connTemp.execute(query)

pGCs="0"
if pOrderStatus="3" then
	do while not rstemp.eof
		query="select pcprod_GC from products where idproduct=" & rstemp("idproduct")
		set rs=connTemp.execute(query)
		if not rs.eof then
			pGC=rs("pcprod_GC")
			if (pGC<>"") and (pGC="1") then
				pGCs="1"
			end if
		end if
		set rs=nothing
	rstemp.moveNext
	loop
end if
set rstemp=nothing
'GGG Add-on end
'// END - Order is processed when placed -> Find out if it contains gift certificates


' change order status & payment status (2=pending, 3=processed and enter gateway codes
Todaydate=Date()
if SQL_Format="1" then
	Todaydate=Day(Todaydate)&"/"&Month(Todaydate)&"/"&Year(Todaydate)
else
	Todaydate=Month(Todaydate)&"/"&Day(Todaydate)&"/"&Year(Todaydate)
end if
pOrderTime=Todaydate&" "&TIME()

if scDB="Access" then
	query="UPDATE orders SET pcOrd_GCs=" & pGCs & ",DPs=" & DPOrder & ",pcOrd_PaymentStatus=" & pPaymentStatus & ",orderstatus="&pOrderStatus&", processDate=#"&Todaydate&"#,gwAuthCode='"&gwAuthCode&"',gwTransID='"&gwTransID&"',pcOrd_CVNResponse='"&gwCVV2Code&"', pcOrd_AVSRespond='"&gwAVSCode&"', gwTransParentID='"&gwTransID&"',paymentCode='"&paymentCode&"',pcOrd_Payer='"&session("Payer")&"', pcOrd_Time=#"&pOrderTime&"# WHERE idOrder=" &pIdOrder
else
	query="UPDATE orders SET pcOrd_GCs=" & pGCs & ",DPs=" & DPOrder & ",pcOrd_PaymentStatus=" & pPaymentStatus & ",orderstatus="&pOrderStatus&", processDate='"&Todaydate&"',gwAuthCode='"&gwAuthCode&"',gwTransID='"&gwTransID&"',pcOrd_CVNResponse='"&gwCVV2Code&"', pcOrd_AVSRespond='"&gwAVSCode&"', gwTransParentID='"&gwTransID&"',paymentCode='"&paymentCode&"',pcOrd_Payer='"&session("Payer")&"', pcOrd_Time='"&pOrderTime&"' WHERE idOrder=" &pIdOrder
end if
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
set rs=nothing

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

qry_ID=pIdOrder
If piRewardPoints > 0 Then
	if ppStatus=0 then
		'even if pending, if a customer uses pts, they must be held as substracted until order is canceled.
		query="SELECT iRewardPointsUsed, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		iRewardPointsUsed=rs("iRewardPointsUsed")
		If IsNull(iRewardPointsUsed) OR iRewardPointsUsed="" Then
			iRewardPointsUsed=0
		end if
		query = "SELECT iRewardValue FROM orders WHERE idOrder=" & qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		piRewardValue=rs("iRewardValue")
		iNewUsed = iRewardPointsUsed + piRewardPoints
		query = "UPDATE customers SET iRewardPointsUsed=" & iNewUsed & " WHERE idCustomer=" & pIdCustomer
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
End If

'// START - Order is processed when placed -> Perform order processing tasks
IF pOrderStatus="3" THEN
	if ppStatus=0 then
		'update reward pts.
		If RewardsActive <> 0 then
			'add points from refferer if any points were awarded.
			If piRewardRefId>0 AND piRewardPointsRef>0 then
				query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & piRewardRefId
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				iAccrued=rs("iRewardPointsAccrued") + piRewardPointsRef
				query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & piRewardRefId
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				set rs=nothing
			End If
			'if points were used, subtract from customer

			'add accrued points from customer if any points were accrued
			If piRewardPointsCustAccrued>0 then
				query="SELECT iRewardPointsAccrued, idCustomer FROM customers WHERE idCustomer=" & pIdCustomer
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				iAccrued=rs("iRewardPointsAccrued") + piRewardPointsCustAccrued
				query="UPDATE customers SET iRewardPointsAccrued=" & iAccrued & " WHERE idCustomer=" & pIdCustomer
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				set rs=nothing
			End If
		End If
	End If

	query="select idcustomer,orderdate,processdate from Orders WHERE idOrder="& qry_ID
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		pIdCustomer=rs("IdCustomer")
		pOrderDate=rs("OrderDate")
		pProcessDate=rs("ProcessDate")
	end if
	set rs=nothing

	'GGG Add-on start

	IF pGCs="1" then
		query="select idproduct,quantity from ProductsOrdered WHERE idOrder="& qry_ID
		set rstemp=connTemp.execute(query)
		DO while not rstemp.eof
			query="select pcGC.pcGC_Exp,pcGC.pcGC_ExpDate,pcGC.pcGC_ExpDays,pcGC.pcGC_CodeGen,pcGC.pcGC_GenFile,products.sku,products.price from pcGC,Products where pcGC.pcGC_idproduct=" & rstemp("idproduct") & " and Products.idproduct=pcGC.pcGC_idproduct and products.pcprod_GC=1"
			set rs=connTemp.execute(query)

			if not rs.eof then
				pIdproduct=rstemp("idproduct")
				pQuantity=rstemp("quantity")
				pGCExp=rs("pcGC_Exp")
				pGCExpDate=rs("pcGC_ExpDate")
				pGCExpDay=rs("pcGC_ExpDays")
				pGCGen=rs("pcGC_CodeGen")
				pGCGenFile=rs("pcGC_GenFile")
				pSku=rs("sku")
				pGCAmount=rs("price")
				if pGCGen<>"" then
				else
				pGCGen="0"
				end if
				if (pGCGen=1) and (pGCGenFile="") then
				pGCGen="0"
				pGCGenFile="DefaultGiftCode.asp"
				end if

				if (pGCGen="0") or (not (pGCGenFile<>"")) then

					pGCGenFile="DefaultGiftCode.asp"

				end if

				if (pGCExp="2") then
				pGCExpDate=Now()+cint(pGCExpDay)
				end if

				if (pGCExp="1") and (year(pGCExpDate)=1900) then
				pGCExp="0"
				pGCExpDate="01/01/1900"
				end if

				if (pGCExp="2") and (pGCExpDay="0") then
				pGCExp="0"
				pGCExpDate="01/01/1900"
				end if

				if SQL_Format="1" then
				pGCExpDate=(day(pGCExpDate)&"/"&month(pGCExpDate)&"/"&year(pGCExpDate))
				else
				pGCExpDate=(month(pGCExpDate)&"/"&day(pGCExpDate)&"/"&year(pGCExpDate))
				end if

				IF (pGCGenFile<>"") THEN

						SPath1=Request.ServerVariables("PATH_INFO")
						mycount1=0
						do while mycount1<1
							if mid(SPath1,len(SPath1),1)="/" then
							mycount1=mycount1+1
							end if
							if mycount1<1 then
							SPath1=mid(SPath1,1,len(SPath1)-1)
							end if
						loop
						SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
						if Right(SPathInfo,1)="/" then
							pGCGenFile=SPathInfo & "licenses/" & pGCGenFile
						else
							pGCGenFile=SPathInfo & "/licenses/" & pGCGenFile
						end if
						pGCGenFile=replace(pGCGenFile,"/pc/","/"&scAdminFolderName&"/")
						L_Action=pGCGenFile

					L_postdata=""
					L_postdata=L_postdata&"idorder=" & pIdOrder
					L_postdata=L_postdata&"&orderDate=" & pOrderDate
					L_postdata=L_postdata&"&ProcessDate=" & pProcessDate
					L_postdata=L_postdata&"&idcustomer=" & pIdCustomer
					L_postdata=L_postdata&"&idproduct=" & pIdproduct
					L_postdata=L_postdata&"&quantity=" & pQuantity
					L_postdata=L_postdata&"&sku=" & pSKU

					For k=1 to Cint(pQuantity)

					DO

					Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
					srvXmlHttp.open "POST", L_Action, False
					srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					srvXmlHttp.send L_postdata
					result1 = srvXmlHttp.responseText

					RArray = split(result1,"<br>")
					GiftCode= RArray(2)

					'If have errors from GiftCode Generator
					IF (IsNumeric(RArray(0))=false) and (IsNumeric(RArray(1))=false) then

					Tn1=""
					For w=1 to 6
					Randomize
					myC=Fix(3*Rnd)
					Select Case myC
						Case 0:
						Randomize
						Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
						Case 1:
						Randomize
						Tn1=Tn1 & Cstr(Fix(10*Rnd))
						Case 2:
						Randomize
						Tn1=Tn1 & Chr(Fix(26*Rnd)+97)
					End Select
					Next

					GiftCode=Tn1 & Day(Now()) & Minute(Now()) & Second(Now())

					END IF

					ReqExist=0

					query="select pcGO_IDProduct from pcGCOrdered where pcGO_GcCode='" & GiftCode & "'"
					set rstemp19=connTemp.execute(query)

					if not rstemp19.eof then
					ReqExist=1
					end if

					LOOP UNTIL ReqExist=0

					'Insert Gift Codes to Database

					if scDB="Access" then
					query="Insert into pcGCOrdered (pcGO_IdOrder,pcGO_IdProduct,pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status) values (" & pIdOrder & "," & pIdProduct & ",'" & GiftCode & "',#" & pGCExpDate & "#," & pGCAmount & ",1)"
					else
					query="Insert into pcGCOrdered (pcGO_IdOrder,pcGO_IdProduct,pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status) values (" & pIdOrder & "," & pIdProduct & ",'" & GiftCode & "','" & pGCExpDate & "'," & pGCAmount & ",1)"
					end if
					set rstemp19=connTemp.execute(query)

					Next

				END IF

			end if
			rstemp.moveNext
		LOOP
	END IF

	'GGG Add-on end

	'Call License Generator for Standard & BTO Products

	Sub CreateDownloadInfo(pIDProduct,pQuantity)
		Dim query,rstemp,pSku,pLicense,pLocalLG,pRemoteLG,k,dd

			query="select sku,License,LocalLG,RemoteLG from Products,DProducts where products.idproduct=" & pIdproduct & " and DProducts.idproduct=Products.idproduct and products.downloadable=1"
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=connTemp.execute(query)
			if not rstemp.eof then
				pSku=rstemp("sku")
				pLicense=rstemp("License")
				pLocalLG=rstemp("LocalLG")
				pRemoteLG=rstemp("RemoteLG")
				set rstemp=nothing

				IF (pLicense<>"") and (pLicense="1") THEN
					if pLocalLG<>"" then
						SPath1=Request.ServerVariables("PATH_INFO")
						mycount1=0
						do while mycount1<1
							if mid(SPath1,len(SPath1),1)="/" then
								mycount1=mycount1+1
							end if
							if mycount1<1 then
								SPath1=mid(SPath1,1,len(SPath1)-1)
							end if
						loop
						SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
						if Right(SPathInfo,1)="/" then
							pLocalLG=SPathInfo & "licenses/" & pLocalLG
						else
							pLocalLG=SPathInfo & "/licenses/" & pLocalLG
						end if
						pLocalLG=replace(pLocalLG,"/pc/","/"&scAdminFolderName&"/")
						L_Action=pLocalLG
					else
						L_Action=pRemoteLG
					end if
					L_postdata=""
					L_postdata=L_postdata&"idorder=" & pIdOrder
					L_postdata=L_postdata&"&orderDate=" & pOrderDate
					L_postdata=L_postdata&"&ProcessDate=" & pProcessDate
					L_postdata=L_postdata&"&idcustomer=" & pIdCustomer
					L_postdata=L_postdata&"&idproduct=" & pIdproduct
					L_postdata=L_postdata&"&quantity=" & pQuantity
					L_postdata=L_postdata&"&sku=" & pSKU

					Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
					srvXmlHttp.open "POST", L_Action, False
					srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					srvXmlHttp.send L_postdata
					result1 = srvXmlHttp.responseText
					AR=split(result1,"<br>")

					rIdOrder=AR(0)
					rIdProduct=AR(1)
					Lic1=split(AR(2),"***")
					Lic2=split(AR(3),"***")
					Lic3=split(AR(4),"***")
					Lic4=split(AR(5),"***")
					Lic5=split(AR(6),"***")

					For k=0 to Cint(pQuantity)-1
						if K<=ubound(Lic1) then
							PLic1=Lic1(k)
						else
							PLic1=""
						end if
						if K<=ubound(Lic2) then
							PLic2=Lic2(k)
						else
							PLic2=""
						end if
						if K<=ubound(Lic3) then
							PLic3=Lic3(k)
						else
							PLic3=""
						end if
						if K<=ubound(Lic4) then
							PLic4=Lic4(k)
						else
							PLic4=""
						end if
						if K<=ubound(Lic5) then
							PLic5=Lic5(k)
						else
							PLic5=""
						end if
						if ppStatus=0 then
							query="Insert into DPLicenses (IdOrder,IdProduct,Lic1,Lic2,Lic3,Lic4,Lic5) values (" & rIdOrder & "," & rIdProduct & ",'" & PLic1 & "','" & PLic2 & "','" & PLic3 & "','" & PLic4 & "','" & PLic5 & "')"
							set rstemp=server.CreateObject("ADODB.RecordSet")
							set rstemp=connTemp.execute(query)
							set rstemp=nothing
						end if
					Next
				END IF

				DO
					Tn1=""
						For dd=1 to 24
							Randomize
							myC=Fix(3*Rnd)
							Select Case myC
								Case 0:
									Randomize
									Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
								Case 1:
									Randomize
									Tn1=Tn1 & Cstr(Fix(10*Rnd))
								Case 2:
									Randomize
									Tn1=Tn1 & Chr(Fix(26*Rnd)+97)
							End Select
						Next

						ReqExist=0

						query="select IDOrder from DPRequests where RequestSTR='" & Tn1 & "'"
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=connTemp.execute(query)

						if not rstemp.eof then
							ReqExist=1
						end if
						set rstemp=nothing
				LOOP UNTIL ReqExist=0

				if ppStatus=0 then
					pTodaysDate=Date()
					if SQL_Format="1" then
						pTodaysDate=(day(pTodaysDate)&"/"&month(pTodaysDate)&"/"&year(pTodaysDate))
					else
						pTodaysDate=(month(pTodaysDate)&"/"&day(pTodaysDate)&"/"&year(pTodaysDate))
					end if

					'Insert Standard & BTO Products Download Requests into DPRequests Table
					if scDB="Access" then
						query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "',#" & pTodaysDate & "#)"
					else
						query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "','" & pTodaysDate & "')"
					end if
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
				end if
			end if
			set rstemp=nothing

	End Sub
	IF DPOrder="1" then
		query="select idproduct,quantity,idconfigSession from ProductsOrdered WHERE idOrder="& qry_ID
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)

		do while not rs.eof
			pIdProduct=rs("idproduct")
			pQuantity=rs("quantity")
			tmpidConfig=rs("idconfigSession")
			Call CreateDownloadInfo(pIDProduct,pQuantity)
			'Find downloadable items in BTO configuration
			if tmpidConfig<>"" AND tmpidConfig>"0" then
				query="SELECT stringProducts,stringQuantity,stringCProducts FROM configSessions WHERE idconfigSession=" & tmpidConfig & ";"
				set rs1=connTemp.execute(query)
				if not rs1.eof then
					stringProducts=rs1("stringProducts")
					stringQuantity=rs1("stringQuantity")
					stringCProducts=rs1("stringCProducts")
					if (stringProducts<>"") and (stringProducts<>"na") then
						PrdArr=split(stringProducts,",")
						QtyArr=split(stringQuantity,",")

						for k=lbound(PrdArr) to ubound(PrdArr)
							if (PrdArr(k)<>"") and (PrdArr(k)<>"0") then
								query="SELECT idproduct FROM Products WHERE idProduct=" & PrdArr(k) & " AND Downloadable=1;"
								set rs1=conntemp.execute(query)
								if not rs1.eof then
									Call CreateDownloadInfo(PrdArr(k),QtyArr(k)*pQuantity)
								end if
								set rs1=nothing
							end if
						next
					end if
					if (stringCProducts<>"") and (stringCProducts<>"na") then
						CPrdArr=split(stringCProducts,",")
						for k=lbound(CPrdArr) to ubound(CPrdArr)
							if (CPrdArr(k)<>"") and (CPrdArr(k)<>"0") then
								query="SELECT idproduct FROM Products WHERE idProduct=" & CPrdArr(k) & " AND Downloadable=1;"
								set rs1=conntemp.execute(query)
								if not rs1.eof then
									Call CreateDownloadInfo(CPrdArr(k),1)
								end if
								set rs1=nothing
							end if
						next
					end if
				end if
				set rs1=nothing
			end if
			rs.moveNext
		loop
		set rs=nothing
	END IF

END IF
'// END - Order is processed when placed -> Perform order processing tasks
%>

<!--#include file="adminNewOrderEmail.asp"-->

<%
if ppStatus=0 then
	dim strNewOrderSubject
	strNewOrderSubject=dictLanguage.Item(Session("language")&"_storeEmail_9")&(scpre + int(pIdOrder))
	if pcOrderKey<>"" then
		storeAdminEmail=storeAdminEmail & vbCrLf
		storeAdminEmail=storeAdminEmail & "----------------------------------------------------------------------------------------------" & vbCrLf
		storeAdminEmail=storeAdminEmail & "ORDER CODE: " & pcOrderKey & vbCrLf
		storeAdminEmail=storeAdminEmail & "----------------------------------------------------------------------------------------------" & vbCrLf
	end if
	call sendmail (scCompanyName, scEmail, scFrmEmail, strNewOrderSubject, replace(storeAdminEmail,"&quot;", chr(34)))
end if

'GGG Add-on start
if pOrderStatus>="2" then
%>
<!--#include file="ggg_UpdateGC.asp"-->
<%
end if
'GGG Add-on end

'// START - Order is processed when placed -> Send order confirmation
IF pOrderStatus="3" THEN
	'order processed
	if ppStatus=0 then
		'Variable to generate Customer Order Confirmation Email
		pcv_CustomerReceived=0 %>
		<!--#include file="customerOrderConfirmEmail.asp"-->
		<%
		pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_2") & " - " & dictLanguage.Item(Session("language")&"_sendMail_1") & (scpre + int(pIdOrder))
		call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerEmail, "&quot;", chr(34)))
		'Start SDBA%>
		<!--#include file="inc_DropShipperNotificationEmail.asp"-->
		<%'End SDBA
	end if
ELSE ' Order is pendng -> Send order received e-mail
	if ppStatus=0 then
		'Variable to generate Customer Order Received Email
		pcv_CustomerReceived=1%>
		<!--#include file="customerOrderConfirmEmail.asp"-->
		<%pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_1") & " - " & dictLanguage.Item(Session("language")&"_sendMail_1") & (scpre + int(pIdOrder))
		call sendmail (scCompanyName, scEmail, pEmail, pcv_strSubject, replace(customerEmail, "&quot;", chr(34)))
	end if
END IF
'// END - Order is processed when placed -> Send order confirmation

'Check for onetime discount session
If strPcOrd_DiscountsUsed<>"" then
	if instr(strPcOrd_DiscountsUsed,",") then
		strPcOrd_DiscountsUsed=replace(strPcOrd_DiscountsUsed,", ",",")
		pDiscountUsedArray=split(strPcOrd_DiscountsUsed,",")
		tempCnt=Ubound(pDiscountUsedArray)
	else
		tempCnt=0
	end if
	for i=0 to tempCnt
		if tempCnt=0 then
			pDiscountUsedVar=strPcOrd_DiscountsUsed
		else
			pDiscountUsedVar=pDiscountUsedArray(i)
		end if
		query="INSERT INTO used_discounts (idDiscount, idcustomer) VALUES ("&pDiscountUsedVar&","&pIdCustomer&");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		set rs=nothing
	Next
end if

' // START - Affiliate Notification
' Send notification to affiliate only if the order is processed when placed
IF pOrderStatus="3" THEN

	query="SELECT orders.idaffiliate, orders.affiliatePay, affiliates.affiliateemail, affiliates.affiliateName FROM orders, affiliates WHERE affiliates.idaffiliate=orders.idaffiliate AND orders.idOrder="& pIdOrder
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	If NOT rs.eof then
		If rs("idaffiliate")<>1 then
			AffiliatePay=rs("affiliatePay")
			AffiliateEmail=rs("affiliateemail")
			AffiliateName=rs("affiliateName")%>
			<!--#include file="affiliateOrderConfirmEmail.asp"-->
			<%pcv_strSubject = dictLanguage.Item(Session("language")&"_storeEmail_3")
			if ppStatus=0 then
				call sendmail (scCompanyName, scEmail, AffiliateEmail, pcv_strSubject, AffiliateOrderEmail)
			end if
		End If
		set rs=nothing
	End If
END IF
' // END - Affiliate Notification


'Start SDBA - Send Low Inventory Notification
%>
<!--#include file="inc_StockNotificationEmail.asp"-->
<%
'End SDBA - Send Low Inventory Notification

'GGG Add-on start%>
<!--#include file="ggg_updGRQty.asp"-->
<%'GGG Add-on end

' clear cart data
dim pcCartArray2(100,45)
Session("pcCartSession")=pcCartArray2
Session("pcCartIndex")=Cint(0)
session("iOrderTotal")=""
session("continueRef")=""
session("pcSFCartRewards")=Cint(0)
session("pcSFUseRewards")=Cint(0)
session("IDRefer")=""
session("specialdiscount")=""
session("EPN_idOrder")=""
session("pc_pidOrder")=""
session("GWAuthCode")=""
session("GWTransId")=""
session("Entered-" & session("GWPaymentId"))=""
session("admin-" & session("GWPaymentId") & "-pCardType")=""
session("admin-" & session("GWPaymentId") & "-pCardNumber")=""
session("admin-" & session("GWPaymentId") & "-expMonth")=""
session("admin-" & session("GWPaymentId") & "-expYear")=""
session("GWPaymentId")=""
session("GWTransType")=""
session("GWOrderId")=""
session("GWSessionID")=""
session("GWOrderDone")=""
session("idGWSubmit")=""
session("idGWSubmit2")=""
session("idGWSubmit3")=""
session("Gateway")=""
session("SaveOrder")=""
session("RefRewardPointsTest")=""
'GGG Add-on start
session("Cust_BuyGift")=""
session("Cust_IDEvent")=""
'GGG Add-on end

call closeDb()
' go to confirmation page
session("idOrder")=pIdOrder
response.redirect "orderComplete.asp"

'Functions
Public Function FixedField(ByVal Width, ByVal Justify, ByVal Text)
	Select Case True
		Case Width < Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Left(Text, Width)
				Case Justify="R"
					FixedField=Right(Text, Width)
				Case Else
			End Select

		Case Width=Len(Text)
			FixedField=Text

		Case Width > Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Text & String(Width - Len(Text), " ")
				Case Justify="R"
					FixedField=String(Width - Len(Text), " ") & Text
				Case Else
			End Select

	End Select

End Function
%>