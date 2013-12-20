<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 'SB S %>
<!--#include file="inc_sb.asp"--> 
<% 'SB E %>
<%
'// ONE PAGE CHECKOUT SETTING
'// Hide company address, customer billing and shipping addresses
'// in the Order Preview section.

Dim pcIntHideAddresses
pcIntHideAddresses = 1 ' Addresses are hidden
'pcIntHideAddresses = 0 ' Addresses are shown

'// ONE PAGE CHECKOUT SETTING - END

response.Buffer=true

if session("idCustomer")=0 OR session("idCustomer")="" then
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
end if

if session("pcSFIdDbSession")="" OR session("pcSFRandomKey")=""  then
	response.clear
	Call SetContentType()
	response.write ""
	response.End
end if

Response.Expires = 60
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"--> 
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/shipFromSettings.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/bto_language.asp"--> 
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/GCConstants.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/sendmail.asp" -->
<!--#include file="../includes/dateinc.asp" -->
<!--#include file="opc_contentType.asp" -->
<% Session.LCID = 1033 %>
<% On Error Resume Next

Dim TurnOffDiscountCodesWhenHasSale, HavePrdsOnSale

TurnOffDiscountCodesWhenHasSale=scDisableDiscountCodes
'=1: True - Default
'=0: False

HavePrdsOnSale=0

Call SetContentType()

'GGG Add-on start%>
<!--#include file="ggg_inc_calGW.asp" -->
<% intGCIncludeShipping=GC_INCSHIPPING

intTaxExemptZoneFlag="1" 'Change to 0 if you want to tax any tax zone exempt products when they are added to the cart with taxable products. "1" will ensure that tax zone exempt products are never taxed for that zone.  
Dim GiftWrapPaymentTotal
GiftWrapPaymentTotal=0
%>
<%'GGG Add-on end
dim query, connTemp, rs, pcCartArray, ppcCartIndex, paymentTotal

'SB S
Dim pcIsSubscription , StrandSub 
pcIsSubscription = session("pcIsSubscription")

Dim pcv_sbTax
pcv_sbTax=getUserInput(request("sbTax"),0)
'SB E

'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
ppcCartIndex=session("pcCartIndex")

paymentTotal=0

'// GET CUSTOMER SESSION DATA
call openDb()

QUERY="SELECT customers.name, customers.lastName, customers.customerCompany, customers.email, customers.phone, customers.fax, customers.address, customers.address2, customers.zip, customers.stateCode, customers.state, customers.city, customers.countryCode, customers.iRewardPointsAccrued, customers.iRewardPointsUsed, pcCustSession_ShippingFirstName, pcCustSession_ShippingLastName, pcCustSession_ShippingCompany, pcCustSession_ShippingAddress, pcCustSession_ShippingAddress2, pcCustSession_ShippingCity, pcCustSession_ShippingStateCode, pcCustSession_ShippingProvince, pcCustSession_ShippingPostalCode, pcCustSession_ShippingCountryCode, pcCustSession_ShippingPhone, pcCustSession_ShippingNickName, pcCustSession_TaxShippingAlone, pcCustSession_TaxShippingAndHandlingTogether, pcCustSession_TaxLocation, pcCustSession_TaxProductAmount, pcCustSession_OrdPackageNumber, pcCustSession_ShippingArray, pcCustSession_ShippingResidential, pcCustSession_IdPayment, pcCustSession_Comment, pcCustSession_discountcode, pcCustSession_UseRewards, pcCustSession_RewardsBalance,pcCustSession_NullShipper,pcCustSession_NullShipRates,pcCustSession_TF1,pcCustSession_DF1,pcCustSession_OrderName,pcCustSession_ShowShipAddr,pcCustSession_ShippingEmail,pcCustSession_ShippingFax,pcCustSession_GCDetails FROM pcCustomerSessions INNER JOIN customers ON pcCustomerSessions.idCustomer = customers.idcustomer WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&")) ORDER BY pcCustomerSessions.idDbSession DESC;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	'set rs=nothing
	'call closedb()
	'response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pcStrBillingFirstName=rs("name")
pcStrBillingLastName=rs("lastName")
pcStrBillingCompany=rs("customerCompany")
pcStrBillingEmail=rs("email")
pcStrBillingPhone=rs("phone")
pcStrBillingfax=rs("fax")
pcStrBillingAddress=rs("address")
pcStrBillingAddress2=rs("address2")
pcStrBillingPostalCode=rs("zip")
pcStrBillingStateCode=rs("stateCode")
pcStrBillingProvince=rs("state")
pcStrBillingCity=rs("city")
pcStrBillingCountryCode=rs("CountryCode")
pcIntRewardPointsAccrued = rs("iRewardPointsAccrued")
pcIntRewardPointsUsed = rs("iRewardPointsUsed")
pcStrShippingFirstName=rs("pcCustSession_ShippingFirstName")
pcStrShippingLastName=rs("pcCustSession_ShippingLastName")
pcStrShippingCompany=rs("pcCustSession_ShippingCompany")
pcStrShippingAddress=rs("pcCustSession_ShippingAddress")
pcStrShippingAddress2=rs("pcCustSession_ShippingAddress2")
pcStrShippingCity=rs("pcCustSession_ShippingCity")
pcStrShippingStateCode=rs("pcCustSession_ShippingStateCode")
pcStrShippingProvince=rs("pcCustSession_ShippingProvince")
pcStrShippingPostalCode=rs("pcCustSession_ShippingPostalCode")
pcStrShippingCountryCode=rs("pcCustSession_ShippingCountryCode")
pcStrShippingPhone=rs("pcCustSession_ShippingPhone")
pcStrShippingNickName=rs("pcCustSession_ShippingNickName")
TAX_SHIPPING_ALONE=rs("pcCustSession_TaxShippingAlone")
TAX_SHIPPING_AND_HANDLING_TOGETHER=rs("pcCustSession_TaxShippingAndHandlingTogether")
ptaxLoc=Cdbl(rs("pcCustSession_TaxLocation"))
ptaxPrdAmount =ccur(rs("pcCustSession_TaxProductAmount"))
pcIntOrdPackageNumber=rs("pcCustSession_OrdPackageNumber")
pcShippingArray=rs("pcCustSession_ShippingArray")
pOrdShipType=rs("pcCustSession_ShippingResidential")
pcIdPayment=rs("pcCustSession_IdPayment")
savOrderComments=rs("pcCustSession_Comment")
savdiscountcode=rs("pcCustSession_discountcode")
savUseRewards=rs("pcCustSession_UseRewards")
savNullShipper=rs("pcCustSession_NullShipper")
savNullShipRates=rs("pcCustSession_NullShipRates")
savTF1=rs("pcCustSession_TF1")
savDF1=rs("pcCustSession_DF1")
savOrderNickName=rs("pcCustSession_OrderName")
pcShowShipAddr=rs("pcCustSession_ShowShipAddr")
pcStrShippingEmail=rs("pcCustSession_ShippingEmail")
pcStrShippingFax=rs("pcCustSession_ShippingFax")
savGCs=rs("pcCustSession_GCDetails")
if savGCs<>"" then
	GCArr=split(savGCs,"|g|")
	savGCs=""
	for y=0 to ubound(GCArr)
		if GCArr(y)<>"" then
			GCInfo=split(GCArr(y),"|s|")
			if savGCs<>"" then
				savGCs=savGCs & ","
			end if
			savGCs=savGCs & GCInfo(0)
		end if
	next
	if savdiscountcode<>"" then
		if Right(savdiscountcode,1)<>"," then
			savdiscountcode=savdiscountcode & ","
		end if
	end if
	savdiscountcode=savdiscountcode & savGCs
end if				


set rs=nothing

'// GET TAX ZONE DATA
%> <!--#include file="pcTaxZone.asp"--> <%

'// Check if the Customer is European Union 
Dim pcv_IsEUMemberState
pcv_IsEUMemberState = pcf_IsEUMemberState(pcStrShippingCountryCode)

pcSFIntBalance = 0
If RewardsActive = 1 Then
	If IsNull(pcIntRewardPointsAccrued) or pcIntRewardPointsAccrued="" Then 
		pcIntRewardPointsAccrued = 0
	End if
	If IsNull(pcIntRewardPointsUsed) or pcIntRewardPointsUsed="" Then 
		pcIntRewardPointsUsed = 0
	End if
	pcIntBalance = pcIntRewardPointsAccrued - pcIntRewardPointsUsed
	pcIntDollarValue = pcIntBalance * (RewardsPercent / 100)
	pcSFIntBalance = pcIntBalance
End If

if request("rtype")="1" then
	pcUseRewards=request("UseRewards")

	if IsNull(pcUseRewards) OR pcUseRewards="" then
		pcUseRewards=0
	end if
else
	pcUseRewards=savUseRewards
	if IsNull(pcUseRewards) OR pcUseRewards="" then
		pcUseRewards=0
	end if
end if

if validNum(pcUseRewards) then
	pcSFCartRewards=0
	If (Cint(pcUseRewards) > pcSFIntBalance) AND (pcSFIntBalance>0) Then
		pcUseRewards = pcSFIntBalance
	End If
	If pcSFIntBalance=0 Then
		pcUseRewards = 0
	End If
	pcSFUseRewards=pcUseRewards
	if session("customerType")="1" AND RewardsIncludeWholesale=0 then 
		pcSFUseRewards=0
	end if
else
	if request("rtype")<>"1" then
		if not validNum(pcSFUseRewards) then
			pcSFUseRewards=0
		end if
	else
		pcSFUseRewards=0
	end if
end if

if pcSFUseRewards="" or pcSFUseRewards="0" then
	pcSFUseRewards=""
	'This customer will accrue the points since they are not using any for the purchase
	pcIntCartRewards=Int(calculateCartRewards(pcCartArray, ppcCartIndex))
	pcSFCartRewards=pcIntCartRewards
	'if customer is wholesale and wholesale is not included in rewards
	if session("customerType")="1" AND RewardsIncludeWholesale=0 then 
		pcSFCartRewards=0
		pcIntCartRewards=0
	end if
end if

'//Request Discount Code from Input to recalculate
if request("rtype")<>"1" then
	displayDiscountCode=savdiscountcode
else
	displayDiscountCode=request("discountcode")
end if
if displayDiscountCode<>"" then
displayDiscountCode=replace(displayDiscountCode,", ",",")
displayDiscountCode=replace(displayDiscountCode," ,",",")
end if

'//Set AutoDiscount flag to 0
pcIntADCnt=0

IF (TurnOffDiscountCodesWhenHasSale="1") AND (scDB="SQL") AND (not scHideDiscField="1") THEN
	Dim tmpPrdList
	tmpPrdList=""
	for f=1 to ppcCartIndex
		if pcCartArray(f,10)=0 then
			if tmpPrdList<>"" then
				tmpPrdList=tmpPrdList & ","
			end if
			tmpPrdList=tmpPrdList & pcCartArray(f,0)
		end if
	next
	if tmpPrdList="" then
		tmpPrdList="0"
	end if
	tmpPrdList="(" & tmpPrdList & ")"
	query="SELECT idProduct FROM Products WHERE idProduct IN " & tmpPrdList & " AND pcSC_ID>0;"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		HavePrdsOnSale=1
	end if
	set rsQ=nothing
END IF

'//If this is first visit, check for Auto discounts
IF HavePrdsOnSale=0 THEN
	query="SELECT discountcode FROM discounts WHERE pcDisc_Auto=1 AND active=-1 ORDER BY percentagetodiscount DESC,pricetodiscount DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if NOT rs.eof then
		pcStrAutoDiscCode=""
		do until rs.eof
			pcIntADCnt=pcIntADCnt+1
			if pcIntADCnt=1 then
				pcStrAutoDiscCode=pcStrAutoDiscCode&rs("discountcode")
			else
				pcStrAutoDiscCode=pcStrAutoDiscCode&","&rs("discountcode")
			end if
			rs.movenext
		loop
	end if
	if len(displayDiscountCode)>0 then
		displayDiscountCode= pcStrAutoDiscCode & "," & trim(displayDiscountCode)
	else
		displayDiscountCode= pcStrAutoDiscCode	
	end if
	set rs=nothing
END IF

if displayDiscountCode<>"" then
pDiscountCode=URLDecode(getUserInput(displayDiscountCode,0))
end if

if session("pcSFIdPayment")<>"" AND session("pcSFIdPayment")<>"0" then
pidPayment=session("pcSFIdPayment")
else
pidPayment=pcIdPayment
end if
if request("idpayment")<>"" then
	pidPayment=URLDecode(getUserInput(request("idpayment"),0))
end if
if pidPayment="" then
	pidPayment=pcIdPayment
else
	if not IsNumeric(pidPayment) then
		pidPayment=pcIdPayment
	end if
end if
pcIdPayment=pidPayment
session("pcSFIdPayment")=pidPayment

'RP ADDON-S
if session("customerType")="1" AND RewardsIncludeWholesale=0 then 
	pcSFUseRewards=""
end if
	
if session("customerType")=1 AND ptaxwholesale=0 AND (ptaxCanada<>"1" OR (ptaxCanada="1" AND session("SFTaxZoneRateCnt")=0)) then
	ptaxPrdAmount=ccur(0)
end if

' if cart has no items cancel the order
'if countCartRows(pcCartArray, ppcCartIndex)=0 then
	'response.redirect "msg.asp?message=201"
'end if

' calculate total price of the order, total weight and product total quantities
Dim pTaxableTotal, pSubTotal, pCartTotalWeight, pCartQuantity, pEryPassword
'SB S
If len(pcv_sbTax)>0 Then
	pTaxableTotal=ccur(calculateTaxableTotal_SB(pcCartArray, ppcCartIndex))
Else
pTaxableTotal=ccur(calculateTaxableTotal(pcCartArray, ppcCartIndex))
End If
'SB S
pSubTotal=ccur(calculateCartTotal(pcCartArray, ppcCartIndex))
pCartTotalWeight=Int(calculateCartWeight(pcCartArray, ppcCartIndex))
pCartQuantity=Int(calculateCartQuantity(pcCartArray, ppcCartIndex))

err.clear

'VALIDATE CART - Used on SaveOrd.asp
pSFSubTotal=pSubTotal

'GET SHIPMENT DATA
If savNullShipper="Yes" then
	pcStrShipmentDesc=ship_dictLanguage.Item(Session("language")&"_noShip_a")
	pcShipmentPriceToAdd="0"
Else
	if savNullShipRates="Yes" then
		pcStrShipmentDesc=ship_dictLanguage.Item(Session("language")&"_noShip_b")
		pcShipmentPriceToAdd="0"
	else
		TempStrNewShipping=""
		pcSplitShipping=split(pcShippingArray,",")
		TempStrShipper=pcSplitShipping(0)
		TempStrService=pcSplitShipping(1)
		TempDblPostage=pcSplitShipping(2)
		
		if ubound(pcSplitShipping)>4 then
			query="SELECT idshipservice, serviceFreeOverAmt FROM shipService WHERE serviceCode='"&pcSplitShipping(5)&"';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			
			if err.number<>0 then
				call LogErrorToDatabase()
				'set rs=nothing
				'call closedb()
				'response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
			if not rs.eof then
				pcIntIdShipService=rs("idshipservice")
				serviceFreeOverAmt=rs("serviceFreeOverAmt")
			end if
			
			set rs=nothing
			
		end if
		TempStrNewShipping=TempStrNewShipping&TempStrShipper&","&TempStrService&","&TempDblPostage
		
		if TempStrService="" then
			pcStrShipmentDesc=TempStrShipper
			if pcIntIdShipService="" then
				query="SELECT idshipservice, serviceFreeOverAmt FROM shipService WHERE serviceDescription like '%" & TempStrShipper & "%'"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				
				if err.number<>0 then
					call LogErrorToDatabase()
					'set rs=nothing
					'call closedb()
					'response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
				if not rs.eof then
					pcIntIdShipService=rs("idshipservice")
					serviceFreeOverAmt=rs("serviceFreeOverAmt")
				end if
				
				set rs=nothing
			end if
		else
			if pcIntIdShipService="" then
				query="SELECT idshipservice, serviceFreeOverAmt FROM shipService WHERE serviceDescription LIKE '%" & TempStrService & "%'"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				
				if err.number<>0 then
					call LogErrorToDatabase()
					'set rs=nothing
					'call closedb()
					'response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
				if not rs.eof then
					pcIntIdShipService=rs("idshipservice")
					serviceFreeOverAmt=rs("serviceFreeOverAmt")
				end if
				
				set rs=nothing
			end if

			If TempStrShipper="UPS" then
				serviceVar=TempStrService
				select case serviceVar
				case "UPS Next Day Air "
					TempStrService="UPS Next Day Air&reg;"
				case "UPS 2nd Day Air "
					TempStrService="UPS 2nd Day Air&reg;"
				case "UPS Ground"
					TempStrService="UPS Ground"
				case "UPS Worldwide Express "
					TempStrService="UPS Worldwide Express<sup>SM</sup>"
				case "UPS Worldwide Expedited "
					TempStrService="UPS Worldwide Expedited<sup>SM</sup>"
				case "UPS Standard To Canada"
					TempStrService="UPS Standard To Canada"
				case "UPS 3 Day Select "
					TempStrService="UPS 3 Day Select<sup>SM</sup>"
				case "UPS Next Day Air Saver "
					TempStrService="UPS Next Day Air Saver&reg;"
				case "UPS Next Day Air Early A.M.&reg;"
					TempStrService="UPS Next Day Air Early A.M.&reg;"
				case "UPS Next Day Air Early A.M. "
					TempStrService="UPS Next Day Air&reg; Early A.M.&reg;"
				case "UPS 2nd Day Air A.M. "
					TempStrService="UPS 2nd Day Air A.M.&reg;"
				case "UPS Express Saver "
					TempStrService="UPS Express Saver <sup>SM</sup>"
				end select	
			End If
			pcStrShipmentDesc=TempStrService
		end if
		
		pcShipmentPriceToAdd=TempDblPostage

		if ubound(pcSplitShipping)=3 OR ubound(pcSplitShipping)>3 then
			pcDblServiceHandlingFee=pcSplitShipping(3)
			TempStrNewShipping=TempStrNewShipping&","&pcSplitShipping(3)
			if ubound(pcSplitShipping)=4 then
				pcDblIncHandlingFee=pcSplitShipping(4)
			else
				pcDblIncHandlingFee=0
			end if
		else
			pcDblServiceHandlingFee=0
			pcDblIncHandlingFee=0
		end if
	end if
End If
'END GET SHIPMENT DATA

'GET PAYMENT DATA
'SB S
strAndSub = ""
if pcIsSubscription = True Then
	strAndSub = " pcPayTypes_Subscription <> 0 "
else
	strAndSub =" idPayment=" & pidPayment
End if 
'SB E
if pidPayment<>0 and pidPayment<>"" and pidPayment<>999999 then
	'SB S
	query="SELECT paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE " & StrandSub
	'SB E
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		'set rs=nothing
		'call closedb()
		'response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	'This condition was checked by opc_checkpayment.asp
	'if rs.eof then
		'set rs=nothing
		'call closeDb() 
		'response.redirect "msg.asp?message=200"
	'end if
	
	pPaymentDesc=rs("paymentDesc")
	pPaymentPriceToAdd=rs("priceToAdd")
	pPaymentpercentageToAdd=rs("percentageToAdd")
	
	set rs=nothing
elseif pidPayment=0 then

	'SB S
	strAndSub = ""
	if pcIsSubscription = True Then
	   strAndSub = " AND pcPayTypes_Subscription = 1 ORDER by pcPayTypes_Subscription, paymentPriority"
	else
	   strAndSub = " ORDER by paymentPriority"
	End if 
	'SB E

	if session("customerType")=1 then
		query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE active=-1 AND (payTypes.pcPayTypes_PPAB <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
	else
		query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE active=-1 AND Cbtob=0 AND (payTypes.pcPayTypes_PPAB <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
	end if 
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		'set rs=nothing
		'call closedb()
		'response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	'This condition was checked by opc_checkpayment.asp
	'if rs.eof then
		'set rs=nothing
		'call closeDb() 
		'response.redirect "msg.asp?message=200"
	'end if
	
	If NOT rs.eof Then
		pidPayment=rs("idPayment")
		pPaymentDesc=rs("paymentDesc")
		pPaymentPriceToAdd=rs("priceToAdd")
		pPaymentpercentageToAdd=rs("percentageToAdd")
	End If
	set rs=nothing
	
	pcIdPayment=pidPayment
else
	if pidPayment=999999 then
		query="SELECT paymentDesc,priceToAdd,percentageToAdd FROM paytypes WHERE gwcode=46 OR gwcode=53 OR gwcode=999999;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
	
		if err.number<>0 then
			call LogErrorToDatabase()
			'set rs=nothing
			'call closedb()
			'response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		if rs.eof then
			pPaymentDesc="Paypal Express Checkout"
			pPaymentPriceToAdd=0
			pPaymentpercentageToAdd=0
		else
			pPaymentDesc=rs("paymentDesc")
			pPaymentPriceToAdd=rs("priceToAdd")
			pPaymentpercentageToAdd=rs("percentageToAdd")
		end if
		set rs=nothing
	end if
end if
'END GET PAYMENT DATA

intCalPaymnt=pSubTotal
' add payment amount
if ccur(pPaymentPriceToAdd)<>0 or ccur(pPaymentpercentageToAdd)<>0 then 
	tempPercentageToAdd=(pPaymentpercentageToAdd*intCalPaymnt/100)
	tempPercentageToAdd=roundTo(tempPercentageToAdd,.01)
	tempTaxPercentageToAdd=(pPaymentpercentageToAdd*pTaxableTotal/100)
	tempTaxPercentageToAdd=roundTo(tempTaxPercentageToAdd,.01)
	paymentTotal=pPaymentPriceToAdd + tempPercentageToAdd 
	taxPaymentTotal=pPaymentPriceToAdd + tempTaxPercentageToAdd 'processing fees on taxable total (only if percentage)           
end if

pSubTotal=pSubTotal + paymentTotal

' add shipment
if pcShipmentPriceToAdd>0 then 
	pcDblShipmentTotal=pcShipmentPriceToAdd     
else
	pcDblShipmentTotal=0
end if

'// Start Reward Points
If RewardsActive=1 And pcSFUseRewards<>"" Then 
	iDollarValue=pcSFUseRewards * (RewardsPercent / 100)
	if pSubTotal<>0 then
		pSubTotal=pSubTotal - iDollarValue
	else
		pSubTotal=0
	end if
	pTaxableTotal=pTaxableTotal-iDollarValue
	if pTaxableTotal<0 then
		pTaxableTotal=0
	end if
	if session("customerType")=1 AND ptaxwholesale=0 then
		pTaxableTotal=0
	end if
	if pSubTotal<-1 then
		xVar=(pSubTotal+iDollarValue)/(RewardsPercent/100)
		pcIntUseRewards=Round(xVar)
		pcSFUseRewards=pcIntUseRewards
		iDollarValue=pcSFUseRewards * (RewardsPercent / 100)
		pSubTotal=0
	end if
End If
'// End Reward Points
%>
	<table class="pcShowContent">
    <%
	'// SHOW/HIDE ADDRESSES
	IF pcIntHideAddresses=0 THEN
	%>
        <tr> 
            <td colspan="4">
                <table class="pcShowContent">
                    <tr>
                        <td width="70%">
                            <p>
                            <b><% response.write replace(scCompanyName,"''","'")%></b><br>
                            <% response.write replace(scCompanyAddress,"''","'")%><br>
                            <% response.write replace(scCompanyCity,"''","'")&", "&scCompanyState&" "&scCompanyZip&" - "& scCompanyCountry%>
                            </p>
                        </td>
                        <td width="30%" align="right">
                            <p>
                            <b><%=dictLanguage.Item(Session("language")&"_orderverify_6")%></b> 
                            <%=showDateFrmt(Date())	%>
                            </p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td colspan="4" class="pcSpacer"></td>
        </tr>
        <tr>
            <td colspan="4">
                <table class="pcShowContent">
                    <tr> 
                        <th colspan="2">
                            <p><%=dictLanguage.Item(Session("language")&"_orderverify_23")%></p>
                        </th>
                        <th>
                            <%if (pcShowShipAddr="1") AND session("gHideAddress")<>"1" then
                                response.write("<p>" & dictLanguage.Item(Session("language")&"_orderverify_24") & "</p>")
                            end if%>
                        </th>
                    </tr>
                    <tr> 
                        <td width="17%"><p><% response.write(replace(dictLanguage.Item(Session("language")&"_orderverify_7"),"''","'"))%></p></td>
                        <td width="41%"><p><% response.write(pcStrBillingFirstName&" "&pcStrBillingLastName) %></p></td>
                        <td width="42%"> 
                            <%if (pcShowShipAddr="1") AND session("gHideAddress")<>"1" then%>
                                <p><%=pcStrShippingFirstName&" "&pcStrShippingLastName %></p>
                            <% end if%>
                        </td>
                    </tr>
                    <%if pcStrBillingCompany<>"" OR pcStrShippingCompany <>"" then%>
                    <tr> 
                        <td><p><%=dictLanguage.Item(Session("language")&"_orderverify_8")%></p></td>
                        <td><p><%=pcStrBillingCompany%></p></td>
                        <td>
                            <% if (pcShowShipAddr="1") AND session("gHideAddress")<>"1" then
                                if pcStrShippingCompany<>"" then
                                response.write("<p>" & pcStrShippingCompany & "</p>")
                                end if
                            end if %>
                        </td>
                    </tr>
                    <%end if%>
                    <%if pcStrBillingEmail<>pcStrShippingEmail AND pcStrShippingEmail<>"" then%>
                    <tr> 
                        <td><p><%=dictLanguage.Item(Session("language")&"_opc_5")%></p></td>
                        <td><p><%=pcStrBillingEmail%></p></td>
                        <td>
                            <%if (pcShowShipAddr="1") AND session("gHideAddress")<>"1" then
                                response.write("<p>" & pcStrShippingEmail & "</p>")
                            end if %>
                        </td>
                    </tr>
                    <%end if%>
                    <tr> 
                        <td><p><%=dictLanguage.Item(Session("language")&"_orderverify_9")%></p></td>
                        <td><p><%=pcStrBillingPhone%></p></td>
                        <td>
                            <%if (pcShowShipAddr="1") AND session("gHideAddress")<>"1" then
                                response.write("<p>" & pcStrShippingPhone & "</p>")
                            end if %>
                        </td>
                    </tr>
                    <%if pcStrBillingFax<>"" OR pcStrShippingFax<>"" then%>
                    <tr> 
                        <td><p><%=dictLanguage.Item(Session("language")&"_opc_18")%></p></td>
                        <td><p><%=pcStrBillingFax%></p></td>
                        <td>
                            <%if (pcShowShipAddr="1") AND session("gHideAddress")<>"1" then
                                response.write("<p>" & pcStrShippingFax & "</p>")
                            end if %>
                        </td>
                    </tr>
                    <%end if%>
                    <tr> 
                        <td valign="top"><p><%=dictLanguage.Item(Session("language")&"_orderverify_10")%></p></td>
                        <td>
                            <p>
                            <%=pcStrBillingAddress%><br>
                            <%=pcStrBillingAddress2%><br>
                            <% response.write pcStrBillingCity
                            if pcStrBillingProvince="" then
                                response.write(", "&pcStrBillingStateCode)
                            else
                                response.write(", "&pcStrBillingProvince)
                            end if
                            response.write("&nbsp;"&pcStrBillingPostalCode&"<br>")
                            response.write pcStrBillingCountryCode											
                            %>
                            </p>
                        </td>
                        <td valign="top">
                        <%if (pcShowShipAddr="1") AND session("gHideAddress")<>"1" then %>
                            <p>
                            <%=pcStrShippingAddress %><br>
                            <%=pcStrShippingAddress2 %><br>
                            <% response.write pcStrShippingCity
                            If pcStrShippingProvince="" then
                                response.write(", "&pcStrShippingStateCode)
                            Else
                                response.write(", "&pcStrShippingProvince)
                            End If
                            response.write(" "&pcStrShippingPostalCode&"<br>")
                            response.write pcStrShippingCountryCode
                            %>
                            </p>
                        <% end if %>
                        </td>
                    </tr>						
                </table>
            </td>
        </tr>
        <%
		END IF
		' SHOW/HIDE ADDRESSES - END
						
				' ------------------------------------------------------
				'Start SDBA - Notify Drop-Shipping
				' ------------------------------------------------------
				if scShipNotifySeparate="1" and ppcCartIndex>1 then
					
					tmp_showmsg=0
					for f=1 to ppcCartIndex
						tmp_idproduct=pcCartArray(f,0)
						query="SELECT pcProd_IsDropShipped FROM products WHERE idproduct=" & tmp_idproduct & " AND pcProd_IsDropShipped=1;"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						
						if err.number<>0 then
							call LogErrorToDatabase()
							'set rs=nothing
							'call closedb()
							'response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
						
						if not rs.eof then
							tmp_showmsg=1
							exit for
						end if
						
						set rs=nothing
					next
					
					if tmp_showmsg=1 then%>
					<tr> 
						<td colspan="4" class="pcSpacer"></td>
					</tr>
					<tr>
						<td colspan="4">
							<div class="pcTextMessage"><img src="images/sds_boxes.gif" alt="<%response.write ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%>" align="left" vspace="5" hspace="10"><%response.write ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%></div>
						</td>
					</tr>
					<tr> 
						<td colspan="4" class="pcSpacer"></td>
					</tr>
					<%end if
					
				end if
				' ------------------------------------------------------
				'End SDBA - Notify Drop-Shipping
				' ------------------------------------------------------
						
				'START ORDER NAME
				' Display horizontal line, if needed
				If (savOrderNickName<>"" AND savOrderNickName<>"No Name") or savDF1<>"" or len(savOrderComments)>3 then %>
					<tr> 
						<td colspan="4"><hr></td>
					</tr>
				<% End If %>
					
				<% ' Show order name, if any
				If savOrderNickName<>"" AND savOrderNickName<>"No Name" then %>
					<tr>
						<td colspan="4">
						<p><b><%=dictLanguage.Item(Session("language")&"_CustviewOrd_40")%></b>&nbsp;
						<%=savOrderNickName%></p></td>
					</tr>
				<% End If
				'END ORDER NAME

				' START Show order delivery date/time, if any
				If savDF1<>"" then %>
					<tr>
						<td colspan="4">
							<p><b><%=dictLanguage.Item(Session("language")&"_orderverify_34")%></b>&nbsp;
							<%=showDateFrmt(savDF1)%> <%If savTF1<>"" then%>&nbsp;&nbsp;-&nbsp;&nbsp;<%= savTF1 %><% End If %></p>
						</td>
					</tr>
				<% End If
				' END Show order delivery date/time, if any					
					
				' START Show order comments, if any
				if len(savOrderComments)>3 then %>
					<tr> 
						<td colspan="2">
							<p><b><%=dictLanguage.Item(Session("language")&"_orderverify_11")%></b>&nbsp;
							<%=savOrderComments%></p>
						</td>
					</tr>
				<% end if
				' END Show order comments, if any %>
						
				<tr> 
					<td colspan="4" class="pcSpacer"></td>
				</tr>					
				<tr> 
					<th width="4%"><p><%=dictLanguage.Item(Session("language")&"_orderverify_25")%></p></th>
					<th width="62%"><p><%=dictLanguage.Item(Session("language")&"_orderverify_27")%></p></th>
					<th width="12%" nowrap align="right"><p><%=dictLanguage.Item(Session("language")&"_orderverify_32")%></p></th>
					<th width="12%" nowrap align="right"><p><%=dictLanguage.Item(Session("language")&"_orderverify_28")%></p></th>
				</tr>
				
				<% 'START GET PRODUCTS ORDERING
				strBundleArray=""
				pSFstrBundleArray=""
				Dim pcProductList(100,5)
				for f=1 to ppcCartIndex
					pcProductList(f,0)=pcCartArray(f,0)
					pcProductList(f,1)=pcCartArray(f,10)
					pcProductList(f,3)=pcCartArray(f,2)
					pcProductList(f,4)=0
					
					'SB S
					if (pcCartArray(f,38)) > 0  then
						'// Get the Sub Details
						pSubscriptionID = (pcCartArray(f,38)) %>				
						<!--#include file="../includes/pcSBDataInc.asp" --> 	
				  	<% end if 
					'SB E
					
					if pcCartArray(f,10)=0 then
							
						'BTO ADDON-S
						pBTOValues=0
						if trim(pcCartArray(f,16))<>"" then 
							
							query="SELECT stringProducts, stringValues, stringCategories,stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
							
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rs=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							
							stringProducts=rs("stringProducts")
							stringValues=rs("stringValues")
							stringCategories=rs("stringCategories")
							ArrProduct=Split(stringProducts, ",")
							ArrValue=Split(stringValues, ",")
							ArrCategory=Split(stringCategories, ",")
							Qstring=rs("stringQuantity")
							ArrQuantity=Split(Qstring,",")
							Pstring=rs("stringPrice")
							ArrPrice=split(Pstring,",")
							set rs=nothing
							
							
							if ArrProduct(0)="na" then
							else
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
									query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
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
									pBTOValues=pBTOValues+ccur((ArrValue(i)+UPrice)*pcCartArray(f,2))
									end if
									set rsObj=nothing
								next
							end if						
						End if
						'BTO ADDON-E
					End if
											
					if pcCartArray(f,10)=0 then
						
						if pcv_IsEUMemberState=0 then
							tmpRowPrice=ccur( pcCartArray(f,2) * pcCartArray(f,17) )
						end if

						pRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))
						pExtRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))-ccur(pBTOvalues) %>
						<% 'Validate for multiple of N
						query="SELECT pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty FROM products WHERE idproduct=" & pcCartArray(f,0)
						set rs=server.CreateObject("ADODB.RecordSet") 									
						set rs=connTemp.execute(query)
								
						if err.number<>0 then
							call LogErrorToDatabase()
							'set rs=nothing
							'call closedb()
							'response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
								
						pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
						if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
							pcv_intHideBTOPrice="0"
						end if
						pcv_intQtyValidate=rs("pcprod_QtyValidate")
						if isNULL(pcv_intQtyValidate) OR pcv_intQtyValidate="" then
							pcv_intQtyValidate="0"
						end if				
						pcv_lngMinimumQty=rs("pcprod_MinimumQty")
						if isNULL(pcv_lngMinimumQty) OR pcv_lngMinimumQty="" then
							pcv_lngMinimumQty="0"
						end if
						set rs=nothing 
						
						%>
						<tr valign="top"> 
							<td>
								<p><%=pcCartArray(f,2)%></p>
							</td>
							<td>
								<p><%=pcCartArray(f,1) %>&nbsp;<span class="opcSku">(<%=pcCartArray(f,7)%>)</span></p>
							</td>
							<td align="right">
							<% if pcv_intHideBTOPrice<>"1" then
								if pcCartArray(f,17) > 0 then %>
									<%=scCurSign & money(pcCartArray(f,17)-ccur(ccur(pBTOvalues)/pcCartArray(f,2)))%>
								<% 	end if
							end if %>
							</td>
							<td align="right" nowrap>
								<p><% if pExtRowPrice > 0 then response.write(scCurSign & money(pExtRowPrice)) end if %></p>
							</td>
						</tr>
					
						<% 'BTO ADDON-S
						if trim(pcCartArray(f,16))<>"" then
							
							query="SELECT stringProducts,stringValues,stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
							
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rs=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
						
							stringProducts=rs("stringProducts")
							stringValues=rs("stringValues")
							stringCategories=rs("stringCategories")
							ArrProduct=Split(stringProducts, ",")
							ArrValue=Split(stringValues, ",")
							ArrCategory=Split(stringCategories, ",")
							Qstring=rs("stringQuantity")
							ArrQuantity=Split(Qstring,",")
							Pstring=rs("stringPrice")
							ArrPrice=split(Pstring,",")
							set rs=nothing
							%>
								
							<tr> 
								<td>&nbsp;</td>
								<td colspan="3" class="pcShowBTOconfiguration"> 
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr> 
											<td colspan="2"><p><%=bto_dictLanguage.Item(Session("language")&"_viewcart_1")%></p></td>
										</tr>
                      
										<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
											query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))"
											set rs=server.CreateObject("ADODB.RecordSet") 
											set rs=conntemp.execute(query)
											
											if err.number<>0 then
												call LogErrorToDatabase()
												'set rs=nothing
												'call closedb()
												'response.redirect "techErr.asp?err="&pcStrCustRefID
											end if
											
											strCategoryDesc=rs("categoryDesc")
											strDescription=rs("description")
											set rs=nothing
													
											query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i)&" and specProduct=" & pcCartArray(f,0) 
											set rs=server.CreateObject("ADODB.RecordSet") 
											set rs=conntemp.execute(query)
														
											if err.number<>0 then
												call LogErrorToDatabase()
												'set rs=nothing
												'call closedb()
												'response.redirect "techErr.asp?err="&pcStrCustRefID
											end if
												
											btDisplayQF=rs("displayQF")
											set rs=nothing
											
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
											query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
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
											set rsQ=nothing %>
											<tr> 
												<td width="85%" valign="top">
													<p><%=strCategoryDesc%>:&nbsp;
													<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%>(<%=ArrQuantity(i)%>)&nbsp;<%end if%>
													<%=strDescription%>
													</p>
												</td>
												<td width="15%" valign="top">
													<p align="right">
													<%if (ArrValue(i)<>"") and (ArrQuantity(i)<>"") and (ArrPrice(i)<>"") then 
														if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
															if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
																if tmpDefault=1 then
																	UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
																else
																	UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
																end if
															else
																UPrice=0
															end if %>
															<%=scCurSign & money(ccur((ArrValue(i)+UPrice)*pcCartArray(f,2)))%>
														<%else
															if tmpDefault=1 then%>
																<%=dictLanguage.Item(Session("language")&"_defaultnotice_1")%>
															<%end if
														end if
													end if%>
													</p>
												</td>
											</tr>
										 <% next %>
									</table>
								</td>
							</tr>
						<% End if 
						'BTO ADDON-E
								
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: SHOW PRODUCT OPTIONS
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

						if trim(pcCartArray(f,4))<>"" then
						
							Dim pcv_strOptionsArray, pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice, tAprice
							Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions

							pcv_strOptionsArray = trim(pcCartArray(f,4))
						
							if len(pcv_strOptionsArray)>0 then %>
								<tr valign="top">
									<td>&nbsp;</td>
									<td colspan="2">
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<%
										'#####################
										' START LOOP
										'#####################	
										
										'// Generate Our Local Arrays from our Stored Arrays  
										
										' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
										pcArray_strSelectedOptions = ""					
										pcArray_strSelectedOptions = Split(trim(pcCartArray(f,11)),chr(124))
										
										' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
										pcArray_strOptionsPrice = ""
										pcArray_strOptionsPrice = Split(trim(pcCartArray(f,25)),chr(124))
										
										' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
										pcArray_strOptions = ""
										pcArray_strOptions = Split(trim(pcv_strOptionsArray),chr(124))
										
										' Get Our Loop Size
										pcv_intOptionLoopSize = 0
										pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
										
										' Start in Position One
										pcv_intOptionLoopCounter = 0
										
										' Display Our Options
										For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize %>
											<tr>
												<td width="67%"><p><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></p></td>
												<td align="right" width="33%">									
												<% tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
												
												if tempPrice="" or tempPrice=0 then
													response.write "&nbsp;"
												else %>
													<table width="100%" cellpadding="0" cellspacing="0" border="0">
														<tr>
															<td align="left" width="60%">
																<%=scCurSign&money(tempPrice)%>
															</td>
															<td align="right" width="40%">
																<%									
																tAprice=(tempPrice*ccur(pcCartArray(f,2)))
																response.write scCurSign&money(tAprice) 
																%>
															</td>
														</tr>
													</table>
												<% end if %>			
												</td>
											</tr>
										<% Next
										'#####################
										' END LOOP
										'#####################	
									
										%>
										</table>
									</td>
									<td>&nbsp;</td>
								</tr>															
							<% end if
						end if
								
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: SHOW PRODUCT OPTIONS
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
						pRowPrice=pRowPrice + ccur(pcCartArray(f,2) * pcCartArray(f,5)) %>
								
						<% if trim(pcCartArray(f,21))<>"" then %>
							<tr> 
								<td>&nbsp;</td>
								<td colspan="2"><p><% response.write(replace(pcCartArray(f,21),"''","'"))%></p></td>
								<td>&nbsp;</td>
							</tr>
						<%end if %>
							
						<% 'if items quantities discounts apply to this product, show the total applied amount here
						if trim(pcCartArray(f,16))<>"" then
							if ccur(pcCartArray(f,30))>0 then
								pRowPrice=pRowPrice-ccur(pcCartArray(f,30)) %>
								<tr> 							
									<td>&nbsp;</td>
									<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_showcart_23")%></p></td>
									<td nowrap align="right">
										<p>- 
										<% response.write scCurSign &  money(ccur(pcCartArray(f,30))) %>
										</p>
									</td>
								</tr>
							<% end if
						End if%>
							
						<% 'BTO Additional Charges
						if trim(pcCartArray(f,16))<>"" then
							query="SELECT stringCProducts,stringCValues,stringCCategories FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
									
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rs=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							
							stringCProducts=rs("stringCProducts")
							stringCValues=rs("stringCValues")
							stringCCategories=rs("stringCCategories")
							ArrCProduct=Split(stringCProducts, ",")
							ArrCValue=Split(stringCValues, ",")
							ArrCCategory=Split(stringCCategories, ",")
							set rs=nothing
									
							if ArrCProduct(0)<>"na" then
								pRowPrice=pRowPrice+ccur(pcCartArray(f,31))%>
								<tr> 
									<td>&nbsp;</td>
									<td colspan="3" valign="top" class="pcShowBTOconfiguration"> 
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr> 
												<td><p><b><%=bto_dictLanguage.Item(Session("language")&"_viewcart_3")%></b></p></td>
												<td></td>
											</tr>
											<% for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
												query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))"
												set rs=server.CreateObject("ADODB.RecordSet") 
												set rs=conntemp.execute(query)
												
												if err.number<>0 then
													call LogErrorToDatabase()
													'set rs=nothing
													'call closedb()
													'response.redirect "techErr.asp?err="&pcStrCustRefID
												end if
												
												strCategoryDesc=rs("categoryDesc") 
												strDescription=rs("description") 
												set rs=nothing %>
												<tr> 
													<td width="85%" valign="top">
														<p><%=strCategoryDesc%>:&nbsp;<%=strDescription%></p>
													</td>
													<td width="15%" align="right" valign="top">
													<p> 
													<%if (ccur(ArrCValue(i))>0)then %>
														<%=scCurSign & money(ArrCValue(i))%>
													<%end if%>
													</p>
													</td>
												</tr>
											<% next %>
										</table>
									</td>
								</tr>
							<% End if
							'Have Charges 
							
						End if 
						'BTO Additional Charges %>
														
						<% 'if quantity discounts apply to this product, show the total applied amount here
						if trim(pcCartArray(f,15))<>"" AND trim(pcCartArray(f,15))>0 then
							pRowPrice=pRowPrice-ccur(pcCartArray(f,15)) %>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2">
									<%=dictLanguage.Item(Session("language")&"_showcart_20")%>
									<%=dictLanguage.Item(Session("language")&"_showcart_20b")%>
								</td>
								<td nowrap align="right">
									<p>-<% response.write scCurSign & money(pcCartArray(f,15)) %></p>
								</td>
							</tr>
						<% End if 
								
						if pExtRowPrice<>pRowPrice then %>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right"><p><%=dictLanguage.Item(Session("language")&"_showcart_22")%></td>
								<td nowrap align="right"><p><%=scCurSign & money(pRowPrice) %></p></td>
							</tr>
						<% end if %>
						
						<% 
						'SB S
						if (pcCartArray(f,38)) > 0  then
						
					 		'// Get the data 
					  		pSubscriptionID = (pcCartArray(f,38)) 

                            '// If there's a trial set the line total to the trial price
                            if pcv_intIsTrial = "1" Then
                            	pRowPrice = "8" '// pcv_curTrialAmount
                            else
                            	pRowPrice = "8" '// pExtRowPrice
                            end if 
							 
						end if 
						'SB E 
						%>
						
						<% 'START 10th Row - Cross Sell Bundle Discount %>	
						<% if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then 	%>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right">
								<%=dictLanguage.Item(Session("language")&"_showcart_26")%>
								</td>
								<td align="right">
								<% =scCurSign &  money( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) %>
								</td>
							</tr>
							<% strBundleArray=strBundleArray&pcCartArray(f,0)&","&pcCartArray(f,27)&","&pcCartArray(f,28)&","&((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)&"||"
						end if %>
						<% 'END 10th Row - Cross Sell Bundle Discount %>	
						
						<% 'START 11th Row - Cross Sell Bundle Subtotal %>
						<% if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then 
						    pRowPrice = ( ccur(pRowPrice) + ccur(pcProductList(cint(pcCartArray(f,27)),2)) ) - ( ( ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28) ) ) * pcCartArray(f,2) )%>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right">
								<%= dictLanguage.Item(Session("language")&"_showcart_22")%>
								</td>
								<td align="right">
								<%= scCurSign &  money(pRowPrice) %>
								</td>
							</tr>
						<% end if %>
						<% 'END 11th Row - Cross Sell Bundle Subtotal %>	


						<% 'GGG Add-on start
						if Session("Cust_GW")="1" then
							
							GWmsg="<u>" & dictLanguage.Item(Session("language")&"_orderverify_36a") & "</u>: "
							gIDPro=pcCartArray(f,0)
							gMode=1
							query="select pcPE_IDProduct from pcProductsExc where pcPE_IDProduct=" & gIDPro
							set rsG=server.CreateObject("ADODB.RecordSet")
							set rsG=connTemp.execute(query)
							
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rsG=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
										
							if not rsG.eof then
								GWmsg=GWmsg & dictLanguage.Item(Session("language")&"_orderverify_38a")
								gMode=0
							else
								if (pcCartArray(f,34)="") or (pcCartArray(f,34)="0") then
									GWmsg=GWmsg & dictLanguage.Item(Session("language")&"_orderverify_37a")
									gMode=1
								else
									gIDOpt=pcCartArray(f,34)
									query="select pcGW_OptName,pcGW_OptPrice from pcGWOptions where pcGW_IDOpt=" & gIDOpt
									set rsG=server.CreateObject("ADODB.RecordSet")
									set rsG=connTemp.execute(query)
									
									if err.number<>0 then
										call LogErrorToDatabase()
										'set rsG=nothing
										'call closedb()
										'response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
												
									if NOT rsG.eof then
										pcv_strOptName = rsG("pcGW_OptName")
										pcv_strOptPrice = rsG("pcGW_OptPrice")
										GWmsg=GWmsg & pcv_strOptName & " - " & scCurSign & money(pcv_strOptPrice)
										GiftWrapPaymentTotal=GiftWrapPaymentTotal+pcv_strOptPrice
									end if 

									gMode=1
								end if
							end if %>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="3"><p><%=GWmsg%>&nbsp;</p></td>
							</tr> 
						<%end if
						'GGG end%>
						<tr> 
							<td colspan="4"><hr></td>
						</tr>
					<% end if %>
					
					<% 
					if pcv_IsEUMemberState = 0 then
						pcProductList(f,2) = tmpRowPrice
					else
						pcProductList(f,2) = pRowPrice
					end if
				next
				pSFstrBundleArray=strBundleArray %>
				
				<%'Product Promotions
				TotalPromotions=0
				if Session("pcPromoIndex")<>"" and Session("pcPromoIndex")>"0" then
					PromoArr1=Session("pcPromoSession")
					PromoIndex=Session("pcPromoIndex")
					For m=1 to PromoIndex
						TotalPromotions=TotalPromotions+cdbl(PromoArr1(m,2))
					Next
				end if
				pSubTotal=pSubTotal-TotalPromotions
				Session("PromotionTotal")=TotalPromotions
				%>

				<%
				'// Discounts by Categories
				Dim pcv_strApplicableProducts
				pcv_strApplicableProducts=""
				CatDiscTotal=0

				query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)								
				if err.number<>0 then
					call LogErrorToDatabase()
				end if

				Do While not rs.eof
					CatSubQty=0
					CatSubTotal=0
					CatSubDiscount=0
					ApplicableCategoryID = rs("IDCat")
					CanNotRun=0
					IDCat=rs("IDCat")
					
					query="SELECT categories_products.idcategory FROM categories_products INNER JOIN pcPrdPromotions ON categories_products.idproduct=pcPrdPromotions.idproduct WHERE categories_products.idcategory=" & IDCat & ";"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						CanNotRun=1
					end if
					set rsQ=nothing
					
					IF CanNotRun=0 THEN
							
						For f=1 to ppcCartIndex
							if (pcProductList(f,1)=0) and (pcProductList(f,4)=0) then 
								
								query="select idproduct from categories_products where idcategory=" & IDCat & " and idproduct=" & pcProductList(f,0)
								set rstemp=server.CreateObject("ADODB.RecordSet")
								set rstemp=connTemp.execute(query)							
								if err.number<>0 then
									call LogErrorToDatabase()
								end if								
								if not rstemp.eof then
									CatSubQty=CatSubQty+pcProductList(f,3)
									CatSubTotal=CatSubTotal+pcProductList(f,2)
									pcProductList(f,4)=1
									pcv_strApplicableProducts = pcv_strApplicableProducts & pcProductList(f,0) & chr(124) &  ApplicableCategoryID & ","								
								end if
								set rstemp=nothing
								
							end if
							
						Next
						
						pcv_strrApplicableCategories = pcv_strrApplicableCategories & CatSubTotal & chr(124) &  ApplicableCategoryID & ","
						
						if CatSubQty>0 then
	
							query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & IDCat & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
							set rstemp=server.CreateObject("ADODB.RecordSet")
							set rstemp=conntemp.execute(query)								
							if err.number<>0 then
								call LogErrorToDatabase()
							end if								
							if not rstemp.eof then
								'// There are quantity discounts defined for that quantity 
								pDiscountPerUnit=rstemp("pcCD_discountPerUnit")
								pDiscountPerWUnit=rstemp("pcCD_discountPerWUnit")
								pPercentage=rstemp("pcCD_percentage")
								pbaseproductonly=rstemp("pcCD_baseproductonly")
								if session("customerType")<>1 then  'customer is a normal user
									if pPercentage="0" then 
										CatSubDiscount=pDiscountPerUnit*CatSubQty
									else
										CatSubDiscount=(pDiscountPerUnit/100) * CatSubTotal
									end if
								else  'customer is a wholesale customer
									if pPercentage="0" then 
										CatSubDiscount=pDiscountPerWUnit*CatSubQty
									else
										CatSubDiscount=(pDiscountPerWUnit/100) * CatSubTotal
									end if
								end if							
							end if						
							set rstemp=nothing		
											
						end if '// if CatSubQty>0 then
	
						CatDiscTotal=CatDiscTotal+CatSubDiscount
					
					END IF 'CanNotRun
					
					rs.MoveNext
				loop
				set rs=nothing

				'// Round the Category Discount to two decimals
				if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
					CatDiscTotal = RoundTo(CatDiscTotal,.01)
				end if
				
				pSubTotal=pSubTotal-CatDiscTotal
				pSFCatDiscTotal=CatDiscTotal

				'//////////////////////////////////////////////////////////////
				' START - Discounts by code
				'//////////////////////////////////////////////////////////////
				
				pcGlobalDiscError=Cstr("")
				pDiscountError=Cstr("")
				pGCError=Cstr("")
				pDiscountShowCode=Cstr("")
				discountTotal=ccur(0)
				passDiscountCnt=-1
				noCode=""
				intCodeCnt=-1
				pTempDiscountCode=""
				intGCCnt=-1
				pTempGC=""
				pCodeTotal=""
				
				if pDiscountCode="" then
					noCode="1"
				end if
				
				IF noCode="" THEN
					'****************************************************
					' START - Split out GCs and Discount Codes
					'****************************************************
					
					DiscountCodeArry=Split(pDiscountCode,",")
					for i=0 to ubound(DiscountCodeArry)
						if DiscountCodeArry(i)<>"" then
							
							query="SELECT pcGCOrdered.pcGO_ExpDate, pcGCOrdered.pcGO_Amount, pcGCOrdered.pcGO_Status, products.Description FROM pcGCOrdered, products WHERE pcGCOrdered.pcGO_GcCode='"&DiscountCodeArry(i)&"' AND products.idproduct=pcGCOrdered.pcGO_IDProduct"
							set rsQ=server.CreateObject("ADODB.RecordSet")
							set rsQ=conntemp.execute(query)
							if not rsQ.eof then
								if pTempGC<>"" then
									pTempGC=pTempGC & ","
								end if
								pTempGC=pTempGC & DiscountCodeArry(i)
							else
								if pTempDiscountCode<>"" then
									pTempDiscountCode=pTempDiscountCode & ","
								end if
								pTempDiscountCode=pTempDiscountCode & DiscountCodeArry(i)
							end if
							set rsQ=nothing
							
						end if
					next
					
					pDiscountCode=pTempDiscountCode

					pCodeTotal=pTempDiscountCode
					if pTempGC<>"" then
						if pCodeTotal<>"" then
							pCodeTotal=pCodeTotal & ","
						end if
						pCodeTotal=pCodeTotal & pTempGC
					end if
						
					if displayDiscountCode<>pCodeTotal then
						displayDiscountCode=pCodeTotal
						session("DCODE")=displayDiscountCode
					end if
					
					'****************************************************
					' END - Split out GCs and Discount Codes
					'****************************************************
				END IF 'noCode=""
				
				'****************************************************
				' START - Check Discount codes
				'****************************************************
				dim UsedDiscountCodes, intArryCnt
					UsedDiscountCodes=""
					intArryCnt=0
					
				'set filter variables 
				CatCount=1
				CatFound=0
				
				IF pDiscountCode<>"" THEN
					DiscountTableRow=""
					
					DiscountCodeArry=Split(pDiscountCode,",")
					intCodeCnt=ubound(DiscountCodeArry)
					
					DiscountCodeArryO=Split(pDiscountCode,",")
					intCodeCntO=ubound(DiscountCodeArry)
					
					'ORDER to check Discount codes
					IF (pDiscountCode<>"") AND (InStr(pDiscountCode,",")>0) THEN
						tmpDC=""
						For i=0 to intCodeCnt
							if trim(DiscountCodeArry(i))<>"" then
								if tmpDC<>"" then
									tmpDC=tmpDC & ","
								end if
								tmpDC=tmpDC & "'" & trim(DiscountCodeArry(i)) & "'"
							end if
						Next
						if tmpDC<>"" then
							query="SELECT discountcode FROM discounts WHERE discountcode IN (" & tmpDC & ") ORDER BY pcDisc_Auto DESC,pcSeparate DESC;"
							set rsQ=connTemp.execute(query)
							if not rsQ.eof then
								pDiscountCode=""
								tmpDCArr=rsQ.getRows()
								intCountD=ubound(tmpDCArr,2)
								For i=0 to intCountD
									if pDiscountCode<>"" then
										pDiscountCode=pDiscountCode & ","
									end if
									pDiscountCode=pDiscountCode & tmpDCArr(0,i)
								Next
								DiscountCodeArry=Split(pDiscountCode,",")
								intCodeCnt=ubound(DiscountCodeArry)
							end if
							set rsQ=nothing
							if pDiscountCode="" then
								intCodeCnt=-1
							end if
						end if
					END IF
					
					'Check Invalid Discount Codes
					Dim FoundInArr
					For ik=0 to intCodeCntO
						IF trim(DiscountCodeArryO(ik))<>"" then
							FoundInArr=0
							For i=0 to intCodeCnt
								if trim(ucase(DiscountCodeArryO(ik)))=trim(ucase(DiscountCodeArry(i))) then
									FoundInArr=1
								end if
							Next
							if FoundInArr=0 then
								pcGlobalDiscError=pcGlobalDiscError & "<li>" & dictLanguage.Item(Session("language")&"_orderverify_4") & " (<b>"&DiscountCodeArryO(ik)&"</b>)</li>"
							end if
						END IF
					Next
					
					pcv_HaveSeparateCode=0

						For i=0 to intCodeCnt
							pcv_Filters=0
							pcv_FResults=0
							pcv_ProTotal=0
							
							IF trim(DiscountCodeArry(i))<>"" THEN

							pTempDiscCode=DiscountCodeArry(i)
							Session("DiscountTotal"&pTempDiscCode)=0
							Session("DiscountRow"&pTempDiscCode)=""
							
							'see if discount code has already been used for this store
							intDiscMatchFound=0
							
							if UsedDiscountCodes<>"" then
								UsedDiscountCodeArry=split(UsedDiscountCodes,",")
								for t=0 to (ubound(UsedDiscountCodeArry)-1)
									if pTempDiscCode=UsedDiscountCodeArry(t) then
										intDiscMatchFound=1
										pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_40") 
									end if
								next
							end if
							
							if intDiscMatchFound=0 then
								UsedDiscountCodes=UsedDiscountCodes&pTempDiscCode&","
							end if
							
							query="SELECT iddiscount, onetime,expDate, idProduct, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, DiscountDesc, priceToDiscount, percentageToDiscount, pcSeparate, pcDisc_Auto, pcDisc_StartDate, pcRetailFlag, pcWholesaleFlag, pcDisc_PerToFlatCartTotal, pcDisc_PerToFlatDiscount,pcDisc_IncExcPrd,pcDisc_IncExcCat,pcDisc_IncExcCust,pcDisc_IncExcCPrice FROM discounts WHERE discountcode='" &pTempDiscCode& "' AND active=-1;"
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=connTemp.execute(query)
	
							if rs.eof then
								pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_4")
								pDiscountDesc=""
							else
								pcv_IDDiscount=rs("iddiscount")
								pcv_IDDiscount1=rs("iddiscount")
								pcv_OneTime=rs("onetime")
								expDate=rs("expDate")
								dcIdProduct=rs("idProduct")
								dcQuantityFrom=rs("quantityFrom")
								dcQuantityUntil=rs("quantityUntil")
								dcWeightFrom=rs("weightFrom")
								dcWeightUntil=rs("weightUntil")
								dcPriceFrom=rs("priceFrom")
								dcPriceUntil=rs("priceUntil")
								pDiscountDesc=rs("DiscountDesc")
								pPriceToDiscount=ccur(rs("priceToDiscount"))
								ppercentageToDiscount=rs("percentageToDiscount")
								intPcSeparate=rs("pcSeparate")
								intPcAuto=rs("pcDisc_Auto")
								pcv_startDate=rs("pcDisc_StartDate")
								pcv_retail = rs("pcRetailFlag")
								pcv_wholeSale = rs("pcWholeSaleFlag")
								pcv_PerToFlatCartTotal = rs("pcDisc_PerToFlatCartTotal")
								pcv_PerToFlatDiscount = rs("pcDisc_PerToFlatDiscount")
								pcIncExcPrd=rs("pcDisc_IncExcPrd")
								pcIncExcCat=rs("pcDisc_IncExcCat")
								pcIncExcCust=rs("pcDisc_IncExcCust")
								pcIncExcCPrice=rs("pcDisc_IncExcCPrice")

								if intPcSeparate="" OR IsNull(intPcSeparate) then
								else
									if clng(intPcSeparate)=0 then
										pcv_HaveSeparateCode=1
									end if
								end if
								
								if (clng(pcv_HaveSeparateCode)=1) AND (passDiscountCode<>"") then
									pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_39")
								end if
								
								'check to see if discount has been used for one use only for this customer specified
								If pcv_OneTime<>0 Then
									
									'check used discounts in database with iddiscount
									query="SELECT * FROM used_discounts WHERE idcustomer="&session("IDCustomer")&" AND iddiscount=" & pcv_IDDiscount1
									set rsCheck=server.CreateObject("ADODB.RecordSet")
									set rsCheck=connTemp.execute(query)									
									if err.number<>0 then
										call LogErrorToDatabase()
									end if
									
									varOneTimePresent=0
									if NOT rsCheck.eof then
										'discount has been used already by the customer
										pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
										varOneTimePresent=1
									end if
									set rsCheck=nothing
									
									If expDate<>"" then
										If datediff("d", Now(), expDate) <= 0 Then
											if varOneTimePresent=0 then
												pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
											end if
										end if
									end if
									
									'check to see if discount has start date
									If pcv_startDate<>"" then
										StartDate=pcv_startDate
										If datediff("d", Now(), StartDate) > 0 Then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_43")
										End If
									end if
								Else
									'check to see if discount code has expired
									If expDate<>"" then
										If datediff("d", Now(), expDate) <= 0 Then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_21")
										end if
									end if
									
									'check to see if discount has start date
									If pcv_startDate<>"" then
										StartDate=pcv_startDate
										If datediff("d", Now(), StartDate) > 0 Then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_43")
										End If
									end if
								end if

								Dim pcv_dblSubTotalAdjusted
								pcv_dblSubTotalAdjusted = pSubTotal - paymentTotal - discountTotal
								If pcv_dblSubTotalAdjusted<0 Then
									pcv_dblSubTotalAdjusted=0
								End If
								
								If pDiscountError="" Then
									if Int(pCartQuantity)>=Int(dcQuantityFrom) and Int(pCartQuantity)<=Int(dcQuantityUntil) and Int(pCartTotalWeight)>=Int(dcWeightFrom) and Int(pCartTotalWeight)<=Int(dcWeightUntil) and ccur(pcv_dblSubTotalAdjusted)>=ccur(dcPriceFrom) and ccur(pcv_dblSubTotalAdjusted)<=ccur(dcPriceUntil) then

									else
									
										if NOT (Int(pCartQuantity)>=Int(dcQuantityFrom) and Int(pCartQuantity)<=Int(dcQuantityUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5b")
										elseif NOT (Int(pCartTotalWeight)>=Int(dcWeightFrom) and Int(pCartTotalWeight)<=Int(dcWeightUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5a")
										elseif NOT (ccur(pSubTotal)>=ccur(dcPriceFrom) and ccur(pSubTotal)<=ccur(dcPriceUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5c")
										else
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5")
										end if
										
									end if
								End If
							
							end if
							set rs=nothing


							IF pcv_IDDiscount<>"" AND pDiscountError="" THEN
								
								'// START: Filter by Products
								pcv_ProductFilter = 0
								query="select pcFPro_IDProduct from PcDFProds where pcFPro_IDDiscount=" & pcv_IDDiscount1
								set rs=server.CreateObject("ADODB.RecordSet")	
								set rs=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
								end if
								if not rs.eof then									
									pcv_Filters=pcv_Filters+1
									tmpIDArr=rs.getRows()									
									intIDCount=ubound(tmpIDArr,2)
									pcv_ProductFilter = 1
								end if
								set rs=nothing
								
								If pcv_ProductFilter=1 Then
								
									for f=1 to ppcCartIndex
										if pcProductList(f,1)=0 then
											tmpgotit=0
											for ik=0 to intIDCount
												if clng(pcProductList(f,0))=clng(tmpIDArr(0,ik)) then
													tmpgotit=1
													exit for
												end if
											next
											if (pcIncExcPrd="0") AND (tmpgotit=1) then
												pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
												pcv_FResults=1
											else
												if (pcIncExcPrd="1") AND (tmpgotit=0) then
													pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
													pcv_FResults=1
												end if
											end if
										end if
									next '// for f=1 to ppcCartIndex
									
								End If '// If intIDCount>0 Then								
								'// END: Filter by Products


								'// START: Filter by Categories
								If pcv_Filters=0 Then
									
									pcv_CatFilter = 0
									query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount1
									set rs=server.CreateObject("ADODB.RecordSet")
									set rs=connTemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
									end if						
									if not rs.eof then
										pcv_CatFilter = 1
									end if 
									set rs=nothing 
									
									If pcv_CatFilter=1 Then
										
										pcv_Filters=pcv_Filters+1
										
										for f=1 to ppcCartIndex
											
											if pcProductList(f,1)=0 then
												
												query="SELECT idcategory FROM categories_products WHERE idproduct=" & pcProductList(f,0)
												set rs2=server.CreateObject("ADODB.RecordSet")
												set rs2=connTemp.execute(query)
												if err.number<>0 then
													call LogErrorToDatabase()
												end if
												intCatCount=-1
												if not rs2.eof then                                                	
													tmpCatArr=rs2.getRows()
                                                    intCatCount=ubound(tmpCatArr,2)
                                                    tmpgotit=0													
												end if
                                                set rs2=nothing
												
												If intCatCount>=0 Then
												
													'Check assigned categories
                                                    For ik=0 to intCatCount
													
														pcv_IDCat=tmpCatArr(o,ik)
														query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount1 & " and pcFCat_IDCategory=" & pcv_IDCat
														set rstemp=server.CreateObject("ADODB.RecordSet")
														set rstemp=connTemp.execute(query)
														if err.number<>0 then
															call LogErrorToDatabase()
														end if
														if not rstemp.eof then															
                                                        	set rstemp=nothing
															tmpgotit=1
                                                            exit for
														end if
                                                        set rstemp=nothing
														
                                                        'Check parent-categories
                                                        if (tmpgotit=0) AND (pcv_IDCat<>"1") then
                                                        	pcv_ParentIDCat=pcv_IDCat
															do while (tmpgotit=0) and (pcv_ParentIDCat<>"1")
																
																query="select idParentCategory from categories where idcategory=" & pcv_ParentIDCat
																set rstemp=server.CreateObject("ADODB.RecordSet")
																set rstemp=connTemp.execute(query)
																if err.number<>0 then
																	call LogErrorToDatabase()
																end if														
																if not rstemp.eof then																	
																	pcv_ParentIDCat=rstemp("idParentCategory")
																	if pcv_ParentIDCat<>"1" then
																		
																		query="select pcFCat_IDCategory from PcDFCats where pcFCat_IDDiscount=" & pcv_IDDiscount1 & " and pcFCat_IDCategory=" & pcv_ParentIDCat & " and pcFCat_SubCats=1;"
																		set rsFCat=server.CreateObject("ADODB.RecordSet")
																		set rsFCat=connTemp.execute(query)
																		if err.number<>0 then
																			call LogErrorToDatabase()
																		end if
																		if not rsFCat.eof then
																			tmpgotit=1
																		end if
																		set rsFCat=nothing
																		
																	end if
																end if
                                                                set rstemp=nothing
																
															loop '// do while (tmpgotit=0) and (pcv_ParentIDCat<>"1")
                                                        end if

                                                        if tmpgotit=1 then
                                                            exit for
														end if
														
													Next '//  For ik=0 to intCatCount
													
													if (pcIncExcCat="0") AND (tmpgotit=1) then
															pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
														pcv_FResults=1
													else
														if (pcIncExcCat="1") AND (tmpgotit=0) then
																pcv_ProTotal=pcv_ProTotal+ccur(pcProductList(f,2))
															pcv_FResults=1
														end if
													end if

												End If '// If intCatCount>0 Then
											end if '// if pcProductList(f,1)=0 then  (Not deleted product)											
										next '// for f=1 to ppcCartIndex
									End If '// If pcv_CatFilter=1 Then
								End If '// If pcv_Filters=0 Then
								
								'// END: Filter by Categories


								'// START: Filter by Customers
								pcv_CustFilter=0
								query="select pcFCust_IDCustomer from PcDFCusts where pcFCust_IDDiscount=" & pcv_IDDiscount1
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
								end if								
								if not rs.eof then
									pcv_Filters=pcv_Filters+1
									pcv_CustFilter=1
								end if
								set rs=nothing
								
								if pcv_CustFilter=1 then
		
									query="select pcFCust_IDCustomer from PcDFCusts where pcFCust_IDDiscount=" & pcv_IDDiscount1 & " and pcFCust_IDCustomer=" & session("IDCustomer")
									set rs=server.CreateObject("ADODB.RecordSet")
									set rs=connTemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
									end if							
									if not rs.eof then
										if (pcIncExcCust="0") then
											pcv_FResults=pcv_FResults+1
										end if
									else
										if (pcIncExcCust="1") then
											pcv_FResults=pcv_FResults+1
										end if
									end if
									set rs=nothing
								
								end if
								'// END: Filter by Customers

								'// START: Customer Categories
								pcv_CustCatFilter=0
								
								query="select pcFCPCat_IDCategory from pcDFCustPriceCats where pcFCPCat_IDDiscount=" & pcv_IDDiscount1
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)
								if err.number<>0 then
									call LogErrorToDatabase()
								end if								
								if not rs.eof then
									pcv_Filters=pcv_Filters+1
									pcv_CustCatFilter=1
								end if
								set rs=nothing
								
								if pcv_CustCatFilter=1 then
		
									query="select pcDFCustPriceCats.pcFCPCat_IDCategory from pcDFCustPriceCats, Customers where pcDFCustPriceCats.pcFCPCat_IDDiscount=" & pcv_IDDiscount1 & " and pcDFCustPriceCats.pcFCPCat_IDCategory = Customers.idCustomerCategory and Customers.idcustomer=" & session("IDCustomer")
									set rs=server.CreateObject("ADODB.RecordSet")
									set rs=connTemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
									end if							
									if not rs.eof then
										if (pcIncExcCPrice="0") then
											pcv_FResults=pcv_FResults+1
										end if
									else
										if (pcIncExcCPrice="1") then
											pcv_FResults=pcv_FResults+1
										end if
									end if
									set rs=nothing
								
								end if
								'// END: Filter by Customer Categories


								'// START: Filter by reatil or wholesale
		                        if (pcv_retail ="0" and pcv_wholeSale ="1") or (pcv_retail ="1" and pcv_wholeSale ="0") Then
							    	pcv_Filters=pcv_Filters+1
								   	if pcv_wholeSale = "1" and session("customertype") = 1 then
								   		pcv_FResults=pcv_FResults+1	
								   	end if 
								   	if pcv_retail = "1" and 	session("customertype") <> 1 Then
								    	pcv_FResults=pcv_FResults+1
								   	end if    
							    end if 
								'// END: Filter by reatil or wholesale

								if pcv_Filters<>pcv_FResults then
									pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_38")
								end if
								
							END IF
		

							
							if pDiscountError="" then     
								pTempPriceToDiscount=pPriceToDiscount
								pTempPercentageToDiscount=pPercentageToDiscount
								pTempIdDiscount=pcv_IDDiscount

								' calculate discount. Note: percentage does not affect shipment and payment prices
								if pTempPriceToDiscount>0 or pTempPercentageToDiscount>0 then
									if pcv_ProTotal=0 then
										pcv_ProTotal=pSubTotal-paymentTotal-CatDiscTotal
									else
										pcv_ProTotal=pcv_ProTotal-CatDiscTotal
									end if
									if pcv_PerToFlatCartTotal<>0 AND pcv_ProTotal>pcv_PerToFlatCartTotal then
										tempPercentageToDiscount=pcv_PerToFlatDiscount
									else
										tempPercentageToDiscount=(pTempPercentageToDiscount*(pcv_ProTotal)/100)
										tempPercentageToDiscount=RoundTo(tempPercentageToDiscount,.01)
									end if
									pcv_ProTotal=0
									tempDiscountAmount=pTempPriceToDiscount + tempPercentageToDiscount
									discountTotal=discountTotal + tempDiscountAmount
									Session("DiscountTotal"&pTempDiscCode)=tempDiscountAmount
									pCheckSubtotal=pSubtotal-discountTotal
									if pCheckSubTotal<0 then
										tempDiscountAmount=tempDiscountAmount+pChecksubTotal
									end if
									if discountTotal<=0 then
										pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5")
									else
										if intArryCnt=0 then
											discountAmount=tempDiscountAmount
											passDiscountCode=pTempDiscCode
											passDiscountCnt=passDiscountCnt+1
											intArryCnt=intArryCnt+1
										else
											discountAmount=discountAmount&","&tempDiscountAmount
											passDiscountCode=passDiscountCode&","&pTempDiscCode
											passDiscountCnt=passDiscountCnt+1
											intArryCnt=intArryCnt+1
										end if
										pSFDiscountCodeTotal = discountTotal
									end if
										
								else '// else is "Free Shipping Coupon"
									if pcv_ProTotal=0 then
										pcv_ProTotal=pSubTotal-paymentTotal
									else
										if pcv_FResults=1 then
											pcv_ProTotal=pcv_ProTotal 
										else
											pcv_ProTotal=pcv_ProTotal-CatDiscTotal  '// If no exclusions remove CD first							
										end if
									end if

									if Int(pCartQuantity)>=Int(dcQuantityFrom) and Int(pCartQuantity)<=Int(dcQuantityUntil) and Int(pCartTotalWeight)>=Int(dcWeightFrom) and Int(pCartTotalWeight)<=Int(dcWeightUntil) and ccur(pcv_ProTotal)>=ccur(dcPriceFrom) and ccur(pcv_ProTotal)<=ccur(dcPriceUntil) then

										if pcIntIdShipService<>"" then
											
											query="select pcFShip_IDShipOpt from pcDFShip where pcFShip_IDDiscount=" & pTempIdDiscount & " and pcFShip_IDShipOpt=" & pcIntIdShipService
											set rs=server.CreateObject("ADODB.RecordSet")
											set rs=connTemp.execute(query)
											if err.number<>0 then
												call LogErrorToDatabase()
											end if								
											if not rs.eof then
												if intArryCnt=0 then
													discountAmount=ccur(pcDblShipmentTotal)
													passDiscountCode=pTempDiscCode
													passDiscountCnt=passDiscountCnt+1
													intArryCnt=intArryCnt+1
												else
													discountAmount=discountAmount&","&ccur(pcDblShipmentTotal)
													passDiscountCode=passDiscountCode&","&pTempDiscCode
													passDiscountCnt=passDiscountCnt+1
													intArryCnt=intArryCnt+1
												end if
												Session("DiscountTotal"&pTempDiscCode)=discountAmount
												pcDblShipmentTotal=0
												pcv_FREESHIP="ok"
											else
												pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_36")
											end if
											set rs=nothing
											
										else
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_36")
										end if
									
									else
									
										if NOT (Int(pCartQuantity)>=Int(dcQuantityFrom) and Int(pCartQuantity)<=Int(dcQuantityUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5b")
										elseif NOT (Int(pCartTotalWeight)>=Int(dcWeightFrom) and Int(pCartTotalWeight)<=Int(dcWeightUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5a")
										elseif NOT (ccur(pcv_ProTotal)>=ccur(dcPriceFrom) and ccur(pcv_ProTotal)<=ccur(dcPriceUntil)) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5c")
										else
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5")
										end if
										
									end if
									
								end if
							end if
						
							if ((pDiscountDesc <> "") OR (pDiscountError<>"")) AND (noCode<>"1") Then
								if pcIntADCnt>0 AND intPcAuto=1 and pDiscountError<>"" then
									pDiscountDesc=""
								else
									TableRowStr=""
									if pDiscountError<>"" then
										pcGlobalDiscError=pcGlobalDiscError & "<li>" & pDiscountError & " (<b>"&pTempDiscCode&"</b>)</li>"
									else
										TableRowStr="<tr><td colspan=""3""><p><b>"
										TableRowStr=TableRowStr&dictLanguage.Item(Session("language")&"_orderverify_14")&"</b>"
										TableRowStr=TableRowStr&"&nbsp;"&pDiscountDesc
										TableRowStr=TableRowStr&"</p></td><td nowrap align=""right""><p>"
									end if
									
									if pDiscountDesc <> "" then
										if tempDiscountAmount>0 then
											TableRowStr=TableRowStr&"-"&scCurSign & money(tempDiscountAmount)
										end if
									else
										if TableRowStr<>"" then
											TableRowStr=TableRowStr&"&nbsp;"
										end if
									end If
									if TableRowStr<>"" then
										TableRowStr=TableRowStr&"</p></td></tr>"
									end if
									if NOT pDiscountError<>"" then
										pDiscountShowCode=pDiscountShowCode&pTempDiscCode&","
									end if
									DiscountTableRow=DiscountTableRow&TableRowStr
									Session("DiscountRow"&pTempDiscCode)=TableRowStr
								end if
							end if
							pDiscountError=""
							tempDiscountAmount=0
							
							END IF 'DiscountCodeArry(i)<>""
						
						Next

				END IF
				
				'// Start: Double check the discounts are still valid after all discounts have been applied
				if pDiscountError="" then
					dim AdjustedSubTotal
					AdjustedSubTotal=pSubTotal 
					AdjustedSubTotal=AdjustedSubTotal - discountTotal
					if AdjustedSubTotal<0 then
						AdjustedSubTotal=0
						discountTotal=pSubTotal						
					end if
				end if

				IF (pDiscountCode<>"" AND pcGlobalDiscError="") AND (len(passDiscountCode)>0) THEN
					tmpDiscountCodeArry = split(passDiscountCode,",")
					pcvCodeCnt=ubound(tmpDiscountCodeArry)
					If pcvCodeCnt > 0 Then 
						For i=0 to pcvCodeCnt			
							IF trim(tmpDiscountCodeArry(i))<>"" THEN							
								pTempDiscCode=tmpDiscountCodeArry(i)								
								
								query="SELECT priceFrom, priceUntil, priceToDiscount, PercentageToDiscount, pcDisc_Auto FROM discounts WHERE discountcode='" & pTempDiscCode & "' AND active=-1;"
								set rs2=server.CreateObject("ADODB.RecordSet")
								set rs2=connTemp.execute(query)	
								if NOT rs2.eof then
									
									dcPriceFrom=rs2("priceFrom")
									dcPriceUntil=rs2("priceUntil")
									pPriceToDiscount=ccur(rs2("priceToDiscount"))
									pPercentageToDiscount=ccur(rs2("PercentageToDiscount"))	
									tmpPcAuto=rs2("pcDisc_Auto")									
									tmpPcAuto=clng(tmpPcAuto)
									
									'// Only double check the discount code if it free shipping								
									if NOT ( (ccur(AdjustedSubTotal)>=ccur(dcPriceFrom)) AND (ccur(AdjustedSubTotal)<=ccur(dcPriceUntil)) ) then	
						
										if NOT (pPriceToDiscount>0 or pPercentageToDiscount>0) then
											pDiscountError=dictLanguage.Item(Session("language")&"_orderverify_5c")

											if pcIntADCnt>0 AND tmpPcAuto=1 and pDiscountError<>"" then
												pDiscountDesc=""
											else
												pcGlobalDiscError=pcGlobalDiscError & "<li>" & pDiscountError & " (<b>"&pTempDiscCode&"</b>)</li>"										
												pcv_FREESHIP="" '// disable free shipping
												pcDblShipmentTotal=Session("DiscountTotal"&pTempDiscCode) '// add shipping total back									
												'// remove the invalid discount
												discountAmount=replace(discountAmount, ","&tempDiscountAmount, "")	
												passDiscountCode=replace(passDiscountCode, ","&pTempDiscCode, "")	
												discountAmount=replace(discountAmount, tempDiscountAmount, "")	
												passDiscountCode=replace(passDiscountCode, pTempDiscCode, "")									
												DiscountTableRow=replace(DiscountTableRow, Session("DiscountRow"&pTempDiscCode), "")
												passDiscountCnt=passDiscountCnt-1
												intArryCnt=intArryCnt-1		
											end if
										end if
									end if	
									
								end if	
								set rs2=nothing	
								pDiscountError=""
								tempDiscountAmount=0
								Session("DiscountTotal"&pTempDiscCode)=""
								Session("DiscountRow"&pTempDiscCode)=""										
							END IF					
						Next
					End If
				END IF
				'// END: Double check the discounts are still valid after all discounts have been applied
				
				if pDiscountError="" then
					dim tSubTotal
					tSubTotal=pSubTotal 
					pSubTotal=pSubTotal - discountTotal
					if pSubTotal<0 then
						pSubTotal=0
						discountTotal=tSubTotal						
					end if
				end if

				session("SF_DiscountTotal")= discountTotal
				'****************************************************
				' END - Check Discount codes
				'****************************************************
				
				'//////////////////////////////////////////////////////////////
				' END - Discounts by code
				'//////////////////////////////////////////////////////////////

				'GGG Add-on start
				if Session("Cust_GW")="1" then
					GWTotal=calGWTotal()
					pTotal=pTotal+ccur(GWTotal)
				end if
				'GGG Add-on end
					
				' tax calculations.
				'Include Payment/Shipping charges for Tax calculation
				dim taxCalcAmt
				taxCalAmt=0

				if Session("customerType")<>1 OR (Session("customerType")=1 AND ptaxwholesale=1) then
					if TAX_SHIPPING_ALONE="NA" then
						If pTaxonCharges=1 then
							taxCalAmt=taxCalAmt+pcDblShipmentTotal
						End If
						If pTaxonFees=1 then
							taxCalAmt=taxCalAmt+pcDblServiceHandlingFee
						End If
					else
						if TAX_SHIPPING_AND_HANDLING_TOGETHER="Y" then
							taxCalAmt=taxCalAmt+pcDblShipmentTotal+pcDblServiceHandlingFee
						else
							if TAX_SHIPPING_ALONE="Y" then
								taxCalAmt=taxCalAmt+pcDblShipmentTotal
							end if
						end if
					end if
					
					if ptaxCanada="1" and session("SFTaxZoneRateCnt")>0 then
						taxCalAmt=0
					end if
					if cdbl(taxCalAmt)=0 AND cdbl(pTaxableTotal)=0 AND (ptaxCanada<>"1" OR (ptaxCanada="1" AND session("SFTaxZoneRateCnt")=0)) then
						ptaxLocAmount=0
					else
						'GGG Add-on start
						if Session("Cust_GW")="1" then
							pTaxableTotal=pTaxableTotal+ccur(GWTotal)
						end if
						'GGG Add-on end
						'if VAT
						VATTotal=0
						if taxPaymentTotal="" then
							taxPaymentTotal=0
						end if
						
						'////////////////////////////////////////////////////////////////////////////////////////
						'// START: VAT
						'////////////////////////////////////////////////////////////////////////////////////////
						if ptaxVAT="1" then
							


							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// Start: Discount Distribution %
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							Dim ApplicableDisountTotal
							ApplicableDisountTotal = pTaxableTotal + taxCalAmt + taxPaymentTotal
							
							'// Shipping and Handling represents what % of the Total Discount?  							
							Proportional_taxCalAmt = RoundTo((taxCalAmt/ApplicableDisountTotal),.01)
							
							'// Payment Charges represents what % of the Total Discount? 							
							Proportional_taxPaymentTotal = RoundTo((taxPaymentTotal/ApplicableDisountTotal),.01)
							
							'// Product Pricing represents what % of the Total Discount?
							'	NOTE: Product Level Distributions are calculated at the line item via "calculateVATTotal"
							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// End: Discount Distribution %
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// Start: Distribute Discounts based off % above
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
							'// Shipping and Handling after discount
							ApplicableDisount_taxCalAmt = (discountTotal * Proportional_taxCalAmt)
							taxCalAmt = (taxCalAmt - ApplicableDisount_taxCalAmt)
							
							'// Payment Charges after discount
							ApplicableDisount_taxPaymentTotal = (discountTotal * Proportional_taxPaymentTotal)
							taxPaymentTotal = (taxPaymentTotal - ApplicableDisount_taxPaymentTotal)
							
							'// Products Price after discount
							'	NOTE: Discount are distributed to Products at the line item via "calculateVATTotal"
							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// End: Distribute Discounts based off % above
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							

							
							'// VAT TAXED AMOUNT - ORDER LEVEL ("Discounts" Removed Proportionately)
							VatTaxedAmount_OrderLevel = taxCalAmt + taxPaymentTotal	
							
							'// VAT TAXED AMOUNT - PRODUCT LEVEL ("Discounts" and "Category Discounts" Removed Proportionately)						
							VatTaxedAmount_ProductLevel = ccur( calculateVATTotal(pcCartArray, ppcCartIndex, discountTotal, CatDiscTotal) )
							
							'// VAT TAXED AMOUNT
							VatTaxedAmount = VatTaxedAmount_OrderLevel + VatTaxedAmount_ProductLevel
							
							'// Shipping and Handling "VATable" Total - Uses Default Rate							
							taxCalAmtNoVAT = RoundTo(pcf_RemoveVAT(taxCalAmt,""),.01)
							CalAmtTotal = RoundTo(taxCalAmt-taxCalAmtNoVAT,.01)
										
							'// Payment Charges "Always VATable" Total - Uses Default Rate
							taxPaymentTotalNoVAT = RoundTo(pcf_RemoveVAT(taxPaymentTotal,""),.01)
							tPaymentTotal = RoundTo(taxPaymentTotal-taxPaymentTotalNoVAT,.01)
							
							'// Gift Wrapping Charges "Always VATable" Total - Uses Default Rate
							taxGiftWrapNoVAT = RoundTo(pcf_RemoveVAT(GiftWrapPaymentTotal,""),.01)
							tGiftWrapTotal = RoundTo(GiftWrapPaymentTotal-taxGiftWrapNoVAT,.01)							

							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// Start: VAT Totals
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
							'// Order Level Total VAT ("Discounts" Removed)
							VATTotal_OrderLevel = CalAmtTotal + tPaymentTotal + tGiftWrapTotal	'// Payment VAT + Shipping/Handling VAT					
							
							'// Product Level Total VAT ("Discounts" and "Category Discounts" Removed)
							NoVATTotal=ccur( calculateNoVATTotal(pcCartArray, ppcCartIndex, discountTotal, CatDiscTotal) )	
							NoVATTotal=RoundTo(NoVATTotal,.01)
							VATTotal_ProductLevel=RoundTo(VatTaxedAmount_ProductLevel-NoVATTotal,.01)
							
							'// Total VAT
							VATTotal=VATTotal_OrderLevel+VATTotal_ProductLevel							
							
							'// NOTE: CalAmtTotal is included in the Subtotal for display purposes, but technically is applied to Order Level.
							
							'// Display the correct Sub Total when outside the EU
							if pcv_IsEUMemberState = 0 then
								pSubTotal = pSubTotal - VATTotal_ProductLevel - tPaymentTotal - tGiftWrapTotal - CalAmtTotal '// Remove VAT charges from the total.
							end if
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							'// END: VAT Totals
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							
							'// For reference determine the Total VAT Taxed Amount
							'	NOTE: Includes "Shipping and Handling" + "Payment Charges" - "Discounts" - "Cat Discounts" (discounts applied proportionately)
							VATTaxedAmount=RoundTo(VatTaxedAmount,.01)
							if VATTaxedAmount<0 then
								VATTaxedAmount=0
							end if
							
							'// If outside EU then specifiy how much VAT was removed
							VATRemovedTotal=0
							if pcv_IsEUMemberState=0 then
								VATRemovedTotal=VATTotal
								VATTaxedAmount=0
							end if	
							'////////////////////////////////////////////////////////////////////////////////////////
							'// END: VAT
							'////////////////////////////////////////////////////////////////////////////////////////
							
						else
						
							if ptaxCanada="1" and session("SFTaxZoneRateCnt")>0 then
								'//Calculate Zone Taxes
								ptaxDetailsString=""
								ptaxLocAmount=0
								pTempTaxableCanadaTotal=0
								for u=1 to session("SFTaxZoneRateCnt")
									taxCalAmt=0
									session("taxAmount"&u)=0
									pcv_IntTaxZoneRateID=session("SFTaxZoneRateID"&u)
									dim pTaxZoneExemption
									if intTaxExemptZoneFlag="1" then
										'recalculate taxable total
										if pTempTaxableCanadaTotal>0 then
											pTaxableCanadaTotal=pTempTaxableCanadaTotal
										else
											pTaxableCanadaTotal=ccur(calculateTaxableZoneTotal(pcCartArray, ppcCartIndex,pcv_IntTaxZoneRateID))
										end if
										'change per rewards points
										if pTaxableCanadaTotal>0 then
											ptaxLoc=session("SFTaxZoneRateRate"&u)
											if session("SFTaxZoneRateApplyToSH"&u) then
												taxCalAmt=taxCalAmt+pcDblShipmentTotal+pcDblServiceHandlingFee
											end if
												tempTAmt=((pTaxableCanadaTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal) * ptaxLoc)
											tempTAmt=roundTo(tempTAmt,.01)
											if tempTAmt<0 then
												tempTAmt=0
											end if
											ptaxLocAmount=ptaxLocAmount+tempTAmt
											session("taxAmount"&u)=tempTAmt
											session("taxDesc"&u)=session("SFTaxZoneRateName"&u)
										else
											tempTAmt=0
										end if
										if session("SFTaxZoneRateTaxable"&u)="1" then
											pTempTaxableCanadaTotal=pTaxableCanadaTotal+tempTAmt
										end if
									else
										pTaxZoneExemption=checkTaxExempt(pcCartArray, ppcCartIndex,pcv_IntTaxZoneRateID)
										if pTaxZoneExemption=0 then
											ptaxLoc=session("SFTaxZoneRateRate"&u)
											if session("SFTaxZoneRateApplyToSH"&u) then
												taxCalAmt=taxCalAmt+pcDblShipmentTotal+pcDblServiceHandlingFee
											end if
												tempTAmt=((pTaxableTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal-TotalPromotions) * ptaxLoc)
											tempTAmt=roundTo(tempTAmt,.01)
											if tempTAmt<0 then
												tempTAmt=0
											end if
											ptaxLocAmount=ptaxLocAmount+tempTAmt
											session("taxAmount"&u)=tempTAmt
											session("taxDesc"&u)=session("SFTaxZoneRateName"&u)
											if session("SFTaxZoneRateTaxable"&u)="1" then
												pTempTaxableCanadaTotal=pTaxableCanadaTotal+tempTAmt
											end if
										else
											session("taxAmount"&u)=0
											session("taxDesc"&u)=session("SFTaxZoneRateName"&u)
										end if
									end if
									ptaxDetailsString=ptaxDetailsString&replace(session("taxDesc"&u),",","")&"|"&session("taxAmount"&u)&","
								next

								session("taxCnt")=session("SFTaxZoneRateCnt")
							else
								if session("taxCnt")>"0" then
									ptaxDetailsString=""
									ptaxLocAmount=0
									for i=1 to session("taxCnt")
										ptaxLoc=session("tax"&i)
										tempTAmt=((pTaxableTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal-TotalPromotions) * ptaxLoc)
										tempTAmt=roundTo(tempTAmt,.01)
										if tempTAmt<0 then
											tempTAmt=0
										end if
										ptaxLocAmount=ptaxLocAmount+tempTAmt
										session("taxAmount"&i)=tempTAmt
										ptaxDetailsString=ptaxDetailsString&replace(session("taxDesc"&i),",","")&"|"&session("taxAmount"&i)&","
									next
								else
									ptaxLocAmount=((pTaxableTotal+taxCalAmt+taxPaymentTotal-discountTotal-CatDiscTotal-TotalPromotions) * ptaxLoc)
									ptaxLocAmount = CCur(pTaxLocAmount)
									ptaxLocAmount=RoundTo(ptaxLocAmount,.01)
									if ptaxLocAmount<0 then
										ptaxLocAmount=0
									end if
								end if
							end if 'pTaxCanada
						end if
					end if
				else
					ptaxLocAmount=0
				end if

				pTaxAmount=ptaxPrdAmount + ptaxLocAmount
			   	%>
				
				<%'Show Product Promotions
				if Session("pcPromoIndex")<>"" and Session("pcPromoIndex")>"0" then
					PromoArr1=Session("pcPromoSession")
					PromoIndex=Session("pcPromoIndex")
					For m=1 to PromoIndex%>
					<tr>
						<td colspan="3">
							<%=PromoArr1(m,1)%>
						</td>
						<td nowrap align="right">
							-<%=scCurSign  & money(PromoArr1(m,2))%>
						</td>
					</tr>
					<%Next
				end if%>
					
				<%if CatDiscTotal>0 then%>
					<tr> 
						<td colspan="3" align="right"><p><b><%=dictLanguage.Item(Session("language")&"_catdisc_2")%></b></p></td>
						<td nowrap align="right"><p><%response.write "-" & scCurSign & money(CatDiscTotal)%></p></td>
					</tr>
				<%end if%>
													
				<% if session("customerType")=1 then %>
					<tr> 
						<td colspan="3" align="right">
							<p><b><%=dictLanguage.Item(Session("language")&"_orderverify_12")%></b> <%=pPaymentDesc%></b></p>
						</td>
						<td nowrap align="right">
							<% if paymentTotal > "0" then %>
								<p><% response.write scCurSign & money(paymentTotal) %></p>
							<% end if %>
						</td>
					</tr>
				<% else %>
					<% if paymentTotal > "0" then %>
						<tr> 
							<td colspan="3" align="right">
								<p><b><%=dictLanguage.Item(Session("language")&"_orderverify_20")%></b></p>
							</td>
							<td nowrap align="right">
								<% if paymentTotal > "0" then %>
									<p><% response.write scCurSign & money(paymentTotal)%></p>
								<% end if %>
							</td>
					</tr>
					<% end if %>
				<% end if %>
          
				<%'GGG Add-on start
				 if (DiscountTableRow<>"") Then %>
				<%=DiscountTableRow%>
				<%end if
				'GGG Add-on end%>
					
				<% '// Start Reward Points
				RewardsDollarValue=0
				If RewardsActive=1 And pcSFUseRewards<>"" Then ' Reward Points are being used against the purchase
				%>
					<tr> 
						<td colspan="3" align="right">
							<p><b><%response.write pcSFUseRewards &" "&RewardsLabel %></b><%=dictLanguage.Item(Session("language")&"_orderverify_31")%></p>
						</td>
						<td align="right" nowrap>
							<p>
							<% response.write "-" & scCurSign & money(iDollarValue)
							RewardsDollarValue=iDollarValue							
							%>
							</p>
						</td>
					</tr>
				<% 
				End If 
				session("SF_RewardPointTotal") = RewardsDollarValue
				%>
					
				<%
				If RewardsActive=1 And pcIntUseRewards=0 And pCartRewards > 0 Then ' Reward points are being accrued.
				%>
					<tr> 
						<td colspan="4" align="right">
							<p><b><%=dictRewardsLanguage.Item(Session("rewards_language")&"_orderverify")%></b>
							<% response.write dictLanguage.Item(Session("language")&"_orderverify_30") & pCartRewards %></p>
						</td>
					</tr>
				<% End If 
				'// End Reward Points
				%>				

                            
				<tr> 
					<td colspan="3" align="right"><p><b><%=dictLanguage.Item(Session("language")&"_orderverify_15")%></b></p></td>
					<td nowrap align="right">
						<p><%response.write scCurSign & money(pSubTotal)%></p>
					</td>
				</tr>
				
				<% 
				'GGG Add-on start
				If Session("Cust_GW")="1" then %>
				<tr> 
					<td colspan="3" align="right"><p><%response.write dictLanguage.Item(Session("language")&"_orderverify_36a")%></p></td>
					<td align="right"><p><%response.write scCurSign & money(GWTotal)%></p></td>
				</tr>
				<%
				End If
				'GGG Add-on end %>
				
				<% If savNullShipper="Yes" then %>
				<% Else %>
					<tr> 
						<td colspan="3" align="right">
							<p><b><%=dictLanguage.Item(Session("language")&"_orderverify_13")%></b> <%=pcStrShipmentDesc%></p>
						</td>
						<td nowrap align="right">
							<p>
							<%if pcDblShipmentTotal > 0 then %>
								<%response.write scCurSign & money(pcDblShipmentTotal)%>
							<%else
								if pcv_FREESHIP="ok" then%>
									<%response.write dictLanguage.Item(Session("language")&"_orderverify_37")%>
									<%shiparr=split(TempStrNewShipping,",")
									TempStrNewShipping=""
									for i=lbound(shiparr) to ubound(shiparr)
										if i=2 then
											TempStrNewShipping=TempStrNewShipping & "0,"
										else
											if i=ubound(shiparr) then
												TempStrNewShipping=TempStrNewShipping & shiparr(i)
											else
												TempStrNewShipping=TempStrNewShipping & shiparr(i) & ","
											end if
										end if
									next
								else
									response.write dictLanguage.Item(Session("language")&"_orderverify_37")
								end if
							end if %>
							</p>
						</td>
					</tr>
				<% End If %>
                            
				<% if pcDblServiceHandlingFee<>0 then %>
					<tr> 
						<td colspan="3" align="right"><p><b><%=dictLanguage.Item(Session("language")&"_orderverify_18")%></b></p></td>
						<td nowrap align="right">
							<p><%response.write scCurSign & money(pcDblServiceHandlingFee)%></p>
						</td>
					</tr>
				<% end if %>
					
				<% 'start taxes %>
				
				<% if ptaxVAT<>"1" AND pTaxAmount>0 then
					if (ptaxseparate="1" OR (ptaxCanada="1" AND session("SFTaxZoneRateCnt")>0)) AND session("taxCnt")<>0 then
							for i=1 to session("taxCnt") %>
							<% if (session("taxAmount"&i) > 0) then%>
									<tr> 
										<td align="right" colspan="3"><p><b><%=session("taxDesc"&i)%></b></p></td>
										<td nowrap align="right">
											<p><% response.write scCurSign & money(session("taxAmount"&i))%></p>
										</td>
									</tr>
							<% end if
								if ccur(ptaxPrdAmount)>0 then %>
									<tr> 
										<td align="right" colspan="3"><p><b><%=dictLanguage.Item(Session("language")&"_orderverify_44")%></b></p></td>
										<td nowrap align="right">
											<p><% response.write scCurSign & money(ptaxPrdAmount)%></p>
										</td>
									</tr>
								<% end if
							next %>
					<% else %>
						<tr> 
							<td colspan="3" align="right">
								<p><b><%=dictLanguage.Item(Session("language")&"_orderverify_16")%></b></p>
							</td>
							<td nowrap align="right">
								<p><% response.write scCurSign & money(pTaxAmount)%></p>
							</td>
						</tr>
					<% end if %>
				<% end if %>
				<% 'end taxes %>

				<%'GGG Add-on start
				if Session("Cust_GW")="1" then
					pSubTotal=pSubTotal + ccur(GWTotal)
				end if 
				'GGG Add-on end %>
				<%'GGG Add-on start
					'****************************************************
					' START - Check Gift Certificates
					'****************************************************
					ListGCs=""
					ListUsedGCs=""
					TotalGCAmount=0
					passGCCode=""
					IF pTempGC<>"" THEN
						if intGCIncludeShipping=1 then
							pGCSubTotal=pSubTotal + pcDblShipmentTotal + pcDblServiceHandlingFee + RoundTo(pTaxAmount,.01)
						else
							pGCSubTotal=pSubTotal + pcDblServiceHandlingFee + RoundTo(pTaxAmount,.01)
						end if
						pSubTotal=pGCSubTotal
					
						GCArr=split(pTempGC,",")
						pTempGC=""
						For i=0 to ubound(GCArr)
							if GCArr(i)<>"" AND cdbl(pSubTotal)>0 then
								intDiscMatchFound=0
								if ListUsedGCs<>"" then
									UsedGCArry=split(ListUsedGCs,",")
									for t=0 to (ubound(UsedGCArry)-1)
										if GCArr(i)=UsedGCArry(t) then
											intDiscMatchFound=1
										end if
									next
								end if
							
								if intDiscMatchFound=0 then
							
									ListUsedGCs=ListUsedGCs&GCArr(i)&","
							
									query="SELECT pcGCOrdered.pcGO_ExpDate, pcGCOrdered.pcGO_Amount, pcGCOrdered.pcGO_Status, products.Description FROM pcGCOrdered, products WHERE pcGCOrdered.pcGO_GcCode='"&GCArr(i)&"' AND products.idproduct=pcGCOrdered.pcGO_IDProduct"
									set rsQ=Server.CreateObject("ADODB.Recordset")
									set rsQ=conntemp.execute(query)
					
									IF rsQ.eof then
										pGCError=pGCError & "<li>" & dictLanguage.Item(Session("language")&"_orderverify_4") & ": " & "<b>" & GCArr(i) & "</b></li>"
									ELSE
										mTest=0
										pGCExpDate=rsQ("pcGO_ExpDate")
										pGCAmount=rsQ("pcGO_Amount")
										if len(pGCAmount)<0 then
											pGCAmount=0
										end if
						
										pGCStatus=rsQ("pcGO_Status")
										pDiscountDesc=rsQ("Description")
										if mTest=0 AND ccur(pGCAmount)<=0 then
											mTest=1
											pGCError=pGCError & "<li>" & dictLanguage.Item(Session("language")&"_msg_1") & "<b>" & pDiscountDesc & "</b>" & " (<b>" & GCArr(i) & "</b>)" & dictLanguage.Item(Session("language")&"_msg_3") & "</li>"
										end if
										if mTest=0 AND cint(pGCStatus)<>1 then
											mTest=1
											pGCError=pGCError & "<li>" & dictLanguage.Item(Session("language")&"_msg_1") & "<b>" & pDiscountDesc & "</b>" & " (<b>" & GCArr(i) & "</b>)" & dictLanguage.Item(Session("language")&"_msg_1a") & "</li>"
										end if
										if mTest=0 AND year(pGCExpDate)<>"1900" then
											if Date()>pGCExpDate then
												mTest=1
												pGCError=pGCError & "<li>" & dictLanguage.Item(Session("language")&"_msg_1") & "<b>" & pDiscountDesc & "</b>" & " (<b>" & GCArr(i) & "</b>)" & dictLanguage.Item(Session("language")&"_msg_2") & "</li>"
											end if
										end if
										if mTest=0 then
										'Have Available Amount
											GCAmount=pGCAmount
											pTempSubTotal1=pSubTotal - GCAmount
											tempGCAmount=0
											if pTempSubTotal1<0 then
												pGCAmount=pGCAmount-pSubTotal
												TotalGCAmount=TotalGCAmount+pSubTotal
												if ListGCs<>"" then
													ListGCs=ListGCs & "|g|"
												end if
												ListGCs=ListGCs & GCArr(i) & "|s|" & pDiscountDesc & "|s|" & pSubTotal
												tempGCAmount=pSubTotal
												if passGCCode<>"" then
													passGCCode=passGCCode & ","
												end if
												pSubTotal=0
												passGCCode=passGCCode & GCArr(i)
											else
												pSubTotal=pTempSubTotal1
												TotalGCAmount=TotalGCAmount+pGCAmount
												pGCAmount=0
												if ListGCs<>"" then
													ListGCs=ListGCs & "|g|"
												end if
												ListGCs=ListGCs & GCArr(i) & "|s|" & pDiscountDesc & "|s|" & GCAmount
												tempGCAmount=GCAmount
												if passGCCode<>"" then
													passGCCode=passGCCode & ","
												end if
												passGCCode=passGCCode & GCArr(i)
											end if
											if pTempGC<>"" then
												pTempGC=pTempGC & ","
											end if
											pTempGC=pTempGC & GCArr(i)%>
											<tr> 
											<td colspan="3" align="right">
											<p>
												<% 
												if pDiscountDesc<>"" then
													response.write "<strong>" & dictLanguage.Item(Session("language")&"_orderverify_46") & "</strong>" & pDiscountDesc & " (" & GCArr(i) & ")"
												end if
												%>
											</p>
											</td>
											<td nowrap align="right"><p>
											<% if pDiscountDesc <> "" then
											if tempGCAmount>0 then  %>
												-<%response.write scCurSign & money(tempGCAmount)%>
											<%end if
											end If %>
											</p>
											</td>
											</tr>
										<%end if
									END IF
								set rs=nothing
								end if 'intDiscMatchFound
							end if
						Next 'GCArr
					
						if pGCError<>"" then
							pGCError="<ul>" & pGCError & "</ul>"
						end if
					
						GCAmount=TotalGCAmount
					END IF
				
					'****************************************************
					' END - Check Gift Certificates
					'****************************************************
				'GGG Add-on end%>
				<tr> 
					<td colspan="3" align="right"><p><b><%=dictLanguage.Item(Session("language")&"_orderverify_17")%></b></p></td>
					<td nowrap align="right">
						<p>
						<% if pSubTotal<0 then
								pSubTotal=0
						end if
						pSubTotal=RoundTo(pSubTotal,.01) 'pSubTotal=Round(pSubTotal,2)
						'GGG Add-on start
						IF GCAmount=0 then
							pSubTotal=pSubTotal + pcDblShipmentTotal + pcDblServiceHandlingFee + RoundTo(pTaxAmount,.01) 'Round(pTaxAmount,2)
						Else
							if intGCIncludeShipping=0 then
								pSubTotal=pSubTotal + pcDblShipmentTotal
							end if
						END IF
						'GGG Add-on end
						response.write scCurSign & money(pSubTotal)
						%>
						</p>
					</td>
				</tr>
				<% if ptaxVAT="1" and VATTotal>0 then %>
					<tr> 
						<td colspan="4" align="right" class="pcSmallText">
							<% if VATRemovedTotal=0 then %>
								<p><% response.write dictLanguage.Item(Session("language")&"_orderverify_35") & scCurSign & money(VATTotal) %></p>
							<% else %>
								<p><% response.write dictLanguage.Item(Session("language")&"_orderverify_42") & scCurSign & money(VATTotal) %></p>
							<% end if %>
						</td>
					</tr>
				<% end if %>
				<% 
				'SB S
				if pSubTotal=0 AND Not pcIsSubscription then
				'SB E
					chkPayment="FREE"
				else 
					if pidPayment<>999999 then %>
							<% 
							'SB S
							strAndSub = ""
							if pcIsSubscription = True Then
							   strAndSub = " AND pcPayTypes_Subscription = 1 ORDER by pcPayTypes_Subscription, paymentPriority"
							else
							   strAndSub = " ORDER by paymentPriority"
							End if 
							'SB E
							
							' get available paytypes
							if session("customerType")=1 then
								query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd, type, gwcode, paymentNickName,sslURL FROM paytypes WHERE active=-1 AND (payTypes.pcPayTypes_PPAB <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
							else
								query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd, type, gwcode, paymentNickName,sslURL FROM paytypes WHERE active=-1 AND Cbtob=0 AND (payTypes.pcPayTypes_PPAB <> 1) AND (gwcode<>50 AND gwcode<>999999)" & strAndSub
							end if
							set rs=server.CreateObject("Adodb.recordset")
							set rs=conntemp.execute(query)
						
							if err.number<>0 then
								call LogErrorToDatabase()
								'set rs=nothing
								'call closedb()
								'response.redirect "techErr.asp?err="&pcStrCustRefID
							end if

							if not rs.eof then %>
								<% while not rs.eof
									tempidPayment=rs("idPayment")
									temppaymentDesc=rs("paymentDesc")
									temppriceToAdd=rs("priceToAdd")
									temppercentageToAdd=rs("percentageToAdd")
									tempType=rs("Type")
									tempgwCode=rs("gwCode")
									tempPaymentNickName=rs("paymentNickName")
									if isNull(temppriceToAdd) OR temppriceToAdd="" then
										temppriceToAdd=0
									end if
									if isNull(temppercentageToAdd) OR temppercentageToAdd="" then
										temppercentageToAdd=0
									end if
									if ccur(temppriceToAdd)<>0 or ccur(temppercentageToAdd)<>0 then 
										HowMuch=temppriceToAdd + (temppercentageToAdd*intCalPaymnt/100)
										HowMuch=roundTo(HowMuch,.01)           
									else
										HowMuch=""
									end if
									CustomPayType=0
									payURL=rs("sslURL")
									if payURL<>"" then
										if Instr(UCase(payURL),UCASE("paymnta_"))=1 then
											CustomPayType=1
											payURL="opc_" & payURL
										end if
									end if
									if int(pidPayment)=int(tempidPayment) then 
										if CustomPayType=0 then
											session("NeedToUpdatePay")="0"									
										end if
									end if
									rs.movenext
								wend
								set rs=nothing %>
							<%end if%>
					<% end if
				end if
				'// START REWARD POINTS
				If RewardsActive = 1 then
					Dim pcIntHideRPField
					pcIntHideRPField = 0
					If pcSFCartRewards<>0 AND pcSFUseRewards="" then
				%>
						<tr>
							<td colspan="4"><hr></td>
						</tr>

						<tr> 
							<td colspan="4" align="right">
								<p><b><%=dictRewardsLanguage.Item(Session("rewards_language")&"_orderverify")%></b>
								<% response.write dictLanguage.Item(Session("language")&"_orderverify_30") & pcSFCartRewards %></p>
							</td>
						</tr>
					<%
					end if
				else 
					pcIntHideRPField = 1
				end if
				'// END REWARD POINTS%>
               	<%
				'// START: FREE SHIPPING Eligibility check

				Dim pSubTotalCheckFreeShipping
				pSubTotalCheckFreeShipping = pSubTotal-pcDblServiceHandlingFee

				Dim IsFreeShippingAvailable
				IsFreeShippingAvailable = ((cdbl(pSubTotalCheckFreeShipping) + cdbl(GCAmount)) < cdbl(serviceFreeOverAmt))

				'// Make sure Free Shipping is still available
				If (pcDblShipmentTotal=0) AND (IsFreeShippingAvailable) AND (pcv_FREESHIP<>"ok") Then
					tmpErrorOPCReady="<div class=pcErrorMessage>" & dictLanguage.Item(Session("language")&"_opc_43") & "</div>"
					session("OPCReady")="NO"
				Else	
					session("OPCReady")="YES"
				End If 
				
				'// END: FREE SHIPPING Eligibility check
				%>

</table>
<%
IF session("OPCReady")="NO" THEN
	session("OPCstep")="3"
	response.clear()
	Call SetContentType()
	response.Write(tmpErrorOPCReady)
END IF
'// Update Customer Session with new changes (if any): Discounts, Reward Points, Payment Methods
if pcSFUseRewards="" then
	pcSFUseRewards=0
end if
if pTaxAmount="" then
	pTaxAmount=0
end if
if pSubTotal="" then
	pSubTotal=0
end if
if discountAmount="" then
	discountAmount=0
end if
if passDiscountCnt="" then
	passDiscountCnt=0
end if
if VATTotal="" then
	VATTotal =0
end if
if RewardsDollarValue="" then
	RewardsDollarValue=0
end if

if pSFDiscountCodeTotal="" then
	pSFDiscountCodeTotal=0
end if

if pSFSubTotal="" then
	pSFSubTotal=0
end if

if GWTotal="" then
	GWTotal="0"
end if

if pSFCatDiscTotal="" then
	pSFCatDiscTotal=0
end if

if pcSFCartRewards="" then
	pcSFCartRewards=0
end if

if pcIntBalance="" then
	pcIntBalance=0
end if

if GCAmount="" then
	GCAmount=0
end if

'SB S
If len(pcv_sbTax)>0 Then
	query="UPDATE pcCustomerSessions SET pcCustSession_SB_taxAmount=" & pTaxAmount & ",pcCustSession_VATTotal=" & VATTotal & ",pcCustSession_taxDetailsString='" & ptaxDetailsString & "' WHERE pcCustomerSessions.idDbSession="&session("pcSFIdDbSession")&" AND pcCustomerSessions.randomKey="&session("pcSFRandomKey")&" AND pcCustomerSessions.idCustomer="&session("idCustomer")&";"
Else
query="UPDATE pcCustomerSessions SET pcCustSession_GCDetails='" & replace(ListGCs,"'","''") & "',pcCustSession_GCTotal=" & GCAmount & ",pcCustSession_strBundleArray='" & pSFstrBundleArray & "',pcCustSession_CatDiscTotal=" & pSFCatDiscTotal & ",pcCustSession_DiscountCodeTotal=" & pSFDiscountCodeTotal & ",pcCustSession_pSubTotal=" & pSFSubTotal & ",pcCustSession_chkPayment='" & chkPayment & "',pcCustSession_RewardsDollarValue=" & RewardsDollarValue & ",pcCustSession_GWTotal=" & GWTotal & ",pcCustSession_taxAmount=" & pTaxAmount & ",pcCustSession_total=" & pSubTotal & ",pcCustSession_discountAmount='" & discountAmount & "',pcCustSession_intCodeCnt=" & passDiscountCnt & ",pcCustSession_VATTotal=" & VATTotal & ",pcCustSession_taxDetailsString='" & ptaxDetailsString & "',pcCustSession_discountcode='" & passDiscountCode & "',pcCustSession_UseRewards=" & pcSFUseRewards & ",pcCustSession_RewardsBalance=" & pcIntBalance & ",pcCustSession_IdPayment=" & pidPayment & ",pcCustSession_CartRewards=" & pcSFCartRewards & " WHERE pcCustomerSessions.idDbSession="&session("pcSFIdDbSession")&" AND pcCustomerSessions.randomKey="&session("pcSFRandomKey")&" AND pcCustomerSessions.idCustomer="&session("idCustomer")&";"
End If
set rs=connTemp.execute(query)
set rs=nothing
'SB E

tmpStr=""
pcGlobalDiscError=pGCError & pcGlobalDiscError
if pcGlobalDiscError<>"" then
	tmpStr="|***|ERROR"
	pcGlobalDiscError=dictLanguage.Item(Session("language")&"_opc_44") & "<ul>" & pcGlobalDiscError & "</ul>"
else
	tmpStr="|***|OK"
end if
if passGCCode<>"" then
	if passDiscountCode<>"" then
		if Right(passDiscountCode,1)<>"," then
			passDiscountCode=passDiscountCode & ","
		end if
	end if
	passDiscountCode=passDiscountCode & passGCCode
end if
	
tmpStr=tmpStr & "|***|" & passDiscountCode & "|***|" & pcGlobalDiscError & "|***|" & pcSFUseRewards & "|***|" & chkPayment & "|***|" & pIdPayment & "|***|" & scCurSign & money(pSubTotal) & "|***|" & session("OPCReady") & "|***|" & pSubTotalCheckFreeShipping
response.write tmpStr
call closedb() %>