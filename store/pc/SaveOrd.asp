<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "SaveOrd.asp"
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/pcAffConstants.asp"-->
<% 'SB S %>
<!--#include file="inc_sb.asp"-->
<% 'SB E %>
<%
On Error Resume Next
Response.Buffer = True

if session("idCustomer")=0 OR session("idCustomer")="" then
	response.redirect "OnePageCheckout.asp?msg=1"
end if

dim pcIsOPC
pcIsOPC=getUserInput(request("opc"),0)

Dim query, connTemp, rs

%><!--#include file="inc_checkPrdQtyCart.asp"--><%
If len(pcIsOPC)=0 Then
	Call CheckALLCartStock()
End If

'Get info from sessions and customers
call opendb()

'SB S
session("pcIsSubTrial") = False
'SB E
'arrService=Session("Service")
'pServiceIndex=Session("serviceIndex")

Dim pcCartArray
'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
ppcCartIndex=Session("pcCartIndex")

query="SELECT customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email, customers.address, customers.zip, customers.stateCode, customers.state, customers.city, customers.countryCode, customers.address2, customers.fax, pcCustSession_ShippingFirstName, pcCustSession_ShippingLastName, pcCustSession_ShippingCompany, pcCustSession_ShippingAddress, pcCustSession_ShippingAddress2, pcCustSession_ShippingCity, pcCustSession_ShippingStateCode, pcCustSession_ShippingProvince, pcCustSession_ShippingPostalCode, pcCustSession_ShippingCountryCode, pcCustSession_ShippingPhone, pcCustSession_ShippingResidential, pcCustSession_ShippingNickName, pcCustSession_TaxShippingAlone, pcCustSession_TaxShippingAndHandlingTogether, pcCustSession_TaxCountyCode, pcCustSession_TaxProductAmount, customers.IDRefer, pcCustSession_RewardsBalance, pcCustSession_IdPayment, pcCustSession_OrdPackageNumber, pcCustSession_ShippingArray, pcCustSession_Comment, pcCustomerSessions.idcustomer, pcCustomerSessions.idDbSession, pcCustomerSessions.randomKey, pcCustSession_ShippingReferenceId, pcCustSession_ShippingFax, pcCustSession_ShippingEmail,pcCustSession_VATTotal,pcCustSession_IdPayment,pcCustSession_chkPayment,pcCustSession_discountcode,pcCustSession_intCodeCnt,pcCustSession_discountAmount,pcCustSession_taxDetailsString,pcCustSession_total,pcCustSession_GWTotal,pcCustSession_taxAmount, pcCustSession_SB_taxAmount, pcCustSession_RewardsDollarValue,pcCustSession_pSubTotal,pcCustSession_NullShipper,pcCustSession_NullShipRates,pcCustSession_DiscountCodeTotal,pcCustSession_UseRewards,pcCustSession_TF1,pcCustSession_DF1,pcCustSession_OrderName,pcCustSession_CatDiscTotal,pcCustSession_CartRewards,pcCustSession_GcReName,pcCustSession_GcReEmail,pcCustSession_GcReMsg,pcCustSession_strBundleArray,pcCustSession_ShowShipAddr,pcCustSession_GCDetails,pcCustSession_GCTotal FROM customers INNER JOIN pcCustomerSessions ON customers.idcustomer = pcCustomerSessions.idCustomer WHERE (((pcCustomerSessions.idcustomer)="&session("idCustomer")&") AND ((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&"));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pFirstName=getUserInput(rs("name"),0)
pLastName=getUserInput(rs("lastName"),0)
pCustomerCompany=getUserInput(rs("customerCompany"),100)
pPhone=rs("phone")
pEmail=rs("email")
pAddress=getUserInput(rs("address"),0)
pZip=rs("zip")
pStateCode=rs("stateCode")
pState=getUserInput(rs("state"),0)
pCity=getUserInput(rs("city"),0)
pCountryCode=rs("countryCode")
pAddress2=getUserInput(rs("address2"),0)
pFax=rs("fax")
pShippingFirstName=getUserInput(rs("pcCustSession_ShippingFirstName"),0)
pShippingLastName=getUserInput(rs("pcCustSession_ShippingLastName"),0)
pShippingCompany=getUserInput(rs("pcCustSession_ShippingCompany"),0)
pShippingAddress=getUserInput(rs("pcCustSession_ShippingAddress"),0)
pShippingAddress2=getUserInput(rs("pcCustSession_ShippingAddress2"),0)
pShippingCity=getUserInput(rs("pcCustSession_ShippingCity"),0)
pShippingStateCode=getUserInput(rs("pcCustSession_ShippingStateCode"),0)
pShippingState=getUserInput(rs("pcCustSession_ShippingProvince"),0)
pShippingZip=getUserInput(rs("pcCustSession_ShippingPostalCode"),0)
pShippingCountryCode=getUserInput(rs("pcCustSession_ShippingCountryCode"),0)
if pShippingCountryCode="" OR IsNull(pShippingCountryCode) then
	pShippingCountryCode=pCountryCode
end if
pShippingPhone=getUserInput(rs("pcCustSession_ShippingPhone"),0)
if isNULL(rs("pcCustSession_ShippingResidential")) then
	pOrdShipType=1
else
	pOrdShipType=getUserInput(rs("pcCustSession_ShippingResidential"),0)
end if
pShippingNickName=getUserInput(rs("pcCustSession_ShippingNickName"),0)
pTaxShippingAlone=getUserInput(rs("pcCustSession_TaxShippingAlone"),0)
pTaxShipppingAndHandlingTogether=getUserInput(rs("pcCustSession_TaxShippingAndHandlingTogether"),0)
pTaxCountyCode=getUserInput(rs("pcCustSession_TaxCountyCode"),0)
pTaxProductAmount=getUserInput(rs("pcCustSession_TaxProductAmount"),0)
pIDRefer=rs("IDRefer")
pRewardsBalance=rs("pcCustSession_RewardsBalance")
pOrdPackageNum=rs("pcCustSession_OrdPackageNumber")
pShipping=rs("pcCustSession_ShippingArray")
if NOT isNULL(pShipping) AND pShipping<>"" then
	pShipping=getUserInput(pShipping,0)
end if
pComments=getUserInput(rs("pcCustSession_Comment"),0)
pShippingReferenceId=rs("pcCustSession_ShippingReferenceId")
pShippingFax=getUserInput(rs("pcCustSession_ShippingFax"),0)
pShippingEmail=getUserInput(rs("pcCustSession_ShippingEmail"),0)
pShippingFullName=pShippingFirstName& " "&pShippingLastName

'discountcode, chkPayment, discountc, idPayment, taxDetailsString, VATTotal, discountAmount
pVATTotal=getUserInput(rs("pcCustSession_VATTotal"),0)
pIdPayment=getUserInput(rs("pcCustSession_IdPayment"),0)
chkPayment=getUserInput(rs("pcCustSession_chkPayment"),0)
pDiscountCode=rs("pcCustSession_discountcode")
pDiscountEntry=rs("pcCustSession_discountcode")
intCodeCnt=getUserInput(rs("pcCustSession_intCodeCnt"),0)
discountAmount=getUserInput(rs("pcCustSession_discountAmount"),0)
ptaxDetailsString=getUserInput(rs("pcCustSession_taxDetailsString"),0)

pTotal=getUserInput(rs("pcCustSession_total"),0)

'GGG Add-on start
pGWTotal=getUserInput(rs("pcCustSession_GWTotal"),0)
'GGG Add-on end

pTaxAmount=getUserInput(rs("pcCustSession_taxAmount"),0)
pSBTaxAmount=getUserInput(rs("pcCustSession_SB_taxAmount"),0)

pcSFRewardsDollarValue=getUserInput(rs("pcCustSession_RewardsDollarValue"),0)

If pcSFRewardsDollarValue<>"" then
	piRewardValue = pcSFRewardsDollarValue
	pcSFRewardsDollarValue=""
End if

pSFSubTotal=getUserInput(rs("pcCustSession_pSubTotal"),0)

pcNullShipper=getUserInput(rs("pcCustSession_NullShipper"),0)
pcNullShipRates=getUserInput(rs("pcCustSession_NullShipRates"),0)
pcDiscountCodeTotal=getUserInput(rs("pcCustSession_DiscountCodeTotal"),0)
pcUseRewards=getUserInput(rs("pcCustSession_UseRewards"),0)
pcTF1=getUserInput(rs("pcCustSession_TF1"),0)
pcDF1=getUserInput(rs("pcCustSession_DF1"),0)
pcOrderName=getUserInput(rs("pcCustSession_OrderName"),0)
pcCatDiscTotal=getUserInput(rs("pcCustSession_CatDiscTotal"),0)
pcCartRewards=getUserInput(rs("pcCustSession_CartRewards"),0)
pcGcReName=getUserInput(rs("pcCustSession_GcReName"),0)
pcGcReEmail=getUserInput(rs("pcCustSession_GcReEmail"),0)
pcGcReMsg=getUserInput(rs("pcCustSession_GcReMsg"),0)
pcstrBundleArray=rs("pcCustSession_strBundleArray")
pcShowShipAddr=rs("pcCustSession_ShowShipAddr")
if IsNull(pcShowShipAddr) OR pcShowShipAddr="" then
	pcShowShipAddr=0
end if

GCDetails=rs("pcCustSession_GCDetails")
GCAmount=rs("pcCustSession_GCTotal")
if IsNull(GCAmount) OR GCAmount="" then
	GCAmount=0
end if

set rs=nothing

'// Check if the Customer is European Union
Dim pcv_IsEUMemberState
pcv_IsEUMemberState = pcf_IsEUMemberState(pShippingCountryCode)

call closedb()

pDiscountUsed=""
if pZip = "" then
	pZip = "NA"
end if
if pShippingZip = "" then
	pShippingZip = "NA"
end if
if pOrdShipType="0" then 'flagged as commercial
	pOrdShipType=1 'enter it as commercial
else
	pOrdShipType=0 'residential (default)
end if
if pOrdPackageNum="" then 'flagged as commercial
	pOrdPackageNum=1 'enter it as commercial
end if

if session("ExpressCheckoutPayment")="YES" then
	pShippingReferenceID="0"
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: Save Recipient
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If pShippingReferenceId="0" Then '// This is a new address location, or the existing default location.

	'// Check if they have already have a default location. If not, save it now.
	call opendb()
	query="SELECT customers.shippingCity FROM customers WHERE idCustomer=" &session("idCustomer")& ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if NOT rs.eof then
		pcv_tmpShippingCity = rs("shippingCity")
		if (pcv_tmpShippingCity="") OR (isNULL(pcv_tmpShippingCity)=True) then
			query="UPDATE customers SET shippingAddress='" & pShippingAddress & "', shippingCity='" & pShippingCity & "', shippingState='" & pShippingState & "', shippingStateCode='" & pShippingStateCode & "', shippingZip='" & pShippingZip & "', shippingCountryCode='" & pShippingCountryCode & "', shippingCompany='" & pShippingCompany & "', shippingAddress2='" & pShippingAddress2 & "', shippingEmail='" &pShippingEmail& "', shippingPhone='" &pShippingPhone& "',shippingFax='" & pShippingFax & "' WHERE IDCustomer=" & session("idCustomer") &";"
			set rsShippingObj=server.CreateObject("ADODB.RecordSet")
			set rsShippingObj=conntemp.execute(query)
			set rsShippingObj=nothing
		end if
	end if
	set rs=nothing
	call closedb()

Else '// This is an existing location

	'// They are a current customer, do nothing.

End If

'// A nickname was specified. Check if its an existing location, or a new location.
If session("OPCstep")="" then 'The OPC no need this step
	If pShippingNickName<>"" Then

		call opendb()
		query="SELECT * FROM recipients WHERE recipient_NickName='"&pShippingNickName&"' AND idCustomer="&session("idCustomer")&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if rs.eof then
			'// Insert new record
			query="INSERT INTO recipients (idCustomer, recipient_FullName,recipient_FirstName,recipient_LastName,recipient_Company,recipient_Address,recipient_Address2,recipient_City,recipient_StateCode,recipient_State,recipient_Zip,recipient_CountryCode,recipient_phone,recipient_NickName, recipient_Fax, recipient_Email) VALUES ("&session("idCustomer")&",'"&pShippingFirstName&" "&pShippingLastName&"', '"&pShippingFirstName&"','"&pShippingLastName&"','"&pShippingCompany&"','"&pShippingAddress&"','"&pShippingAddress2&"','"&pShippingCity&"','"&pShippingStateCode&"','"&pShippingState&"','"&pShippingZip&"','"&pShippingCountryCode&"','"&pShippingPhone&"','"&pShippingNickName&"','"&pShippingFax&"','"&pShippingEmail&"');"
		else
			'// Update
			query="UPDATE recipients SET recipient_FullName='"&pShippingFirstName&" "&pShippingLastName&"', recipient_FirstName='"&pShippingFirstName&"',recipient_LastName='"&pShippingLastName&"',recipient_Company='"&pShippingCompany&"',recipient_Address='"&pShippingAddress&"',recipient_Address2='"&pShippingAddress2&"',recipient_City='"&pShippingCity&"',recipient_StateCode='"&pShippingStateCode&"',recipient_State='"&pShippingState&"',recipient_Zip='"&pShippingZip&"',recipient_CountryCode='"&pShippingCountryCode&"',recipient_phone='"&pShippingPhone&"',recipient_NickName='"&pShippingNickName&"',recipient_Fax='"&pShippingFax&"', recipient_Email='"&pShippingEmail&"' WHERE recipient_NickName='"&pShippingNickName&"' AND idCustomer="&session("idCustomer")&";"
		end if

		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		call closedb()

	End If
End if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: Save Recipient
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

session("idOrderConfirm")=""

'// AFFILIATES - START
	'// Retrieve affiliate ID from session
	pIdAffiliate=session("idAffiliate")
	if pIdAffiliate="" then
		pIdAffiliate=1
	end if

	IF pIdAffiliate<>1 THEN
		'// Determine whether the affiliate should be associated with this order
		pcInt_AllowedAffOrders=session("pcInt_AllowedAffOrders")
		'// If 0, then unlimited orders are allowed.
		If pcInt_AllowedAffOrders = 0 then
			pcInt_AffiliateOK = 1
		else
			'// Find out if the same affiliate has referred this customer before
			query="SELECT idOrder FROM orders WHERE idAffiliate="&pIdAffiliate&" AND idCustomer="&session("idCustomer")&" AND orderStatus>1 AND orderStatus<>5"
			call opendb()
			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
			totalAffiliateOrders=clng(rs.RecordCount)
			set rs=nothing
			call closedb()
			'// Check the number of orders against the max that the affiliate can earn
			'// commissions on
			if clng(totalAffiliateOrders) <= clng(pcInt_AllowedAffOrders) then
				pcInt_AffiliateOK = 1
			else
				pcInt_AffiliateOK = 0
			end if
		end if

		'// Check for customer type and exclude wholesale customers, if feature is active
		Dim pcInt_ExcludeWholesaleAff
		pcInt_ExcludeWholesaleAff = scExcludeWholesaleAff
		if pcInt_ExcludeWholesaleAff="" or not validNum(pcInt_ExcludeWholesaleAff) then pcInt_ExcludeWholesaleAff = 1
		if session("customerType")=1 and scExcludeWholesaleAff="1" then pcInt_AffiliateOK = 0

		if pcInt_AffiliateOK=0 then
			pIdAffiliate=1
		end if

		'// START Troubleshooting Area: write useful affiliate variables to the page
		'response.write "totalAffiliateOrders="&totalAffiliateOrders&"<br>"
		'response.write "pcInt_AffiliateOK="&pcInt_AffiliateOK&"<br>"
		'response.write "pcInt_ExcludeWholesaleAff="&pcInt_ExcludeWholesaleAff&"<br>"
		'response.write "customerType="&session("customerType")&"<br>"
		'response.write "pIdAffiliate="&pIdAffiliate
		'response.End()
		'// END Troubleshooting Area: write useful affiliate variables to the page

	END IF
'// AFFILIATES - END

pGeneratePassword=0

' check for some hidden fields - Does not have them in OPC
' or (pIdPayment<>999999 AND pPhone="") - no in OPC
if pFirstName="" or pEmail="" then
	response.redirect "onepagecheckout.asp?msg=1"
end if


pDetails=Cstr("")

'// ===============================================
'// START - Calculate total price of the order,
'// total weight and product total quantities
'// ===============================================

pSubtotal=ccur(calculateCartTotal(pcCartArray, ppcCartIndex))
pCartTotalWeight=int(calculateCartWeight(pcCartArray, ppcCartIndex))
pCartQuantity=int(calculateCartQuantity(pcCartArray, ppcCartIndex))
pShipWeight=Cdbl(calculateShipWeight(pcCartArray, ppcCartIndex))
pAffiliateSubTotal=pSubtotal

if ccur(pSFSubTotal)<>pSubtotal then
	response.redirect "onepagecheckout.asp?msg=1"
end if
pSFSubTotal=""

'// ===============================================
'// END - Calculate total price of the order, ...
'// ===============================================


'// ===============================================
'// START - Calculate referred Reward Points
'// ===============================================

	'// Make sure this code only runs once to avoid issues
	IF RewardsReferral=1 AND session("RefRewardPointsTest")="" THEN
	pRewardReferral=0
	pRewardRefId=0

	' Validate that ID of referring customer is a number
	If not validNum(Session("ContinueRef")) then Session("ContinueRef")=""

	' ID validated, proceed with additional tests
	If Session("ContinueRef")<>"" then
	call opendb()

		' Prepare info to save in "Admin Comments" field
		Dim pcvStrReasonNoRPF, pcvIntReferringCustID
		pcvStrReasonNoRPF=""
		pcvIntReferringCustID=Session("ContinueRef")

		' START - Customers cannot refer themselves
			' Check the customer ID
			If Session("ContinueRef") = session("idCustomer") then
				Session("ContinueRef") = 0
				pcvStrReasonNoRPF="Customer ID " & pcvIntReferringCustID & " referred this order. However, referral points cannot be awarded to him/her because the customer is referring himself/herself (same customer ID)."
			end if

			' New customer could be a guest (new ID), so validate based on the e-mail address
			Dim rsCustCheck
			query="SELECT email FROM customers WHERE idCustomer = " & Session("ContinueRef")
			set rsCustCheck=server.CreateObject("ADODB.RecordSet")
			set rsCustCheck=conntemp.execute(query)
			if not rsCustCheck.eof then
				pcCustCheckEmail=trim(rsCustCheck("email"))
			end if
			if pEmail = pcCustCheckEmail then
				Session("ContinueRef") = 0
				pcvStrReasonNoRPF="Customer ID " & pcvIntReferringCustID & " referred this order. However, referral points cannot be awarded to him/her because the customer is referring himself/herself (same e-mail address)."
			end if
			set rsCustCheck=nothing
		' END - Customers cannot refer themselves

		' START - Referral points are only awarded on the first order
			' Check to see if the customer that's checking out has already ordered
			Dim rsOrdersCheck, tmpCountOrdersCheck
			' Ignore the current order being saved
			query="SELECT idorder FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND idCustomer = " & session("idCustomer")
			set rsOrdersCheck=server.CreateObject("ADODB.RecordSet")
			set rsOrdersCheck=conntemp.execute(query)
			if not rsOrdersCheck.eof then
				tmpCountOrdersCheck=0
				do while NOT rsOrdersCheck.eof
					tmpCountOrdersCheck = tmpCountOrdersCheck + 1
				rsOrdersCheck.movenext
				loop
				if tmpCountOrdersCheck>0 then
					Session("ContinueRef") = 0
					pcvStrReasonNoRPF="Customer ID " & pcvIntReferringCustID & " referred this order. However, referral points cannot be awarded to him/her because this was not the referred customer's first order."
				end if
			end if
			set rsOrdersCheck=nothing
		' END - Referral points are only awarded on the first order

		If Session("ContinueRef") > 0 Then
			pRewardRefId=Session("ContinueRef")
			If RewardsFlat=1 Then
				pRewardReferral=RewardsFlatValue
			End If
			If RewardsPerc=1 Then
				pRewardReferral=(pSubtotal * (RewardsPercValue / 100))
			End If
		End If

	call closedb()
	End if
	' This code should not run again
	session("RefRewardPointsTest")="DONE"
	END IF

	IF RewardsReferral<>1 then
		pRewardReferral=0
		pRewardRefId=0
	end if

	If pRewardRefId&""="" Then
		pRewardReferral=0
		pRewardRefId=0
	End If
'// ===============================================
'// END - Calculate referred Reward Points
'// ===============================================

'<<==== create date order var====>>
pDateOrder=Date()
if SQL_Format="1" then
	pDateOrder=Day(pDateOrder)&"/"&Month(pDateOrder)&"/"&Year(pDateOrder)
else
	pDateOrder=Month(pDateOrder)&"/"&Day(pDateOrder)&"/"&Year(pDateOrder)
end if

'<<==== prevent state and province to both be saved to the DB ====>>
If pStateCode <> "" and (pCountryCode="US" or pCountryCode="CA") then
	pState=""
end if
If pShippingStateCode <> "" and (pShippingCountryCode="US" or pShippingCountryCode="CA") then
	pShippingState=""
end if

call openDb()

' compile order details memo field
for f=1 to ppcCartIndex
	' if item is not deleted from cart
	if pcCartArray(f,10) = 0 then
		tempAmt=Cdbl( pcCartArray(f,2) * (pcCartArray(f,5)+pcCartArray(f,3)) )
		if scDecSign="," then
			tempAmt=replace(tempAmt,",",".")
		end if
		pDetails	= pDetails & "  Amount: ||"& tempAmt & " Qty:" &pcCartArray(f,2)& "  SKU #:" &pcCartArray(f,7) & " - " &pcCartArray(f,1)& " " & pcCartArray(f,4) & Vbcrlf
		pDetails = replace(pDetails,"'","''")
		pDetails=replace(pDetails,"''''","''")
	end if ' item deleted
next

' get shipment data
If pcNullShipper="Yes" then
	pShipmentDesc=ship_dictLanguage.Item(Session("language")&"_noShip_a")
	pShipmentPriceToAdd="0"
else
	if pcNullShipRates="Yes" then
		pShipmentDesc=ship_dictLanguage.Item(Session("language")&"_noShip_b")
		pShipmentPriceToAdd="0"
	else
		shipping=split(pShipping,",")

		Shipper=shipping(0)
		Service=shipping(1)
		Postage=shipping(2)
		pShipmentDesc=Shipper&" - "&Service
		pShipmentPriceToAdd=Postage

		if ubound(shipping)=>3 then
			pserviceHandlingFee=shipping(3)
		else
			pserviceHandlingFee="0"
		end if
	end if
end if

' calculate shipment price
if pShipmentPriceToAdd>0 then
	shipmentTotal=pShipmentPriceToAdd
end if
pPSubtotal=pSubtotal
pSubtotal=pSubtotal + shipmentTotal+pserviceHandlingFee
pSRF="0"
If pcNullShipper="Yes" then
	pShipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
else
	if pcNullShipRates="Yes" then
		pSRF="1"
		pShipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_b")
	else
		pShipmentDetails=pShipping
		pShipmentDetails=replace(pShipmentDetails,"<SUP>SM</SUP>","")
		pShipmentDetails=replace(pShipmentDetails,"<SUP>&reg;</SUP>","")
		pShipmentDetails=replace(pShipmentDetails,"&reg;","")
	end if
end if

if chkPayment="FREE" then
	pPaymentDetails = "FREE || 0.00"
else
	query="SELECT gwCode FROM paytypes WHERE idPayment=" &pIdPayment
	set rsCPP=server.CreateObject("ADODB.RecordSet")
	set rsCPP=conntemp.execute(query)
	if NOT rsCPP.eof then
		pgwCode = rsCPP("gwCode")
	end if
	set rsCPP = Nothing
	if (pgwCode=999999) OR (session("ExpressCheckoutPayment")="YES") then
		pPaymentDetails = "PayPal Express Checkout || 0.00"
		pPaymentDesc="PayPal Express Checkout"
		pPaymentPriceToAdd=0
		pPaymentpercentageToAdd=0
		If session("ExpressPayPPA") = "YES" OR session("ExpressPayPPL") = "YES" Then
			if session("ExpressPayPPA") = "YES" then
				pSslUrl="gwExpressPPADo.asp"
			else
				pSslUrl="gwExpressPPLDo.asp"
			end if
		Else
		'// Determine which API to use (US or UK)
		query="SELECT pcPay_PayPal.pcPay_PayPal_Partner, pcPay_PayPal.pcPay_PayPal_Vendor FROM pcPay_PayPal WHERE (((pcPay_PayPal.pcPay_PayPal_ID)=1));"
		set rsPayPalType=Server.CreateObject("ADODB.Recordset")
		set rsPayPalType=conntemp.execute(query)
		pcPay_PayPal_Partner=rsPayPalType("pcPay_PayPal_Partner")
		pcPay_PayPal_Vendor=rsPayPalType("pcPay_PayPal_Vendor")
		if isNULL(pcPay_PayPal_Partner)=True then pcPay_PayPal_Partner=""
		if isNULL(pcPay_PayPal_Vendor)=True then pcPay_PayPal_Vendor=""
		if pcPay_PayPal_Partner<>"" AND pcPay_PayPal_Vendor<>"" then
			pcPay_PayPal_Version = "UK"
		else
			pcPay_PayPal_Version = "US"
		end if
		set rsPayPalType=nothing
		if pcPay_PayPal_Version = "US" then
			pSslUrl="gwExpressDo.asp"
		else
			pSslUrl="gwExpressUKDo.asp"
			end if
		end if
	else
		query="SELECT paymentDesc,priceToAdd,percentageToAdd,sslUrl FROM paytypes WHERE idPayment=" &pIdPayment
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
			response.redirect "techErr.asp?error="&Server.Urlencode("Could not locate selected paytypes in database")
		end if

		pPaymentDesc=rs("paymentDesc")
		if pPaymentDesc = "PxPay" Then
			pPaymentDesc = "DPS"
		end if
		pPaymentPriceToAdd=rs("priceToAdd")
		pPaymentpercentageToAdd=rs("percentageToAdd")
		pSslUrl=rs("sslUrl")
		set rs=nothing

		' calculate payment price
		if Cdbl(pPaymentPriceToAdd)<>0 or Cdbl(pPaymentpercentageToAdd)<>0 then
			tempPercentageToAdd=(pPaymentpercentageToAdd*pPSubtotal/100)
			tempPercentageToAdd=roundTo(tempPercentageToAdd,.01)
			paymentTotal=pPaymentPriceToAdd + tempPercentageToAdd
		end if

		pSubtotal=pSubtotal + paymentTotal

		pPaymentDetails = pPaymentDesc & " || "& paymentTotal
		if scDecSign="," then
			pPaymentDetails=replace(pPaymentDetails,",",".")
		end if
	end if
end if

if pDiscountCode<>"" OR GCDetails<>"" OR Session("pcPromoIndex")<>"" and Session("pcPromoIndex")>"0" then
	'GGG Add-on start
	myTest=0
	'GGG Add-on end
	pDiscountDetails=Cstr("")
	discountTotal=Cdbl(0)
	discountApplied=0

	'GGG Add-on start
	if instr(pDiscountCode,",")>0 then
		myTest=0
	else
		query="SELECT quantityFrom,quantityUntil,weightFrom,weightUntil,priceFrom,priceUntil,idDiscount,oneTime,discountDesc,priceToDiscount,percentageToDiscount FROM discounts WHERE discountcode='" &pDiscountCode &"'"
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=conntemp.execute(query)

		if err.number <> 0 then
			call closeDb()
			response.redirect "techErr.asp?error="&Server.Urlencode("Error in saveorder 6. Error: "&err.description)
		end if
		if rstemp.eof then
			myTest=1
			pDiscountDetails=dictLanguage.Item(Session("language")&"_saveorder_10") & " - || 0"
		end if
	end if
	'GGG Add-on end

	'GGG Add-on start
	'There are discount code(s)
	IF myTest=0 THEN
		'GGG Add-on end

		if intCodeCnt>0 then
			DiscountCodeArry=Split(pDiscountCode,",")
			DiscountAmountArry=split(discountAmount,",")
		end if

		dim intDiscountUsedCnt
		intDiscountUsedCnt=0
		intDiscountArryCnt=0
		For i=0 to intCodeCnt
			if intCodeCnt=0 then
				pTempDiscCode=pDiscountCode
				pTempDiscAmount=discountAmount
			else
				pTempDiscCode=DiscountCodeArry(i)
				pTempDiscAmount=DiscountAmountArry(i)
			end if

			query="SELECT quantityFrom,quantityUntil,weightFrom,weightUntil,priceFrom,priceUntil,idDiscount,oneTime,discountDesc,priceToDiscount,percentageToDiscount FROM discounts WHERE discountcode='" &pTempDiscCode&"'"
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
				if intCodeCnt=0 then
					pDiscountDetails=dictLanguage.Item(Session("language")&"_saveorder_10") & " - || 0"
				end if
			else
				intQuantityFrom=rs("quantityFrom")
				intQuantityUntil=rs("quantityUntil")
				intWeightFrom=rs("weightFrom")
				intWeightUntil=rs("weightUntil")
				dblPriceFrom=rs("priceFrom")
				dblPriceUntil=rs("priceUntil")
				pIdDiscount=rs("idDiscount")
				pOneTime=rs("oneTime")
				pDiscountDesc=rs("discountDesc")
				pPriceToDiscount=rs("priceToDiscount")
				pPercentageToDiscount=rs("percentageToDiscount")

				if pCartQuantity>=intQuantityFrom and pCartQuantity<=intQuantityUntil and pCartTotalWeight>=intWeightFrom and pCartTotalWeight<=intWeightUntil and pSubtotal>=dblPriceFrom and pSubtotal<=dblPriceUntil then
				' update discount used for customer
					if pOneTime=-1 then
						'Create Session
						if intDiscountUsedCnt=0 then
							pDiscountUsed=pDiscountUsed&pIdDiscount
							intDiscountUsedCnt=intDiscountUsedCnt+1
							intDiscountArryCnt=intDiscountArryCnt+1
						else
							pDiscountUsed=pDiscountUsed&","&pIdDiscount
							intDiscountArryCnt=intDiscountArryCnt+1
						end if
					end if ' one time
				else
					pDiscountDetails=dictLanguage.Item(Session("language")&"_saveorder_11") & " - || 0"
					intDiscountArryCnt=intDiscountArryCnt+1
				end if
			end if

			if pPriceToDiscount>0 or ppercentageToDiscount>0 then
				discountTotal=pcDiscountCodeTotal
			end if

			if intDiscountArryCnt=0 then
				pDiscountDetails=pDiscountDetails&pDiscountDesc & " ("&pTempDiscCode&") - || "& pTempDiscAmount
				intDiscountArryCnt=intDiscountArryCnt+1
				discountApplied=1
			else
				pDiscountDetails=pDiscountDetails&","&pDiscountDesc & " ("&pTempDiscCode&") - || "& pTempDiscAmount
				intDiscountArryCnt=intDiscountArryCnt+1
				discountApplied=1
			end if

		Next
		if len(pDiscountDetails)>0 then
			if left(pDiscountDetails,1)="," then
				pDiscountDetails=Mid(pDiscountDetails, 2)
			end if
		end if
	END IF

	discountTotal=Cdbl(0)

	if pPriceToDiscount>0 or ppercentageToDiscount>0 then
		discountTotal=pcDiscountCodeTotal
	end if

	discountTotal=discountTotal+cdbl(GCAmount)

	TotalPromotions=Session("PromotionTotal")
	discountTotal=discountTotal+cdbl(TotalPromotions)

	if TotalPromotions>"0" then
		PromoArr1=Session("pcPromoSession")
		PromoIndex=Session("pcPromoIndex")
		if pDiscountDetails=dictLanguage.Item(Session("language")&"_saveorder_10") & " - || 0" OR pDiscountDetails=dictLanguage.Item(Session("language")&"_saveorder_11") & " - || 0" THEN
			pDiscountDetails=""
		end if
		For m=1 to PromoIndex
			if intDiscountArryCnt=0 then
				pDiscountDetails=pDiscountDetails & PromoArr1(m,3) & " - || " & PromoArr1(m,2)
				intDiscountArryCnt=intDiscountArryCnt+1
			else
				pDiscountDetails=pDiscountDetails & "," & PromoArr1(m,3) & " - || " & PromoArr1(m,2)
				intDiscountArryCnt=intDiscountArryCnt+1
			end if
		Next
	end if

	pSubtotal=pSubtotal - discountTotal

	if discountApplied=0 AND (Session("pcPromoIndex")="" OR Session("pcPromoIndex")="0") then
		pDiscountDetails=dictLanguage.Item(Session("language")&"_saveorder_10")
	end if

	'pDiscountCode is Null
else
	pDiscountDetails=dictLanguage.Item(Session("language")&"_saveorder_10")
end if

pAffiliateValid =Cint(1)
pAffiliatePay=0
if pIdAffiliate<>1 then
	query="SELECT commission FROM affiliates WHERE idAffiliate=" &pIdAffiliate
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rs.eof then
	pAffiliateValid=0
	else
		'calculate affiliatePay
		pAffiliateSubTotal=pAffiliateSubTotal - discountTotal
		pAffiliatePay=pAffiliateSubTotal * (rs("commission")/100)
	end if
	set rs=nothing
end if

if pAffiliateValid=0 then
	pIdAffiliate=1
	pAffiliatePay=0
end if

If piRewardValue="" then
	piRewardValue="0"
End If
If pcUseRewards="" OR IsNull(pcUseRewards) then
	pcUseRewards="0"
End If
if pcUseRewards="0" then
	piRewardValue="0"
end if

' save order temporarily
IDrefer=pIDRefer
if isNull(IDrefer) OR IDrefer="" then
	IDrefer="0"
end if

pord_DeliveryDate=pcDF1
if pord_DeliveryDate<>"" then
	if isDate(pord_DeliveryDate) then
		if SQL_Format="1" then
			expDateArray=split(pord_DeliveryDate,"/")
			pord_DeliveryDate=(expDateArray(1)&"/"&expDateArray(0)&"/"&expDateArray(2))
		end if
	end if
end if
pord_DeliveryDate=pord_DeliveryDate & " " & pcTF1
pord_DeliveryDate=trim(pord_DeliveryDate)
if not isDate(pord_DeliveryDate) then
	pord_DeliveryDate=""
end if

if pcOrderName<>"" then
pord_OrderName=pcOrderName
else
pord_OrderName=""
end if

pcv_CatDiscounts=pcCatDiscTotal
if isNull(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
	pcv_CatDiscounts="0"
end if

'Create query string
strUpdateQuery="UPDATE orders SET pcOrd_GCDetails='" & replace(GCDetails,"'","''") & "',pcOrd_GCAmount=" & GCAmount & ",IDrefer=" & IDrefer & ","

if scDB="SQL" then
	strUpdateQuery=strUpdateQuery&"orderDate='" & pDateOrder  & "',"
else
	strUpdateQuery=strUpdateQuery&"orderDate=#" & pDateOrder  & "#,"
end if

strUpdateQuery=strUpdateQuery&"idCustomer=" & int(Session("idCustomer"))& ", details='" &pDetails &"', total=" &DecimalFormatter(pTotal)& ", taxAmount=" &DecimalFormatter(pTaxAmount)& ", comments='" &pComments& "', address='" &paddress & "', zip='" &pzip& "',state='" &pState& "',stateCode='" &pStateCode& "',city='" &pCity& "',CountryCode='" &pCountryCode& "',shippingAddress='" &pShippingAddress & "',shippingZip='" &pShippingZip& "',shippingState='" &pShippingState& "',shippingStateCode='" &pShippingStateCode& "', shippingCity='" &pShippingCity& "', shippingCountryCode='" &pShippingCountryCode& "',shipmentDetails='" &pShipmentDetails& "', paymentDetails='" &replace(pPaymentDetails,"'","''")& "',discountDetails='" &replace(pDiscountDetails,"'","''")& "',randomNumber=" &session("pcSFIdDbSession")& ",orderStatus=1,pcOrd_shippingPhone=' " &pShippingPhone& "',idAffiliate=" &pIdAffiliate& ", affiliatePay=" &DecimalFormatter(pAffiliatePay)&",shippingFullName='"&pShippingFullName&"', iRewardPoints="&pcUseRewards&",iRewardValue= " &piRewardValue&", iRewardPointsCustAccrued=" & pcCartRewards & ", address2='" &paddress2 & "', shippingCompany='" &pShippingCompany & "', shippingAddress2='" &pShippingAddress2 & "',taxDetails='"&ptaxDetailsString&"',SRF="&pSRF&",ordShipType="&pOrdShipType&", ordPackageNum="&pOrdPackageNum&", ord_OrderName='"&pord_OrderName&"'"

if DFShow="1"  and pord_DeliveryDate <> "" then
	if scDB="SQL" then
		strUpdateQuery=strUpdateQuery&",ord_DeliveryDate='" & pord_DeliveryDate  & "'"
	else
		strUpdateQuery=strUpdateQuery&",ord_DeliveryDate=#" & pord_DeliveryDate  & "#"
	end if
end if

'SB S
If pSBTaxAmount>0 Then
	strUpdateQuery=strUpdateQuery&",pcOrd_SubTax=" & DecimalFormatter(pSBTaxAmount) '// Tax
	strUpdateQuery=strUpdateQuery&",pcOrd_SubTrialTax=" & DecimalFormatter(pTaxAmount) '// Trial Tax
Else
	strUpdateQuery=strUpdateQuery&",pcOrd_SubTax=" & DecimalFormatter(pTaxAmount) '// Tax
	strUpdateQuery=strUpdateQuery&",pcOrd_SubTrialTax=" & 0 '// Trial Tax
End If
strUpdateQuery=strUpdateQuery&",pcOrd_SubShipping=" & cdbl(shipmentTotal) + cdbl(pserviceHandlingFee) '// Shipping
strUpdateQuery=strUpdateQuery&",pcOrd_SubTrialShipping=" & cdbl(shipmentTotal) + cdbl(pserviceHandlingFee) '// Trial Shipping (same as standard shipping)
'SB E


if pVATTotal="" then
	pVATTotal=0
end if

'GGG Add-on start
if GCAmount<>"" then
else
GCAmount=0
end if

if session("Cust_IDEvent")<>"" then
gIDEvent=session("Cust_IDEvent")
else
gIDEvent="0"
end if

pcv_GcReName=pcGcReName
pcv_GcReEmail=pcGcReEmail
pcv_GcReMsg=pcGcReMsg

if GCAmount=0 then
	pDiscountCode=""
end if
'GGG Add-on end

'Generate Order Key
pcOrderKey=""
TestedOrderKey=0
do while (TestedOrderKey=0)
	pcOrderKey=generateABC(3) & generate123(10)
	query="SELECT idOrder FROM Orders WHERE pcOrd_OrderKey like '" & pcOrderKey & "';"
	set rs=connTemp.execute(query)
	if rs.eof then
		TestedOrderKey=1
	end if
	set rs=nothing
loop

strUpdateQuery=strUpdateQuery&",ord_VAT="&DecimalFormatter(pVATTotal)&",pcord_CatDiscounts=" & pcv_CatDiscounts & ",pcOrd_DiscountsUsed='"&pDiscountUsed&"',pcOrd_GcCode='" & pDiscountCode & "',pcOrd_GcUsed=" & GCAmount & ",pcOrd_GCs=0,pcOrd_IDEvent=" & gIDEvent & ",pcOrd_GWTotal=" & pGWTotal & ",pcOrd_GcReName='" & pcv_GcReName & "',pcOrd_GcReEmail='" & pcv_GcReEmail & "',pcOrd_GcReMsg='" & pcv_GcReMsg & "',pcOrd_shippingFax='"&pShippingFax&"', pcOrd_ShippingEmail='"&pShippingEmail&"', pcOrd_ShipWeight="&pShipWeight&" WHERE idOrder="&session("idOrderSaved")&";"

strInsertQuery="INSERT INTO orders (pcOrd_GCDetails,pcOrd_GCAmount,pcOrd_ShowShipAddr,pcOrd_OrderKey,IDrefer,orderDate,idCustomer, details, total, taxAmount, comments, address, zip, state, stateCode, city, CountryCode, shippingAddress, shippingZip, shippingState, shippingStateCode, shippingCity, shippingCountryCode, shipmentDetails, paymentDetails, discountDetails, randomNumber, orderStatus, pcOrd_shippingPhone, idAffiliate, affiliatePay,shippingFullName, iRewardPoints, iRewardValue,iRewardRefid,iRewardPointsRef,iRewardPointsCustAccrued, address2, shippingCompany, shippingAddress2,taxDetails,SRF,ordShipType, ordPackageNum, ord_OrderName"

if DFShow="1"  and pord_DeliveryDate <> "" then
	strInsertQuery=strInsertQuery&",ord_DeliveryDate"
end if

strInsertQuery=strInsertQuery&",ord_VAT,pcord_CatDiscounts,pcOrd_DiscountsUsed,pcOrd_GcCode,pcOrd_GcUsed,pcOrd_GCs,pcOrd_IDEvent,pcOrd_GWTotal,pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg,pcOrd_shippingFax, pcOrd_ShippingEmail, pcOrd_ShipWeight) VALUES ('" & replace(GCDetails,"'","''") & "'," & GCAmount & "," & pcShowShipAddr & ",'" & pcOrderKey & "'," & IDrefer & ","

if scDB="SQL" then
	strInsertQuery=strInsertQuery&"'" & pDateOrder  & "'"
else
	strInsertQuery=strInsertQuery&"#" & pDateOrder  & "#"
end if

strInsertQuery=strInsertQuery&"," & int(Session("idCustomer"))& ",'" &pDetails &"'," &DecimalFormatter(pTotal)& "," &DecimalFormatter(pTaxAmount)& ",'" &pComments& "','" &paddress & "','" &pzip& "','" &pState& "','" &pStateCode& "','" &pCity& "','" &pCountryCode& "','" &pShippingAddress & "','" &pShippingZip& "','" &pShippingState& "','" &pShippingStateCode& "','" &pShippingCity& "','" &pShippingCountryCode& "','" &pShipmentDetails& "','" &replace(pPaymentDetails,"'","''")& "','" &replace(pDiscountDetails,"'","''")& "'," &session("pcSFIdDbSession")& ",1,' " &pShippingPhone& "'," &pIdAffiliate& ", " &DecimalFormatter(pAffiliatePay)&",'"&pShippingFullName&"', "&pcUseRewards&", " &piRewardValue&", " &pRewardRefId&", " &pRewardReferral&", " &pcCartRewards&", '" &paddress2 & "', '" &pShippingCompany & "', '" &pShippingAddress2 & "','"&ptaxDetailsString&"',"&pSRF&","&pOrdShipType&","&pOrdPackageNum&",'"&pord_OrderName&"'"

if DFShow="1" and pord_DeliveryDate <> "" then
	if scDB="SQL" then
		strInsertQuery=strInsertQuery&",'" & pord_DeliveryDate & "'"
	else
		strInsertQuery=strInsertQuery&",#" & pord_DeliveryDate & "#"
	end if
end if

if pVATTotal="" then
	pVATTotal=0
end if

strInsertQuery=strInsertQuery&","&DecimalFormatter(pVATTotal)&"," & pcv_CatDiscounts & ",'"&pDiscountUsed&"','" & pDiscountCode & "'," & GCAmount & ",0," & gIDEvent & "," & pGWTotal & ",'" & pcv_GcReName & "','" & pcv_GcReEmail & "','" & pcv_GcReMsg & "', '"&pShippingFax&"', '"&pShippingEmail&"', "&pShipWeight&")"

if session("idOrderSaved")<>"" then
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute("SELECT idOrder FROM orders WHERE idOrder="&session("idOrderSaved")&";")
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if NOT rs.eof then
		set rs=conntemp.execute(strUpdateQuery)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		'delete current from ProductsOrdered
		query="DELETE FROM ProductsOrdered WHERE idOrder="&session("idOrderSaved")&";"
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	else
		set rs=conntemp.execute(strInsertQuery)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	end if
	set rs=nothing
else

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(strInsertQuery)
	set rs=nothing
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
end if

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

' get id of the saved order
query="SELECT idOrder FROM orders WHERE randomNumber=" &session("pcSFIdDbSession")& " AND idCustomer=" &session("idCustomer")
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
	response.redirect "techErr.asp?error="&Server.Urlencode("Error in saveorder.asp. Could not locate order.")
end if

pIdorder=rs("idOrder")
session("idOrderSaved")=pIdorder
set rs=nothing

' save ProductsOrdered
strBundleArray=pcstrBundleArray
tmpArrIdx=0
for f=1 to ppcCartIndex
	if pcCartArray(f,10)=0 then

		pcPrdOrd_BundledDisc=0
		if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
			pcvBundleArray=split(strBundleArray, "||")
			for b=tmpArrIdx to ubound(pcvBundleArray)-1
				pcvSplitArray=split(pcvBundleArray(b),",")
				if pcvSplitArray(0)=pcCartArray(f,0) then
					'Save to new fields in database!!
					'response.write "27: "&pcvSplitArray(1)&"<BR>"
					'response.write "28: "&pcvSplitArray(2)&"<BR>"
					pcPrdOrd_BundledDisc=pcvSplitArray(3)*-1
					tmpArrIdx=b+1
					exit for
				end if
			next
		end if
		if pcCartArray(f,11)="" or isNull(pcCartArray(f,11)) then
			pcCartArray(f,11)="NULL"
		end if

		if pcCartArray(f,12)="" or isNull(pcCartArray(f,12)) then
			pcCartArray(f,12)="NULL"
		end if

		if pcCartArray(f,14)="" or isNull(pcCartArray(f,14)) then
			pcCartArray(f,14)=0
		end if

		' replace , by .
		pcCartArray(f,14)=replace(pcCartArray(f,14),",",".")
		if pcCartArray(f,16)<>"" or pcCartArray(f,15)<>"0" then
			tempVar1=(pcCartArray(f,5) + pcCartArray(f,17))
		else
			tempVar1=(pcCartArray(f,5) + pcCartArray(f,3))
		end if

		If pcCartArray(f,16)="" then
			pcCartArray(f,16)=0
		end If

		if pcCartArray(f,15)<>"" then
			QDiscounts=pcCartArray(f,15)
		else
			QDiscounts="0"
		end if
		if pcCartArray(f,30)<>"" then
			ItemsDiscounts=pcCartArray(f,30)
		else
			ItemsDiscounts="0"
		end if

		'SB S
		'// Get BTO additional charges for SB
		if pcCartArray(f,31)<>"" then
			AddCharges=pcCartArray(f,31)
		else
			AddCharges="0"
		end if
		'SB E

		'GGG Add-on start
		if pcCartArray(f,33)<>"" then
		geID=pcCartArray(f,33)
		else
		geID="0"
		end if

		if pcCartArray(f,34)<>"" then
		pGWOpt=pcCartArray(f,34)
		else
		pGWOpt="0"
		end if

		if pcCartArray(f,35)<>"" then
			pGWOptText=Server.HTMLEncode(pcCartArray(f,35))
			pGWOptText=replace(pGWOptText,"'","''")
			if len(pGWOptText)>240 then
				pGWOptText=left(pGWOptText,240)
			end if
		else
			pGWOptText=""
		end if

		if pGWOpt<>"0" then
			query="select pcGW_OptPrice from pcGWOptions where pcGW_IDOpt=" & pGWOpt
			set rsG=server.CreateObject("ADODB.RecordSet")
			set rsG=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsG=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			If NOT rsG.eof Then
				pGWPrice=rsG("pcGW_OptPrice")
			Else
				pGWPrice="0"
			End If
			Set rsG = nothing
			if pGWPrice<>"" then
			else
				pGWPrice="0"
			end if
		else
			pGWPrice="0"
		end if
		'GGG Add-on end

		pcv_xdetails=pcCartArray(f,21)
		if pcv_xdetails<>"" then
			pcv_xdetails=replace(pcv_xdetails,"<br>","|")
			pcv_xdetails=replace(pcv_xdetails,"'","''")
			pcv_xdetails=replace(pcv_xdetails,"''''","''")
		end if

		'Start SDBA
		query="SELECT serviceSpec,stock,nostock,pcProd_BackOrder,pcDropShipper_ID FROM Products WHERE idproduct=" & pcCartArray(f,0)
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		if not rs.eof then
			pcv_serviceSpec=rs("serviceSpec")
			if IsNull(pcv_pserviceSpec) or pcv_pserviceSpec="" then
				pcv_pserviceSpec="0"
			end if
			pcv_Stock=rs("stock")
			if IsNull(pcv_Stock) or pcv_Stock="" then
				pcv_Stock="0"
			end if
			pcv_NoStock=rs("nostock")
			if IsNull(pcv_NoStock) or pcv_NoStock="" then
				pcv_NoStock="0"
			end if
			pcv_intBackOrder=rs("pcProd_BackOrder")
			if IsNull(pcv_intBackOrder) or pcv_intBackOrder="" then
				pcv_intBackOrder="0"
			end if
			pcv_IDDropShipper=rs("pcDropShipper_ID")
			if IsNull(pcv_IDDropShipper) or pcv_IDDropShipper="" then
				pcv_IDDropShipper="0"
			end if
		else
			pcv_pserviceSpec="0"
			pcv_Stock="0"
			pcv_NoStock="0"
			pcv_intBackOrder="0"
			pcv_IDDropShipper="0"
		end if
		set rs=nothing

		If (scOutofStockPurchase=-1 AND CLng(pcv_Stock)<1 AND pcv_serviceSpec=0 AND pcv_NoStock=0 AND pcv_intBackOrder=1) OR (pcv_serviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pcv_Stock)<1 AND pcv_NoStock=0 AND pcv_intBackOrder=1) Then
			tmp_BackOrder="1"
		Else
			tmp_BackOrder="0"
		End if
		'End SDBA

		'Add new record to database if product quantity > 0
		if pcCartArray(f,2)>"0" then

			'SB S
			if pcCartArray(f,38) > 0 Then

				pcv_Subscription_ID = pcCartArray(f,38)

				'// Get the Sub Details
				pSubscriptionID = pcCartArray(f,38)
				%>
				<!--#include file="../includes/pcSBDataInc.asp" -->
				<%
				If NOT pcv_intIsTrial="1" Then
					pcv_SubType = "FULL"
					pcv_SubTrialAmt = 0
				Else
					pcv_SubType = "TRIAL"
					session("pcIsSubTrial") = True
					pcv_SubTrialAmt = ((pcv_curTrialAmount+ AddCharges)-QDiscounts ) - ItemsDiscounts
					pcv_SubTrialAmt = round(pcv_SubTrialAmt, 2)
					If pcv_SubTrialAmt < 0 Then
						pcv_SubTrialAmt = 0
					End If
				End if

				pcv_SubPrice = ((tempVar1+ AddCharges)-QDiscounts) - ItemsDiscounts
				pcv_SubPrice = round(pcv_SubPrice, 2)

				'// If coupon exists we must be in single mode... apply the coupon.
				If discountTotal>0 OR piRewardValue>0 Then
					If pcv_SubType = "TRIAL" Then
						pcv_SubTrialAmt = pcv_SubTrialAmt - discountTotal - piRewardValue
						If pcv_SubTrialAmt < 0 Then
							pcv_SubTrialAmt = 0
						End If
					Else
						pcv_SubPrice = pcv_SubPrice - discountTotal - piRewardValue
					End If
				End If
				'SM-S
				pcSCID=pcCartArray(f,39)
				if IsNull(pcSCID) OR len(pcSCID)=0 then
					pcSCID=0
				end if
				'SM-E

				query="INSERT INTO ProductsOrdered (idOrder, idProduct, quantity, unitPrice, unitCost, idconfigSession, xfdetails, QDiscounts, ItemsDiscounts, pcPackageInfo_ID, pcDropShipper_ID, pcPrdOrd_Shipped, pcPrdOrd_BackOrder, pcPrdOrd_SelectedOptions, pcPrdOrd_OptionsPriceArray, pcPrdOrd_OptionsArray, pcPO_EPID, pcPO_GWOpt, pcPO_GWNote, pcPO_GWPrice, pcSubscription_ID, pcPO_SubDetails, pcPO_SubAmount, pcPO_SubType, pcPO_SubFrequency, pcPO_SubPeriod, pcPO_SubCycles, pcPO_SubStartDate, pcPO_SubTrialFrequency, pcPO_SubTrialPeriod, pcPO_SubTrialCycles, pcPO_SubTrialAmount, pcPO_NoShipping, pcPrdOrd_BundledDisc, pcPO_LinkID, pcPO_IsTrial, pcPO_SubActive,pcSC_ID) VALUES (" &pIdOrder& "," &pcCartArray(f,0)& "," &pcCartArray(f,2)& "," & DecimalFormatter(tempVar1) & "," & DecimalFormatter(pcCartArray(f,14))& "," &pcCartArray(f,16)& ",'" &pcv_xdetails& "'," & QDiscounts & "," & ItemsDiscounts & ",0," & pcv_IDDropShipper & ",0," & tmp_BackOrder & ",'" & replace(pcCartArray(f,11),"'","''") & "','" & replace(pcCartArray(f,25),"'","''") & "','" & replace(pcCartArray(f,4),"'","''") &"'," & geID & "," & pGWOpt & ",'" & pGWOptText & "'," & pGWPrice & ","&pSubscriptionID&",'',"&pcv_SubPrice&",'"&pcv_SubType&"',"&pcv_intBillingFrequency&",'"&pcv_strBillingPeriod&"',"&pcv_intBillingCycles & ",'" & pSubStartDate & "'," & pcv_intTrialFrequency & ", '" & pcv_strTrialPeriod & "'," & pcv_intTrialCycles &"," & pcv_SubTrialAmt &"," &pcCartArray(f,20)&","&pcPrdOrd_BundledDisc&",'"&pcv_strLinkID&"',"&pcv_intIsTrial&",1," & pcSCID & ")"

				set rsSub=server.CreateObject("ADODB.RecordSet")
				set rsSub=conntemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rsSub=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

			else

			'SM-S
			pcSCID=pcCartArray(f,39)
			if IsNull(pcSCID) OR len(pcSCID)=0 then
				pcSCID=0
			end if
			'SM-E

			query="INSERT INTO ProductsOrdered (idOrder, idProduct, quantity, unitPrice, unitCost, idconfigSession, xfdetails, QDiscounts,ItemsDiscounts, pcPackageInfo_ID, pcDropShipper_ID, pcPrdOrd_Shipped, pcPrdOrd_BackOrder, pcPrdOrd_SelectedOptions, pcPrdOrd_OptionsPriceArray, pcPrdOrd_OptionsArray, pcPO_EPID,pcPO_GWOpt, pcPO_GWNote, pcPO_GWPrice,pcPrdOrd_BundledDisc,pcSC_ID) VALUES (" &pIdOrder& "," &pcCartArray(f,0)& "," &pcCartArray(f,2)& "," & DecimalFormatter(tempVar1) & "," & DecimalFormatter(pcCartArray(f,14))& "," &pcCartArray(f,16)& ",'" &pcv_xdetails& "'," & QDiscounts & "," & ItemsDiscounts & ",0," & pcv_IDDropShipper & ",0," & tmp_BackOrder & ",'" & replace(pcCartArray(f,11),"'","''") & "','" & replace(pcCartArray(f,25),"'","''") & "','" & replace(pcCartArray(f,4),"'","''") &"'," & geID & "," & pGWOpt & ",'" & pGWOptText & "'," & pGWPrice & ","&pcPrdOrd_BundledDisc&"," & pcSCID & ")"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing

			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		end if
			'SB E

		end if
	end if
next

'//Check if this customer agreed to the terms of this checkout and save to db
if scTerms = 1 and Session("pcCustomerTermsAgreed")="1" then
	if scDB="SQL" then
		strDtDelim="'"
	else
		strDtDelim="#"
	end if

	query="SELECT pcCustomerTermsAgreed.pcCustomerTermsAgreed_ID, pcCustomerTermsAgreed.idCustomer, pcCustomerTermsAgreed.idOrder  FROM pcCustomerTermsAgreed WHERE pcCustomerTermsAgreed.idCustomer=" & int(Session("idCustomer"))& " AND pcCustomerTermsAgreed.idOrder=" &pIdOrder& ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if rs.eof then
		query="INSERT INTO pcCustomerTermsAgreed (idCustomer, idOrder, pcCustomerTermsAgreed_InsertDate) VALUES (" & int(Session("idCustomer"))& ", " &pIdOrder& ", "&strDtDelim&Now()&strDtDelim&" );"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	end if
	set rs=nothing
end if

'// Update "Admin Comments" with the reason why Reward Points were not awarded to the referring customer (if any)
if pcvStrReasonNoRPF<>"" then
	pcvStrReasonNoRPF=replace(pcvStrReasonNoRPF,"'","''")
	query="UPDATE orders SET adminComments='"&pcvStrReasonNoRPF&"' WHERE idOrder="& pIdOrder
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)
	set rs=nothing
end if

if err.number<>0 then
	call LogErrorToDatabase()
end if
call closeDb()

If len(pcIsOPC)=0 Then
  ' go to SSL form or not, depending on payment settings
  if chkPayment="FREE" then
	  response.redirect "gwSubmit.asp?purchasemode=FREE&idOrder=" &(cLng(scpre)+cLng(pIdOrder))
  end if
  response.redirect "gwSubmit.asp?pSsL="&scSSL&"&scSslURL="&scSslURL&"&psslurl="&psslurl&"&idPayment="&pIdPayment&"&idOrder="&(cLng(scpre)+cLng(pIdorder))
Else
	%> <!--#include file="opc_GateWayData.asp" --> <%
	response.Clear
	response.Write("OK")
End If
session("GWOrderID") = cLng(scpre)+cLng(pIdorder)
response.end

function DecimalFormatter(pricenumber)
	Dim testnum,testnum1
	testnum=cdbl(pricenumber)
	testnum=round(testnum,2)
	testnum1=cstr(testnum)
	if Instr(testnum1,",")>0 then
		testnum1=replace(testnum1,",",".")
	end if
	DecimalFormatter=testnum1
end function

Function generateABC(keyLength)
Dim sDefaultChars
Dim iCounter
Dim sMyKeys
Dim iPickedChar
Dim iDefaultCharactersLength
Dim ikeyLength

	sDefaultChars="ABCDEFGHIJKLMNOPQRSTUVXYZ"
	ikeyLength=keyLength
	iDefaultCharactersLength = Len(sDefaultChars)
	Randomize
	For iCounter = 1 To ikeyLength
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1)
		sMyKeys = sMyKeys & Mid(sDefaultChars,iPickedChar,1)
	Next
	generateABC = sMyKeys
End Function

Function generate123(keyLength)
Dim sDefaultChars
Dim iCounter
Dim sMyKeys
Dim iPickedChar
Dim iDefaultCharactersLength
Dim ikeyLength

	sDefaultChars="0123456789"
	ikeyLength=keyLength
	iDefaultCharactersLength = Len(sDefaultChars)
	Randomize
	For iCounter = 1 To ikeyLength
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1)
		sMyKeys = sMyKeys & Mid(sDefaultChars,iPickedChar,1)
	Next
	generate123 = sMyKeys
End Function


%>