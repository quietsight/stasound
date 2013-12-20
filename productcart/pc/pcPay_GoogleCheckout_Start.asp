<%@ LANGUAGE = VBScript.Encode %>
<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<% Response.Buffer = True %>
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/dimensionsformatinc.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/FedEXconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/FedEXWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="../includes/CPconstants.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/pcProductOptionsCode.asp"-->
<!--#include file="../includes/USPSCountry.asp"-->

<!--#include file="pcPay_GoogleCheckout_Global.asp"-->
<!--#include file="pcPay_GoogleCheckout_Checkout.asp"-->
<!--#include file="pcPay_GoogleCheckout_Handler.asp"-->

<%
dim conntemp, query, rs
dim f, total, totalDeliveringTime

Dim objFEDEXXmlDoc, objFedExStream, strFileName, GraphicXML, errstr
Dim fedex_postdata, objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEX_result, FEDEX_URL, pcv_strErrorMsg, pcv_strAction

Dim pcv_strAccountNameWS, pcv_strMeterNumberWS, pcv_strCarrierCodeWS
Dim pcv_strMethodNameWS, pcv_strMethodReplyWS, fedex_postdataWS, objFedExWSClass, objOutputXMLDocWS, srvFEDEXWSXmlHttp, FEDEXWS_result, FEDEXWS_URL, pcv_strErrorMsgWS

Session("UsedDiscountCodes")=""

'////////////////////////////////
'// Start: Activate Debug Mode
'////////////////////////////////
Dim pcv_strDebug
pcv_strDebug = "2"
'////////////////////////////////
'// End: Activate Debug Mode
'////////////////////////////////


'////////////////////////////////
'// Start: Shipping Zones
'////////////////////////////////
Dim pcv_strZones, pcv_ArrStates, pcv_ArrZips, pcv_ArrCountries, pcv_strType
pcv_strZones = "AUTO" '// Ex: "AUTO", "ALL", "CONTINENTAL_48", "FULL_50_STATES"
pcv_ArrStates = Array() '// Ex: Array("PA", "CA")
pcv_ArrZips = Array() '// Ex: Array("90210", "90212", "90214")
pcv_ArrCountries = Array() '// Ex: Array("US", "GB")
pcv_strType = "allowed"
'////////////////////////////////
'// End: Shipping Zones
'////////////////////////////////


'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************


'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************

total=Cint(0)
totalDeliveringTime=Cint(0)

if countCartRows(pcCartArray, pcCartIndex)=0 then
	response.redirect "msg.asp?message=1"
end if


'// Open the Database
call opendb()


'////////////////////////////////////////////////////////////////////////////////////////////
'// START: GOOGLE CHECKOUT
'////////////////////////////////////////////////////////////////////////////////////////////
If Request.ServerVariables("REQUEST_METHOD") = "POST" OR Request.QueryString("action")="checkout" Then

'// Google Analytics
pcv_strAnalyticsData = Request("analyticsdata")

'// Always set prices to no VAT - Let Google add VAT
Dim pcv_IsEUMemberState
pcv_IsEUMemberState = 0

'***********************************************************************************
' START: SAVE CART ARRAY
'***********************************************************************************
ppcCartIndex=Session("pcCartIndex")
if countCartRows(pcCartArray, ppcCartIndex)=0 then
	response.redirect "msg.asp?message=9"
end if
If session("customerType")=1 Then
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then
		'// Wholesale minimum not met, so customer cannot checkout -> show message
		response.redirect "msg.asp?message=205"
	end if
Else
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then
		'// Retail minimum not met, so customer cannot checkout -> show message
		response.redirect "msg.asp?message=206"
	end if
End If
'// Check if user already logged in
if session("idCustomer")<>0 and session("idCustomer")<>"" then
	'// See if customer is allowed to purchase
	query="SELECT suspend FROM customers WHERE idCustomer="&session("idCustomer")&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if rs("suspend")="1" then
		set rs=nothing
		response.redirect "msg.asp?message=131"
		response.end
	end if
end if
function fixstring(x)
	fixstring=replace(x,"'","''")
	fixstring=replace(fixstring,",","")
end function
function randomNumber(limit)
	randomize
	randomNumber=int(rnd*limit)+2
end function
dim pcv_intGoogleRandomKey, queryDate
pcv_intGoogleRandomKey=randomNumber(99999999)
'// Ensure the key is unique
pcv_intGoogleRandomKey=pcf_ValidateKey(pcv_intGoogleRandomKey)

pcCustSession_Date=Date()
if SQL_Format="1" then
	pcCustSession_Date=Day(pcCustSession_Date)&"/"&Month(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
else
	pcCustSession_Date=Month(pcCustSession_Date)&"/"&Day(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
end if
if scDB="Access" then
	queryDate=" #" &pcCustSession_Date& "# "
else
	queryDate=" '" &pcCustSession_Date& "' "
end if
' save the cart to the database
pcs_SaveCartArrayToDB
'***********************************************************************************
' END: SAVE CART ARRAY
'***********************************************************************************



'// Define objects used to create the shopping cart
Dim domTaxArea
Dim domShippingRestrictions
Dim elemItemName
Dim elemItemDescription
Dim elemQuantity
Dim elemUnitPrice
Dim elemTaxTableSelector
Dim elemShippingTaxed
Dim elemRate
Dim elemTaxAreaState
Dim elemTaxAreaCountry
Dim elemTaxAreaZip
Dim elemPrice
Dim elemMerchantCalculationsUrl
Dim elemAcceptMerchantCoupons
Dim elemAcceptGiftCertificates
Dim elemEditCartUrl
Dim elemContinueShoppingUrl
Dim elemMerchantPrivateItemData
Dim elemMerchantPrivateData
Dim attrName
Dim attrMerchantCalculated
Dim attrStandalone
Dim attrAllowedCountryArea
Dim attrExcludedCountryArea
Dim dtmCartExpiration
Dim strAllowedState
Dim strAllowedZip
Dim strExcludedState
Dim strExcludedZip
Dim arrayAllowedState
Dim arrayAllowedZip
Dim arrayExcludedState
Dim arrayExcludedZip
Dim checkoutPostData
Dim diagnoseResponse


'***********************************************************************************
' START: MAKE CART
'***********************************************************************************

Public Function pcf_ReverseHTML(HTMLcode)
	HTMLcode=replace(HTMLcode,"&quot;","""")
	HTMLcode=replace(HTMLcode,"&amp;","&")
	HTMLcode=replace(HTMLcode,"&lt; ","<")
	HTMLcode=replace(HTMLcode,"&gt;",">")
	HTMLcode=replace(HTMLcode,"&copy;","©")
	HTMLcode=replace(HTMLcode,"&reg;","®")
	HTMLcode=replace(HTMLcode,"&#8482;","™")
	HTMLcode=replace(HTMLcode,"&#8220;","“")
	HTMLcode=replace(HTMLcode,"&#8221;","”")
	HTMLcode=replace(HTMLcode,"&#8216;","‘")
	HTMLcode=replace(HTMLcode,"&#8217;","’")
	HTMLcode=replace(HTMLcode,"&#8218;","‚")
	pcf_ReverseHTML=HTMLcode
End Function

dim totalRowWeight
totalRowWeight=0

Dim pcv_strNoShipOrder, pnoshipping, pcv_strShipOrder
pcv_strShipOrder=1
pcv_strNoShipOrder=0
pnoshipping=0

Dim ProList(100,5)

for f=1 to pcCartIndex
	ProList(f,0)=pcCartArray(f,0)
	ProList(f,1)=pcCartArray(f,10)
	ProList(f,2)=pcCartArray(f,20)
	ProList(f,3)=pcCartArray(f,2)
	ProList(f,4)=0

	if pcCartArray(f,10)=0 then

		'//////////////////////////////
		'// START: ITEM NAME
		'//////////////////////////////
		elemItemName = pcf_ReverseHTML(pcCartArray(f,1))
		'//////////////////////////////
		'// END: ITEM NAME
		'//////////////////////////////




		'//////////////////////////////
		'// START: ITEM DESCRIPTION
		'//////////////////////////////
		'// Get product description
		psDesc = pcCartArray(f,1)
		psDesc = trim(psDesc)
		if psDesc<>"" AND isNULL(psDesc)=False then
			psDesc = ClearHTMLTags2(psDesc, 2)
		else
			psDesc = " "
		end if
		set rsImg = nothing
		elemItemDescription = psDesc
		'//////////////////////////////
		'// END: ITEM DESCRIPTION
		'//////////////////////////////




		'//////////////////////////////
		'// START: QUANTITY
		'//////////////////////////////
		elemQuantity = pcCartArray(f,2)
		'//////////////////////////////
		'// END: QUANTITY
		'//////////////////////////////



		'//////////////////////////////
		'// START: PRICE
		'//////////////////////////////

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start:  Pricing Calculations
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// BTO ADDON-S
		pBTOValues=0
		if trim(pcCartArray(f,16))<>"" then
			query="SELECT stringProducts, stringValues, stringCategories, stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
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
					if (ccur(ArrValue(i))<>0) or ((ArrQuantity(i)-1<>0) and (ArrPrice(i)<>0)) then
						if (ArrQuantity(i)-1)>=0 then
							UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
						else
							UPrice=0
						end if
						pBTOValues=pBTOValues+ccur((ArrValue(i)+UPrice)*pcCartArray(f,2))
					end if
					set rs=nothing
				next
			end if
		End if

		Dim pRowPrice, pRowWeight, pExtRowPrice
		pRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))
		pExtRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))-ccur(pBTOvalues)
		if pcCartArray(f,20)=0 then
			pRowWeight=pcCartArray(f,2)*pcCartArray(f,6)
		else
			pRowWeight=0
		end if
		totalRowWeight=totalRowWeight+pRowWeight
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End:  Pricing Calculations
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start:  Addtional Pricing Calculations
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' 1) Options
		pRowPrice=pRowPrice + ccur(pcCartArray(f,2) * pcCartArray(f,5))

		' 2) Item Disounts
		If pcCartArray(f,16)<>"" Then
			itemsDiscounts=0
			for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
				query="select quantityFrom, quantityUntil, discountperUnit, percentage, discountperWUnit from discountsPerQuantity where IDProduct=" & ArrProduct(i)
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				TempDiscount=0
				do while not rs.eof
					QFrom=rs("quantityFrom")
					QTo=rs("quantityUntil")
					DUnit=rs("discountperUnit")
					QPercent=rs("percentage")
					DWUnit=rs("discountperWUnit")
					if (DWUnit=0) and (DUnit>0) then
						DWUnit=DUnit
					end if


					TempD1=0
					if (clng(ArrQuantity(i)*pcCartArray(f,2))>=clng(QFrom)) and (clng(ArrQuantity(i)*pcCartArray(f,2))<=clng(QTo)) then
						if QPercent="-1" then
							if session("customerType")=1 then
								TempD1=ArrQuantity(i)*pcCartArray(f,2)*ArrPrice(i)*0.01*DWUnit
							else
								TempD1=ArrQuantity(i)*pcCartArray(f,2)*ArrPrice(i)*0.01*DUnit
							end if
						else
							if session("customerType")=1 then
								TempD1=ArrQuantity(i)*pcCartArray(f,2)*DWUnit
							else
								TempD1=ArrQuantity(i)*pcCartArray(f,2)*DUnit
							end if
						end if
					end if
					TempDiscount=TempDiscount+TempD1
					rs.movenext
				loop
				set rs=nothing
				itemsDiscounts=ItemsDiscounts+TempDiscount
			next
			if ItemsDiscounts>0 then
				pcCartArray(f,30)=ItemsDiscounts
				pRowPrice=pRowPrice-ItemsDiscounts
			end if
		End if


		' 3) BTO Additional Charges
		If trim(pcCartArray(f,16))<>"" Then
			query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			stringCProducts=rs("stringCProducts")
			stringCValues=rs("stringCValues")
			stringCCategories=rs("stringCCategories")
			ArrCProduct=Split(stringCProducts, ",")
			ArrCValue=Split(stringCValues, ",")
			ArrCCategory=Split(stringCCategories, ",")
			set rs=nothing
			if ArrCProduct(0)<>"na" then
				pRowPrice=pRowPrice+ccur(pcCartArray(f,31))
			end if
		End if

		' 4) Quantity Discounts
		If trim(pcCartArray(f,15))<>"" AND trim(pcCartArray(f,15))>0 Then
			pRowPrice=pRowPrice-ccur(pcCartArray(f,15))
		End if

		' 5) Cross Sell Bundle Discount
		if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) then
			pRowPrice = ( ccur(pRowPrice) + ccur(ProList(cint(pcCartArray(f,27)),2)) ) - ( ( ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28) ) ) * pcCartArray(f,2) )
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End:  Addtional Pricing Calculations
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


		if GOOGLECURRENCY="GBP" then
			elemUnitPrice = (pRowPrice/elemQuantity)
			elemUnitPrice=money(pcf_VAT(elemUnitPrice, pcCartArray(f,0) ))
			elemUnitPrice=pcf_CurrencyField(elemUnitPrice)
		else
			elemUnitPrice = (pRowPrice/elemQuantity)
			elemUnitPrice=money(elemUnitPrice)
			elemUnitPrice=pcf_CurrencyField(elemUnitPrice)
		end if
		'//////////////////////////////
		'// END: PRICE
		'//////////////////////////////



		'//////////////////////////////
		'// START: TAX
		'//////////////////////////////
		query="SELECT pcProductsVATRates.pcVATRate_ID FROM pcProductsVATRates WHERE pcProductsVATRates.idProduct="&pcCartArray(f,0)
		Set rsVAT=Server.CreateObject("ADODB.Recordset")
		set rsVAT=conntemp.execute(query)
		if not rsVAT.eof then
			elemTaxTableSelector = rsVAT("pcVATRate_ID")
		else
			elemTaxTableSelector = ""
		end if
		set rsVAT=nothing
		'//////////////////////////////
		'// END: TAX
		'//////////////////////////////



		'//////////////////////////////
		'// START: OPTIONS, BTO CONFIGURATION, CROSS-SELLING, CUSTOMER INPUT FIELDS
		'//////////////////////////////
		pcv_strItemNote = ""
		elemMerchantPrivateItemData = "<item-note>"& pcv_strItemNote &"</item-note>"
		'//////////////////////////////
		'// END: OPTIONS, BTO CONFIGURATION, CROSS-SELLING, CUSTOMER INPUT FIELDS
		'//////////////////////////////



		'//////////////////////////////
		'// START: CREATE ITEM
		'//////////////////////////////
		createItem elemItemName, elemItemDescription, elemQuantity, elemUnitPrice, elemTaxTableSelector, elemMerchantPrivateItemData
		'//////////////////////////////
		'// END: CREATE ITEM
		'//////////////////////////////



		'//////////////////////////////
		'// START: ORDER IS NO SHIP
		'//////////////////////////////
		if pcCartArray(f,20)="" OR isNULL(pcCartArray(f,20))=True then
			pnoshipping=0
		else
			pnoshipping=Abs(CInt(pcCartArray(f,20)))
		end if
		pcv_strNoShipOrder=pcv_strNoShipOrder+pnoshipping

		if pcv_strNoShipOrder=f then
			pcv_strShipOrder=0
		else
			pcv_strShipOrder=1
		end if
		'//////////////////////////////
		'// END: ORDER IS NO SHIP
		'//////////////////////////////

	end if
next

'***********************************************************************************
' END: MAKE CART
'***********************************************************************************



'***********************************************************************************
' START: NUMBER OF PACKAGES
'***********************************************************************************
'// Set Variables
pcv_intPackageNum=0
pcv_intTotPackageNum=0
intPackageCnt=0
intWeightCnt=0
pShipWeight=Cdbl(calculateShipWeight(pcCartArray, ppcCartIndex))
intUniversalWeight=pShipWeight

'// Verify Oversized Values
pcv_intOSCheck=oversizecheck(pcCartArray, ppcCartIndex)
if pcv_intOSCheck<>"" then
	pcv_arrOSCheckArray=split(pcv_intOSCheck,",")
	for i=0 to Ubound(pcv_arrOSCheckArray)-1
		pcv_arrOSArray=split(pcv_arrOSCheckArray(i),"|||")
		if pcv_arrOSArray(0)>pcv_intOSStatus then
			pcv_intOSStatus=1
		end if
	next
end if

'// Oversized Status = "1"
if pcv_intOSStatus<>0 then
	'// Loop through OS packages
	for i=0 to Ubound(pcv_arrOSCheckArray)-1
		pcv_arrOSArray=split(pcv_arrOSCheckArray(i),"|||")
		if pcv_arrOSArray(0)>pcv_intOSStatus then
			pcv_arrOSArray2=pcv_arrOSArray(1)
			pcv_strOSString=split(pcv_arrOSArray2,"||")
			if ubound(pcv_strOSString)=-1 then
			else
				intPackageCnt=intPackageCnt+1
				pcv_intTotPackageNum=pcv_intTotPackageNum+1
				intOSweight=pcv_strOSString(5)
				intWeightCnt=intWeightCnt+intOSweight
			end if
		end if
	next 'End loop through OS packages

	intOSpackageCnt=intPackageCnt
else
	pcv_intOSStatus=0
end if

intCustomShipWeight=intUniversalWeight
pShipWeight=intUniversalWeight-intWeightCnt

if pcv_intOSStatus=0 then
	intPackageCnt=0
end if

'// Start: Weight > 0
if pShipWeight>0 then
	if scShipFromWeightUnit="KGS" then
		intPounds=Int(pShipWeight/1000)
		intUniversalOunces=pShipWeight-(intPounds*1000) 'intUniversalOunces used for USPS
	else
		intPounds=Int(pShipWeight/16) 'intPounds used for USPS
		intUniversalOunces=pShipWeight-(intPounds*16) 'intUniversalOunces used for USPS
	end if
	intUniversalWeight=intPounds
	if intUniversalWeight<1 AND intUniversalOunces<1 then
		intUniversalWeight=0
	end if
	if intUniversalWeight<1 AND intUniversalOunces>0 then 'if total weight is less then a pound, make UPS/FedEX weight 1 pound
		intUniversalWeight=1
	else  'total weight is not less then a pound and ounces exist, round weight up one more pound.
		If intUniversalWeight>0 AND intUniversalOunces>0 then
			intUniversalWeight=(intUniversalWeight+1)
		End if
	end if

	'// Check to see if there is a weight limit set for packages >0
	if int(scPackageWeightLimit)<>0 then
		'// How many package this should be if over the limit
		if int(intUniversalWeight)>int(scPackageWeightLimit) then '// There are more package after OS
			intTempPackageNum=(intUniversalWeight/int(scPackageWeightLimit))
			pcv_intPackageNum=int(intUniversalWeight/int(scPackageWeightLimit))
			if intTempPackageNum>pcv_intPackageNum then
				pcv_intPackageNum=pcv_intPackageNum+1
			end if
			for r=1 to (pcv_intPackageNum-1)
				intPackageCnt=intPackageCnt+1
				pcv_intTotPackageNum=pcv_intTotPackageNum+1
			next
			'// Last package
			intPackageCnt=intPackageCnt+1
			pcv_intTotPackageNum=pcv_intTotPackageNum+1
		else '// There are more package after OS
			intPackageCnt=intPackageCnt+1
			pcv_intTotPackageNum=pcv_intTotPackageNum+1
		end if '// There are more package after OS
	else '// There is a package Weight limit set
		'// No weight limit set
		intPackageCnt=intPackageCnt+1
		pcv_intTotPackageNum=pcv_intTotPackageNum+1
	end if '// There is a package Weight limit set
end if
'// End: Weight > 0


pcv_intPackageNum=intPackageCnt
'***********************************************************************************
' END: NUMBER OF PACKAGES
'***********************************************************************************




'***********************************************************************************
' START: SET CART
'***********************************************************************************
pcv_strDateTime = dateAdd ("d", GOOGLEEXPIREDAYS, now())
pcv_strDateTime = Year(pcv_strDateTime) &"-"& Right(Cstr(Month(pcv_strDateTime) + 100),2) &"-"& Right(Cstr(Day(pcv_strDateTime) + 100),2)
dtmCartExpiration = pcv_strDateTime & "T23:59:59"
if NOT len(session("idCustomer"))>0 then
	session("idCustomer") = Cint(0)
end if
if NOT len(session("idAffiliate"))>0 then
	session("idAffiliate") = Cint(1)
end if
elemMerchantPrivateData = "<merchant-note>"& pcv_intGoogleRandomKey & chr(124) & queryDate & chr(124) & pcv_intPackageNum & chr(124) & session("idCustomer") & chr(124) & session("idAffiliate") & "</merchant-note>"
createShoppingCart dtmCartExpiration, elemMerchantPrivateData
'***********************************************************************************
' END: SET CART
'***********************************************************************************




'***********************************************************************************
' START: SHIPPING DEFAULTS
'***********************************************************************************
if pcv_strShipOrder=1 then

' Create list of areas where a particular shipping option is available
attrAllowedCountryArea = pcv_strZones
arrayAllowedState = pcv_ArrStates
arrayAllowedZip = pcv_ArrZips
arrayAllowedCountries = pcv_ArrCountries
Set domShippingRestrictions = addAllowedAreas(attrAllowedCountryArea, arrayAllowedState, arrayAllowedZip, arrayAllowedCountries, pcv_strType)

'// Retrieve the city
Session("pcSFcity") = scShipFromCity
'// Retrieve the state
Session("pcSFStateCode") = scShipFromState
'// Retrieve the zip
Session("pcSFzip") = scShipFromPostalCode
'// Retrieve the county code
Session("pcSFCountryCode") = scShipFromPostalCountry
%>

<!--#include file="pcPay_GoogleCheckout_Shipping.asp"-->

<%
'// Open the Database
call opendb()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: Calculate the "Default" Rates
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
pcv_ArrayGoogleShippingMethods = split(pcv_strGoogleShippingPrices, "||||")
For i=lbound(pcv_ArrayGoogleShippingMethods) to ubound(pcv_ArrayGoogleShippingMethods)
	if trim(pcv_ArrayGoogleShippingMethods(i))<>"" then
		temparray = split(pcv_ArrayGoogleShippingMethods(i),"|?|")
		attrName = temparray(0)
		pServiceCode = temparray(2)
		if pServiceCode = "" then pServiceCode = 0
		'// Check for a default
		c=0
		query="SELECT serviceDefaultRate FROM shipService WHERE serviceActive=-1 AND serviceCode='"& pServiceCode &"';"
		set rs2=connTemp.execute(query)
		if NOT rs2.eof then
			c=rs2("serviceDefaultRate")
			if isNULL(c)=True OR c="" then
				c=0
			end if
		end if
		if c=0 then
			elemPrice2x = ((temparray(1)*2)+2)
			elemPrice = temparray(1)
		else
			elemPrice2x = c
			elemPrice = c
		end if
		Session(pServiceCode)=elemPrice
		Session(pServiceCode & "2x")=elemPrice2x
	end if
Next
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: Calculate the "Default" Rates
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: Get ALL Shipping Options & Add the "Default" Rates
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim pDefaultRate, pAttributeName

query="SELECT shipService.serviceDefaultRate, shipService.serviceDescription, shipService.serviceCode FROM shipService WHERE (((shipService.serviceActive)=-1)) ORDER BY shipService.servicePriority;"
set rsActive=Server.CreateObject("ADODB.RecordSet")
set rsActive=connTemp.execute(query)
if NOT rsActive.eof then

	pcv_ArrayActiveOptions = rsActive.getRows()
	intCount=ubound(pcv_ArrayActiveOptions,2)

	For i=0 to intCount

		pDefaultRate= pcv_ArrayActiveOptions(0,i)
		if isNULL(pDefaultRate)=True OR pDefaultRate="" then
			pDefaultRate=0
		end if
		pAttributeName= pcv_ArrayActiveOptions(1,i)
		pAttributeName= replace(pAttributeName,"®","")
		pAttributeName= replace(pAttributeName,"&lt;sup&gt;SM&lt;/sup&gt;","")
		pAttributeName= replace(pAttributeName,"&lt;sup&gt;","")
		pAttributeName= replace(pAttributeName,"&lt;/sup&gt;","")
		pAttributeName= replace(pAttributeName,"&reg;","")
		pAttributeName= replace(pAttributeName,"<sup>SM</sup>","")
		pAttributeName= replace(pAttributeName,"<sup>","")
		pAttributeName= replace(pAttributeName,"</sup>","")
		pAttributeName= trim(pAttributeName)
		pServiceCode= pcv_ArrayActiveOptions(2,i)
		pcv_strPerformCalc=1
		select case pServiceCode
			'// FedEx
			case "PRIORITYOVERNIGHT"
				pAttributeName="FedEx Priority Overnight"
			case "STANDARDOVERNIGHT"
				pAttributeName="FedEx Standard Overnight"
			case "FIRSTOVERNIGHT"
				pAttributeName="FedEx First Overnight"
			case "FEDEX2DAY"
				pAttributeName="FedEx 2Day"
			case "FEDEXEXPRESSSAVER"
				pAttributeName="FedEx Express Saver"
			case "INTERNATIONALPRIORITY"
				pAttributeName="FedEx International Priority"
			case "INTERNATIONALECONOMY"
				pAttributeName="FedEx International Economy"
			case "INTERNATIONALFIRST"
				pAttributeName="FedEx International First"
			case "FEDEX1DAYFREIGHT"
				pAttributeName="FedEx 1Day Freight"
			case "FEDEX2DAYFREIGHT"
				pAttributeName="FedEx 2Day Freight"
			case "FEDEX3DAYFREIGHT"
				pAttributeName="FedEx 3Day Freight"
			case "FEDEXGROUND"
				pAttributeName="FedEx Ground"
			case "GROUNDHOMEDELIVERY"
				pAttributeName="FedEx Home Delivery"
			case "INTERNATIONALPRIORITYFREIGHT"
				pAttributeName="FedEx International Priority Freight"
			case "INTERNATIONALECONOMYFREIGHT"
				pAttributeName="FedEx International Economy Freight"

			'// FedEx WS
			case "PRIORITY_OVERNIGHT"
				pAttributeName="FedEx Priority Overnight"
			case "STANDARD_OVERNIGHT"
				pAttributeName="FedEx Standard Overnight"
			case "FIRST_OVERNIGHT"
				pAttributeName="FedEx First Overnight"
			case "FEDEX_2_DAY"
				pAttributeName="FedEx 2Day"
			case "FEDEX_EXPRESS_SAVER"
				pAttributeName="FedEx Express Saver"
			case "INTERNATIONAL_PRIORITY"
				pAttributeName="FedEx International Priority"
			case "INTERNATIONAL_ECONOMY"
				pAttributeName="FedEx International Economy"
			case "INTERNATIONAL_FIRST"
				pAttributeName="FedEx International First"
			case "FEDEX_1_DAY_FREIGHT"
				pAttributeName="FedEx 1Day Freight"
			case "FEDEX_2_DAY_FREIGHT"
				pAttributeName="FedEx 2Day Freight"
			case "FEDEX_3_DAY_FREIGHT"
				pAttributeName="FedEx 3Day Freight"
			case "FEDEX_GROUND"
				pAttributeName="FedEx Ground"
			case "GROUND_HOME_DELIVERY"
				pAttributeName="FedEx Home Delivery"
			case "INTERNATIONAL_PRIORITY_FREIGHT"
				pAttributeName="FedEx International Priority Freight"
			case "INTERNATIONAL_ECONOMY_FREIGHT"
				pAttributeName="FedEx International Economy Freight"
			case "INTERNATIONAL_GROUND"
				pAttributeName="FedEx International Ground"
			case "FEDEX_FREIGHT"
				pAttributeName="FedEx Freight"
			case "FEDEX_NATIONAL_FREIGHT"
				pAttributeName="FedEx National Freight"
			case "SMART_POST"
				pAttributeName="FedEx SmartPost"

			'// UPS
			case "01"
				pAttributeName="UPS Next Day Air"
			case "02"
				pAttributeName="UPS 2nd Day Air"
			case "03"
				pAttributeName="UPS Ground"
			case "07"
				pAttributeName="UPS Worldwide Express"
			case "08"
				pAttributeName="UPS Worldwide Expedited"
			case "11"
				pAttributeName="UPS Standard To Canada"
			case "12"
				pAttributeName="UPS 3 Day Select"
			case "13"
				pAttributeName="UPS Next Day Air Saver"
			case "14"
				pAttributeName="UPS Next Day Air Early A.M."
			case "54"
				pAttributeName="UPS Worldwide Express Plus"
			case "59"
				pAttributeName="UPS 2nd Day Air A.M."
			case "65"
				pAttributeName="UPS Express Saver"
			'// USPS
			case "9902"
				pAttributeName="Express Mail"
			case "9901"
				pAttributeName="Priority Mail"
			case "9904"
				pAttributeName= "First-Class Mail"
			case "9903"
				pAttributeName="Standard Post"
			case "9915"
				pAttributeName="Bound Printed Matter"
			case "9916"
				pAttributeName="Media Mail"
			case "9917"
				pAttributeName="Library Mail"
			case "9914"
				pAttributeName="Global Express Guaranteed"
			case "9905"
				pAttributeName="Global Express Guaranteed Non-Document Rectangular"
			case "9906"
				pAttributeName="Express Mail International (EMS)"
			case "9907"
				pAttributeName="Priority Mail International"
			case "9908"
				pAttributeName= "Priority Mail International Flat-Rate Envelope" '// "Priority Mail International Flat Rate Envelope"
			case "9909"
				pAttributeName= "Priority Mail International Flat-Rate Box" '// "Priority Mail International Flat Rate Box"
			case "9910"
				pAttributeName="Global Express Guaranteed Non-Document Non-Rectangular"
			case "9911"
				pAttributeName="Express Mail International (EMS) Flat Rate Envelope"
			case "9912"
				pAttributeName= "First Class Mail International" '// "First Class Mail International Package"
			'// Canadian Post
			case "1010"
				pAttributeName="Canada Post - REGULAR"
			case "1020"
				pAttributeName="Canada Post - EXPEDITED"
			case "1030"
				pAttributeName="Canada Post - XPRESSPOST"
			case "1040"
				pAttributeName="Canada Post - PRIORITY COURIER"
			case "1120"
				pAttributeName="Canada Post - EXPEDITED EVENING"
			case "1130"
				pAttributeName="Canada Post - XPRESSPOST EVENING"
			case "1220"
				pAttributeName="Canada Post - EXPEDITED SATURDAY"
			case "1230"
				pAttributeName="Canada Post - XPRESSPOST SATURDAY"
			case "2010"
				pAttributeName="Canada Post - SURFACE US"
			case "2020"
				pAttributeName="Canada Post - AIR US"
			case "2030"
				pAttributeName="Canada Post - XPRESSPOST US"
			case "2040"
				pAttributeName="Canada Post - PUROLATOR US"
			case "2050"
				pAttributeName="Canada Post - PUROPAK US"
			case "3010"
				pAttributeName="Canada Post - SURFACE INTERNATIONAL"
			case "3020"
				pAttributeName="Canada Post - AIR INTERNATIONAL"
			case "3040"
				pAttributeName="Canada Post - PUROLATOR INTERNATIONAL"
			case "3050"
				pAttributeName="Canada Post - PUROPAK INTERNATIONAL"
			case else
				pcv_strPerformCalc=0
		end select

		If Session(pServiceCode)="" OR isNULL(Session(pServiceCode))=True Then
			createMerchantCalculatedShipping pAttributeName, pDefaultRate, domShippingRestrictions
		Else
			if pcv_strPerformCalc=0 then
				createMerchantCalculatedShipping pAttributeName, Session(pServiceCode), domShippingRestrictions
			else
				createMerchantCalculatedShipping pAttributeName, Session(pServiceCode & "2x"), domShippingRestrictions
			end if
		End If

		Session(pServiceCode)=""
		Session(pServiceCode & "2x")=""

	Next

end if
set rsActive = nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: Get ALL Shipping Options & Add the "Default" Rates
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




Session("pcSFcity") = ""
Session("pcSFStateCode") = ""
Session("pcSFzip") = ""
Session("pcSFCountryCode") = ""

end if
'***********************************************************************************
' END: SHIPPING DEFAULTS
'***********************************************************************************



'***********************************************************************************
' START: TAX TABLES
'***********************************************************************************

if GOOGLECURRENCY="GBP" then
	elemRate = (ptaxVATrate/100)
	attrMerchantCalculated = "false"
	query="SELECT pcVATCountries.pcVATCountry_Code From pcVATCountries Order By pcVATCountry_State ASC;"
	set rsVAT=Server.CreateObject("ADODB.Recordset")
	set rsVAT=conntemp.execute(query)
	if not rsVAT.eof then
		pcArr=rsVAT.getRows()
		set rsVAT=nothing
		intCount=0
		intCount=ubound(pcArr,2)
		if intCount>=0 then

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: Default VAT Table
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			For vatCounter=0 to intCount
				elemTaxAreaCountry = pcArr(0,vatCounter)
				Set domTaxArea = createTaxArea("country", elemTaxAreaCountry)
				elemShippingTaxed = GOOGLETAXSHIPPING
				'// Default Tax Rules
				createDefaultTaxRule elemRate, domTaxArea, elemShippingTaxed
			next
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'//End: Default VAT Table
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: Alternate VAT Table
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			query="SELECT pcVATRates.pcVATRate_Rate, pcVATRates.pcVATRate_ID FROM pcVATRates"
			set rsVAT=Server.CreateObject("ADODB.Recordset")
			set rsVAT=conntemp.execute(query)
			if not rsVAT.eof then
				pcRateArr=rsVAT.getRows()
				set rsVAT=nothing
				intRateCount=0
				intRateCount=ubound(pcRateArr,2)
				if intRateCount>=0 then
					For vatRateCounter=0 to intRateCount
						'// Alternate Tax Rules
						elemRate = (pcRateArr(0,vatRateCounter)/100)
						elemRateID = pcRateArr(1,vatRateCounter)
						For vatCounter=0 to intCount
							Set domTaxArea = createTaxArea("country", pcArr(0,vatCounter))
							createAlternateTaxRule elemRate, domTaxArea
						next
						'// Set Alternate Table
						createAlternateTaxTable "false", elemRateID
					Next
				end if
			end if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: Alternate VAT Table
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		end if '// if intCount>0 then
	end if '// if not rsVAT.eof then
else
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Start: Default TAX Table
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	elemRate = "0.00"
	attrMerchantCalculated = "true"
	elemTaxAreaCountry = "ALL"
	Set domTaxArea = createTaxArea("country", elemTaxAreaCountry)
	elemShippingTaxed = GOOGLETAXSHIPPING
	createDefaultTaxRule elemRate, domTaxArea, elemShippingTaxed
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// End: Default TAX Table
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
end if

createTaxTables attrMerchantCalculated

'***********************************************************************************
' END: TAX TABLES
'***********************************************************************************



'***********************************************************************************
' START: MERCHANT CALCULATIONS
'***********************************************************************************
CalculationsUrl=scSslURL & "/" & scPcFolder & "/pc/"
CalculationsUrl=replace(CalculationsUrl,"///","/")
CalculationsUrl=replace(CalculationsUrl,"//","/")
CalculationsUrl=replace(CalculationsUrl,"https:/","https://")
CalculationsUrl=replace(CalculationsUrl,"http:/","http://")
elemMerchantCalculationsUrl=CalculationsUrl&"pcPay_GoogleCheckout_Callback.asp"

'// Check for Coupons
query="SELECT discounts.iddiscount FROM discounts;"
set rstemp=server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
if rstemp.eof then
	elemAcceptMerchantCoupons = "false"
else
	elemAcceptMerchantCoupons = "true"
end if
set rstemp=nothing
'// Check for Discounts
query="SELECT Products.IDProduct from Products where pcprod_GC=1;"
set rstemp=server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
if rstemp.eof then
	elemAcceptGiftCertificates = "false"
else
	elemAcceptGiftCertificates = "true"
end if
set rstemp=nothing
createMerchantCalculations elemMerchantCalculationsUrl, elemAcceptMerchantCoupons, elemAcceptGiftCertificates
'***********************************************************************************
' END: MERCHANT CALCULATIONS
'***********************************************************************************



'***********************************************************************************
' START: ROUNDING POLICY
'***********************************************************************************
if GOOGLECURRENCY="GBP" then
	createRoundingPolicy "HALF_UP", "PER_LINE"
end if
'***********************************************************************************
' END: ROUNDING POLICY
'***********************************************************************************




'***********************************************************************************
' START: MISC VARIABLES
'***********************************************************************************
CalculationsUrlPath=Request.ServerVariables("PATH_INFO")
CalculationsUrlPath = Left( CalculationsUrlPath,(InStrRev(CalculationsUrlPath,"/")))
elemEditCartUrl = "http://" & Request.ServerVariables("SERVER_NAME") & CalculationsUrlPath & "viewcart.asp"
elemContinueShoppingUrl = "http://" & Request.ServerVariables("SERVER_NAME") & CalculationsUrlPath & "CustLOb.asp"
elemPlatformID = "956144146810286"
'***********************************************************************************
' END: MISC VARIABLES
'***********************************************************************************



'***********************************************************************************
' START: SEND THE CART
'***********************************************************************************
' API request.
createMerchantCheckoutFlowSupport elemEditCartUrl, elemContinueShoppingUrl, elemPlatformID

Dim xmlCart
Dim b64signature
Dim b64cart

' Get <checkout-shopping-cart> XML
xmlCart = createCheckoutShoppingCart
b64cart = Base64_Encode(xmlCart)
b64signature = b64_hmac_sha1(strMerchantKey, xmlCart)

' Post the cart XML via a server-to-server POST and
' capture the <checkout-redirect> response.
Dim transmitResponse

If pcv_strDebug = "1" Then
	response.ContentType="text/xml"
	response.Write(xmlCart)
	response.End()
End If

if errstr="" then
	transmitResponse = SendRequest(xmlCart, checkoutUrl)
	ProcessXmlData(transmitResponse)
end if

' Free object
Set domTaxArea = Nothing
Set domShippingRestrictions = Nothing
'***********************************************************************************
' END: SEND THE CART
'***********************************************************************************



End If
'////////////////////////////////////////////////////////////////////////////////////////////
'// END: GOOGLE CHECKOUT
'////////////////////////////////////////////////////////////////////////////////////////////
%>

<%
call closedb()
response.redirect "msgb.asp?message="&Server.URLEncode(errstr)
%>


