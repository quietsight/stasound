<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/rc4.asp" --> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Global.asp"-->
<!--#include file="../includes/GoogleCheckout_APIFunctions.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Handler.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/languages_ship.asp" --> 
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<style>

	#pcCPmain ul {
		margin: 0px;
		padding: 0;
	}

	#pcCPmain ul li {
		margin: 0px;
	}

	div.menu ul {
	text-align:left;
	margin:0 0 0 60px;
	padding:0;
	cursor:pointer;
	}

	div.menu ul li {
	display:inline;
	list-style:none;
	margin:0 0.3em;
	cursor:pointer;
	font-size:12px;
	}

	div.menu ul li a {
	position:relative;
	z-index:0;
	font-weight:bold;
	border:solid 2px #e1e1e1;
	border-bottom-width:0;
	padding:0.3em;
	background-color:#ffffcc;
	color:black;
	text-decoration:none;
	cursor:pointer;
	font-size:12px;
	}

	div.menu ul li a.current {
	background-color:#F5F5F5;
	border:solid 2px #CCCCCC;
	border-bottom-width:0;
	position:relative;z-index:2;
	cursor:pointer;
	font-size:12px;
	}
	
	div.menu ul li a.current:hover {
	background-color:#F5F5F5;
	cursor:pointer;
	font-size:12px;
	}

	div.menu ul li a:hover {
	z-index:2;
	background-color:#F5F5F5;
	border-bottom:0;
	cursor:pointer;
	font-size:12px;
	}
	
	div.menu a span {display:none;}
	
	div.menu a:hover span {
		display:block;
		position:absolute;
		top:2.3em;
		background-color:#F5F5F5;
		border-bottom:thin dotted gray;
		border-top:thin dotted gray;
		font-weight:normal;
		left:0;
		padding:1px 2px;
		cursor:pointer;
		font-size:12px;
	}
	
	div.panes {
		padding: 1em;
		border: dashed 2px #CCCCCC;
		background-color: #F5F5F5;
		display: none;
		text-align:left;
		position:relative;z-index:1;
		margin-top:0.15em;
	}
	
	div.navbox {
		display: table-cell;
		padding: .3em;
		font-size: 12px;
		font-weight:bold;
		border: solid 2px #CCCCCC;
		background-color: #F5F5F5;
		text-align:left;
	}
	
	div.NavOrderClass1 {
		padding: 0.2em;
		font-size:12px;
		font-weight:bold;
		background-color: #B0E0E6;
		display: none;
		text-align:left;
		margin-top:0em;
		border-bottom: 1px solid #CCCCCC;
		border-left: 1px solid #CCCCCC;
		border-right: 1px solid #CCCCCC;
	}
	
	div.NavOrderClass2 {
		padding: 0.2em;
		font-size:12px;
		font-weight:bold;
		background-color: #FFFFFF;
		display: none;
		text-align:left;
		margin-top:0em;
		border-bottom: 1px solid #CCCCCC;
		border-left: 1px solid #CCCCCC;
		border-right: 1px solid #CCCCCC;
	}
</style>

<%
Dim objFEDEXXmlDoc, objFedExStream, strFileName, GraphicXML
Dim iPageCurrent, varFlagIncomplete, uery, strORD, pcv_intOrderID
Dim pcv_strMethodName, pcv_strMethodReply, CustomerTransactionIdentifier, pcv_strAccountNumber, pcv_strMeterNumber, pcv_strCarrierCode
Dim pcv_strTrackingNumber, pcv_strShipmentAccountNumber
Dim pcv_strDestinationCountryCode, pcv_strDestinationPostalCode, pcv_strLanguageCode, pcv_strLocaleCode, pcv_strDetailScans, pcv_strPagingToken
Dim fedex_postdata, objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEX_result, FEDEX_URL, pcv_strErrorMsg, pcv_strAction

'// Define objects used to create and send Google Checkout Order Processing API requests
Dim xmlRequest
Dim xmlResponse
Dim attrGoogleOrderNumber
Dim elemAmount 
Dim elemReason
Dim elemComment
Dim elemCarrier
Dim elemTrackingNumber
Dim elemMessage
Dim elemSendEmail
Dim elemMerchantOrderNumber
Dim transmitResponse

function fnStripPhone(PhoneField)
	PhoneField=replace(PhoneField," ","")
	PhoneField=replace(PhoneField,"-","")
	PhoneField=replace(PhoneField,".","")
	PhoneField=replace(PhoneField,"(","")
	PhoneField=replace(PhoneField,")","")
	fnStripPhone = PhoneField
end function

function sanitizeField(UserInput)
	if UserInput<>"" AND isNULL(UserInput)=False then
		UserInput=replace(UserInput,"&"," ")
	end if
	sanitizeField=UserInput
end function

'// If there is no Tracking Number, provide a random number for the log file. 
'// This is due to the fact there could be multiple errors for the same ID.
function randomNumber(limit)
	randomize
	randomNumber=int(rnd*limit)+2
end function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// OPEN DATABASE
call openDb()

'// PACKAGE COUNT
pcv_strPackageCount=request("PackageCount")
pcv_strSessionPackageCount=Session("pcAdminPackageCount")
if pcv_strSessionPackageCount="" OR len(pcv_strPackageCount)>0 then
	pcPackageCount=pcv_strPackageCount
	Session("pcAdminPackageCount")=pcPackageCount
else
	pcPackageCount=pcv_strSessionPackageCount
end if
if pcPackageCount="" then
	pcPackageCount=1
end if
pcArraySize = (pcPackageCount -1)
						
'// GET ORDER ID
pcv_strOrderID=request("idorder")
pcv_strSessionOrderID=Session("pcAdminOrderID")
if pcv_strSessionOrderID="" OR len(pcv_strOrderID)>0 then
	pcv_intOrderID=pcv_strOrderID
	Session("pcAdminOrderID")=pcv_intOrderID
else
	pcv_intOrderID=pcv_strSessionOrderID
end if

'// REDIRECT
if pcv_intOrderID="" then
	response.redirect "menu.asp"
end if

query="SELECT orders.pcOrd_GoogleIDOrder FROM orders WHERE idOrder="& pcv_intOrderID
set rs=server.CreateObject("ADODB.RecordSet")
Set rs=conntemp.execute(query)
if Not rs.eof then
	pcv_strGoogleIDOrder = rs("pcOrd_GoogleIDOrder") '// determine if this is a google order
end if
set rs=nothing

'// ITEM COUNT
pcv_count=Request("count")
if pcv_count="" then
	pcv_count=0
end if

'// CREATE THE ARRAY
Dim pcLocalArray()

'// SIZE THE ARRAY
ReDim pcLocalArray(pcArraySize)

'// POPULATE THE ARRAY
if request.form("submit")<>"" OR request.form("submit1")<>"" then
	For xPackageCount=0 to pcArraySize
		pcLocalArray(xPackageCount) = Request("pcAdminPrdList" & (xPackageCount+1))	
	Next 
else
	if Session("pcGlobalArray")<>"" then
		pcArray_TmpGlobalReturn = split(Session("pcGlobalArray"), chr(124))
		For xPackageCount = LBound(pcArray_TmpGlobalReturn) TO UBound(pcArray_TmpGlobalReturn)
			pcLocalArray(xPackageCount) = pcArray_TmpGlobalReturn(xPackageCount)	 
		Next
	end if
end if

'// UPDATE ARRAY
If pcv_count <> 0 Then	
	For i=1 to pcv_count
		if request("C" & i)="1" then
			pcv_strTmpList=pcv_strTmpList & request("IDPrd" & i) & ","
		end if		
	Next
	pcLocalArray((pcPackageCount-1)) = pcv_strTmpList
End If

'// CONVERT ARRAY TO SESSIONS
For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
	Session("pcAdminPrdList"&(xArrayCount+1)) = pcLocalArray(xArrayCount)  
Next

'// ARRAY TO PASS TO OTHER PAGES
pcv_strItemsList = join(pcLocalArray, chr(124))

'// SESSION FOR REDIRECTS
Session("pcGlobalArray") = pcv_strItemsList

'////////////////////////////////////////////
'// END: PRODUCT ID LIST FOUR
'////////////////////////////////////////////

'// PAGE NAME
pcPageName="FedEx_ManageShipmentsRequest.asp"
ErrPageName="FedEx_ManageShipmentsRequest.asp"

'// ACTION
pcv_strAction = request("Action")

'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExClass
	
'// FEDEX CREDENTIALS
query = 		"SELECT ShipmentTypes.userID, ShipmentTypes.password, ShipmentTypes.AccessLicense "
query = query & "FROM ShipmentTypes "
query = query & "WHERE (((ShipmentTypes.idShipment)=1));"	
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if NOT rs.eof then
	pcv_strAccountNumber=rs("userID")
	pcv_strMeterNumber=rs("password")
	pcv_strEnvironment=rs("AccessLicense")
end if
set rs=nothing

'// DATE FUNCTION
function ShowDateFrmt(x)
	ShowDateFrmt = x
end function

'// FedEx Ship Preferences
query="SELECT pcUPSPref_Service, pcUPSPref_PackageType, pcUPSPref_PaymentMethod, pcUPSPref_AccountNumber, pcUPSPref_ReadyHours, pcUPSPref_ReadyMinutes, pcUPSPref_ReadyAMPM, pcUPSPref_PUHours, pcUPSPref_PUMinutes, pcUPSPref_RefNumber1, pcUPSPref_RefNumber2, pcUPSPref_RefData1, pcUPSPref_RefData2, pcUPSPref_CODPackage, pcUPSPref_CODAmount, pcUPSPref_CODCurrency, pcUPSPref_CODFunds, pcUPSPref_ShipmentNotification, pcUPSPref_NotifiCode1, pcUPSPref_NotifiCode2, pcUPSPref_NotifiCode3, pcUPSPref_NotifiCode4, pcUPSPref_NotifiCode5, pcUPSPref_NotifiEmail1, pcUPSPref_NotifiEmail2, pcUPSPref_NotifiEmail3, pcUPSPref_NotifiEmail4, pcUPSPref_NotifiEmail5, pcUPSPref_SaturdayDelivery, pcUPSPref_InsuredValue, pcUPSPref_VerbalConfirmation FROM pcUPSPreferences WHERE pcUPSPref_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then
	'/////////////////////////////////////////////////////
	'// Set Local Variables for Setting
	'/////////////////////////////////////////////////////
	
	Session("pcAdminUPSServiceCode")=rs("pcUPSPref_Service")
	Session("pcAdminUPSPackageType")=rs("pcUPSPref_PackageType")
	Session("pcAdminUPSPayorType")=rs("pcUPSPref_PaymentMethod")
	Session("pcAdminUPSAccountNumber")=rs("pcUPSPref_AccountNumber")
	Session("pcAdminUPSReadyHours")=rs("pcUPSPref_ReadyHours")
	Session("pcAdminUPSReadyMinutes")=rs("pcUPSPref_ReadyMinutes")
	Session("pcAdminUPSReadyAMPM")=rs("pcUPSPref_ReadyAMPM")
	Session("pcAdminUPSPUHours")=rs("pcUPSPref_PUHours")
	Session("pcAdminUPSPUMinutes")=rs("pcUPSPref_PUMinutes")
	Session("pcAdminUPSRefNumber1")=rs("pcUPSPref_RefNumber1")
	Session("pcAdminUPSRefNumber2")=rs("pcUPSPref_RefNumber2")
	Session("pcAdminUPSRefData1")=rs("pcUPSPref_RefData1")
	Session("pcAdminUPSRefData2")=rs("pcUPSPref_RefData2")
	Session("pcAdminUPSCODPackage")=rs("pcUPSPref_CODPackage")
	Session("pcAdminUPSCODAmount")=rs("pcUPSPref_CODAmount")
	Session("pcAdminUPSCODCurrency")=rs("pcUPSPref_CODCurrency")
	Session("pcAdminUPSCODFunds")=rs("pcUPSPref_CODFunds")
	Session("pcAdminUPSShipmentNotification")=rs("pcUPSPref_ShipmentNotification")
	Session("pcAdminUPSNotifiCode1")=rs("pcUPSPref_NotifiCode1")
	Session("pcAdminUPSNotifiCode2")=rs("pcUPSPref_NotifiCode2")
	Session("pcAdminUPSNotifiCode3")=rs("pcUPSPref_NotifiCode3")
	Session("pcAdminUPSNotifiCode4")=rs("pcUPSPref_NotifiCode4")
	Session("pcAdminUPSNotifiCode5")=rs("pcUPSPref_NotifiCode5")
	Session("pcAdminUPSNotifiEmail1")=rs("pcUPSPref_NotifiEmail1")
	Session("pcAdminUPSNotifiEmail2")=rs("pcUPSPref_NotifiEmail2")
	Session("pcAdminUPSNotifiEmail3")=rs("pcUPSPref_NotifiEmail3")
	Session("pcAdminUPSNotifiEmail4")=rs("pcUPSPref_NotifiEmail4")
	Session("pcAdminUPSNotifiEmail5")=rs("pcUPSPref_NotifiEmail5")
	Session("pcAdminUPSSaturdayDelivery")=rs("pcUPSPref_SaturdayDelivery")
	Session("pcAdminUPSInsuredValue")=rs("pcUPSPref_InsuredValue")
	Session("pcAdminUPSVerbalConfirmation")=rs("pcUPSPref_VerbalConfirmation")
end if

'// GET CONSTANTS
if Session("pcAdminUPSPackageType")="" then
	Session("pcAdminUPSPackageType")=UPS_PICKUP_TYPE
end if
'// SELECT DATA SET
' >>> Tables: pcPackageInfo
query="SELECT orders.idCustomer, orders.ShippingFullname, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, "
query = query & "orders.shippingCountryCode, orders.shippingZip, orders.shippingCompany, orders.shippingAddress2, orders.pcOrd_shippingPhone, orders.pcOrd_ShippingEmail, "
query = query & "orders.SRF, orders.shipmentDetails, orders.OrdShipType, orders.OrdPackageNum, orders.pcOrd_ShipWeight, orders.pcOrd_ShippingFax "
query = query & "FROM orders "
query = query & "WHERE orders.idOrder=" & pcv_intOrderID &" "

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then		
	Dim pidorder, pidcustomer, pshippingAddress, pshippingCity, pshippingStateCode, pshippingState, pshippingZip, pshippingPhone, pshippingCountryCode, pshippingCompany, pshippingAddress2, pShippingEmail, SRF
	
	'// ORDER INFO
	pidorder=scpre+int(pcv_intOrderID)
	pidcustomer=rs("idcustomer")
	
	'// DESTINATION ADDRESS
	pShippingFullname=rs("ShippingFullname")
	pshippingAddress=rs("shippingAddress")
	pshippingCity=rs("shippingCity")
	pshippingStateCode=rs("shippingStateCode")
	pshippingState=rs("shippingState")
	pshippingZip=rs("shippingZip")
	pshippingPhone=rs("pcOrd_shippingPhone")
	pshippingCountryCode=rs("shippingCountryCode")	
	pshippingCompany=rs("shippingCompany")
	pshippingAddress2=rs("shippingAddress2")	
	pShippingEmail=rs("pcOrd_ShippingEmail")
	pSRF=rs("SRF")
	pshipmentDetails=rs("shipmentDetails")
	pOrdShipType=rs("ordShipType")
	pOrdPackageNum=rs("ordPackageNum")
	pcOrd_ShipWeight=rs("pcOrd_ShipWeight")
	pshippingFax=rs("pcOrd_ShippingFax")
	
end if
set rs=nothing	


' Shipment
If pSRF="1" then
	pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_b")
else
	shipping=split(pshipmentDetails,",")
	if ubound(shipping)>1 then
		if NOT isNumeric(trim(shipping(2))) then
			varShip="0"
			pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
		else
			Shipper=shipping(0)
			Service=shipping(1)
			Postage=trim(shipping(2))
			if ubound(shipping)=3 then
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
		pshipmentDetails=ship_dictLanguage.Item(Session("language")&"_noShip_a")
	end if 
end if

'// SHIPPER EMAIL
query="SELECT * FROM emailsettings WHERE id=1;"	
set rs=Server.CreateObject("ADODB.Recordset")     
set rs=conntemp.execute(query)	
if err.number <> 0 then
	set rs = nothing
	call closeDb()
end if 
ownerEmail=rs("ownerEmail")
set rs = nothing	

'// SHIP TIME
if Session("pcAdminShipTime") = "" then
	todaysDate=now()
	pcv_strShipTime = FormatDateTime(todaysDate,4) & ":00"
	Session("pcAdminShipTime") = pcv_strShipTime
end if	


'// Residential Delivery
if Session("pcAdminResidentialDelivery") = "" then
	if pOrdShipType=0 then
		pcv_strResidentialDelivery = "1"
	else
		pcv_strResidentialDelivery = "0" 
	end if	
	Session("pcAdminResidentialDelivery") = pcv_strResidentialDelivery
end if

'// SHIP CONSTANTS

'// DropType
if Session("pcAdminDropoffType") = "" then
	pcv_strDropoffType = FEDEX_DROPOFF_TYPE
	Session("pcAdminDropoffType") = pcv_strDropoffType
end if

'// PackageType
if Session("pcAdminPackaging1") = "" then
	pcv_strPackaging = FEDEX_FEDEX_PACKAGE
	Session("pcAdminPackaging1") = pcv_strPackaging
	Session("pcAdminPackaging2") = pcv_strPackaging
	Session("pcAdminPackaging3") = pcv_strPackaging
	Session("pcAdminPackaging4") = pcv_strPackaging
end if

'// L
if Session("pcAdminLength1") = "" then
	pcv_strLength = FEDEX_LENGTH
	Session("pcAdminLength1") = pcv_strLength
	Session("pcAdminLength2") = pcv_strLength
	Session("pcAdminLength3") = pcv_strLength
	Session("pcAdminLength4") = pcv_strLength
end if

'// W
if Session("pcAdminWidth1") = "" then
	pcv_strWidth = FEDEX_WIDTH
	Session("pcAdminWidth1") = pcv_strWidth
	Session("pcAdminWidth2") = pcv_strWidth
	Session("pcAdminWidth3") = pcv_strWidth
	Session("pcAdminWidth4") = pcv_strWidth
end if

'// H
if Session("pcAdminHeight1") = "" then
	pcv_strHeight = FEDEX_HEIGHT
	Session("pcAdminHeight1") = pcv_strHeight
	Session("pcAdminHeight2") = pcv_strHeight
	Session("pcAdminHeight3") = pcv_strHeight
	Session("pcAdminHeight4") = pcv_strHeight
end if

'// U
if Session("pcAdminUnits1") = "" then
	pcv_strUnits = FEDEX_DIM_UNIT
	Session("pcAdminUnits1") = pcv_strUnits
	Session("pcAdminUnits2") = pcv_strUnits
	Session("pcAdminUnits3") = pcv_strUnits
	Session("pcAdminUnits4") = pcv_strUnits
end if

if Session("pcAdminWeightUnits1") = "" then
	pcv_strWeightUnits = scShipFromWeightUnit
	Session("pcAdminWeightUnits1") = pcv_strWeightUnits
end if	

'// SHIPPER INFO
if Session("pcAdminOriginPersonName") = "" then
	pcv_strOriginPersonName = scOriginPersonName
	Session("pcAdminOriginPersonName") = pcv_strOriginPersonName
end if

if Session("pcAdminOriginCompanyName") = "" then
	pcv_strOriginCompanyName = scShipFromName
	Session("pcAdminOriginCompanyName") = pcv_strOriginCompanyName
end if

if Session("pcAdminOriginDepartment") = "" then
	pcv_strOriginDepartment = scOriginDepartment
	Session("pcAdminOriginDepartment") = pcv_strOriginDepartment
end if

if Session("pcAdminOriginPhoneNumber") = "" then
	pcv_strOriginPhoneNumber = scOriginPhoneNumber
	Session("pcAdminOriginPhoneNumber") = pcv_strOriginPhoneNumber
end if
if Session("pcAdminOriginPagerNumber") = "" then
	pcv_strOriginPagerNumber = scOriginPagerNumber
	Session("pcAdminOriginPagerNumber") = pcv_strOriginPagerNumber
end if
if Session("pcAdminOriginFaxNumber") = "" then
	pcv_strOriginFaxNumber = scOriginFaxNumber
	Session("pcAdminOriginFaxNumber") = pcv_strOriginFaxNumber
end if					
if Session("pcAdminOriginEmailAddress") = "" then
	pcv_strOriginEmailAddress = ownerEmail
	Session("pcAdminOriginEmailAddress") = pcv_strOriginEmailAddress
end if

'// ORIGIN ADDRESS
if Session("pcAdminOriginLine1") = "" then
	pcv_strOriginLine1 = scShipFromAddress1
	Session("pcAdminOriginLine1") = pcv_strOriginLine1
end if
if Session("pcAdminOriginLine2") = "" then
	pcv_strOriginLine2 = scShipFromAddress2
	Session("pcAdminOriginLine2") = pcv_strOriginLine2
end if
if Session("pcAdminOriginCity") = "" then
	pcv_strOriginCity = scShipFromCity
	Session("pcAdminOriginCity") = pcv_strOriginCity
end if
if Session("pcAdminOriginStateOrProvinceCode") = "" then
	pcv_strOriginStateOrProvinceCode = scShipFromState
	Session("pcAdminOriginStateOrProvinceCode") = pcv_strOriginStateOrProvinceCode
end if
if Session("pcAdminOriginPostalCode") = "" then
	pcv_strOriginPostalCode = scShipFromPostalCode
	Session("pcAdminOriginPostalCode") = pcv_strOriginPostalCode
end if
if Session("pcAdminOriginCountryCode") = "" then
	pcv_strOriginCountryCode = scShipFromPostalCountry
	Session("pcAdminOriginCountryCode") = pcv_strOriginCountryCode
end if

'// RECIPIENT
if Session("pcAdminRecipPersonName") = "" then
	pcv_strRecipPersonName = pShippingFullname
	Session("pcAdminRecipPersonName") = pcv_strRecipPersonName
end if
if Session("pcAdminRecipCompanyName") = "" then
	pcv_strRecipCompanyName = pshippingCompany
	Session("pcAdminRecipCompanyName") = pcv_strRecipCompanyName
end if

if Session("pcAdminRecipPhoneNumber") = "" then
	pcv_strRecipPhoneNumber = pshippingPhone
	Session("pcAdminRecipPhoneNumber") = pcv_strRecipPhoneNumber
end if

if Session("pcAdminRecipPhoneNumber") = "" then
	Session("pcAdminDeliveryPhone") = Session("pcAdminRecipPhoneNumber")
end if

if Session("pcAdminRecipFaxNumber") = "" then
	pcv_strRecipFaxNumber = pshippingFax
	Session("pcAdminRecipFaxNumber") = pcv_strRecipFaxNumber
end if	
			
if Session("pcAdminRecipEmailAddress") = "" then
	pcv_strRecipEmailAddress = pShippingEmail
	Session("pcAdminRecipEmailAddress") = pcv_strRecipEmailAddress
end if

'   >>> Origin Address Conditionals
'// Use the Request object to toggle State (based of Country selection)
isRequiredState =  true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	isRequiredState=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
isRequiredProvince = false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	isRequiredProvince=pcv_strProvinceCodeRequired
end if

'// DESTINATION ADDRESS	
if Session("pcAdminRecipLine1") = "" then
	pcv_strRecipLine1 = pshippingAddress
	Session("pcAdminRecipLine1") = pcv_strRecipLine1
end if
if Session("pcAdminRecipLine2") = "" then
	pcv_strRecipLine2 = pshippingAddress2
	Session("pcAdminRecipLine2") = pcv_strRecipLine2
end if
if Session("pcAdminRecipCity") = "" then
	pcv_strRecipCity = pshippingCity
	Session("pcAdminRecipCity") = pcv_strRecipCity
end if
if Session("pcAdminRecipStateOrProvinceCode") = "" then
	pcv_strRecipStateOrProvinceCode = pshippingStateCode
	Session("pcAdminRecipStateOrProvinceCode") = pcv_strRecipStateOrProvinceCode
end if
if Session("pcAdminRecipPostalCode") = "" then
	pcv_strRecipPostalCode = pshippingZip
	Session("pcAdminRecipPostalCode") = pcv_strRecipPostalCode
end if
if Session("pcAdminRecipCountryCode") = "" then
	pcv_strRecipCountryCode = pshippingCountryCode
	Session("pcAdminRecipCountryCode") = pcv_strRecipCountryCode
end if	

'   >>> Recipient Address Conditionals
'// Use the Request object to toggle State (based of Country selection)
isRequiredState2 =  true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired2")
if  len(pcv_strStateCodeRequired)>0 then
	isRequiredState2=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
isRequiredProvince2 = false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired2")
if  len(pcv_strProvinceCodeRequired)>0 then
	isRequiredProvince2=pcv_strProvinceCodeRequired
end if

if Session("pcAdminShipToCountryCode") = "US" OR Session("pcAdminShipToCountryCode") = "CA" then
	isRequiredShipToPostal = true
end if	

	
if Session("pcAdminCustomerReference") = "" then
	pcv_strCustomerReference = pidcustomer
	Session("pcAdminCustomerReference") = pcv_strCustomerReference
end if	
if Session("pcAdminCustomerInvoiceNumber") = "" then
	CustomerInvoiceNumber = pidorder
	Session("pcAdminCustomerInvoiceNumber") = CustomerInvoiceNumber
end if
if Session("pcAdminType") = "" then
	pcv_strType = "2DCOMMON"
	Session("pcAdminType") = pcv_strType
end if


'// SET REQUIRED VARIABLES
pcv_strMethodName = "FDXShipRequest"
pcv_strMethodReply = "FDXShipReply"	
CustomerTransactionIdentifier = "ProductCart_Test"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2">Order ID#: <b><%=(scpre+int(pcv_intOrderID))%></b></td>
	</tr>
	<tr>
		<th colspan="2">FedEx<sup>&reg;</sup> Shipment Request</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
			<p>
			This flexible service allows a customer to request shipments and print return labels.  Simply fillout all required fields from each of the console's tabs.
			Then click the "Process Shipment" button to send your request to FedEx.  If any error or warning occurs it will be displayed on your screen.
			Once your order is confirmed you will be redirected back to the Shipping Wizard for FedEx.
			</p>
		</td>
	</tr>
</table>
<table class="pcCPcontent">
	<tr>
		<td>
			<%		
			'*******************************************************************************
			' START: ON POSTBACK
			'*******************************************************************************		 
			if request.form("submit")<>"" then
										
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' ServerSide Validate the Required Fields and Formatting.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
				'// Generic error for page
				pcv_strGenericPageError = "At least one required field was empty."
				
				'// Clear error string
				pcv_strSecondaryErrors = ""
				pcv_strErrorMsg = ""
				
				'// Get all the dynamic package details and validate
				pcv_xCounter = 1
				pcv_strTotalDeclaredValue = 0
				pcv_strTotalWeight = 0
				For pcv_xCounter = 1 to pcPackageCount
				
					' If its shipped the field is no longer required
					if pcLocalArray(pcv_xCounter-1) = "shipped" then	
						pcv_strToggle = false
					else
						pcv_strToggle = true
					end if	
					
					pcs_ValidateTextField	"FaxLetter"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"Service"&pcv_xCounter, pcv_strToggle, 0
					pcs_ValidateTextField	"Length"&pcv_xCounter, false, 5
					pcs_ValidateTextField	"Width"&pcv_xCounter, false, 5
					pcs_ValidateTextField	"Height"&pcv_xCounter, false, 5 
					
					pcs_ValidateTextField	"Units"&pcv_xCounter, false, 3 						
					pcs_ValidateTextField	"Packaging"&pcv_xCounter, pcv_strToggle, 0
					pcs_ValidateTextField	"WeightUnits"&pcv_xCounter, pcv_strToggle, 0
					pcs_ValidateTextField	"Weight"&pcv_xCounter, pcv_strToggle, 0
					pcs_ValidateTextField	"declaredvalue"&pcv_xCounter, pcv_strToggle, 0	
					
					Session("pcAdminLength"&pcv_xCounter)=int(Session("pcAdminLength"&pcv_xCounter))
					Session("pcAdminWidth"&pcv_xCounter)=int(Session("pcAdminWidth"&pcv_xCounter))
					Session("pcAdminHeight"&pcv_xCounter)=int(Session("pcAdminHeight"&pcv_xCounter))
					
					if Session("pcAdminWeight"&pcv_xCounter)="" then Session("pcAdminWeight"&pcv_xCounter) = 0
					if Session("pcAdmindeclaredvalue"&pcv_xCounter)="" then Session("pcAdmindeclaredvalue"&pcv_xCounter) = 0
					pcv_strTotalDeclaredValue = pcv_strTotalDeclaredValue + Session("pcAdmindeclaredvalue"&pcv_xCounter)
					pcv_strTotalWeight = pcv_strTotalWeight + Session("pcAdminWeight"&pcv_xCounter)
						
				Next			
				Session("pcAdminTotalWeight")=FormatNumber(pcv_strTotalWeight,1)
				Session("pcAdminTotalDeclaredValue")=FormatNumber(pcv_strTotalDeclaredValue,2)				
				pcs_ValidateTextField	"CarrierCode", true, 10				
				pcs_ValidateTextField	"ShipDate", true, 0
				pcs_ValidateTextField	"ShipTime", true, 0
				pcs_ValidateTextField	"ReturnShipmentIndicator", false, 0				
				if Session("pcAdminCarrierCode") = "FDXE" then
					isRequiredDropoffType = true
				else
					isRequiredDropoffType = false
				end if
				pcs_ValidateTextField "DropoffType", isRequiredDropoffType, 0
				pcs_ValidateTextField	"CurrencyCode", false, 3
				pcs_ValidateTextField	"ListRate", false, 1
				
				'// Origin
				pcs_ValidateTextField	"OriginPersonName", true, 35
				pcs_ValidateTextField	"OriginCompanyName", true, 35
				pcs_ValidateTextField	"OriginDepartment", false, 10
				pcs_ValidatePhoneNumber	"OriginPhoneNumber", true, 16
				pcs_ValidatePhoneNumber	"OriginPagerNumber", false, 16
				pcs_ValidatePhoneNumber	"OriginFaxNumber", false, 16
				pcs_ValidateEmailField	"OriginEmailAddress", true, 0								
				pcs_ValidateTextField	"OriginLine1", true, 35
				pcs_ValidateTextField	"OriginLine2", false, 35
				pcs_ValidateTextField	"OriginCity", true, 0
				pcs_ValidateTextField	"OriginStateOrProvinceCode", isRequiredState, 2
				pcs_ValidateTextField	"OriginProvinceCode", isRequiredProvince, 2
				pcs_ValidateTextField	"OriginPostalCode", true, 16
				pcs_ValidateTextField	"OriginCountryCode", true, 2
				
				Session("pcAdminOriginPostalCode")=replace(Session("pcAdminOriginPostalCode"),"-","")
				
				'   >>> Merge Province or State into one variable before we send to FedEx
				if Session("pcAdminOriginProvinceCode") <> "" then
					Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminOriginProvinceCode")
				else
					Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminOriginStateOrProvinceCode")
				end if
				
				' Recipient	
				pcs_ValidateTextField	"RecipPersonName", true, 35
				pcs_ValidateTextField	"RecipCompanyName", false, 35
				pcs_ValidateTextField	"RecipDepartment", false, 10
				pcs_ValidatePhoneNumber	"RecipPhoneNumber", true, 16
				pcs_ValidatePhoneNumber	"RecipPagerNumber", false, 16
				pcs_ValidatePhoneNumber	"RecipFaxNumber", false, 16
				pcs_ValidateEmailField	"RecipEmailAddress", false, 0
				
				'// International
				pcs_ValidateTextField	"DutiesAccountNumber", false, 12
				pcs_ValidateTextField	"DutiesCountryCode", false, 2
				pcs_ValidateTextField	"DutiesPayorType", false, 10
				
				'// Recipient Address
				pcs_ValidateTextField	"RecipCountryCode", true, 2			
				
				'   >>> Recipient Address Conditionals
				if Session("pcAdminRecipCountryCode") = "US" OR Session("pcAdminRecipCountryCode") = "CA" then
					isRequiredRecipPostal = true
				else
					isRequiredRecipPostal = false
				end if				
				pcs_ValidateTextField	"RecipLine1", true, 35
				pcs_ValidateTextField	"RecipLine2", false, 35
				pcs_ValidateTextField	"RecipCity", true, 0 '// FDXE-35, FDXG-20				
				pcs_ValidateTextField	"RecipStateOrProvinceCode", isRequiredState2, 2
				pcs_ValidateTextField	"RecipProvinceCode", isRequiredProvince2, 2					
				
				'   >>> Merge Province or State into one variable before we send to FedEx
				if Session("pcAdminRecipProvinceCode") <> "" then
					Session("pcAdminRecipStateOrProvinceCode")=Session("pcAdminRecipProvinceCode")
				else
					Session("pcAdminRecipStateOrProvinceCode")=Session("pcAdminRecipStateOrProvinceCode")
				end if
			
				pcs_ValidateTextField	"RecipPostalCode", isRequiredRecipPostal, 16	
				
				Session("pcAdminRecipPostalCode")=replace(Session("pcAdminRecipPostalCode"),"-","")				
							
				pcs_ValidateTextField	"PayorType", false, 0 '// Required if PayorType is RECIPIENT or THIRDPARTY. 				
				pcs_ValidateTextField	"PayorAccountNumber", false, 0
				pcs_ValidateTextField	"PayorCountryCode", false, 2
				
				'// Customer Reference
				pcs_ValidateTextField	"CustomerReference", true, 0 '// FDXE-40, FDXG-30 
				pcs_ValidateTextField	"CustomerPONumber", false, 30
				pcs_ValidateTextField	"CustomerInvoiceNumber", false, 30 
				
				'// Declared Value
				' >>> This is hardcoded into the xml transaction for now.
								
				'// COD
				pcs_ValidateTextField	"AddTransportationCharges", false, 45 
				pcs_ValidateTextField	"CollectionAmount", false, 45 '// Required if COD element is provided.  Format: Two explicit decimals (e.g.5.00). 
				pcs_ValidateTextField	"CollectionType", false, 45 '// ANY, GUARANTEEDFUNDS, CASH 
				
				'// CODReturn 
				' >>> will default to the shipper if not specified	
				pcs_ValidateTextField	"TrackingNumber", false, 20 '// Required for a COD multiple-piece shipments only. / Tracking number assigned to the COD remittance 
				pcs_ValidateTextField	"ReferenceIndicator", false, 0 '// TRACKING, REFERENCE, PO, INVOICE
				
				'// Expandable Regions
				pcs_ValidateTextField	"bOrder", false, 0
				pcs_ValidateTextField	"bShip", false, 0
				pcs_ValidateTextField	"bAdditional", false, 0
				pcs_ValidateTextField	"bInternational", false, 0
				pcs_ValidateTextField	"bGround", false, 0
												
				'// International
				pcs_ValidateTextField	"TotalCustomsValue", false, 0
				pcs_ValidateTextField	"TermsOfSale", false, 0
				pcs_ValidateTextField	"AdmissibilityPackageType", false, 0
				if Session("pcAdminRecipCountryCode")<>"US" then
					pcs_ValidateTextField	"RecipientTIN", true, 0
				else
					pcs_ValidateTextField	"RecipientTIN", false, 0			
				end if
				pcs_ValidateTextField	"SenderTINOrDUNS", false, 0
				pcs_ValidateTextField	"SenderTINOrDUNSType", false, 0
				pcs_ValidateTextField	"AESOrFTSRExemptionNumber", false, 0
				pcs_ValidateTextField	"NumberOfPieces", false, 0
				pcs_ValidateTextField	"Description", false, 0
				pcs_ValidateTextField	"CountryOfManufacture", false, 0
				pcs_ValidateTextField	"HarmonizedCode", false, 0
				pcs_ValidateTextField	"CommodityWeight", false, 0
				pcs_ValidateTextField	"CommodityQuantity", false, 0
				pcs_ValidateTextField	"CommodityQuantityUnits", false, 0
				pcs_ValidateTextField	"CommodityUnitPrice", false, 0
				pcs_ValidateTextField	"CommodityCustomsValue", false, 0
				pcs_ValidateTextField	"ExportLicenseNumber", false, 0
				pcs_ValidateTextField	"ExportLicenseExpirationDate", false, 0
				pcs_ValidateTextField	"CIMarksAndNumbers", false, 0
				
				'// Hold At Location
				pcs_ValidateTextField	"HALPhone", false, 0
				pcs_ValidateTextField	"HALLine1", false, 0
				pcs_ValidateTextField	"HALCity", false, 0
				pcs_ValidateTextField	"HALStateOrProvinceCode", false, 0
				pcs_ValidateTextField	"HALPostalCode", false, 0			
				
				'// Thermal Support				
				pcs_ValidateTextField	"Type", false, 0		
				pcs_ValidateTextField	"ImageType", true, 0
				
				'// Freight
				pcs_ValidateTextField "BookingConfirmationNumber", false, 12
				
				'//Special Services
				pcs_ValidateTextField "ResidentialDelivery", false, 0
				pcs_ValidateTextField "SaturdayPickup", false, 0
				pcs_ValidateTextField "SaturdayDelivery", false, 0
				pcs_ValidateTextField "SignatureOption", false, 0
				pcs_ValidateTextField "HoldAtLocation", false, 0
				
				'//Shipper Notification
				pcs_ValidateTextField	"ShipperNotification", false, 0
				pcs_ValidateTextField "ShipperShipmentNotification", false, 0 'value="1"
				pcs_ValidateTextField "ShipperDeliveryNotification", false, 0 'value="1"
				pcs_ValidateTextField "ShipperExceptionNotification", false, 0 'value="1"
				
				'//Recipient Notification
				pcs_ValidateTextField	"RecipientNotification", false, 0
				pcs_ValidateTextField "RecipientShipmentNotification", false, 0 'value="1"
				pcs_ValidateTextField "RecipientDeliveryNotification", false, 0 'value="1"
				pcs_ValidateTextField "RecipientExceptionNotification", false, 0 'value="1"
				
				'//Other Notification
				pcs_ValidateTextField "OtherShipmentNotification", false, 0 'value="1"
				pcs_ValidateTextField "OtherDeliveryNotification", false, 0 'value="1"
				pcs_ValidateTextField "OtherExceptionNotification", false, 0 'value="1"
				
				pcs_ValidatePhoneNumber "DeliveryType", false, 11
				pcs_ValidatePhoneNumber "DeliveryInstructions", false, 74
				pcs_ValidatePhoneNumber "DeliveryDate", false, 0
				pcs_ValidatePhoneNumber	"DeliveryPhone", false, 16	
				
				
				'// Additional Validation for Numerics
				if NOT validNum(Session("pcAdminResidentialDelivery")) OR Session("pcAdminResidentialDelivery")<>"1" then
					Session("pcAdminResidentialDelivery")="0"
				end if
				if NOT validNum(Session("pcAdminSaturdayPickup")) OR Session("pcAdminSaturdayPickup")<>"1" then
					Session("pcAdminSaturdayPickup")="0"
				end if
				if NOT validNum(Session("pcAdminSaturdayDelivery")) OR Session("pcAdminSaturdayDelivery")<>"1" then
					Session("pcAdminSaturdayDelivery")="0"
				end if
				if NOT validNum(Session("pcAdminShipperShipmentNotification")) OR Session("pcAdminShipperShipmentNotification")<>"1" then
					Session("pcAdminShipperShipmentNotification")="0"
				end if
				if NOT validNum(Session("pcAdminShipperDeliveryNotification")) OR Session("pcAdminShipperDeliveryNotification")<>"1" then
					Session("pcAdminShipperDeliveryNotification")="0"
				end if
				if NOT validNum(Session("pcAdminShipperExceptionNotification")) OR Session("pcAdminShipperExceptionNotification")<>"1" then
					Session("pcAdminShipperExceptionNotification")="0"
				end if
				if NOT validNum(Session("pcAdminRecipientShipmentNotification")) OR Session("pcAdminRecipientShipmentNotification")<>"1" then
					Session("pcAdminRecipientShipmentNotification")="0"
				end if
				if NOT validNum(Session("pcAdminRecipientDeliveryNotification")) OR Session("pcAdminRecipientDeliveryNotification")<>"1" then
					Session("pcAdminRecipientDeliveryNotification")="0"
				end if
				if NOT validNum(Session("pcAdminRecipientExceptionNotification")) OR Session("pcAdminRecipientExceptionNotification")<>"1" then
					Session("pcAdminRecipientExceptionNotification")="0"
				end if
				if NOT validNum(Session("pcAdminOtherShipmentNotification")) OR Session("pcAdminOtherShipmentNotification")<>"1" then
					Session("pcAdminOtherShipmentNotification")="0"
				end if
				if NOT validNum(Session("pcAdminOtherDeliveryNotification")) OR Session("pcAdminOtherDeliveryNotification")<>"1" then
					Session("pcAdminOtherDeliveryNotification")="0"
				end if
				if NOT validNum(Session("pcAdminOtherExceptionNotification")) OR Session("pcAdminOtherExceptionNotification")<>"1" then
					Session("pcAdminOtherExceptionNotification")="0"
				end if
				if NOT validNum(Session("pcAdminHoldAtLocation")) OR Session("pcAdminHoldAtLocation")<>"1" then
					Session("pcAdminHoldAtLocation")="0"
				end if
				If Session("pcAdminOtherShipmentNotification")="1" OR Session("pcAdminOtherDeliveryNotification")="1" OR Session("pcAdminOtherExceptionNotification")="1" then
					pcs_ValidateEmailField "OtherNotification1", true, 0
				else
					pcs_ValidateEmailField "OtherNotification1", false, 0
				end if
				if Session("pcAdminCommodityWeight")="" then
					Session("pcAdminCommodityWeight")=0
				end if
				if Session("pcAdminCommodityQuantity")="" then
					Session("pcAdminCommodityQuantity")=1
				end if
				if Session("pcAdminNumberOfPieces")="" then
					Session("pcAdminNumberOfPieces")=0
				end if
				
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Check for Validation Errors. Do not proceed if there are errors.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				If pcv_intErr>0 Then
					response.redirect pcPageName & "?sub=1&msg=" & pcv_strGenericPageError
				Else								
					
					'///////////////////////////////////////////////////////////////////////
					'// START LOOP
					'///////////////////////////////////////////////////////////////////////
					pcv_xCounter = 1
					pcv_strTotalDeclaredValue = 0
					pcv_strTotalWeight = 0
					errnum = 0
					For pcv_xCounter = 1 to pcPackageCount
					
					
					'// Reverse Address if Return Shipment
						if Session("pcAdminReturnShipmentIndicator")="PRINTRETURNLABEL" then
							pcv_a=Session("pcAdminOriginPersonName")
							pcv_b=Session("pcAdminOriginCompanyName")
							pcv_c=Session("pcAdminOriginDepartment")
							pcv_d=Session("pcAdminOriginPhoneNumber")
							pcv_e=Session("pcAdminOriginPagerNumber")
							pcv_f=Session("pcAdminOriginFaxNumber")
							pcv_g=Session("pcAdminOriginEmailAddress")
							pcv_h=Session("pcAdminOriginLine1")
							pcv_i=Session("pcAdminOriginLine2")
							pcv_j=Session("pcAdminOriginCity")
							pcv_k=Session("pcAdminOriginStateOrProvinceCode")
							pcv_l=Session("pcAdminOriginPostalCode")						
							pcv_m=Session("pcAdminOriginCountryCode")
							
							Session("pcAdminOriginPersonName")=Session("pcAdminRecipPersonName")
							Session("pcAdminOriginCompanyName")=Session("pcAdminRecipCompanyName")
							Session("pcAdminOriginDepartment")=Session("pcAdminRecipDepartment")
							Session("pcAdminOriginPhoneNumber")=Session("pcAdminRecipPhoneNumber")
							Session("pcAdminOriginPagerNumber")=Session("pcAdminRecipPagerNumber")
							Session("pcAdminOriginFaxNumber")=Session("pcAdminRecipFaxNumber")
							Session("pcAdminOriginEmailAddress")=Session("pcAdminRecipEmailAddress")
							Session("pcAdminOriginLine1")=Session("pcAdminRecipLine1")
							Session("pcAdminOriginLine2")=Session("pcAdminRecipLine2")
							Session("pcAdminOriginCity")=Session("pcAdminRecipCity")
							Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminRecipStateOrProvinceCode")
							Session("pcAdminOriginPostalCode")=Session("pcAdminRecipPostalCode")
							Session("pcAdminOriginCountryCode")=Session("pcAdminRecipCountryCode")						
							
							Session("pcAdminRecipPersonName")=pcv_a
							Session("pcAdminRecipCompanyName")=pcv_b
							Session("pcAdminRecipDepartment")=pcv_c
							Session("pcAdminRecipPhoneNumber")=pcv_d
							Session("pcAdminRecipPagerNumber")=pcv_e
							Session("pcAdminRecipFaxNumber")=pcv_f
							Session("pcAdminRecipEmailAddress")=pcv_g
							Session("pcAdminRecipLine1")=pcv_h
							Session("pcAdminRecipLine2")=pcv_i
							Session("pcAdminRecipCity")=pcv_j
							Session("pcAdminRecipStateOrProvinceCode")=pcv_k
							Session("pcAdminRecipPostalCode")=pcv_l
							Session("pcAdminRecipCountryCode")=pcv_m	
						end if	
						
						'// If the package was processed, skip it.
						if pcLocalArray(pcv_xCounter-1) <> "shipped" then	
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Build Our Transaction.
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						objFedExClass.NewXMLTransaction pcv_strMethodName, pcv_strAccountNumber, pcv_strMeterNumber, Session("pcAdminCarrierCode"), sanitizeField(Session("pcAdminRecipPersonName"))
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: SHIPMENT SETTINGS
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								objFedExClass.WriteSingleParent "ShipDate", Session("pcAdminShipDate")
								objFedExClass.WriteSingleParent "ShipTime", Session("pcAdminShipTime")
								objFedExClass.WriteSingleParent "DropoffType", Session("pcAdminDropoffType")
								objFedExClass.WriteSingleParent "Service", Session("pcAdminService"&pcv_xCounter)
								objFedExClass.WriteSingleParent "Packaging", Session("pcAdminPackaging"&pcv_xCounter)
								objFedExClass.WriteSingleParent "WeightUnits", Session("pcAdminWeightUnits"&pcv_xCounter)
								objFedExClass.WriteSingleParent "Weight", FormatNumber(Session("pcAdminWeight"&pcv_xCounter),1)
								objFedExClass.WriteSingleParent "CurrencyCode", Session("pcAdminCurrencyCode")
								objFedExClass.WriteSingleParent "ListRate", Session("pcAdminListRate") '// optional, CP
								objFedExClass.WriteSingleParent "ReturnShipmentIndicator", Session("pcAdminReturnShipmentIndicator")														
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: SHIPMENT SETTINGS
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: ORIGIN
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								objFedExClass.WriteParent "Origin", ""							
									objFedExClass.WriteParent "Contact", ""
										objFedExClass.AddNewNode "PersonName", sanitizeField(Session("pcAdminOriginPersonName"))
										objFedExClass.AddNewNode "CompanyName", sanitizeField(Session("pcAdminOriginCompanyName"))
										objFedExClass.AddNewNode "Department", sanitizeField(Session("pcAdminOriginDepartment"))
										objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminOriginPhoneNumber"))
										objFedExClass.AddNewNode "PagerNumber", fnStripPhone(Session("pcAdminOriginPagerNumber"))
										objFedExClass.AddNewNode "FaxNumber", fnStripPhone(Session("pcAdminOriginFaxNumber"))
										objFedExClass.AddNewNode "E-MailAddress", Session("pcAdminOriginEmailAddress")
									objFedExClass.WriteParent "Contact", "/"	
									objFedExClass.WriteParent "Address", ""
										objFedExClass.AddNewNode "Line1", Session("pcAdminOriginLine1")
										objFedExClass.AddNewNode "Line2", Session("pcAdminOriginLine2")
										objFedExClass.AddNewNode "City", Session("pcAdminOriginCity")
										objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminOriginStateOrProvinceCode")
										objFedExClass.AddNewNode "PostalCode", Session("pcAdminOriginPostalCode")
										objFedExClass.AddNewNode "CountryCode", Session("pcAdminOriginCountryCode")
									objFedExClass.WriteParent "Address", "/"				
								objFedExClass.WriteParent "Origin", "/"
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: ORIGIN
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: RECIPIENT
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								objFedExClass.WriteParent "Destination", ""							
									objFedExClass.WriteParent "Contact", ""
										objFedExClass.AddNewNode "PersonName", sanitizeField(Session("pcAdminRecipPersonName"))
										objFedExClass.AddNewNode "CompanyName", sanitizeField(Session("pcAdminRecipCompanyName"))
										objFedExClass.AddNewNode "Department", sanitizeField(Session("pcAdminRecipDepartment"))
										objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminRecipPhoneNumber"))
										objFedExClass.AddNewNode "PagerNumber", fnStripPhone(Session("pcAdminRecipPagerNumber"))
										objFedExClass.AddNewNode "FaxNumber", fnStripPhone(Session("pcAdminRecipFaxNumber"))
										objFedExClass.AddNewNode "E-MailAddress", Session("pcAdminRecipEmailAddress")
									objFedExClass.WriteParent "Contact", "/"	
									objFedExClass.WriteParent "Address", ""
										objFedExClass.AddNewNode "Line1", Session("pcAdminRecipLine1")
										objFedExClass.AddNewNode "Line2", Session("pcAdminRecipLine2")
										objFedExClass.AddNewNode "City", Session("pcAdminRecipCity")
										if Session("pcAdminRecipCountryCode")="US" OR Session("pcAdminRecipCountryCode")="CA" then
											objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminRecipStateOrProvinceCode")
										end if
										objFedExClass.AddNewNode "PostalCode", Session("pcAdminRecipPostalCode")
										objFedExClass.AddNewNode "CountryCode", Session("pcAdminRecipCountryCode")
									objFedExClass.WriteParent "Address", "/"				
								objFedExClass.WriteParent "Destination", "/"
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: RECIPIENT
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: PAYMENT
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								objFedExClass.WriteParent "Payment", ""
									objFedExClass.AddNewNode "PayorType", Session("pcAdminPayorType")
									if Session("pcAdminPayorAccountNumber") <> "" then
										objFedExClass.WriteParent "Payor", ""
											objFedExClass.AddNewNode "AccountNumber", Session("pcAdminPayorAccountNumber")
											objFedExClass.AddNewNode "CountryCode", Session("pcAdminPayorCountryCode")
										objFedExClass.WriteParent "Payor", "/"
									end if
								objFedExClass.WriteParent "Payment", "/"
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: PAYMENT
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: REFERENCE INFORMATION
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~								
								objFedExClass.WriteParent "ReferenceInfo", ""
									objFedExClass.AddNewNode "CustomerReference", Session("pcAdminCustomerReference")
									objFedExClass.AddNewNode "PONumber", Session("pcAdminCustomerPONumber")
									objFedExClass.AddNewNode "InvoiceNumber", Session("pcAdminCustomerInvoiceNumber")
								objFedExClass.WriteParent "ReferenceInfo", "/"
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: REFERENCE INFORMATION
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: DIMENSIONS
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
								if Session("pcAdminPackaging"&pcv_xCounter) = "YOURPACKAGING" then
								objFedExClass.WriteParent "Dimensions", ""
									objFedExClass.AddNewNode "Length", Session("pcAdminLength"&pcv_xCounter)
									objFedExClass.AddNewNode "Width", Session("pcAdminWidth"&pcv_xCounter)
									objFedExClass.AddNewNode "Height", Session("pcAdminHeight"&pcv_xCounter)
									objFedExClass.AddNewNode "Units", Session("pcAdminUnits"&pcv_xCounter)
								objFedExClass.WriteParent "Dimensions", "/"
								end if
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: DIMENSIONS
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: INTERNATIONAL
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
								if Session("pcAdminRecipCountryCode")<>"US" OR Session("pcAdminRecipStateOrProvinceCode")="PR" then
								objFedExClass.WriteParent "International", ""								
									objFedExClass.WriteSingleParent "TotalCustomsValue", Session("pcAdminTotalCustomsValue") '// two decimals
									objFedExClass.WriteSingleParent "TermsOfSale", Session("pcAdminTermsOfSale")
									objFedExClass.WriteSingleParent "RecipientTIN", Session("pcAdminRecipientTIN")	
									objFedExClass.WriteSingleParent "AdmissibilityPackageType", Session("pcAdminAdmissibilityPackageType")		
												 
									objFedExClass.WriteParent "DutiesPayment", ""
										if Session("pcAdminDutiesPayorType") <> "SENDER" then
											objFedExClass.AddNewNode "PayorType", Session("pcAdminDutiesPayorType")
											objFedExClass.WriteParent "DutiesPayor", ""
												objFedExClass.AddNewNode "CountryCode", Session("pcAdminDutiesCountryCode")
												objFedExClass.AddNewNode "AccountNumber", Session("pcAdminDutiesAccountNumber")
											objFedExClass.WriteParent "DutiesPayor", "/"
										else
											objFedExClass.AddNewNode "PayorType", "SENDER"
										end if										
									objFedExClass.WriteParent "DutiesPayment", "/"
									
									'// INTERNATIONAL / COMMERCIALINVOICE	
									if pcv_strCommercialInvoice = "Not currently an active feature" then																	
										objFedExClass.WriteParent "CommercialInvoice", ""
											objFedExClass.AddNewNode "Comments", ""
											objFedExClass.AddNewNode "FreightCharge", ""
											objFedExClass.AddNewNode "InsuranceCharge", ""
											objFedExClass.AddNewNode "TaxesOrMiscellaneousCharge", ""
											objFedExClass.AddNewNode "Purpose", ""
											objFedExClass.AddNewNode "CustomerInvoiceNumber", ""
										objFedExClass.WriteParent "CommercialInvoice", "/"									
									end if
									
									'// INTERNATIONAL / COMMODITY	
									if Session("pcAdminNumberOfPieces")>0 then											
										if Session("pcAdminDescription")<>"" then
											objFedExClass.WriteParent "SED", ""
												objFedExClass.AddNewNode "SenderTINOrDUNS", Session("pcAdminSenderTINOrDUNS")
												objFedExClass.AddNewNode "SenderTINOrDUNSType", Session("pcAdminSenderTINOrDUNSType")
												objFedExClass.AddNewNode "AESOrFTSRExemptionNumber", Session("pcAdminAESOrFTSRExemptionNumber")
											objFedExClass.WriteParent "SED", "/"	
											objFedExClass.WriteParent "Commodity", ""
												objFedExClass.AddNewNode "NumberOfPieces", Session("pcAdminNumberOfPieces")
												objFedExClass.AddNewNode "Description", Session("pcAdminDescription")
												objFedExClass.AddNewNode "CountryOfManufacture", Session("pcAdminCountryOfManufacture")
												objFedExClass.AddNewNode "HarmonizedCode", Session("pcAdminHarmonizedCode")
												objFedExClass.AddNewNode "Weight", Session("pcAdminCommodityWeight") '// one decimal
												objFedExClass.AddNewNode "Quantity", Session("pcAdminCommodityQuantity")
												objFedExClass.AddNewNode "QuantityUnits", Session("pcAdminCommodityQuantityUnits")
												objFedExClass.AddNewNode "UnitPrice", Session("pcAdminCommodityUnitPrice") '// six decimals
												objFedExClass.AddNewNode "CustomsValue", Session("pcAdminCommodityCustomsValue")
												objFedExClass.AddNewNode "ExportLicenseNumber", Session("pcAdminExportLicenseNumber")
												objFedExClass.AddNewNode "ExportLicenseExpirationDate", Session("pcAdminExportLicenseExpirationDate") '// valid FedEx Date format
												objFedExClass.AddNewNode "CIMarksAndNumbers", Session("pcAdminCIMarksAndNumbers")
											objFedExClass.WriteParent "Commodity", "/"	
										end if								
									end if	
								objFedExClass.WriteParent "International", "/"
								end if
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: INTERNATIONAL
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
								
								'// Declared Value
								if pcv_xCounter = 1 then '// if first package in shipment
									' >>> On multipiece shipments you must put the declared value in the first package.																							
									objFedExClass.WriteSingleParent "DeclaredValue", replace(Session("pcAdminTotalDeclaredValue"),",","")
								end if
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: SHIPMENT CONTENT
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								'// Insight/ Inbound Visibility 
								' If = true or 1, means block all but sender from seeing shipment content data. 
								objFedExClass.WriteParent "ShipmentContent", ""
									objFedExClass.WriteSingleParent "BlockShipmentData", 1
								objFedExClass.WriteParent "ShipmentContent", "/"
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: SHIPMENT CONTENT
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: SPECIAL SERVICES
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								objFedExClass.WriteParent "SpecialServices", ""		
									
									'// Hold At Location
									if Session("pcAdminHoldAtLocation")<>"0" then
										objFedExClass.WriteParent "HoldAtLocation", ""	
											objFedExClass.WriteSingleParent "PhoneNumber", fnStripPhone(Session("pcAdminHALPhone"))	
											objFedExClass.WriteParent "Address", ""
												objFedExClass.AddNewNode "Line1", Session("pcAdminHALLine1")  
												objFedExClass.AddNewNode "City", Session("pcAdminHALCity")
												objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminHALStateOrProvinceCode")
												objFedExClass.AddNewNode "PostalCode", Session("pcAdminHALPostalCode")
											objFedExClass.WriteParent "Address", "/"
										objFedExClass.WriteParent "HoldAtLocation", "/"	
									end if													
									
									'// Shipper Notification
									if Session("pcAdminShipperShipmentNotification")=1 OR Session("pcAdminShipperDeliveryNotification")=1 OR Session("pcAdminShipperExceptionNotification")=1 OR Session("pcAdminRecipientShipmentNotification")=1 OR Session("pcAdminRecipientDeliveryNotification")=1 OR Session("pcAdminRecipientExceptionNotification")=1 OR len(Session("pcAdminOtherNotification1"))>0 then
										objFedExClass.WriteParent "EMailNotification", ""
											if Session("pcAdminShipperNotification")="FAX" then
												objFedExClass.AddNewNode "ShipAlertFaxNumber", fnStripPhone(Session("pcAdminOriginFaxNumber")) 
											end if
											objFedExClass.AddNewNode "ShipAlertOptionalMessage", "ProductCart Shipping Alert"										 
												'// Notify Shipper
												if Session("pcAdminShipperShipmentNotification")=1 OR Session("pcAdminShipperDeliveryNotification")=1 OR Session("pcAdminShipperExceptionNotification")=1 then
												objFedExClass.WriteParent "Shipper", ""
													objFedExClass.AddNewNode "ShipAlert", Session("pcAdminShipperShipmentNotification") 
													objFedExClass.AddNewNode "DeliveryNotification", Session("pcAdminShipperDeliveryNotification")  
													objFedExClass.AddNewNode "ExceptionNotification", Session("pcAdminShipperExceptionNotification")  
													objFedExClass.AddNewNode "Format", "HTML" '// Text, HTML, WIRELESS
												objFedExClass.WriteParent "Shipper", "/"
												end if
												'// Notify Recipient
												if Session("pcAdminRecipientShipmentNotification")=1 OR Session("pcAdminRecipientDeliveryNotification")=1 OR Session("pcAdminRecipientExceptionNotification")=1 then
												objFedExClass.WriteParent "Recipient", ""
													objFedExClass.AddNewNode "ShipAlert", Session("pcAdminRecipientShipmentNotification")  
													objFedExClass.AddNewNode "DeliveryNotification", Session("pcAdminRecipientDeliveryNotification") 
													objFedExClass.AddNewNode "ExceptionNotification", Session("pcAdminRecipientExceptionNotification")  
													objFedExClass.AddNewNode "Format", "HTML" '// Text, HTML, WIRELESS
												objFedExClass.WriteParent "Recipient", "/"
												end if
												'// Other Notifications
												if len(Session("pcAdminOtherNotification1"))>0 then
													objFedExClass.WriteParent "Other", ""
														objFedExClass.AddNewNode "EMailAddress", Session("pcAdminOtherNotification1")   
														objFedExClass.AddNewNode "ShipAlert", Session("pcAdminOtherShipmentNotification") 
														objFedExClass.AddNewNode "DeliveryNotification", Session("pcAdminOtherDeliveryNotification")  
														objFedExClass.AddNewNode "ExceptionNotification", Session("pcAdminOtherExceptionNotification")  
														objFedExClass.AddNewNode "Format", "HTML" '// Text, HTML, WIRELESS
													objFedExClass.WriteParent "Other", "/"
												end if											
										objFedExClass.WriteParent "EMailNotification", "/"	
									end if	
									
									'// Residential Delivery															
									objFedExClass.AddNewNode "ResidentialDelivery", Session("pcAdminResidentialDelivery") 
									
									'// Future Day Shipment									
									if (DateDiff("d", FedExDateFormat(Date()), Session("pcAdminShipDate")) >= 1) then
									objFedExClass.AddNewNode "FutureDayShipment", 1 '// required for ground, defaults to false for express
									end if
									
									'// Inside Pickup
									if Request("InsidePickup") = "1" then
										objFedExClass.AddNewNode "InsidePickup", 0 '// If = true or 1, the shipment is originated from an inside pickup area.
									end if
									
									'// Inside Delivery
									if Request("InsideDelivery") = "1" then
										objFedExClass.AddNewNode "InsideDelivery", 1 '// If = true or 1, the shipment is traveling to an inside delivery area. (freight only)
									end if
									
									'// Saturday Services
									objFedExClass.AddNewNode "SaturdayPickup", Session("pcAdminSaturdayPickup") 
									objFedExClass.AddNewNode "SaturdayDelivery", Session("pcAdminSaturdayDelivery") 
									
									'// NonStandard Container
									'objFedExClass.AddNewNode "NonStandardContainer", 1 '// If = true or 1, the shipment is in a nonstandard container. 
									
									'// Signature Options
									if Session("pcAdminSignatureOption")<>"" then
										objFedExClass.AddNewNode "SignatureOption", Session("pcAdminSignatureOption") '// DELIVERWITHOUTSIGNATURE, INDIRECT, DIRECT, ADULT									
										objFedExClass.AddNewNode "SignatureRelease", Session("pcAdminSignatureRelease") '// used only with DELIVERWITHOUTSIGNATURE
									end if			
								objFedExClass.WriteParent "SpecialServices", "/"
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: SPECIAL SERVICES
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: LABELS
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								objFedExClass.WriteParent "Label", ""
									
									'// Default Label Settings - Non Custom Only
									objFedExClass.AddNewNode "Type", Session("pcAdminType")
									objFedExClass.AddNewNode "ImageType", Session("pcAdminImageType")
									objFedExClass.AddNewNode "LabelStockOrientation", "NONE"
									objFedExClass.AddNewNode "DocTabLocation", ""								
									
									'// This section is required so the origin address prints to the proper section "sender"
									if Session("pcAdminReturnShipmentIndicator") <> "PRINTRETURNLABEL" then									
									objFedExClass.WriteParent "DisplayedOrigin", ""
										objFedExClass.WriteParent "Contact", ""
											objFedExClass.AddNewNode "PersonName", sanitizeField(Session("pcAdminOriginPersonName"))
											objFedExClass.AddNewNode "CompanyName", sanitizeField(Session("pcAdminOriginCompanyName"))
											objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminOriginPhoneNumber"))
										objFedExClass.WriteParent "Contact", "/"	
										objFedExClass.WriteParent "Address", ""
											objFedExClass.AddNewNode "Line1", Session("pcAdminOriginLine1")
											objFedExClass.AddNewNode "Line2", Session("pcAdminOriginLine2")
											objFedExClass.AddNewNode "City", Session("pcAdminOriginCity")
											objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminOriginStateOrProvinceCode")
											objFedExClass.AddNewNode "PostalCode", Session("pcAdminOriginPostalCode")
											objFedExClass.AddNewNode "CountryCode", Session("pcAdminOriginCountryCode")
										objFedExClass.WriteParent "Address", "/"
									objFedExClass.WriteParent "DisplayedOrigin", "/"	
									end if
									
									'// Mask Account Number by Default					
									objFedExClass.AddNewNode "MaskAccountNumber", 1  '// If false, FedEx Account Number will be displayed on label.	
																
								objFedExClass.WriteParent "Label", "/"
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: LABELS
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: Multi-Package Shipment
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								if pcPackageCount > 1 then								
									objFedExClass.WriteParent "MultiPiece", ""
										objFedExClass.AddNewNode "PackageCount", pcPackageCount
										objFedExClass.AddNewNode "PackageSequenceNumber", pcv_xCounter 
										if Session("pcAdminRecipCountryCode")<>"US" then
											'// Required for international multiple-piece shipping. Format: One explicit decimal position (e.g. 5.0). 
											' >>> On multipiece shipments you must put the total weight value in the first package.	
											objFedExClass.AddNewNode "ShipmentWeight",  Session("pcAdminTotalWeight")
										end if
										if pcv_xCounter>1 then
											'// Required for multiple-piece shipping if PackageSequenceNumber value is greater than one.									
											objFedExClass.AddNewNode "MasterTrackingNumber", Session("MasterTrackingNumber")									
											objFedExClass.AddNewNode "MasterFormID", Session("MasterFormID")
										end if
									objFedExClass.WriteParent "MultiPiece", "/"
								end if
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: LABELS
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: FREIGHT
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								if Session("pcAdminService"&pcv_xCounter) = "INTERNATIONALPRIORITYFREIGHT" OR Session("pcAdminService"&pcv_xCounter) = "INTERNATIONALECONOMYFREIGHT" then								
									objFedExClass.WriteParent "Freight", ""										
										objFedExClass.AddNewNode "ShippersLoadAndCount", pcPackageCount
										objFedExClass.AddNewNode "BookingConfirmationNumber", Session("pcAdminBookingConfirmationNumber") 										
										'if Session("pcAdminRecipCountryCode")<>"US" then
										'	objFedExClass.AddNewNode "ShipmentWeight",  Session("pcAdminTotalWeight")
										'end if							
									objFedExClass.WriteParent "Freight", "/"
								end if
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: FREIGHT
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: HOME DELIVERY
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								if Session("pcAdminService"&pcv_xCounter) = "GROUNDHOMEDELIVERY" then
									if Session("pcAdminDeliveryType") = "DATECERTAIN" OR Session("pcAdminDeliveryType") = "APPOINTMENT" OR Session("pcAdminDeliveryType") = "EVENING" then
										objFedExClass.WriteParent "HomeDelivery", ""
											
											'// Required for home delivery shipments if delivery type is set to DATECERTAIN											
											if Session("pcAdminDeliveryType") = "DATECERTAIN" then
												objFedExClass.AddNewNode "Date", Session("pcAdminDeliveryDate")
											end if
											
											'// Applicable only to FedEx Home Delivery shipments.
											objFedExClass.AddNewNode "Instructions", Session("pcAdminDeliveryInstructions") 
											
											'// DeliveryType = DATECERTAIN, EVENING, APPOINTMENT
											objFedExClass.AddNewNode "Type", Session("pcAdminDeliveryType") 
											
											'// Required if home delivery type is set to DATECERTAIN or APPOINTMENT
											if Session("pcAdminDeliveryType") = "DATECERTAIN" OR Session("pcAdminDeliveryType") = "APPOINTMENT" then											
												objFedExClass.AddNewNode "PhoneNumber", "" 
											end if
											
										objFedExClass.WriteParent "HomeDelivery", "/"
									end if
								end if
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: HOME DELIVERY
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
							objFedExClass.EndXMLTransaction pcv_strMethodName
							
							
							
							
							'// Print out our newly formed request xml
							'response.write fedex_postdata
							'response.end
							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' Send Our Transaction.
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							call objFedExClass.SendXMLRequest(fedex_postdata, pcv_strEnvironment)
							
							'// Print out our response
							'response.write FEDEX_result
							'response.end

							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' Load Our Response.
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							call objFedExClass.LoadXMLResults(FEDEX_result)
							
							
							'/////////////////////////////////////////////////////////////
							'// BASELINE LOGGING
							'/////////////////////////////////////////////////////////////
							'// Tracking Number for Logs
							pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//Tracking", "TrackingNumber")
							'// Log our Transaction
							'call objFedExClass.pcs_LogTransaction(fedex_postdata, pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".in", true)
							'// Log our Response
							'call objFedExClass.pcs_LogTransaction(FEDEX_result, pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".out", true)
							
							if trim(FEDEX_result)="" then
								response.redirect ErrPageName & "?msg=FedEx was unable to send a response. There may have been a connection error. Please try again."
							end if
							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' Check for errors from FedEx.
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
							if pcv_xCounter = 1 then
								'// master package error, no processing done			
								call objFedExClass.XMLResponseVerifyCustom(ErrPageName)
								if len(pcv_strErrorMsg)>0 then
									'/////////////////////////////////////////////////////////////
									'// POSTBACK ERROR LOGGING
									'/////////////////////////////////////////////////////////////
									pcv_intRandomNumber = randomNumber(999999999)
									'// Log our Transaction
									call objFedExClass.pcs_LogTransaction(fedex_postdata, "ErrLog__"& pcv_intRandomNumber &".in", true)
									'// Log our Error Response
									call objFedExClass.pcs_LogTransaction(FEDEX_result, "ErrLog__"& pcv_intRandomNumber &".out", true)
									'/////////////////////////////////////////////////////////////					
									'// Display the Error
									response.redirect ErrPageName & "?msg=Your shipment was not processed for the following reason. " & pcv_strErrorMsg
								else																	
									pcLocalArray(pcv_xCounter-1) = "shipped"
									pcv_strItemsList = join(pcLocalArray, chr(124))
									Session("pcGlobalArray") = pcv_strItemsList
									'/////////////////////////////////////////////////////////////
									'// POSTBACK LOGGING
									'/////////////////////////////////////////////////////////////
									'// Tracking Number for Logs
									pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//Tracking", "TrackingNumber")
									if pcv_strTrackingNumber<>"" then
										'// Log our Transaction
										call objFedExClass.pcs_LogTransaction(fedex_postdata, pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".in", true)
										'// Log our Response
										call objFedExClass.pcs_LogTransaction(FEDEX_result, pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".out", true)
									else
										'// Log our Transaction
										call objFedExClass.pcs_LogTransaction(fedex_postdata, pcv_strMethodName&"_noTracking_"&pcv_xCounter&".in", true)
										'// Log our Response
										call objFedExClass.pcs_LogTransaction(FEDEX_result, pcv_strMethodName&"_noTracking_"&pcv_xCounter&".out", true)
									end if
									'/////////////////////////////////////////////////////////////
								end if
							else
								'// tack package errors, same checks with no redirect
								pcv_strErrorMsg = ""
								call objFedExClass.XMLResponseVerifyCustom(ErrPageName)
								if len(pcv_strErrorMsg)>0 then	
									'/////////////////////////////////////////////////////////////
									'// POSTBACK ERROR LOGGING
									'/////////////////////////////////////////////////////////////
									pcv_intRandomNumber = randomNumber(999999999)
									'// Log our Transaction
									call objFedExClass.pcs_LogTransaction(fedex_postdata, "ErrLog__"& pcv_intRandomNumber &".in", true)
									'// Log our Error Response
									call objFedExClass.pcs_LogTransaction(FEDEX_result, "ErrLog__"& pcv_intRandomNumber &".out", true)
									'/////////////////////////////////////////////////////////////					
									'// Pend an error string
									errnum = errnum + 1
									pcv_strSecondaryErrors = pcv_strSecondaryErrors & "<br />" & errnum & ".) " & pcv_strErrorMsg & "<br /> "
								else
									pcLocalArray(pcv_xCounter-1) = "shipped"
									pcv_strItemsList = join(pcLocalArray, chr(124))
									Session("pcGlobalArray") = pcv_strItemsList
									'/////////////////////////////////////////////////////////////
									'// POSTBACK LOGGING
									'/////////////////////////////////////////////////////////////
									'// Tracking Number for Logs
									pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//Tracking", "TrackingNumber")
									if pcv_strTrackingNumber<>"" then
										'// Log our Transaction
										call objFedExClass.pcs_LogTransaction(fedex_postdata, pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".in", true)
										'// Log our Response
										call objFedExClass.pcs_LogTransaction(FEDEX_result, pcv_strMethodName&"_"&pcv_strTrackingNumber&"_"&pcv_xCounter&".out", true)
									else
										'// Log our Transaction
										call objFedExClass.pcs_LogTransaction(fedex_postdata, pcv_strMethodName&"_noTracking_"&pcv_xCounter&".in", true)
										'// Log our Response
										call objFedExClass.pcs_LogTransaction(FEDEX_result, pcv_strMethodName&"_noTracking_"&pcv_xCounter&".out", true)
									end if
									'/////////////////////////////////////////////////////////////
								end if
							end if							

							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' Redirect with a Message OR complete some task.
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							if NOT len(pcv_strErrorMsg)>0 then
							
		
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' Set Our Response Data to Local.
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~							
								'// HEADER
								pcv_strCustomerTransactionIdentifier = objFedExClass.ReadResponseNode("//ReplyHeader", "CustomerTransactionIdentifier")	
								
								'// ERROR
								pcv_strErrorCode = objFedExClass.ReadResponseNode("//Error", "Code")
								pcv_strErrorMessage = objFedExClass.ReadResponseNode("//Error", "Message")
								
								'// SOFT ERROR
								pcv_strSoftError = objFedExClass.ReadResponseNode("//SoftError", "Type")
								
								'// TRACKING
								pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//Tracking", "TrackingNumber")
								pcv_strFormID = objFedExClass.ReadResponseNode("//Tracking", "FormID")						
								pcv_strServiceTypeDescription = objFedExClass.ReadResponseParent(pcv_strMethodReply, "ServiceTypeDescription")
								pcv_strCodReturnTrackingNumber = objFedExClass.ReadResponseNode("//Tracking", "CodReturnTrackingNumber")								
								pcv_strCodReturnFormID = objFedExClass.ReadResponseNode("//Tracking", "CodReturnFormID")
								pcv_strMasterTrackingNumber = objFedExClass.ReadResponseNode("//Tracking", "MasterTrackingNumber")								
								pcv_strMasterFormID = objFedExClass.ReadResponseNode("//Tracking", "MasterFormID")
								Session("CodReturnTrackingNumber") = pcv_strCodReturnTrackingNumber
								Session("MasterTrackingNumber") = pcv_strMasterTrackingNumber
								Session("MasterFormID") = pcv_strMasterFormID
								pcv_strPackagingDescription = objFedExClass.ReadResponseParent(pcv_strMethodReply, "PackagingDescription")
								
								
								'// ESTIMATED CHARGES						
								pcv_strDimWeightUsed = objFedExClass.ReadResponseNode("//EstimatedCharges", "DimWeightUsed")
								pcv_strRateScale = objFedExClass.ReadResponseNode("//EstimatedCharges", "RateScale")
								pcv_strRateZone = objFedExClass.ReadResponseNode("//EstimatedCharges", "RateZone")
								pcv_strCurrencyCode = objFedExClass.ReadResponseNode("//EstimatedCharges", "CurrencyCode")
								pcv_strBilledWeight = objFedExClass.ReadResponseNode("//EstimatedCharges", "BilledWeight")
								pcv_strDimWeight = objFedExClass.ReadResponseNode("//EstimatedCharges", "DimWeight")
								
								'// ESTIMATED CHARGES - DISCOUNTEDCHARGES
								pcv_strBaseCharge = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/BaseCharge")
								pcv_strTotalDiscount = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/TotalDiscount")
								pcv_strTotalSurcharge = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/TotalSurcharge")
								pcv_strNetCharge = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/NetCharge")
								pcv_strTotalRebate = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/TotalRebate")
								
								'// ESTIMATED CHARGES - DISCOUNTEDCHARGES - SURCHARGES
								pcv_strFuel = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/Surcharges/Fuel")
								pcv_strCOD = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/Surcharges/COD")
								pcv_strDeliverySignatureOptions = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/Surcharges/DeliverySignatureOptions")
								pcv_strNonStandardContainer = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/Surcharges/NonStandardContainer")
								pcv_strOversize = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/Surcharges/Oversize")
								pcv_strHomeDeliveryEveningDelivery = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/Surcharges/HomeDeliveryEveningDelivery")
								pcv_strHomeDeliveryDateCertain = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/Surcharges/HomeDeliveryDateCertain")
								pcv_strHomeDelivery = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/Surcharges/HomeDelivery")		
								pcv_strDeclaredValue = objFedExClass.ReadResponseNode("//EstimatedCharges", "DiscountedCharges/Surcharges/DeclaredValue")					
								
								'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
								
								'// LIST CHARGES
								pcv_strListBaseCharge = objFedExClass.ReadResponseNode("//EstimatedCharges", "BaseCharge")
								pcv_strListTotalDiscount = objFedExClass.ReadResponseNode("//EstimatedCharges", "TotalDiscount")
								
								'// ESTIMATED CHARGES - LISTCHARGES - SURCHARGES
								pcv_strListCOD = objFedExClass.ReadResponseNode("//EstimatedCharges", "ListCharges/Surcharges/COD")
								pcv_strListSignatureOptions = objFedExClass.ReadResponseNode("//EstimatedCharges", "ListCharges/Surcharges/DeliverySignatureOptions")
								pcv_strListNonStandardContainer = objFedExClass.ReadResponseNode("//EstimatedCharges", "ListCharges/Surcharges/NonStandardContainer")
								pcv_strListOversize = objFedExClass.ReadResponseNode("//EstimatedCharges", "ListCharges/Surcharges/Oversize")	
								pcv_strListHomeDeliveryEveningDelivery = objFedExClass.ReadResponseNode("//EstimatedCharges", "ListCharges/Surcharges/HomeDeliveryEveningDelivery")
								pcv_strListHomeDeliveryDateCertain = objFedExClass.ReadResponseNode("//EstimatedCharges", "ListCharges/Surcharges/HomeDeliveryDateCertain")
								pcv_strListHomeDelivery = objFedExClass.ReadResponseNode("//EstimatedCharges", "ListCharges/Surcharges/HomeDelivery")	
								pcv_strListDeclaredValue = objFedExClass.ReadResponseNode("//EstimatedCharges", "ListCharges/Surcharges/DeclaredValue")		
								
								
								'// COD
								pcv_strCollectionAmount = objFedExClass.ReadResponseNode("//COD", "CollectionAmount")
								pcv_strCODHandling = objFedExClass.ReadResponseNode("//COD", "Handling")
								pcv_strCODServiceTypeDescription = objFedExClass.ReadResponseNode("//COD", "ServiceTypeDescription")
								pcv_strCODPackagingDescription = objFedExClass.ReadResponseNode("//COD", "PackagingDescription")
								pcv_strCODSecuredDescription = objFedExClass.ReadResponseNode("//COD", "SecuredDescription")
								
								'// ROUTING
								pcv_strUrsaRoutingCode = objFedExClass.ReadResponseNode("//Routing", "UrsaRoutingCode")
								pcv_strServiceCommitment = objFedExClass.ReadResponseNode("//Routing", "ServiceCommitment")
								pcv_strDeliveryDay = objFedExClass.ReadResponseNode("//Routing", "DeliveryDay")
								pcv_strDestinationStationID = objFedExClass.ReadResponseNode("//Routing", "DestinationStationID")
								pcv_strDeliveryDate = objFedExClass.ReadResponseNode("//Routing", "DeliveryDate")
								pcv_strUrsaPrefixCode = objFedExClass.ReadResponseNode("//Routing", "UrsaPrefixCode")
								pcv_strCODReturnServiceCommitment = objFedExClass.ReadResponseNode("//Routing", "CODReturnServiceCommitment")
								
								'// LABEL
								pcv_strOutboundLabel = objFedExClass.ReadResponseNode("//Labels", "OutboundLabel")
								pcv_strCODReturnLabel = objFedExClass.ReadResponseNode("//Labels", "CODReturnLabel") 
								pcv_strLabelsSignatureOption = objFedExClass.ReadResponseNode("//Labels", "SignatureOption")
	
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: SAVE LABEL
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								if pcv_strOutboundLabel <> "" then
									'// Create XML for Label 
									objFedExClass.NewXMLLabel pcv_strTrackingNumber,pcv_strOutboundLabel, "PNG", "Label"
			
									'// Load label from the request stream
									call objFedExClass.LoadXMLLabel(GraphicXML)
			
									'// Use ADO stream to save the binary data
									objFedExClass.SaveBinaryLabel
										
	
									if pcv_strCODReturnLabel<>"" then
										'// Create XML for Label 
										objFedExClass.NewXMLLabel pcv_strTrackingNumber&"R",pcv_strCODReturnLabel, "PNG", "Label"
				
										'// Load label from the request stream
										call objFedExClass.LoadXMLLabel(GraphicXML)
				
										'// Use ADO stream to save the binary data
										objFedExClass.SaveBinaryLabel
									end if
								end if
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: SAVE LABEL
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								
								
								
								
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' START: SAVE PACKAGES
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~								
								
								'// Get Our Required Data
								pcv_method=Session("pcAdminService"&pcv_xCounter)
								pcv_tracking=pcv_strTrackingNumber
								pcv_shippedDate=Session("pcAdminShipDate")
								pcv_AdmComments=""
								
								'// Fix quotes on comments
								if pcv_AdmComments<>"" then
									pcv_AdmComments=replace(pcv_AdmComments,"'","''")
								end if
								
								dim dtShippedDate
								dtShippedDate=pcv_shippedDate
								if pcv_shippedDate<>"" then									
									'dtShippedDate=objFedExClass.pcf_FedExDateFormat(dtShippedDate)
									if SQL_Format="1" then
										dtShippedDate=(day(dtShippedDate)&"/"&month(dtShippedDate)&"/"&year(dtShippedDate))
									else
										dtShippedDate=(month(dtShippedDate)&"/"&day(dtShippedDate)&"/"&year(dtShippedDate))
									end if
								end if
								
								'// Insert Details into Package Info
								if pcv_shippedDate<>"" then
									if scDB="SQL" then
										query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) "
										query=query&"VALUES (" & pcv_intOrderID & ",'" & pcv_method & "','" & dtShippedDate & "','" & pcv_tracking & "','" & pcv_AdmComments & "', 3);"
									else
										query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_ShippedDate,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) " 
										query=query&"VALUES (" & pcv_intOrderID & ",'" & pcv_method & "',#" & dtShippedDate & "#,'" & pcv_tracking & "','" & pcv_AdmComments & "', 3);"
									end if
								else
										query="INSERT INTO pcPackageInfo (idOrder,pcPackageInfo_ShipMethod,pcPackageInfo_TrackingNumber,pcPackageInfo_Comments, pcPackageInfo_MethodFlag) " 
										query=query&"VALUES (" & pcv_intOrderID & ",'" & pcv_method & "','" & pcv_tracking & "','" & pcv_AdmComments & "', 3);"
								end if
								set rs=connTemp.execute(query)
								set rs=nothing
								
								'// Re-Query for the ID
								query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & pcv_intOrderID & " ORDER by pcPackageInfo_ID DESC;"
								set rs=connTemp.execute(query)
								pcv_PackageID=rs("pcPackageInfo_ID")
								set rs=nothing
								qry_ID=pcv_intOrderID
								
								'// Do a Full Update
								if pcv_strNetCharge = "" then
									pcv_strNetCharge = 0
								end if
								
								query=		"UPDATE pcPackageInfo "
								query=query&"SET pcPackageInfo_FDXSPODFlag=0, "
								query=query&"pcPackageInfo_PackageNumber=1, "
								query=query&"pcPackageInfo_PackageWeight=" & Session("pcAdminTotalWeight") & ", "						
								query=query&"pcPackageInfo_ShipToName='" & Session("pcAdminRecipPersonName") & "', "
								query=query&"pcPackageInfo_ShipToAddress1='" & Session("pcAdminRecipLine1") & "', "
								query=query&"pcPackageInfo_ShipToAddress2='" & Session("pcAdminRecipLine2") & "', "						
								query=query&"pcPackageInfo_ShipToCity='" & Session("pcAdminRecipCity") & "', "
								query=query&"pcPackageInfo_ShipToStateCode='" & Session("pcAdminRecipStateOrProvinceCode") & "', "
								query=query&"pcPackageInfo_ShipToZip='" & Session("pcAdminRecipPostalCode") & "', "
								query=query&"pcPackageInfo_ShipToCountry='" & Session("pcAdminRecipCountryCode") & "', "
								query=query&"pcPackageInfo_ShipToPhone='" & Session("pcAdminRecipPhoneNumber") & "', "
								query=query&"pcPackageInfo_ShipToEmail='" & Session("pcAdminRecipEmailAddress") & "', "				
								query=query&"pcPackageInfo_ShipToResidential=" & Session("pcAdminResidentialDelivery") & ", "
								query=query&"pcPackageInfo_PackageDescription='" & pcv_strPackagingDescription & "', "					
								query=query&"pcPackageInfo_ShipFromCompanyName='" & Session("pcAdminOriginCompanyName") & "', "
								query=query&"pcPackageInfo_ShipFromAttentionName='" & Session("pcAdminOriginPersonName") & "', "
								query=query&"pcPackageInfo_ShipFromPhoneNumber='" & Session("pcAdminOriginPhoneNumber") & "', "
								query=query&"pcPackageInfo_ShipFromAddress1='" & Session("pcAdminOriginLine1") & "', "
								query=query&"pcPackageInfo_ShipFromAddress2='" & Session("pcAdminOriginLine2") & "', "
								query=query&"pcPackageInfo_ShipFromCity='" & Session("pcAdminOriginCity") & "', "
								query=query&"pcPackageInfo_ShipFromStateProvinceCode='" & Session("pcAdminOriginStateOrProvinceCode") & "', "
								query=query&"pcPackageInfo_ShipFromPostalCode='" & Session("pcAdminOriginPostalCode") & "', "
								query=query&"pcPackageInfo_ShipFromCountryCode='" & Session("pcAdminOriginCountryCode") & "', "						
								query=query&"pcPackageInfo_UPSServiceCode='" & pcv_strServiceTypeDescription & "', "
								query=query&"pcPackageInfo_UPSPackageType='" & pcv_strPackagingDescription & "', "						
								query=query&"pcPackageInfo_PackageInsuredValue='" & pcv_strDeclaredValue & "', "						
								query=query&"pcPackageInfo_PackageLength='" & Session("pcAdminLength"&pcv_xCounter) & "', "
								query=query&"pcPackageInfo_PackageWidth='" & Session("pcAdminWidth"&pcv_xCounter) & "', "
								query=query&"pcPackageInfo_PackageHeight='" & Session("pcAdminHeight"&pcv_xCounter) & "', "	
								If pcv_strServiceTypeDescription="" Then
									pcv_strServiceTypeDescription = Session("pcAdminService"&pcv_xCounter)
								End If								
								query=query&"pcPackageInfo_ShipMethod='" & "FedEx: " & pcv_strServiceTypeDescription & "', "	
								query=query&"pcPackageInfo_FDXCarrierCode='" & Session("pcAdminCarrierCode") & "', "													
								query=query&"pcPackageInfo_FDXRate=" & pcv_strNetCharge & " "	
								query=query&"WHERE pcPackageInfo_ID=" & pcv_PackageID & " ;"
								
								set rstemp=connTemp.execute(query)
								set rs=nothing														

								
								'// Delete the old comments
								query="DELETE FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=2 AND pcPackageInfo_ID=" & pcv_PackageID & ";"


								set rstemp=connTemp.execute(query)
								
								'// Add the new comments
								query=		"INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments,pcDropShipper_ID,pcACom_IsSupplier,pcPackageInfo_ID) "
								query=query&"VALUES (" & qry_ID & ",2,'" & pcv_AdmComments & "',0,0," & pcv_PackageID & ");"
								set rstemp=connTemp.execute(query)
								
								if trim(Session("pcAdminPrdList"&pcv_xCounter))<>"" then
									pcA=split(Session("pcAdminPrdList"&pcv_xCounter),",")
									For i=lbound(pcA) to ubound(pcA)
										if trim(pcA(i)<>"") then
											query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1, pcPackageInfo_ID=" & pcv_PackageID & " WHERE (idorder=" & qry_ID & " AND idProductOrdered=" & pcA(i) & ");"
											set rs=connTemp.execute(query)
											set rs=nothing
										end if
									Next
								else
									query="UPDATE ProductsOrdered SET pcPackageInfo_ID=" & pcv_PackageID & " WHERE idorder=" & qry_ID & " AND pcPrdOrd_Shipped=0 AND pcDropShipper_ID=0;"
									set rsQ=connTemp.execute(query)
									set rsQ=nothing
								end if
								
								pcv_SendCust="1"
								pcv_SendAdmin="0"	
								pcv_LastShip="0"
								
								query="SELECT ProductsOrdered.pcPrdOrd_Shipped FROM ProductsOrdered INNER JOIN Orders ON (ProductsOrdered.idorder=Orders.idorder AND ProductsOrdered.pcPrdOrd_Shipped=0) WHERE Orders.idorder=" & qry_ID & " AND Orders.orderstatus<>4;"
								set rs=connTemp.execute(query)
								if not rs.eof then
									pcv_LastShip="0"
								else
									pcv_LastShip="1"
								end if
								set rs=nothing
								
								if trim(Session("pcAdminPrdList"&pcv_xCounter))<>"" then
									if pcv_LastShip="1" then
										query="UPDATE Orders SET orderStatus=4 WHERE idorder=" & qry_ID & ";"
									else
										query="UPDATE Orders SET orderStatus=7 WHERE idorder=" & qry_ID & ";"
									end if
									set rs=connTemp.execute(query)
									set rs=nothing
								end if								
								
								If pcv_LastShip="1" Then
									'// Perform a Google Action	
									pcv_strGoogleMethod = "mark" ' // Marks the order shipped at Google 			
									%> <!--#include file="../includes/GoogleCheckout_OrderManagement.asp"--> <%
								End If
								%>
								<!--#include file="../pc/inc_PartShipEmail.asp"-->	
								<%							
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
								' END: SAVE PACKAGES
								'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
								
								
								
							end if	' if NOT len(pcv_strErrorMsg)>0 then				

							
							' Determine if there were any errors. If not, redirect with a message.
							if (NOT len(pcv_strErrorMsg)>0) AND ((pcv_xCounter-1)=UBound(pcLocalArray)) then							
								'// Destroy the Sessions
								pcs_ClearAllSessions
								Session("pcAdminPackageCount")=""
								Session("pcAdminOrderID")=""
								Session("pcGlobalArray")=""
								For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
									Session("pcAdminPrdList"&(xArrayCount+1))
								Next
								
								'// REDIRECT
								response.redirect "FedEx_ManageShipmentsResults.asp?id=" & pcv_intOrderID & "&msg=Your transaction has been completed successfully."
								response.end
							elseif (pcv_xCounter-1)=UBound(pcLocalArray) then
							
								'// Generate an error report and redisplay the page.
								pcv_strSecondaryErrMsg="Your shipment was processed. <br />However, we have found the following errors:  <br />"
								pcv_strSecondarySolution="<br />You must resolve all errors to avoid delivery problems. "
								pcv_strSecondarySolution=pcv_strSecondarySolution&"First find and correct all form fields with errors. <br />Then click the 'Finish Processing' button. "
								pcv_strSecondarySolution = pcv_strSecondarySolution & "Repeat these steps until all packages are successfully shipped."
								
								'// ON ERROR - Reverse Address if Return Shipment
								if Session("pcAdminReturnShipmentIndicator")="PRINTRETURNLABEL" then
									pcv_a=Session("pcAdminOriginPersonName")
									pcv_b=Session("pcAdminOriginCompanyName")
									pcv_c=Session("pcAdminOriginDepartment")
									pcv_d=Session("pcAdminOriginPhoneNumber")
									pcv_e=Session("pcAdminOriginPagerNumber")
									pcv_f=Session("pcAdminOriginFaxNumber")
									pcv_g=Session("pcAdminOriginEmailAddress")
									pcv_h=Session("pcAdminOriginLine1")
									pcv_i=Session("pcAdminOriginLine2")
									pcv_j=Session("pcAdminOriginCity")
									pcv_k=Session("pcAdminOriginStateOrProvinceCode")
									pcv_l=Session("pcAdminOriginPostalCode")						
									pcv_m=Session("pcAdminOriginCountryCode")
									
									Session("pcAdminOriginPersonName")=Session("pcAdminRecipPersonName")
									Session("pcAdminOriginCompanyName")=Session("pcAdminRecipCompanyName")
									Session("pcAdminOriginDepartment")=Session("pcAdminRecipDepartment")
									Session("pcAdminOriginPhoneNumber")=Session("pcAdminRecipPhoneNumber")
									Session("pcAdminOriginPagerNumber")=Session("pcAdminRecipPagerNumber")
									Session("pcAdminOriginFaxNumber")=Session("pcAdminRecipFaxNumber")
									Session("pcAdminOriginEmailAddress")=Session("pcAdminRecipEmailAddress")
									Session("pcAdminOriginLine1")=Session("pcAdminRecipLine1")
									Session("pcAdminOriginLine2")=Session("pcAdminRecipLine2")
									Session("pcAdminOriginCity")=Session("pcAdminRecipCity")
									Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminRecipStateOrProvinceCode")
									Session("pcAdminOriginPostalCode")=Session("pcAdminRecipPostalCode")
									Session("pcAdminOriginCountryCode")=Session("pcAdminRecipCountryCode")						
									
									Session("pcAdminRecipPersonName")=pcv_a
									Session("pcAdminRecipCompanyName")=pcv_b
									Session("pcAdminRecipDepartment")=pcv_c
									Session("pcAdminRecipPhoneNumber")=pcv_d
									Session("pcAdminRecipPagerNumber")=pcv_e
									Session("pcAdminRecipFaxNumber")=pcv_f
									Session("pcAdminRecipEmailAddress")=pcv_g
									Session("pcAdminRecipLine1")=pcv_h
									Session("pcAdminRecipLine2")=pcv_i
									Session("pcAdminRecipCity")=pcv_j
									Session("pcAdminRecipStateOrProvinceCode")=pcv_k
									Session("pcAdminRecipPostalCode")=pcv_l
									Session("pcAdminRecipCountryCode")=pcv_m	
								end if	
								
								response.redirect ErrPageName & "?msg=" & Server.URLEncode(pcv_strSecondaryErrMsg & pcv_strSecondaryErrors & pcv_strSecondarySolution)
								
							end if
							
							
						
						End if '// end skip shipped packages 
						
						
						'// Reverse Address if Return Shipment
						if Session("pcAdminReturnShipmentIndicator")="PRINTRETURNLABEL" then
							pcv_a=Session("pcAdminOriginPersonName")
							pcv_b=Session("pcAdminOriginCompanyName")
							pcv_c=Session("pcAdminOriginDepartment")
							pcv_d=Session("pcAdminOriginPhoneNumber")
							pcv_e=Session("pcAdminOriginPagerNumber")
							pcv_f=Session("pcAdminOriginFaxNumber")
							pcv_g=Session("pcAdminOriginEmailAddress")
							pcv_h=Session("pcAdminOriginLine1")
							pcv_i=Session("pcAdminOriginLine2")
							pcv_j=Session("pcAdminOriginCity")
							pcv_k=Session("pcAdminOriginStateOrProvinceCode")
							pcv_l=Session("pcAdminOriginPostalCode")						
							pcv_m=Session("pcAdminOriginCountryCode")
							
							Session("pcAdminOriginPersonName")=Session("pcAdminRecipPersonName")
							Session("pcAdminOriginCompanyName")=Session("pcAdminRecipCompanyName")
							Session("pcAdminOriginDepartment")=Session("pcAdminRecipDepartment")
							Session("pcAdminOriginPhoneNumber")=Session("pcAdminRecipPhoneNumber")
							Session("pcAdminOriginPagerNumber")=Session("pcAdminRecipPagerNumber")
							Session("pcAdminOriginFaxNumber")=Session("pcAdminRecipFaxNumber")
							Session("pcAdminOriginEmailAddress")=Session("pcAdminRecipEmailAddress")
							Session("pcAdminOriginLine1")=Session("pcAdminRecipLine1")
							Session("pcAdminOriginLine2")=Session("pcAdminRecipLine2")
							Session("pcAdminOriginCity")=Session("pcAdminRecipCity")
							Session("pcAdminOriginStateOrProvinceCode")=Session("pcAdminRecipStateOrProvinceCode")
							Session("pcAdminOriginPostalCode")=Session("pcAdminRecipPostalCode")
							Session("pcAdminOriginCountryCode")=Session("pcAdminRecipCountryCode")						
							
							Session("pcAdminRecipPersonName")=pcv_a
							Session("pcAdminRecipCompanyName")=pcv_b
							Session("pcAdminRecipDepartment")=pcv_c
							Session("pcAdminRecipPhoneNumber")=pcv_d
							Session("pcAdminRecipPagerNumber")=pcv_e
							Session("pcAdminRecipFaxNumber")=pcv_f
							Session("pcAdminRecipEmailAddress")=pcv_g
							Session("pcAdminRecipLine1")=pcv_h
							Session("pcAdminRecipLine2")=pcv_i
							Session("pcAdminRecipCity")=pcv_j
							Session("pcAdminRecipStateOrProvinceCode")=pcv_k
							Session("pcAdminRecipPostalCode")=pcv_l
							Session("pcAdminRecipCountryCode")=pcv_m	
						end if	
												 
					Next
					'///////////////////////////////////////////////////////////////////////
					'// END LOOP
					'///////////////////////////////////////////////////////////////////////
					
				End If ' If pcv_intErr>0 Then
				
			else
			'*******************************************************************************
			' END: ON POSTBACK
			'*******************************************************************************
			%>

			<% 
			'*******************************************************************************
			' START: LOAD HTML FORM
			'*******************************************************************************

msg=request.querystring("msg")
if msg<>"" then %>
	<div class="pcCPmessage">
		<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
	</div>
<% end if %>

<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  FORM VALIDATION
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf
response.write "<!--"&vbcrlf	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf

	pcs_JavaTextField	"CarrierCode", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_1")
	pcs_JavaTextField	"ShipDate", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_2")
	pcs_JavaTextField	"ShipTime", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_3")
	pcs_JavaTextField	"OriginPersonName", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_4")
	pcs_JavaTextField	"OriginCompanyName", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_5")
	pcs_JavaTextField	"OriginPhoneNumber", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_6")
	pcs_JavaTextField	"OriginEmailAddress", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_7")
	pcs_JavaTextField	"OriginLine1", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_8")
	pcs_JavaTextField	"OriginCity", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_9")
	pcs_JavaTextField	"OriginPostalCode", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_10")
	pcs_JavaTextField	"OriginCountryCode", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_11")
	pcs_JavaTextField	"RecipPersonName", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_12")
	pcs_JavaTextField	"RecipPhoneNumber", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_13")
	pcs_JavaTextField	"RecipCountryCode", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_14")
	pcs_JavaTextField	"RecipLine1", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_15")
	pcs_JavaTextField	"RecipCity", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_16")
	pcs_JavaTextField	"CustomerReference", true, dictLanguageCP.Item(Session("language")&"_cpFedEx_17")

response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "//-->"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  FORM VALIDATION
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<form name="form1" method="post" action="<%=pcPageName%>" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">	
	<tr>
		<% 
		dim strJSOnChangeTabCnt, k, pcPackageCount, intTempJSChangeCnt					
		strTabCnt=""
		for k=1 to pcPackageCount
			if k=1 then
				strTabCnt="""tab5"""
			else
				iCnt=4+int(k)
				strTabCnt=strTabCnt&",""tab"&iCnt&""""
			end if
		next
	
	
		strJSOnChangeTabCnt=""
		for k=1 to pcPackageCount
			intTempJSChangeCnt=4+int(k)
			strJSOnChangeTabCnt=strJSOnChangeTabCnt&";change('tabs"&intTempJSChangeCnt&"', '')"
		next %>
		<!--#include file="../includes/javascripts/pcFedExLabelTabs.asp"-->
		<td valign="top">
		<div class="menu">
			<ul>
				<li><a id="tabs1" class="current" onclick="change('tabs1', 'current');change('tabs2', '');change('tabs3', '');change('tabs4', '')<%=strJSOnChangeTabCnt%>;showTab('tab1')">Ship Settings</a></li>
				<li><a id="tabs2" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '');change('tabs4', '')<%=strJSOnChangeTabCnt%>;showTab('tab2')">Ship From</a></li>
				<li><a id="tabs3" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', 'current');change('tabs4', '')<%=strJSOnChangeTabCnt%>;showTab('tab3')">Recipient</a></li>
				<li><a id="tabs4" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current')<%=strJSOnChangeTabCnt%>;showTab('tab4')">Ship Notification</a></li>
				<% strOnclickTabCnt=""
				if pcPackageCount=1 then %>
				<li><a id="tabs5" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '');change('tabs5', 'current');showTab('tab5')">Package Information</a></li>
				<% else %>
					<% for k=1 to pcPackageCount
						intTempPackageCnt=4+int(k)
						strOnclickTabCnt=""
						for l=1 to pcPackageCount
							intCPC=4+int(l)
							if intCPC=intTempPackageCnt then
								strOnclickTabCnt=strOnclickTabCnt&";change('tabs"&intCPC&"', 'current')"
							else
								strOnclickTabCnt=strOnclickTabCnt&";change('tabs"&intCPC&"', '')"
							end if
						next
						%>
						<li><a id="tabs<%=4+int(k)%>" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '')<%=strOnclickTabCnt%>;showTab('tab<%=intTempPackageCnt%>')">Package <%=k%></a></li>
						<%
					next
				end if %>
			</ul>
		</div>
		<!--							
		//////////////////////////////////////////////////////////////////////////////////////////////
		// SHIP SETTINGS
		//////////////////////////////////////////////////////////////////////////////////////////////
		-->
		<div id="tab1" class="panes" style="display:block">		
		<%
		For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
		%>
		<input type="hidden" name="<%="pcAdminPrdList"&(xArrayCount+1)%>" value="<%=pcLocalArray(xArrayCount)%>">
		<% Next %>
		<input type="hidden" name="idorder" value="<%=pcv_intOrderID%>">
		<input type="hidden" name="PackageCount" value="<%=pcPackageCount%>">
		<input type="hidden" name="ItemsList" value="<%=pcv_strItemsList%>">
		
		<input name="CurrencyCode" type="hidden" id="CurrencyCode" value="USD" size="3" maxlength="3">
		<table class="pcCPcontent">
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td class="pcCPshipping"><span class="titleShip">Order Details</span></td>
				<td class="pcCPshipping" align="right">
				<i>(Check box to view)</i>&nbsp;
				<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
				<!--
				function jfOrder(){
				
				var selectValDom = document.forms['form1'];
				if (selectValDom.bOrder.checked == true) {
				document.getElementById('Order').style.display='';
				}else{
				document.getElementById('Order').style.display='none';
				}
				}
				 //-->
				</SCRIPT>
				<%
				if Session("pcAdminbOrder")="true" then
					pcv_strDisplayStyle="style=""display:visible"""
				else
					pcv_strDisplayStyle="style=""display:none"""
				end if
				%>
	<input onClick="jfOrder();" name="bOrder" id="bOrder" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("bOrder", "true")%>>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<div id="Order" <%=pcv_strDisplayStyle%>>
					<table width="100%">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Shipping Information</th>
						</tr>
						<tr> 
							<td colspan="2">
									<% 
									if pcOrd_ShipWeight>0 then
										intTotalWeight=pcOrd_ShipWeight
									end if
									intTotalWeight=round(intTotalWeight,0)
									if scShipFromWeightUnit="KGS" then
									pKilos=Int(intTotalWeight/1000)
									pWeight_g=intTotalWeight-(pKilos*1000) %>
									<div align="left">Total Shipping Weight:&nbsp;&nbsp;<strong><%=pKilos&" kg "%></strong> 
									<% if pWeight_g>0 then 
										response.write "<strong>" & pWeight_g&" g</strong>"
									end if %>
									</div>
								<% else 
									pPounds=Int(intTotalWeight/16)
									pWeight_oz=intTotalWeight-(pPounds*16) %>
									<div align="left">Total Shipping Weight:&nbsp;&nbsp;<strong><%=pPounds&" lbs. "%></strong> 
									<% if pWeight_oz>0 then 
										response.write "<strong>" & pWeight_oz&" oz.</strong>"
									end if %>
									</div>
								<% end if %>
							</td>
						</tr>
						<tr> 
							<td colspan="2">Number of Packages: <strong><%=pOrdPackageNum%></strong></td>
						</tr>
						<tr> 
							<td colspan="2">Shipping Method: 
							<strong>				
							<% 
							if pSRF="1" then
								response.write pshipmentDetails
							else
								if varShip<>"0"  then
									response.write Service
								else
									response.write pshipmentDetails
								end if 
							end if 
							%></strong>
							</td>
						</tr>
						<%
						if varShip<>"0" then
							if pOrdShipType=0 then
								pDisShipType="Residential"
							else
								pDisShipType="Commercial" 
							end if
						%>
						<tr> 
							<td colspan="2">Shipping Type:&nbsp;<strong><%=pDisShipType%></strong></td>
						</tr>
						<% end if %>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
					</table>
					</div>
				</td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td class="pcCPshipping"><span class="titleShip">Ship Settings</span></td>
				<td class="pcCPshipping" align="right">
				<i>(Check box to view)</i>&nbsp;
				<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
				<!--
				function jfShip(){
				
				var selectValDom = document.forms['form1'];
				if (selectValDom.bShip.checked == true) {
				document.getElementById('Ship').style.display='';
				}else{
				document.getElementById('Ship').style.display='none';
				}
				}
				 //-->
				</SCRIPT>
				<%
				if Session("pcAdminbShip")="true" then
					pcv_strDisplayStyle="style=""display:visible"""
				else
					pcv_strDisplayStyle="style=""display:none"""
				end if
				%>
	<input onClick="jfShip();" name="bShip" id="bShip" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("bShip", "true")%>>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<div id="Ship" <%=pcv_strDisplayStyle%>>
					<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
					<!--
					document.getElementById('bShip').checked=true
					jfShip();
					 //-->
					</SCRIPT>
					<table width="100%">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Service Settings  </th>
						</tr>
						<tr>
							<td width="25%" align="right" valign="top"><b>Type of service:</b></td>
							<td width="75%" align="left">
							<%
							'// Set Carrier Code to local
							Select Case Service						
								Case "FedEx First Overnight": ShippingSelector="FDXE"
								Case "FedEx Priority Overnight": ShippingSelector="FDXE"
								Case "FedEx Standard Overnight": ShippingSelector="FDXE"
								Case "FedEx 2Day": ShippingSelector="FDXE"
								Case "FedEx Express Saver": ShippingSelector="FDXE"
								Case "FedEx Ground": ShippingSelector="FDXG"
								Case "FedEx Home Delivery": ShippingSelector="FDXG"
								Case "FedEx International First": ShippingSelector="FDXE"
								Case "FedEx International Priority": ShippingSelector="FDXE"
								Case "FedEx International Economy": ShippingSelector="FDXE"
								Case "FedEx 1Day Freight": ShippingSelector="FDXE"
								Case "FedEx 2Day Freight": ShippingSelector="FDXE"
								Case "FedEx 3Day Freight": ShippingSelector="FDXE"
								Case "FedEx International Priority Freight": ShippingSelector="FDXE"
								Case "FedEx International Economy Freight": ShippingSelector="FDXE"
							End Select
							
							if Session("pcAdminCarrierCode")="" then
								Session("pcAdminCarrierCode")=ShippingSelector
							end if														
							%>
							<select name="CarrierCode" id="CarrierCode" size="1">
								<option value="FDXE" <%=pcf_SelectOption("CarrierCode","FDXE")%>>FedEx Express</option>	
								<option value="FDXG" <%=pcf_SelectOption("CarrierCode","FDXG")%>>FedEx Ground</option>														
							</select>
							<%pcs_RequiredImageTag "CarrierCode", true %>		
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Drop off Type:</b></td>
							<td align="left">
								<select name="DropoffType" id="DropoffType">
									<option value="REQUESTCOURIER" <%=pcf_SelectOption("DropoffType","REQUESTCOURIER")%>>Courier Pickup</option>
									<option value="REGULARPICKUP" <%=pcf_SelectOption("DropoffType","REGULARPICKUP")%>>Regular Pickup</option>
									<option value="DROPBOX" <%=pcf_SelectOption("DropoffType","DROPBOX")%>>FedEx Express Drop Box</option>
									<option value="BUSINESSSERVICECENTER" <%=pcf_SelectOption("DropoffType","BUSINESSSERVICECENTER")%>>Business Service Center</option>
									<option value="STATION" <%=pcf_SelectOption("DropoffType","STATION")%>>FedEx Station</option>								
								</select><%pcs_RequiredImageTag "DropoffType", isRequiredDropoffType %>										
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Return Shipment Indicator:</b></td>
							<td align="left">
								<select name="ReturnShipmentIndicator" id="ReturnShipmentIndicator">
								<option value="NONRETURN" <%=pcf_SelectOption("ReturnShipmentIndicator","NONRETURN")%>>Outgoing Shipment</option>
								<option value="PRINTRETURNLABEL" <%=pcf_SelectOption("ReturnShipmentIndicator","PRINTRETURNLABEL")%>>Return Shipment</option>
								</select>
								<%pcs_RequiredImageTag "ReturnShipmentIndicator", false%>		
								</td>
						</tr>
						<tr>
							<th colspan="2">Date/ Time  </th>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Ship Date:</b></td>
							<td align="left">
								<select name="ShipDate">						
								<% dtTodayDate=Date()
								Function FedExDateFormat (FedExDate)
								FedExDay=Day(FedExDate)
								FedExMonth=Month(FedExDate)
								FedExYear= Year(FedExDate)
								FedExDateFormat=FedExYear&"-"&Right(Cstr(FedExMonth + 100),2)&"-"&Right(Cstr(FedExDay + 100),2)
								End Function %>
								<option value="<%=FedExDateFormat(dtTodayDate)%>" <%=pcf_SelectOption("ShipDate",FedExDateFormat(dtTodayDate))%>>Today</option>
								<% for d=1 to 10
								if DatePart("W", dtTodayDate+d, VBSUNDAY)=1 then
								else %>
								<option value="<%=FedExDateFormat((dtTodayDate+d))%>" <%=pcf_SelectOption("ShipDate",FedExDateFormat(dtTodayDate+d))%>><%=FormatDateTime((dtTodayDate+d), 1)%></option>
								<% end if
								next %>
								</select>						
								<%pcs_RequiredImageTag "ShipDate", true%>		
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Ship Time:</b></td>
							<td align="left">
								<input name="ShipTime" type="text" id="ShipTime" value="<%=pcf_FillFormField("ShipTime", true)%>">
								<%pcs_RequiredImageTag "ShipTime", true%> * hh:mm:ss (e.g. 10:40:00)		
							</td>
						</tr>
						<tr>
							<th colspan="2">Shipment Billing</th>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Payor:</b></td>
							<td align="left">
								<select name="PayorType" id="PayorType">
								<option value="SENDER" <%=pcf_SelectOption("PayorType","SENDER")%>>Sender</option>
								<option value="RECIPIENT" <%=pcf_SelectOption("PayorType","RECIPIENT")%>>Recipient</option>
								<option value="THIRDPARTY" <%=pcf_SelectOption("PayorType","THIRDPARTY")%>>3rd Party</option>
								<option value="COLLECT" <%=pcf_SelectOption("PayorType","COLLECT")%>>Collect</option>
								</select>
								<%pcs_RequiredImageTag "PayorType", true%>
                                (This is optional. Only required for RECIPIENT, THIRDPARTY, and COLLECT)
							</td>
						</tr>
						<tr>
						<td align="right" valign="top"><b>Payor Account Number:</b></td>
						<td align="left">
						<input name="PayorAccountNumber" type="text" id="PayorAccountNumber" value="<%=pcf_FillFormField("PayorAccountNumber", false)%>">
						<%pcs_RequiredImageTag "PayorAccountNumber", false%>  
                        	</td>
						</tr>
						<tr>
						<td align="right" valign="top"><b>Payor Country Code:</b></td>
						<td align="left">
						<input name="PayorCountryCode" type="text" id="PayorCountryCode" value="<%=pcf_FillFormField("PayorCountryCode", false)%>">
						<%pcs_RequiredImageTag "PayorCountryCode", false%> 
						e.g.	US	</td>
						</tr>
					</table>
					</div>
				</td>
			</tr>	
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>				
				<td width="50%" class="pcCPshipping"><span class="titleShip">Additional Settings</span></td>
				<td class="pcCPshipping" align="right">
				<i>(Check box to view)</i>&nbsp;
				<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
				<!--
				function jfAdditional(){
				
				var selectValDom = document.forms['form1'];
				if (selectValDom.bAdditional.checked == true) {
				document.getElementById('Additional').style.display='';
				}else{
				document.getElementById('Additional').style.display='none';
				}
				}
				 //-->
				</SCRIPT>
				<%
				if Session("pcAdminbAdditional")="true" then
					pcv_strDisplayStyle="style=""display:visible"""
				else
					pcv_strDisplayStyle="style=""display:none"""
				end if
				%>
	<input onClick="jfAdditional();" name="bAdditional" id="bAdditional" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("bAdditional", "true")%>>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<div id="Additional" <%=pcv_strDisplayStyle%>>
					<table width="100%">		
						<tr>
						<th colspan="2">Labels</th>
						</tr>	
						<tr>
						<td align="right" valign="top"><b>Type:</b></td>
						<td align="left">
						<input name="Type" type="text" id="Type" value="<%=pcf_FillFormField("Type", false)%>">
						<%pcs_RequiredImageTag "Type", false%>		</td>
						</tr>
						<tr>
						<td align="right" valign="top"><b>Image Type:</b></td>
						<td align="left">
						<select name="ImageType" id="ImageType">
						<option value="PNG" <%=pcf_SelectOption("ImageType","PNG")%>>PNG (Plain Paper)</option>
						<!--
						<option value="PNG4X6" <%'=pcf_SelectOption("ImageType","PNG4X6")%>>PNG4X6 (PNG image)</option>
						<option value="ELTRON" <%'=pcf_SelectOption("ImageType","ELTRON")%>>ELTRON (Thermal)</option>
						<option value="ZEBRA" <%'=pcf_SelectOption("ImageType","ZEBRA")%>>ZEBRA (Thermal)</option>
						<option value="UNIMARK" <%'=pcf_SelectOption("ImageType","UNIMARK")%>>UNIMARK (Thermal)</option>
						 -->
						</select>
						<%pcs_RequiredImageTag "ImageType", true%>		
						</td>
						</tr>
						<tr>
						<th colspan="2">Special Services</th>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Saturday Service:</b></td>
							<td align="left">
								<table cellspacing="0" cellpadding="0">			

									<TR>
										<TD width="79%" height="20">Saturday Delivery:</TD>
										<TD width="21%" height="20">
										<INPUT tabIndex="25" type="checkbox" value="1" name="SaturdayDelivery" class="clearBorder" <%=pcf_CheckOption("SaturdayDelivery", "1")%>>
										</TD>
									</TR>
									<TR>
										<TD width="79%" height="20">Saturday Pickup:</TD>
										<TD width="21%" height="20">
										<INPUT tabIndex="25" type="checkbox" value="1" name="SaturdayPickup" class="clearBorder" <%=pcf_CheckOption("SaturdayPickup", "1")%>>
										</TD>
									</TR>
								</table>
							</td>
						</tr>
						<tr>
						<td align="right" valign="top"><b>Signature:</b></td>
							<td align="left">
								<select name="SignatureOption" id="SignatureOption">
									<option value="" <%=pcf_SelectOption("SignatureOption","")%>>No Signature Options</option>
									<option value="DELIVERWITHOUTSIGNATURE" <%=pcf_SelectOption("SignatureOption","DELIVERWITHOUTSIGNATURE")%>>Deliver Without Signature</option>
									<option value="INDIRECT" <%=pcf_SelectOption("SignatureOption","INDIRECT")%>>Indirect</option>
									<option value="DIRECT" <%=pcf_SelectOption("SignatureOption","DIRECT")%>>Direct</option>
									<option value="ADULT" <%=pcf_SelectOption("SignatureOption","ADULT")%>>Adult Signature Required</option>
								</select>
								<%pcs_RequiredImageTag "ImageType", false%>		
							</td>
						</tr>
						<tr>
							<td align="right"><b>SignatureRelease:</b></td>
							<td align="left">
								<INPUT type="text" name="SignatureRelease" id="SignatureRelease" value="<%=pcf_FillFormField("SignatureRelease", false)%>">
								<%pcs_RequiredImageTag "SignatureRelease", false%>
								(Deliver Without Signature Only)
							</td>
						</tr>	
						<tr>
						<th colspan="2">Hold At Location</th>
						</tr>
						<tr>
							<td colspan="2">
							<input type="checkbox" name="HoldAtLocation" value="1" class="clearBorder" <%=pcf_CheckOption("HoldAtLocation", "1")%>>
							<strong>Check this box to hold at location.</strong>
							</td>
						</tr>
						<tr>
							<td colspan="2">
							<span class="pcCPnotes">*If you check to hold at location you must specify an address below.</span>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Phone:</b></td>
							<td align="left">
								<INPUT type="text" name="HALPhone" id="HALPhone" value="<%=pcf_FillFormField("HALPhone", false)%>">
								<%pcs_RequiredImageTag "HALPhone", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Address:</b></td>
							<td align="left">
								<INPUT type="text" name="HALLine1" id="HALLine1" value="<%=pcf_FillFormField("HALLine1", false)%>">
								<%pcs_RequiredImageTag "HALLine1", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>City:</b></td>
							<td align="left">
								<INPUT type="text" name="HALCity" id="HALCity" value="<%=pcf_FillFormField("HALCity", false)%>">
								<%pcs_RequiredImageTag "HALCity", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>State or Province Code:</b></td>
							<td align="left">
								<INPUT type="text" name="HALStateOrProvinceCode" id="HALStateOrProvinceCode" value="<%=pcf_FillFormField("HALStateOrProvinceCode", false)%>">
								<%pcs_RequiredImageTag "HALStateOrProvinceCode", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Postal Code:</b></td>
							<td align="left">
								<INPUT name="HALPostalCode" type="text" id="HALPostalCode" value="<%=pcf_FillFormField("HALPostalCode", false)%>">
								<%pcs_RequiredImageTag "HALPostalCode", false%>
							</td>

						</tr>						
					</table>
					</div>
				</td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>				
				<td width="50%" class="pcCPshipping"><span class="titleShip">International Settings</span></td>
				<td class="pcCPshipping" align="right">
				<i>(Check box to view)</i>&nbsp;
				<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
				<!--
				function jfInternational(){
				
				var selectValDom = document.forms['form1'];
				if (selectValDom.bInternational.checked == true) {
				document.getElementById('International').style.display='';
				}else{
				document.getElementById('International').style.display='none';
				}
				}
				 //-->
				</SCRIPT>
				<%
				if Session("pcAdminbInternational")="true" then
					pcv_strDisplayStyle="style=""display:visible"""
				else
					pcv_strDisplayStyle="style=""display:none"""
				end if
				%>
	<input onClick="jfInternational();" name="bInternational" id="bInternational" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("bInternational", "true")%>>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<div id="International" <%=pcv_strDisplayStyle%>>
					<table width="100%">
						<tr>
							<th colspan="2">International Settings</th>
						</tr>
						<tr>
							<td width="25%" align="right" valign="top"><b>Terms Of Sale:</b></td>
							<td align="left">
								<select name="TermsOfSale" id="TermsOfSale">
									<option value="FOB_OR_FCA" <%=pcf_SelectOption("TermsOfSale","FOB_OR_FCA")%>>FOB_OR_FCA</option>
									<option value="CIF_OR_CIP" <%=pcf_SelectOption("TermsOfSale","CIF_OR_CIP")%>>CIF_OR_CIP</option>
									<option value="CFR_OR_CPT" <%=pcf_SelectOption("TermsOfSale","CFR_OR_CPT")%>>FOB_OR_FCA</option>
									<option value="EXW" <%=pcf_SelectOption("TermsOfSale","EXW")%>>EXW</option>
									<option value="DDU" <%=pcf_SelectOption("TermsOfSale","DDU")%>>DDU</option>
									<option value="DDP" <%=pcf_SelectOption("TermsOfSale","DDP")%>>DDP</option>
								</select>
								<%pcs_RequiredImageTag "TermsOfSale", false%>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Total Customs Value:</b></td>
							<td align="left">
								<input name="TotalCustomsValue" type="text" id="TotalCustomsValue" value="<%=pcf_FillFormField("TotalCustomsValue", false)%>">
								<%pcs_RequiredImageTag "TotalCustomsValue", false%>		
							(e.g. 500.00) </td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Recipient TIN/ EIN:</b></td>
							<td align="left">
								<input name="RecipientTIN" type="text" id="RecipientTIN" value="<%=pcf_FillFormField("RecipientTIN", false)%>">
								<%pcs_RequiredImageTag "RecipientTIN", false%>	(This is required for International Shipments)		
							</td>
						</tr>							
						<tr>
							<td align="right" valign="top"><b>Admissibility Package Type:</b></td>
							<td align="left">
								<input name="AdmissibilityPackageType" type="text" id="AdmissibilityPackageType" value="<%=pcf_FillFormField("AdmissibilityPackageType", false)%>">
								<%pcs_RequiredImageTag "AdmissibilityPackageType", false%>		
							</td>
						</tr>				
						<tr>
							<th colspan="2">International Duties and Taxes</th>
						</tr>
						<tr>
							<td width="25%" align="right" valign="top"><b>Duties Payor:</b></td>
							<td align="left">
								<select name="DutiesPayorType" id="DutiesPayorType">
								<option value="SENDER" <%=pcf_SelectOption("DutiesPayorType","SENDER")%>>Sender</option>
								<option value="RECIPIENT" <%=pcf_SelectOption("DutiesPayorType","RECIPIENT")%>>Recipient</option>
								<option value="THIRDPARTY" <%=pcf_SelectOption("DutiesPayorType","THIRDPARTY")%>>3rd Party</option>
								</select>
								<%pcs_RequiredImageTag "DutiesPayorType", false%>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Duties Payor Account#:</b></td>
							<td align="left">
								<input name="DutiesAccountNumber" type="text" id="DutiesAccountNumber" value="<%=pcf_FillFormField("DutiesAccountNumber", false)%>">
								<%pcs_RequiredImageTag "DutiesAccountNumber", false%>		
							</td>
						</tr>				
						<tr>
							<td align="right" valign="top"><b>Duties Payor Country Code:</b></td>
							<td align="left">
							<input name="DutiesCountryCode" type="text" id="DutiesCountryCode" value="<%=pcf_FillFormField("DutiesCountryCode", false)%>">
							<%pcs_RequiredImageTag "DutiesCountryCode", false%>		
							(e.g. US) </td>
						</tr>
						<tr>
							<th colspan="2">SED</th>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Sender TIN or DUNS:</b></td>
							<td align="left">
								<input name="SenderTINOrDUNS" type="text" id="SenderTINOrDUNS" value="<%=pcf_FillFormField("SenderTINOrDUNS", false)%>">
								<%pcs_RequiredImageTag "SenderTINOrDUNS", false%>		
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Sender TIN or DUNS Type:</b></td>
							<td align="left">
								<select name="SenderTINOrDUNSType" id="SenderTINOrDUNSType">
								<option value="SSN" <%=pcf_SelectOption("SenderTINOrDUNSType","SSN")%>>SSN</option>
								<option value="EIN" <%=pcf_SelectOption("SenderTINOrDUNSType","EIN")%>>EIN</option>
								<option value="DUNS" <%=pcf_SelectOption("SenderTINOrDUNSType","DUNS")%>>DUNS</option>
								</select>					
								<%pcs_RequiredImageTag "SenderTINOrDUNSType", false%>		
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>AES or FTSR Exemption #:</b></td>
							<td align="left">
								<input name="AESOrFTSRExemptionNumber" type="text" id="AESOrFTSRExemptionNumber" value="<%=pcf_FillFormField("AESOrFTSRExemptionNumber", false)%>">
								<%pcs_RequiredImageTag "AESOrFTSRExemptionNumber", false%>		
							</td>
						</tr>	
						<tr>
							<th colspan="2">Commodity</th>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Number of Pieces:</b></td>
							<td align="left">
								<input name="NumberOfPieces" type="text" id="NumberOfPieces" value="<%=pcf_FillFormField("NumberOfPieces", false)%>">
								<%pcs_RequiredImageTag "NumberOfPieces", false%>		
							</td>
						</tr>				
						<tr>
							<td align="right" valign="top"><b>Description:</b></td>
							<td align="left">
								<input name="Description" type="text" id="Description" value="<%=pcf_FillFormField("Description", false)%>">
								<%pcs_RequiredImageTag "Description", false%>		
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Country Code of Manufacture:</b></td>
							<td align="left">
								<input name="CountryOfManufacture" type="text" id="CountryOfManufacture" value="<%=pcf_FillFormField("CountryOfManufacture", false)%>">
								<%pcs_RequiredImageTag "CountryOfManufacture", false%>		
							(e.g. US) </td>
						</tr>																		
						<tr>
							<td align="right" valign="top"><b>Harmonized Code:</b></td>
							<td align="left">
								<input name="HarmonizedCode" type="text" id="HarmonizedCode" value="<%=pcf_FillFormField("HarmonizedCode", false)%>">
								<%pcs_RequiredImageTag "HarmonizedCode", false%>		
							</td>
						</tr>				
						<tr>
							<td align="right" valign="top"><b>Weight:</b></td>
							<td align="left">
								<input name="CommodityWeight" type="text" id="CommodityWeight" value="<%=pcf_FillFormField("CommodityWeight", false)%>">
								<%pcs_RequiredImageTag "CommodityWeight", false%>		
							(e.g. 3.0) </td>
						</tr>				
						<tr>
							<td align="right" valign="top"><b>Quantity:</b></td>
							<td align="left">
								<input name="CommodityQuantity" type="text" id="CommodityQuantity" value="<%=pcf_FillFormField("CommodityQuantity", false)%>">
								<%pcs_RequiredImageTag "CommodityQuantity", false%>		
							</td>
						</tr>	
						<tr>
							<td align="right" valign="top"><b>Quantity Units:</b></td>
							<td align="left">
								<input name="CommodityQuantityUnits" type="text" id="CommodityQuantityUnits" value="<%=pcf_FillFormField("CommodityQuantityUnits", false)%>">
								<%pcs_RequiredImageTag "CommodityQuantityUnits", false%>		
							(e.g. EA) </td>
						</tr>				
						<tr>
							<td align="right" valign="top"><b>Unit Price:</b></td>
							<td align="left">
								<input name="CommodityUnitPrice" type="text" id="CommodityUnitPrice" value="<%=pcf_FillFormField("CommodityUnitPrice", false)%>">
								<%pcs_RequiredImageTag "CommodityUnitPrice", false%>		
							* Six explicit decimal positions (e.g. 900.000000)</td>
						</tr>				
						<tr>
							<td align="right" valign="top"><b>Total Commodity Value:</b></td>
							<td align="left">
								<input name="CommodityCustomsValue" type="text" id="CommodityCustomsValue" value="<%=pcf_FillFormField("CommodityCustomsValue", false)%>">
								<%pcs_RequiredImageTag "CommodityCustomsValue", false%>		
							(e.g. 500.00)</td>
						</tr>				
						<tr>
							<td align="right" valign="top"><b>Export License #:</b></td>
							<td align="left">
								<input name="ExportLicenseNumber" type="text" id="ExportLicenseNumber" value="<%=pcf_FillFormField("ExportLicenseNumber", false)%>">
								<%pcs_RequiredImageTag "ExportLicenseNumber", false%>		
							</td>
						</tr>				
						<tr>
							<td align="right" valign="top"><b>License Expiration Date:</b></td>
							<td align="left">
								<input name="ExportLicenseExpirationDate" type="text" id="ExportLicenseExpirationDate" value="<%=pcf_FillFormField("ExportLicenseExpirationDate", false)%>">
								<%pcs_RequiredImageTag "ExportLicenseExpirationDate", false%>		
							</td>
						</tr>					
						<tr>
							<td align="right" valign="top"><b>CI Marks and Numbers:</b></td>
							<td align="left">
								<input name="CIMarksAndNumbers" type="text" id="CIMarksAndNumbers" value="<%=pcf_FillFormField("CIMarksAndNumbers", false)%>">
								<%pcs_RequiredImageTag "CIMarksAndNumbers", false%>		
							</td>
						</tr>	
					</table>
					</div>
				</td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td width="50%" class="pcCPshipping"><span class="titleShip">Ground Settings</span></td>
				<td class="pcCPshipping" align="right">				
				<i>(Check box to view)</i>&nbsp;
				<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
				<!--
				function jfGround(){
				
				var selectValDom = document.forms['form1'];
				if (selectValDom.bGround.checked == true) {
				document.getElementById('Ground').style.display='';
				}else{
				document.getElementById('Ground').style.display='none';
				}
				}
				 //-->
				</SCRIPT>
				<%
				if Session("pcAdminbGround")="true" then
					pcv_strDisplayStyle="style=""display:visible"""
				else
					pcv_strDisplayStyle="style=""display:none"""
				end if
				%>
				<input onClick="jfGround();" name="bGround" id="bGround" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("bGround", "true")%>>
				</td>				
			</tr>
			<tr>
				<td colspan="2">
					<div id="Ground" <%=pcv_strDisplayStyle%>>
					<table width="100%">
						<tr>
							<th colspan="2">Home Delivery Options</th>
						</tr>
						<tr>
							<td width="25%" align="right" valign="top"><b>Delivery Type:</b></td>
							<td align="left">
								<select name="DeliveryType" id="DeliveryType">
								<option value="">Please make a selection. (optional)</option>
								<option value="DATECERTAIN" <%=pcf_SelectOption("DeliveryType","DATECERTAIN")%>>Date Certain</option>
								<option value="EVENING" <%=pcf_SelectOption("DeliveryType","EVENING")%>>Evening</option>
								<option value="APPOINTMENT" <%=pcf_SelectOption("DeliveryType","APPOINTMENT")%>>Appointment</option>
								</select> (* Required for Ground Home Delivery shipments.)
								<%pcs_RequiredImageTag "DeliveryType", false%>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Delivery Date:</b></td>
							<td align="left">
								<input name="DeliveryDate" type="text" id="DeliveryDate" value="<%=pcf_FillFormField("DeliveryDate", false)%>">
								<%pcs_RequiredImageTag "DeliveryDate", false%>		
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Delivery Phone:</b></td>
							<td align="left">
								<input name="DeliveryPhone" type="text" id="DeliveryPhone" value="<%=pcf_FillFormField("DeliveryPhone", false)%>">
								<%pcs_RequiredImageTag "DeliveryPhone", false%>		
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Delivery Instructions:</b></td>
							<td align="left">
								<input name="DeliveryInstructions" type="text" id="DeliveryInstructions" value="<%=pcf_FillFormField("DeliveryInstructions", false)%>">
								<%pcs_RequiredImageTag "DeliveryInstructions", false%>		
							</td>
						</tr>
					</table>
					</div>
				</td>
			</tr>	
		</table>		
		</div>
					
					
		<!--							
		//////////////////////////////////////////////////////////////////////////////////////////////
		// SHIPPER
		//////////////////////////////////////////////////////////////////////////////////////////////
		-->
		<div id="tab2" class="panes">
		<table class="pcCPcontent">
			<tr>
			<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
			<th colspan="2">Contact Details</th>
			</tr>
			<tr>
			<td width="25%" align="right"><p>Contact Name:</p></td>
			<td width="75%" align="left"><p>
			<input name="OriginPersonName" type="text" id="OriginPersonName" value="<%=pcf_FillFormField("OriginPersonName", true)%>">
			<%pcs_RequiredImageTag "OriginPersonName", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Company Name:</p></td>
			<td align="left"><p>
			<input name="OriginCompanyName" type="text" id="OriginCompanyName" value="<%=pcf_FillFormField("OriginCompanyName", false)%>">
			<%pcs_RequiredImageTag "OriginCompanyName", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Department:</p></td>
			<td align="left"><p>
			<input name="OriginDepartment" type="text" id="OriginDepartment" value="<%=pcf_FillFormField("OriginDepartment", false)%>">
						<%pcs_RequiredImageTag "OriginDepartment", false%></p>
			</td>
			</tr>
			<% if len(Session("ErrOriginPhoneNumber"))>0 then %>
			<tr>
			<td colspan="2">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> 
			You must enter a valid Phone Number.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>Phone Number:</p></td>
			<td align="left"><p>
			<input name="OriginPhoneNumber" type="text" id="OriginPhoneNumber" value="<%=pcf_FillFormField("OriginPhoneNumber", true)%>">
			<%pcs_RequiredImageTag "OriginPhoneNumber", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Pager Number:</p></td>
			<td align="left"><p>
			<input name="OriginPagerNumber" type="text" id="OriginPagerNumber" value="<%=pcf_FillFormField("OriginPagerNumber", false)%>">
			<%pcs_RequiredImageTag "OriginPagerNumber", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Fax Number:</p></td>
			<td align="left"><p>
			<input name="OriginFaxNumber" type="text" id="OriginFaxNumber" value="<%=pcf_FillFormField("OriginFaxNumber", false)%>">
			<%pcs_RequiredImageTag "OriginFaxNumber", false%></p>
			</td>
			</tr>
			<% if len(Session("ErrOriginEmailAddress"))>0 then %>
			<tr>
			<td colspan="2">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> 
			You must enter a valid Email Address.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>Email Address:</p></td>
			<td align="left"><p>
			<input name="OriginEmailAddress" type="text" id="OriginEmailAddress" value="<%=pcf_FillFormField("OriginEmailAddress", true)%>">
						<%pcs_RequiredImageTag "OriginEmailAddress", true%></p>
			</td>
			</tr>
			<tr>
			<th colspan="2">Location Details</th>
			</tr>
			
			<%
			dim conntemp, query, rs
			'///////////////////////////////////////////////////////////
			'// START: COUNTRY AND STATE/ PROVINCE CONFIG
			'///////////////////////////////////////////////////////////
			' 
			pcv_isStateOrProvinceCodeRequired = isRequiredState '// determines if validation is performed (true or false)
			pcv_isProvinceCodeRequired = isRequiredProvince '// determines if validation is performed (true or false)
			pcv_isCountryCodeRequired = true '// determines if validation is performed (true or false)
			
			'// #3 Additional Required Info
			pcv_strTargetForm = "form1" '// Name of Form
			pcv_strCountryBox = "OriginCountryCode" '// Name of Country Dropdown
			pcv_strTargetBox = "OriginStateOrProvinceCode" '// Name of State Dropdown
			pcv_strProvinceBox =  "OriginProvinceCode" '// Name of Province Field
			
			'// Set local Country to Session
			if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strCountryBox))=True then
				Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
			end if
			
			'// Set local State to Session
			if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strTargetBox))=True then
				Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session("pcAdminOriginStateOrProvinceCode")
			end if
			
			'// Set local Province to Session
			if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strProvinceBox))=True then
				Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  Session("pcAdminOriginStateOrProvinceCode")
			end if
			%>					
			<!--#include file="../includes/javascripts/pcStateAndProvince.asp"-->
			<%
			'///////////////////////////////////////////////////////////
			'// END: COUNTRY AND STATE/ PROVINCE CONFIG
			'///////////////////////////////////////////////////////////
			%>
			
			<%
			'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
			pcs_CountryDropdown
			%>	
			
			<tr>
			<td align="right"><p>Address Line 1:</p></td>
			<td align="left"><p>
			<input name="OriginLine1" type="text" id="OriginLine1" value="<%=pcf_FillFormField("OriginLine1", true)%>">
			<%pcs_RequiredImageTag "OriginLine1", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Address Line 2:</p></td>
			<td align="left"><p>
			<input name="OriginLine2" type="text" id="OriginLine2" value="<%=pcf_FillFormField("OriginLine2", false)%>">
			<%pcs_RequiredImageTag "OriginLine2", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>City:</p></td>
			<td align="left"><p>
			<input name="OriginCity" type="text" id="OriginCity" value="<%=pcf_FillFormField("OriginCity", true)%>">
						<%pcs_RequiredImageTag "OriginCity", true%></p>
			</td>
			</tr>
			
			<%
			'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
			pcs_StateProvince
			%>			

			<tr>
			<td align="right"><p>Postal Code:</p></td>
			<td align="left"><p>
			<input name="OriginPostalCode" type="text" id="OriginPostalCode" value="<%=pcf_FillFormField("OriginPostalCode", true)%>">
			<%pcs_RequiredImageTag "OriginPostalCode", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"></td>
			<td align="left">
			</td>
			</tr>
		</table>
					
		</div>
		
		<!--							
		//////////////////////////////////////////////////////////////////////////////////////////////
		// RECIPIENT
		//////////////////////////////////////////////////////////////////////////////////////////////
		-->
		<div id="tab3" class="panes">
		<table class="pcCPcontent">
			<tr>
			<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
			<th colspan="2">Contact Details</th>
			</tr>
			<tr>
			<td width="25%" align="right"><p>Contact Name:</p></td>
			<td width="75%" align="left"><p>
			<input name="RecipPersonName" type="text" id="RecipPersonName" value="<%=pcf_FillFormField("RecipPersonName", true)%>">
			<%pcs_RequiredImageTag "RecipPersonName", true%></p>
			</td>
			</tr>
			<tr>

			<td align="right"><p>Company Name:</p></td>
			<td align="left"><p>
			<input name="RecipCompanyName" type="text" id="RecipCompanyName" value="<%=pcf_FillFormField("RecipCompanyName", false)%>">
			<%pcs_RequiredImageTag "RecipCompanyName", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Department:</p></td>
			<td align="left"><p>
			<input name="RecipDepartment" type="text" id="RecipDepartment" value="<%=pcf_FillFormField("RecipDepartment", false)%>">
						<%pcs_RequiredImageTag "RecipDepartment", false%></p>
			</td>
			</tr>
			<% if len(Session("ErrRecipPhoneNumber"))>0 then %>
			<tr>
			<td colspan="2">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> 
			You must enter a valid Phone Number.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>Phone Number:</p></td>
			<td align="left"><p>
			<input name="RecipPhoneNumber" type="text" id="RecipPhoneNumber" value="<%=pcf_FillFormField("RecipPhoneNumber", true)%>">
			<%pcs_RequiredImageTag "RecipPhoneNumber", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Pager Number:</p></td>
			<td align="left"><p>
			<input name="RecipPagerNumber" type="text" id="RecipPagerNumber" value="<%=pcf_FillFormField("RecipPagerNumber", false)%>">
			<%pcs_RequiredImageTag "RecipPagerNumber", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Fax Number:</p></td>
			<td align="left"><p>
			<input name="RecipFaxNumber" type="text" id="RecipFaxNumber" value="<%=pcf_FillFormField("RecipFaxNumber", false)%>">
			<%pcs_RequiredImageTag "RecipFaxNumber", false%></p>
			</td>
			</tr>
			<% if len(Session("ErrRecipEmailAddress"))>0 then %>
			<tr>
			<td colspan="2">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> 
			You must enter a valid Email Address.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>Email Address:</p></td>
			<td align="left"><p>
			<input name="RecipEmailAddress" type="text" id="RecipEmailAddress" value="<%=pcf_FillFormField("RecipEmailAddress", false)%>">
						<%pcs_RequiredImageTag "RecipEmailAddress", false%></p>
			</td>
			</tr>
			<tr>
			<th colspan="2">Location Details</th>
			</tr>
							
			<%
			'///////////////////////////////////////////////////////////
			'// START: COUNTRY AND STATE/ PROVINCE CONFIG
			'///////////////////////////////////////////////////////////
			' 
			' 1) Place this section ABOVE the Country field
			' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
			' 3) Additional Required Info
			
			'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
			pcv_isStateOrProvinceCodeRequired = isRequiredState2 '// determines if validation is performed (true or false)
			pcv_isProvinceCodeRequired = isRequiredProvince2 '// determines if validation is performed (true or false)
			pcv_isCountryCodeRequired = true '// determines if validation is performed (true or false)
			
			'// #3 Additional Required Info
			pcv_strTargetForm = "form1" '// Name of Form
			pcv_strCountryBox = "RecipCountryCode" '// Name of Country Dropdown
			pcv_strTargetBox = "RecipStateOrProvinceCode" '// Name of State Dropdown
			pcv_strProvinceBox =  "RecipProvinceCode" '// Name of Province Field
			
			'// Set local Country to Session
			if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strCountryBox))=True then
				Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
			end if
			
			'// Set local State to Session
			if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strTargetBox))=True then
				Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session("pcAdminRecipStateOrProvinceCode")
			end if
			
			'// Set local Province to Session
			if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" OR isNULL(Session(pcv_strSessionPrefix&pcv_strProvinceBox))=True then
				Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  Session("pcAdminRecipStateOrProvinceCode")
			end if
			
			'// Declare the instance number if greater than 1
			pcv_strFormInstance = "2"
			'///////////////////////////////////////////////////////////
			'// END: COUNTRY AND STATE/ PROVINCE CONFIG
			'///////////////////////////////////////////////////////////
			%>
			
			<%
			'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
			pcs_CountryDropdown
			%>	
							
			<tr>
			<td align="right"><p>Address Line 1:</p></td>
			<td align="left"><p>
			<input name="RecipLine1" type="text" id="RecipLine1" value="<%=pcf_FillFormField("RecipLine1", true)%>">
			<%pcs_RequiredImageTag "RecipLine1", true%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Address Line 2:</p></td>
			<td align="left"><p>
			<input name="RecipLine2" type="text" id="RecipLine2" value="<%=pcf_FillFormField("RecipLine2", false)%>">
			<%pcs_RequiredImageTag "RecipLine2", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>City:</p></td>
			<td align="left"><p>
			<input name="RecipCity" type="text" id="RecipCity" value="<%=pcf_FillFormField("RecipCity", true)%>">
						<%pcs_RequiredImageTag "RecipCity", true%></p>
			</td>
			</tr>
						
			<%
			'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
			pcs_StateProvince
			%>
							
			<tr>
			<td align="right"><p>Postal Code:</p></td>
			<td align="left"><p>
			<input name="RecipPostalCode" type="text" id="RecipPostalCode" value="<%=pcf_FillFormField("RecipPostalCode", isRequiredRecipPostal)%>">
			<%pcs_RequiredImageTag "RecipPostalCode", isRequiredRecipPostal %></p>
			</td>
			</tr>

			<tr>
			<td align="right"><p>Customer Reference:</p></td>
			<td align="left"><p>
			<input name="CustomerReference" type="text" id="CustomerReference" value="<%=pcf_FillFormField("CustomerReference", true)%>">
						<%pcs_RequiredImageTag "CustomerReference", true%></p>
			</td>
			</tr>					
			<tr>
			<td align="right"><p>Customer PO Number:</p></td>
			<td align="left"><p>
			<input name="CustomerPONumber" type="text" id="CustomerPONumber" value="<%=pcf_FillFormField("CustomerPONumber", false)%>">
						<%pcs_RequiredImageTag "CustomerPONumber", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right"><p>Customer Invoice Number:</p></td>
			<td align="left"><p>
			<input name="CustomerInvoiceNumber" type="text" id="CustomerInvoiceNumber" value="<%=pcf_FillFormField("CustomerInvoiceNumber", false)%>">
						<%pcs_RequiredImageTag "CustomerInvoiceNumber", false%></p>
			</td>
			</tr>
			<tr>
			<td align="right">    </td>
			<td>
			<input type="checkbox" name="ResidentialDelivery" value="1" class="clearBorder" <%=pcf_CheckOption("ResidentialDelivery", "1")%>>
			<strong>This is a Residential Delivery</strong>
			</td>
			</tr>
			
			<tr>
			<td align="right">    </td>
			<td>
			<input type="checkbox" name="InsideDelivery" value="1" class="clearBorder" <%=pcf_CheckOption("InsideDelivery", "1")%>>
			<strong>Inside Delivery</strong>
			</td>
			</tr>
						
			<tr>
			<td align="right"></td>
			<td align="left">
			</td>
			</tr>
		</table>							
		
		</div>
		
		
		<!--							
		//////////////////////////////////////////////////////////////////////////////////////////////
		// SHIPPING ALERTS
		//////////////////////////////////////////////////////////////////////////////////////////////
		-->
		<div id="tab4" class="panes">
		<table class="pcCPcontent">
			<tr>
			<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
			<td colspan="2"><span class="title">FedEx ShipAlert<sup>&reg;</sup> Notifications</span></td>
			</tr>
			<tr>
			<td colspan="2">
			<strong>Shipment notification</strong> &ndash; Automatically
			send an email message or fax indicating
			the shipment is on the way.<br>
			<strong>Delivery notification</strong> &ndash; receive a
			delivery notification for an express package.
			<br>
			<strong>Email address</strong> &ndash; Enter the email IDs or
			fax number to receive the notifications.
			</td>
			</tr>
			<tr>
			<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
			<th colspan="2">Select Notifications:</th>
			</tr>
			<tr>
			<td width="30%" align="left"><strong>Shipper Notification:</strong></td>
			<td width="70%" align="left">
			<select name="ShipperNotification" id="select">
			<option value="FAX" <%=pcf_SelectOption("ShipperNotification","FAX")%>>Notify by FAX</option>
			<option value="EMAIL" <%=pcf_SelectOption("ShipperNotification","EMAIL")%>>Notify by E-mail</option>
			</select>
			</td>
			</tr>
			<tr>
			<td colspan="2">
			<input type="checkbox" name="ShipperShipmentNotification" value="1" class="clearBorder" <%=pcf_CheckOption("ShipperShipmentNotification", "1")%>>
			Shipment Notification<br>
			<input type="checkbox" name="ShipperDeliveryNotification" value="1" class="clearBorder" <%=pcf_CheckOption("ShipperDeliveryNotification", "1")%>>
			Delivery Notification<br>
			<input type="checkbox" name="ShipperExceptionNotification" value="1" class="clearBorder" <%=pcf_CheckOption("ShipperExceptionNotification", "1")%>>
			Exception Notification
			</td>
			</tr>
			<tr>
			<td align="left"><strong>Recipient Notification:</strong></td>
			<td align="left">&nbsp;</td>
			</tr>
			<tr>
			<td colspan="2">
			<input type="checkbox" name="RecipientShipmentNotification" value="1" class="clearBorder" <%=pcf_CheckOption("RecipientShipmentNotification", "1")%>>
			Shipment Notification<br>
			<input type="checkbox" name="RecipientDeliveryNotification" value="1" class="clearBorder" <%=pcf_CheckOption("RecipientDeliveryNotification", "1")%>>
			Delivery Notification<br>
			<input type="checkbox" name="RecipientExceptionNotification" value="1" class="clearBorder" <%=pcf_CheckOption("RecipientExceptionNotification", "1")%>>
			Exception Notification
			
			<% if len(Session("ErrOtherNotification1"))>0 then %>
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> 
			You must enter a valid Email Address. <br />
			<% end if %>
			</td>
			</tr>
			<tr>
			<td align="left"><strong>Other Notify by Email Address:</strong></td>
			<td align="left">
			<input name="OtherNotification1" type="text" id="OtherNotification1" value="<%=pcf_FillFormField("OtherNotification1", false)%>">
			<%pcs_RequiredImageTag "OtherNotification1", false%>
			</td>
			</tr>
			<tr>
			<td colspan="2">
			<input type="checkbox" name="OtherShipmentNotification" value="1" class="clearBorder" <%=pcf_CheckOption("OtherShipmentNotification", "1")%>>
			Shipment Notification<br>
			<input type="checkbox" name="OtherDeliveryNotification" value="1" class="clearBorder" <%=pcf_CheckOption("OtherDeliveryNotification", "1")%>>
			Delivery Notification<br>
			<input type="checkbox" name="OtherExceptionNotification" value="1" class="clearBorder" <%=pcf_CheckOption("OtherExceptionNotification", "1")%>>
			Exception Notification
			</td>
			</tr>
			</table>						
			
			</div>
					
		<!--							
		//////////////////////////////////////////////////////////////////////////////////////////////
		// PACKAGES
		//////////////////////////////////////////////////////////////////////////////////////////////
		-->
		<% 
		for k=1 to pcPackageCount 
			%>
			<div id="tab<%=4+int(k)%>" class="panes">				
				<table class="pcCPcontent">
					<tr>
					<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<%
					'// If the tab was processed, skip it.
					if pcLocalArray(k-1) <> "shipped" then	
					%>
					<tr>
					<td colspan="2"><span class="title">Package <%=k%> Information</span></td>
					</tr>
					<tr>
					<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td colspan="2">
						<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
						<!--
						function FaxSelected<%=k%>(){

						
						var selectValDom = document.forms['form1'];
						if (selectValDom.FaxLetter<%=k%>.checked == true) {
						document.getElementById('FaxTable<%=k%>').style.display='';
						}else{
						document.getElementById('FaxTable<%=k%>').style.display='none';
						}
						}
						 //-->
						</SCRIPT>
						<%
						if Session("pcAdminFaxLetter"&k)="true" then
							pcv_strDisplayStyle="style=""display:visible"""
						else
							pcv_strDisplayStyle="style=""display:none"""
						end if
						%>
						<input onClick="FaxSelected<%=k%>();" name="FaxLetter<%=k%>" id="FaxLetter<%=k%>" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("FaxLetter"&k, "true")%>>
						Click Here to view <b>package contents</b>.
						
							<table class="pcCPcontent" ID="FaxTable<%=k%>" <%=pcv_strDisplayStyle%>>
								<tr>
									<td colspan="2" valign="top">
										<%
										xProductDisplayArray = split(Session("pcAdminPrdList"&k),",")
										For pcv_xCounter=0 to (ubound(xProductDisplayArray)-1)
											pcv_intPackageInfo_ID = xProductDisplayArray(pcv_xCounter)
											' GET THE PACKAGE CONTENTS
											' >>> Tables: products, ProductsOrdered
											query = 		"SELECT ProductsOrdered.pcPackageInfo_ID , products.description, products.idProduct, products.OverSizeSpec "
											query = query & "FROM ProductsOrdered "
											query = query & "INNER JOIN products "
											query = query & "ON ProductsOrdered.idProduct = products.idProduct "
											query = query & "WHERE ProductsOrdered.idProductOrdered=" & pcv_intPackageInfo_ID &" "  
																
											set rs2=server.CreateObject("ADODB.RecordSet")
											set rs2=conntemp.execute(query)		
											
											if err.number<>0 then
												'// handle admin error
											end if
											
											if NOT rs2.eof then
												Do until rs2.eof	
													pcv_strProductDescription = rs2("description")
													pOverSizeSpec=rs2("OverSizeSpec")
													if pOverSizeSpec="" or isNull(pOverSizeSpec) then
														pOverSizeSpec="NO"
													end if
													if pOverSizeSpec<>"NO" then
														pOSArray=split(pOverSizeSpec,"||")
														if ubound(pOSArray)>2 then
															tOS_width=pOSArray(0)
															tOS_height=pOSArray(1)
															tOS_length=pOSArray(2)
														else
															tOS_width=FEDEX_WIDTH
															tOS_height=FEDEX_HEIGHT
															tOS_length=FEDEX_LENGTH
														end if
													else
														tOS_width=FEDEX_WIDTH
														tOS_height=FEDEX_HEIGHT
														tOS_length=FEDEX_LENGTH
													end if
													'// You only ship one oversized item per package, override dimensions for this tab
													Session("pcAdminLength"&k) = tOS_length
													Session("pcAdminWidth"&k) = tOS_width
													Session("pcAdminHeight"&k) = tOS_height
													%>
													<li><%=pcv_strProductDescription%></li>
													<%
												rs2.movenext
												Loop								
											end if	
										Next				
										%>	
									</td>
								</tr>
							</table>	
						</td>
					</tr>
					<tr>
					<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
					<th colspan="2">Settings <%'=k%></th>
					</tr>
					<tr>
					<td colspan="2">
					<%
					'// Set Service Type to local
					Select Case Service						
						Case "FedEx First Overnight": ServiceSelector="FIRSTOVERNIGHT"
						Case "FedEx Priority Overnight": ServiceSelector="PRIORITYOVERNIGHT"
						Case "FedEx Standard Overnight": ServiceSelector="STANDARDOVERNIGHT"
						Case "FedEx 2Day": ServiceSelector="FEDEX2DAY"
						Case "FedEx Express Saver": ServiceSelector="FEDEXEXPRESSSAVER"
						Case "FedEx Ground": ServiceSelector="FEDEXGROUND"
						Case "FedEx Home Delivery": ServiceSelector="GROUNDHOMEDELIVERY"
						Case "FedEx International First": ServiceSelector="INTERNATIONALFIRST"
						Case "FedEx International Priority": ServiceSelector="INTERNATIONALPRIORITY"
						Case "FedEx International Economy": ServiceSelector="INTERNATIONALECONOMY"
						Case "FedEx 1Day Freight": ServiceSelector="FEDEX1DAYFREIGHT"
						Case "FedEx 2Day Freight": ServiceSelector="FEDEX2DAYFREIGHT"
						Case "FedEx 3Day Freight": ServiceSelector="FEDEX3DAYFREIGHT"
						Case "FedEx International Priority Freight": ServiceSelector="INTERNATIONALPRIORITYFREIGHT"
						Case "FedEx International Economy Freight": ServiceSelector="INTERNATIONALECONOMYFREIGHT"
					End Select
					
					if Session("pcAdminService"&k)="" then
						Session("pcAdminService"&k)=ServiceSelector
					end if	
					%>
					<p>
					<strong>Service Type: </strong>
					<select name="Service<%=k%>" id="Service<%=k%>">
						<option value="FIRSTOVERNIGHT" <%=pcf_SelectOption("Service"&k,"FIRSTOVERNIGHT")%>>FedEx First Overnight&reg;</option>
						<option value="PRIORITYOVERNIGHT" <%=pcf_SelectOption("Service"&k,"PRIORITYOVERNIGHT")%>>FedEx Priority Overnight&reg;</option>
						<option value="STANDARDOVERNIGHT" <%=pcf_SelectOption("Service"&k,"STANDARDOVERNIGHT")%>>FedEx Standard Overnight&reg;</option>									
						<option value="FEDEX2DAY" <%=pcf_SelectOption("Service"&k,"FEDEX2DAY")%>>FedEx 2Day&reg;</option>
						<option value="FEDEXEXPRESSSAVER" <%=pcf_SelectOption("Service"&k,"FEDEXEXPRESSSAVER")%>>FedEx Express Saver&reg;</option>
						<option value="FEDEXGROUND" <%=pcf_SelectOption("Service"&k,"FEDEXGROUND")%>>FedEx Ground&reg;</option>
						<option value="GROUNDHOMEDELIVERY" <%=pcf_SelectOption("Service"&k,"GROUNDHOMEDELIVERY")%>>FedEx Home Delivery&reg;</option>
						<option value="INTERNATIONALFIRST" <%=pcf_SelectOption("Service"&k,"INTERNATIONALFIRST")%>>FedEx International First&reg;</option>
						<option value="INTERNATIONALPRIORITY" <%=pcf_SelectOption("Service"&k,"INTERNATIONALPRIORITY")%>>FedEx International Priority&reg;</option>									
						<option value="INTERNATIONALECONOMY" <%=pcf_SelectOption("Service"&k,"INTERNATIONALECONOMY")%>>FedEx International Economy&reg; </option>
						<option value="INTERNATIONALPRIORITYFREIGHT" <%=pcf_SelectOption("Service"&k,"INTERNATIONALPRIORITYFREIGHT")%>>FedEx International Priority&reg; Freight</option>									
						<option value="INTERNATIONALECONOMYFREIGHT" <%=pcf_SelectOption("Service"&k,"INTERNATIONALECONOMYFREIGHT")%>>FedEx International Economy&reg; Freight</option>
						<option value="FEDEX1DAYFREIGHT" <%=pcf_SelectOption("Service"&k,"FEDEX1DAYFREIGHT")%>>FedEx 1Day&reg; Freight</option>
						<option value="FEDEX2DAYFREIGHT" <%=pcf_SelectOption("Service"&k,"FEDEX2DAYFREIGHT")%>>FedEx 2Day&reg; Freight</option>									
						<option value="FEDEX3DAYFREIGHT" <%=pcf_SelectOption("Service"&k,"FEDEX3DAYFREIGHT")%>>FedEx 3Day&reg; Freight</option>
					</select>

					<%pcs_RequiredImageTag "Service"&k, true%>
					</p>
					<br />
					<span class="pcCPnotes">					
					When using FedEx packaging, select the
					packaging type from the drop-down list.<br>

					When using non-FedEx packaging, select &quot;Your
					Packaging&quot;, and then enter
					the dimensions manually.
					</span>	
					<br /><br />				
					<p>
					<strong>Package Type:</strong> 
					<select name="Packaging<%=k%>" id="Packaging<%=k%>">
					<option value="FEDEXENVELOPE" <%=pcf_SelectOption("Packaging"&k,"FEDEXENVELOPE")%>>FedEx&reg; Envelope</option>
					<option value="FEDEXPAK" <%=pcf_SelectOption("Packaging"&k,"FEDEXPAK")%>>FedEx&reg; Pak</option>									
					<option value="FEDEXBOX" <%=pcf_SelectOption("Packaging"&k,"FEDEXBOX")%>>FedEx&reg; Box</option>
					<option value="FEDEXTUBE" <%=pcf_SelectOption("Packaging"&k,"FEDEXTUBE")%>>FedEx&reg; Tube</option>
					<option value="FEDEX10KGBOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX10KGBOX")%>>FedEx&reg; 10kg Box</option>
					<option value="FEDEX25KGBOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX25KGBOX")%>>FedEx&reg; 25kg Box</option>
					<option value="YOURPACKAGING" <%=pcf_SelectOption("Packaging"&k,"YOURPACKAGING")%>>Customer Package</option>
					</select>
					<%pcs_RequiredImageTag "Packaging"&k, true%>
					</p>
					<br />
					</td>
					</tr>
					<tr>
					<th colspan="2">Dimensions and Weight</th>
					</tr>
					<tr>
					<td colspan="2">
					<p>
					<strong>Package Dimensions:</strong>  					
					<br>
					<span class="pcCPnotes">
					Maximum 274 cm in length (always the longest side)
					<br />
					Maximum 330 cm in length and girth combined. Girth = (2 x height) + (2 x   width)	
					</span>	  
					<br />	
					<p>
					Units: 
					<select name="Units<%=k%>" id="Units<%=k%>">
					<option value="">Select A Unit</option>
					<option value="IN" <%=pcf_SelectOption("Units"&k,"IN")%>>Inches</option>
					<option value="CM" <%=pcf_SelectOption("Units"&k,"CM")%>>Centimeters</option>
					</select>
					<%pcs_RequiredImageTag "Units"&k, false%>
					</p>
					<br />
					<p>	
					Length: 
					<input name="Length<%=k%>" type="text" id="Length<%=k%>" value="<%=pcf_FillFormField("Length"&k, false)%>" width="4">
					<%pcs_RequiredImageTag "Length"&k, false%>
					&nbsp;								
					Width: <input name="Width<%=k%>" type="text" id="Width<%=k%>" value="<%=pcf_FillFormField("Width"&k, false)%>" width="4">
					<%pcs_RequiredImageTag "Width"&k, false%>
					&nbsp;
					Height: <input name="Height<%=k%>" type="text" id="Height<%=k%>" value="<%=pcf_FillFormField("Height"&k, false)%>" width="4">
					<%pcs_RequiredImageTag "Height"&k, false%>
					</p>	
					<br />
					<p><strong>Package Weight:</strong>
					<br />
					<span class="pcCPnotes">
					Enter the weight of the package. If there is more than one package in the shipment, enter the weight of the first package or the total shipment weight.
					</span>
					</p>
					<p>
					Weight Units: 
					<select name="WeightUnits<%=k%>" id="WeightUnits<%=k%>">
					<option value="LBS" <%=pcf_SelectOption("WeightUnits"&k,"LBS")%>>LBS</option>
					<option value="KGS" <%=pcf_SelectOption("WeightUnits"&k,"KGS")%>>KGS</option>
					</select>
					<%pcs_RequiredImageTag "WeightUnits"&k, true
					%>
					<% 
					if scShipFromWeightUnit="KGS" then
						intShipWeightPounds=int(pcOrd_ShipWeight/1000) 
						intShipWeightOunces=pcv_ShipWeight-(intShipWeightPounds*1000)
					else
						intShipWeightPounds=int(pcOrd_ShipWeight/16) 'intPounds used for USPS
						intShipWeightOunces=pcOrd_ShipWeight-(intShipWeightPounds*16) 'intUniversalOunces used for USPS
					end if

					intMPackageWeight=intShipWeightPounds
					if intMPackageWeight<1 AND intShipWeightOunces<1 then
						intMPackageWeight=0
						intShipWeightOunces=0
					end if

					if intMPackageWeight<1 AND intShipWeightOunces>0 then 'if total weight is less then a pound, make UPS/FedEX weight 1 pound
						intMPackageWeight=1
					else  'total weight is not less then a pound and ounces exist, round weight up one more pound.
						If intMPackageWeight>0 AND intShipWeightOunces>0 then
							intMPackageWeight=(intMPackageWeight+1)
						Else
							intMPackageWeight=1
						End if
					end if
					
					if pcPackageCount=1 AND Session("pcAdminWeight"&k)="" then 
						'// Get weight
						Session("pcAdminWeight"&k) = intMPackageWeight
					end if 
					%>
					&nbsp;&nbsp;&nbsp;&nbsp;
					Weight: <input name="Weight<%=k%>" type="text" id="Weight<%=k%>" value="<%=pcf_FillFormField("Weight"&k, true)%>">
					<%pcs_RequiredImageTag "Weight"&k, true%>
					</p>
					<br />
					
					</td>
					</tr>
					<tr>
					<th colspan="2">Package Value</th>
					</tr>
					<tr>
					<td colspan="2">
					<%
					if Session("pcAdmindeclaredvalue"&k)="" then 
						Session("pcAdmindeclaredvalue"&k) = 100
					end if 
					%>
					<p>
					Declared Value: <input name="declaredvalue<%=k%>" type="text" id="declaredvalue<%=k%>" value="<%=pcf_FillFormField("declaredvalue"&k, true)%>">
					<% pcs_RequiredImageTag "declaredvalue"&k, true %>
					<!--					
					&nbsp;&nbsp;&nbsp;&nbsp;
					Custom Carriage Value: <input name="Carriagevalue<%=k%>" type="text" id="Carriagevalue<%=k%>" value="<%=pcf_FillFormField("Carriagevalue"&k, true)%>">
					<% pcs_RequiredImageTag "Carriagevalue"&k, true %> 
					-->
					</p>
					<br />
					</td>
					</tr>
						<% if k = 1 then %>
						<tr>
						<th colspan="2">Freight Settings</th>
						</tr>
						<tr>
						<td colspan="2">
						<p>
						<strong>Booking Confirmation Number:</strong>
						<br>
						When using FedEx International Freight services only.
						<br>
						<input name="BookingConfirmationNumber" type="text" id="BookingConfirmationNumber" value="<%=pcf_FillFormField("BookingConfirmationNumber", false)%>" width="12">
						<%pcs_RequiredImageTag "BookingConfirmationNumber", false%>
						</p>	
						</td>
						</tr>
						<% end if %>
					<% else %>
					<tr>
					<th colspan="2">This package has been shipped.</th>
					</tr>
					<%
					end if
					%>	
				</table>				
			</div>
			<%
		next %>
	
			<br />
			<br />
			
			<%
			pcv_strPreviousPage = "Orddetails.asp?id=" & pcv_intOrderID
			pcv_strAddPackagePage = "sds_ShipOrderWizard1.asp?idorder="&pcv_intOrderID&"&PageAction=FedEx&PackageCount="&pcPackageCount&"&ItemsList="&pcv_strItemsList
			%>
			
			<p>
				<div align="center">
					<input type="button" name="Button" value="Start Over" onclick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
					<% if pcPackageCount<4 then %>
					<input type="button" name="Button" value="Add Another Package" onclick="document.location.href='<%=pcv_strAddPackagePage%>'" class="ibtnGrey">
					<% end if %>
					<input type="submit" name="submit" value="Process Shipment" class="ibtnGrey">
					<br />
					<br />
					<input type="button" name="Button" value="Go Back To Order Details" onclick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
				</div>
			</p>						
		</td>
		</tr>
	<!--End -->
	</table>
</form>
<%
end if 
'*******************************************************************************
' END: LOAD HTML FORM
'*******************************************************************************	
%>
</td>
</tr>
</table>
<%
call closedb() 
'// DESTROY THE SESSIONS
'pcs_ClearAllSessions
'Session("pcAdminPackageCount")=""
'Session("pcAdminOrderID")=""
'Session("pcGlobalArray")=""
'Session("pcAdminTotalWeight")=""
'Session("pcAdminDeclaredValue")=""
For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray)
	Session("pcAdminPrdList"&(xArrayCount+1))
Next
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>
<!--#include file="AdminFooter.asp"-->