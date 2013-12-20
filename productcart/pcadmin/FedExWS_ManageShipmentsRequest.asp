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
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
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
	#smartpost, #expressfreight, #homedelivery {
	display:none;
	}
</style>

<%
Dim objFEDEXXmlDoc, objFedExStream, strFileName, GraphicXML
Dim iPageCurrent, varFlagIncomplete, uery, strORD, pcv_intOrderID
Dim pcv_strMethodName, pcv_strMethodReply, CustomerTransactionIdentifier, pcv_strAccountNumber, pcv_strMeterNumber, pcv_strCarrierCode
Dim pcv_strTrackingNumber, pcv_strShipmentAccountNumber
Dim pcv_strDestinationCountryCode, pcv_strDestinationPostalCode, pcv_strLanguageCode, pcv_strLocaleCode, pcv_strDetailScans, pcv_strPagingToken
Dim fedex_postdataWS, objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEXWS_result, FEDEX_URL, pcv_strErrorMsg, pcv_strAction

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
pcPageName="FedExWS_ManageShipmentsRequest.asp"
ErrPageName="FedExWS_ManageShipmentsRequest.asp"

'// ACTION
pcv_strAction = request("Action")

'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExWSClass

'// FEDEX CREDENTIALS
query = "SELECT ShipmentTypes.userID, ShipmentTypes.password, ShipmentTypes.AccessLicense, ShipmentTypes.FedExKey, ShipmentTypes.FedExPwd "
query = query & "FROM ShipmentTypes "
query = query & "WHERE (((ShipmentTypes.idShipment)=9));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if NOT rs.eof then
	FedExAccountNumber=rs("userID")
	FedExMeterNumber=rs("password")
	pcv_strEnvironment=rs("AccessLicense")
	FedExkey=rs("FedExKey")
	FedExPassword=rs("FedExPwd")
end if
set rs=nothing

'// DATE FUNCTION
function ShowDateFrmt(x)
	ShowDateFrmt = x
end function

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
if Session("pcAdminResidentialDelivery")&"" = "" then
	if pOrdShipType=0 then
		pcv_strResidentialDelivery = "true"
	else
		pcv_strResidentialDelivery = "false"
	end if
	Session("pcAdminResidentialDelivery") = pcv_strResidentialDelivery
end if

'// SHIP CONSTANTS

if Session("pcAdminSMHubID") = "" then
	pcv_strHubId = FDXWS_SMHUBID
	Session("pcAdminSMHubID") = pcv_strHubId
end if

'// DropType
if Session("pcAdminDropoffType") = "" then
	pcv_strDropoffType = FEDEXWS_DROPOFF_TYPE
	Session("pcAdminDropoffType") = pcv_strDropoffType
end if

'//Rate Request Type
if Session("pcAdminRateRequestType") = "" then
	pcv_strRateRequestType = FEDEXWS_LISTRATE
	if pcv_strRateRequestType="-1" then
		pcv_strRateRequestType = "LIST"
	end if
	if pcv_strRateRequestType="0" then
		pcv_strRateRequestType = "ACCOUNT"
	end if
	if pcv_strRateRequestType="-2" then
		pcv_strRateRequestType = "ACCOUNT"
	end if
	Session("pcAdminRateRequestType") = pcv_strRateRequestType
end if
'// PackageType
if Session("pcAdminPackaging1") = "" then
	pcv_strPackaging = FEDEXWS_FEDEX_PACKAGE
	Session("pcAdminPackaging1") = pcv_strPackaging
	Session("pcAdminPackaging2") = pcv_strPackaging
	Session("pcAdminPackaging3") = pcv_strPackaging
	Session("pcAdminPackaging4") = pcv_strPackaging
end if

'// L
if Session("pcAdminLength1") = "" then
	pcv_strLength = FEDEXWS_LENGTH
	Session("pcAdminLength1") = pcv_strLength
	Session("pcAdminLength2") = pcv_strLength
	Session("pcAdminLength3") = pcv_strLength
	Session("pcAdminLength4") = pcv_strLength
end if

'// W
if Session("pcAdminWidth1") = "" then
	pcv_strWidth = FEDEXWS_WIDTH
	Session("pcAdminWidth1") = pcv_strWidth
	Session("pcAdminWidth2") = pcv_strWidth
	Session("pcAdminWidth3") = pcv_strWidth
	Session("pcAdminWidth4") = pcv_strWidth
end if

'// H
if Session("pcAdminHeight1") = "" then
	pcv_strHeight = FEDEXWS_HEIGHT
	Session("pcAdminHeight1") = pcv_strHeight
	Session("pcAdminHeight2") = pcv_strHeight
	Session("pcAdminHeight3") = pcv_strHeight
	Session("pcAdminHeight4") = pcv_strHeight
end if

'// U
if Session("pcAdminUnits1") = "" then
	pcv_strUnits = FEDEXWS_DIM_UNIT
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
	Session("pcAdminNotificationShipperEmail") = pcv_strOriginEmailAddress
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
	Session("pcAdminNotificationRecipientEmail") = pcv_strRecipEmailAddress
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
If Session("pcAdminRecipCountryCode")<>"US" Then
	isRequiredCVAmount = true
	isRequiredCVCurrency = true
	isRequiredNumberOfPieces = true
	isRequiredDescription = true
	isRequiredCountryOfManufacture = true
	isRequiredCommodityWeight = true
	isRequiredCommodityQuantity = true
	isRequiredCommodityQuantityUnits = true
	isRequiredCommodityUnitPrice = true
	isRequiredDutiesAccountNumber = true
	isRequiredDutiesCountryCode = true
Else
	isRequiredCVAmount = false
	isRequiredCVCurrency = false
	isRequiredNumberOfPieces = false
	isRequiredDescription = false
	isRequiredCountryOfManufacture = false
	isRequiredCommodityWeight = false
	isRequiredCommodityQuantity = false
	isRequiredCommodityQuantityUnits = false
	isRequiredCommodityUnitPrice = false
	isRequiredDutiesAccountNumber = false
	isRequiredDutiesCountryCode = false
End If

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
	pcv_strType = "COMMON2D"
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
			This flexible service allows a customer to request shipments and print return labels.  Simply fill out all required fields from each of the console's tabs.
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

				For pcv_xCounter = 1 to pcPackageCount

					' If its shipped the field is no longer required
					if pcLocalArray(pcv_xCounter-1) = "shipped" then
						pcv_strToggle = false
					else
						pcv_strToggle = true
					end if

					pcs_ValidateTextField	"FaxLetter"&pcv_xCounter, false, 0

					pcs_ValidateTextField	"Service"&pcv_xCounter, pcv_strToggle, 0
					select case session("pcAdminService"&pcv_xCounter)
						case 1
							session("pcAdminService"&pcv_xCounter) = "FIRST_OVERNIGHT"
						case 2
							session("pcAdminService"&pcv_xCounter) = "PRIORITY_OVERNIGHT"
						case 3
							session("pcAdminService"&pcv_xCounter) = "STANDARD_OVERNIGHT"
						case 4
							session("pcAdminService"&pcv_xCounter) = "FEDEX_2_DAY"
						case 5
							session("pcAdminService"&pcv_xCounter) = "FEDEX_EXPRESS_SAVER"
						case 6
							session("pcAdminService"&pcv_xCounter) = "FEDEX_GROUND"
						case 7
							session("pcAdminService"&pcv_xCounter) = "GROUND_HOME_DELIVERY"
						case 8
							session("pcAdminService"&pcv_xCounter) = "INTERNATIONAL_FIRST"
						case 9
							session("pcAdminService"&pcv_xCounter) = "INTERNATIONAL_PRIORITY"
						case 10
							session("pcAdminService"&pcv_xCounter) = "INTERNATIONAL_ECONOMY"
						case 11
							session("pcAdminService"&pcv_xCounter) = "INTERNATIONAL_PRIORITY_FREIGHT"
						case 12
							session("pcAdminService"&pcv_xCounter) = "INTERNATIONAL_ECONOMY_FREIGHT"
						case 13
							session("pcAdminService"&pcv_xCounter) = "FEDEX_1_DAY_FREIGHT"
						case 14
							session("pcAdminService"&pcv_xCounter) = "FEDEX_2_DAY_FREIGHT"
						case 15
							session("pcAdminService"&pcv_xCounter) = "FEDEX_3_DAY_FREIGHT"
						case 16
							session("pcAdminService"&pcv_xCounter) = "SMART_POST"
					end select

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
				Next

				pcs_ValidateTextField	"TotalShipmentWeight", true, 10

				pcs_ValidateTextField	"ShipmentWeightUnits", true, 2
				If Session("pcAdminShipmentWeightUnits")&""="" Then
					Session("pcAdminShipmentWeightUnits") = "LB"
				End If
				pcs_ValidateTextField	"RateRequestType", true, 10
				If Session("pcAdminRateRequestType")&""="" Then
					Session("pcAdminRateRequestType") = "LIST"
				End If
				Session("pcAdminTotalDeclaredValue")=FormatNumber(pcv_strTotalDeclaredValue,2)
				pcs_ValidateTextField	"CarrierCode", true, 10
				select case session("pcAdminCarrierCode")
					case 1
						session("pcAdminCarrierCode") = "FDXE"
					case 2
						session("pcAdminCarrierCode") = "FDXG"
				end select

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
				pcs_ValidateTextField	"OriginPersonName", true, 0
				pcs_ValidateTextField	"OriginCompanyName", true, 0
				pcs_ValidateTextField	"OriginDepartment", false, 10
				pcs_ValidatePhoneNumber	"OriginPhoneNumber", true, 16
				pcs_ValidatePhoneNumber	"OriginPagerNumber", false, 16
				pcs_ValidatePhoneNumber	"OriginFaxNumber", false, 16
				pcs_ValidateEmailField	"OriginEmailAddress", true, 0
				pcs_ValidateTextField	"OriginLine1", true, 0
				pcs_ValidateTextField	"OriginLine2", false, 0
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
				pcs_ValidateTextField	"RecipPersonName", true, 0
				pcs_ValidateTextField	"RecipCompanyName", false, 0
				pcs_ValidateTextField	"RecipDepartment", false, 10
				pcs_ValidatePhoneNumber	"RecipPhoneNumber", true, 16
				pcs_ValidatePhoneNumber	"RecipPagerNumber", false, 16
				pcs_ValidatePhoneNumber	"RecipFaxNumber", false, 16
				pcs_ValidateEmailField	"RecipEmailAddress", false, 0

				'// Recipient Address
				pcs_ValidateTextField	"RecipCountryCode", true, 2

				'   >>> Recipient Address Conditionals
				if Session("pcAdminRecipCountryCode") = "US" OR Session("pcAdminRecipCountryCode") = "CA" then
					isRequiredRecipPostal = true
				else
					isRequiredRecipPostal = false
				end if
				pcs_ValidateTextField	"RecipLine1", true, 0
				pcs_ValidateTextField	"RecipLine2", false, 0
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

				If Session("pcAdminRecipCountryCode")<>"US" Then
					isRequiredCVAmount = true
					isRequiredCVCurrency = true
					isRequiredNumberOfPieces = true
					isRequiredDescription = true
					isRequiredCountryOfManufacture = true
					isRequiredCommodityWeight = true
					isRequiredCommodityQuantity = true
					isRequiredCommodityQuantityUnits = true
					isRequiredCommodityUnitPrice = true
					isRequiredDutiesAccountNumber = true
					isRequiredDutiesCountryCode = true
				Else
					isRequiredCVAmount = false
					isRequiredCVCurrency = false
					isRequiredNumberOfPieces = false
					isRequiredDescription = false
					isRequiredCountryOfManufacture = false
					isRequiredCommodityWeight = false
					isRequiredCommodityQuantity = false
					isRequiredCommodityQuantityUnits = false
					isRequiredCommodityUnitPrice = false
					isRequiredDutiesAccountNumber = false
					isRequiredDutiesCountryCode = false
				End If

				'// International
				pcs_ValidateTextField	"DutiesAccountNumber", isRequiredDutiesAccountNumber, 12
				pcs_ValidateTextField	"DutiesCountryCode", isRequiredDutiesCountryCode, 2
				pcs_ValidateTextField	"DutiesPayorType", false, 10

				pcs_ValidateTextField	"PayorType", false, 0 '// Required if PayorType is RECIPIENT or THIRDPARTY.
				pcs_ValidateTextField	"PayorAccountNumber", false, 0
				pcs_ValidateTextField	"PayorCountryCode", false, 2

				'// Customer Reference
				pcs_ValidateTextField	"CustomerReference", true, 0 '// FDXE-40, FDXG-30
				pcs_ValidateTextField	"CustomerPONumber", false, 30
				pcs_ValidateTextField	"CustomerInvoiceNumber", false, 30

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
					pcs_ValidateTextField	"RecipientTIN", false, 0
				else
					pcs_ValidateTextField	"RecipientTIN", false, 0
				end if
				pcs_ValidateTextField	"SenderTINNumber", false, 0
				pcs_ValidateTextField	"SenderTINType", false, 0
				pcs_ValidateTextField	"AESOrFTSRExemptionNumber", false, 0
				pcs_ValidateTextField	"NumberOfPieces", isRequiredNumberOfPieces, 0
				pcs_ValidateTextField	"Description", isRequiredDescription, 0
				pcs_ValidateTextField	"CountryOfManufacture", isRequiredCountryOfManufacture, 0
				pcs_ValidateTextField	"HarmonizedCode", false, 0
				pcs_ValidateTextField	"CommodityWeight", isRequiredCommodityWeight, 0
				pcs_ValidateTextField	"CommodityQuantity", isRequiredCommodityQuantity, 0
				pcs_ValidateTextField	"CommodityQuantityUnits", isRequiredCommodityQuantityUnits, 0
				pcs_ValidateTextField	"CommodityUnitPrice", isRequiredCommodityUnitPrice, 0
				pcs_ValidateTextField	"CommodityCustomsValue", false, 0
				pcs_ValidateTextField	"ExportLicenseNumber", false, 0
				pcs_ValidateTextField	"ExportLicenseExpirationDate", false, 0
				pcs_ValidateTextField	"CIMarksAndNumbers", false, 0
				pcs_ValidateTextField	"B13AFilingOption", false, 0
				pcs_ValidateTextField	"ExportComplianceStatement", false, 0

				'// Hold At Location
				pcs_ValidateTextField	"HALPhone", false, 0
				pcs_ValidateTextField	"HALCompanyName", false, 0
				pcs_ValidateTextField	"HALPersonName", false, 0
				pcs_ValidateTextField	"HALLine1", false, 0
				pcs_ValidateTextField	"HALCity", false, 0
				pcs_ValidateTextField	"HALStateOrProvinceCode", false, 0
				pcs_ValidateTextField	"HALCountryCode", false, 0
				pcs_ValidateTextField	"HALPostalCode", false, 0

				'// Dry Ice Shipment
				pcs_ValidateTextField	"SDIPackageCount", false, 0
				pcs_ValidateTextField	"SDIValue", false, 0
				pcs_ValidateTextField	"SDIUnit", false, 0

				'// Freight
				pcs_ValidateTextField "BookingConfirmationNumber", false, 12

				'//Special Services
				pcs_ValidateTextField "ResidentialDelivery", false, 0
				pcs_ValidateTextField "InsideDelivery", false, 0
				pcs_ValidateTextField "SaturdayPickup", false, 0
				pcs_ValidateTextField "SaturdayDelivery", false, 0
				pcs_ValidateTextField "SignatureOption", false, 0
				pcs_ValidateTextField "PriorityAlert", false, 0
				pcs_ValidateTextField "bSOHAL", false, 0
				pcs_ValidateTextField "bSODryIce", false, 0
				pcs_ValidateTextField "bSODGShip", false, 0
				pcs_ValidateTextField "DGAccessibility", false, 0
				pcs_ValidateTextField "DGAircraftOnly", false, 0
				pcs_ValidateTextField "bISOBrokerSelect", false, 0
				pcs_ValidateTextField "BSOCompanyName", false, 0
				pcs_ValidateTextField "BSOPhoneNumber", false, 0
				pcs_ValidateTextField "BSOStreetLines", false, 0
				pcs_ValidateTextField "BSOCity", false, 0
				pcs_ValidateTextField "BSOStateOrProvinceCode", false, 0
				pcs_ValidateTextField "BSOPostalCode", false, 0
				pcs_ValidateTextField "BSOCountryCode", false, 0
				pcs_ValidateTextField "bSOCODCollection", false, 0
				pcs_ValidateTextField "CODAmount", false, 0
				pcs_ValidateTextField "CODType", false, 0
				pcs_ValidateTextField "CODTinType", false, 0
				pcs_ValidateTextField "CODTinNumber", false, 0
				pcs_ValidateTextField "CODPersonName", false, 0
				pcs_ValidateTextField "CODCompanyName", false, 0
				pcs_ValidateTextField "CODPhoneNumber", false, 0
				pcs_ValidateTextField "CODTitle", false, 0
				pcs_ValidateTextField "CODStreetLines", false, 0
				pcs_ValidateTextField "CODCity", false, 0
				pcs_ValidateTextField "CODState", false, 0
				pcs_ValidateTextField "CODPostalCode", false, 0
				pcs_ValidateTextField "CODCountryCode", false, 0
				'//Customs
				pcs_ValidateTextField "CVAmount", isRequiredCVAmount, 0
				pcs_ValidateTextField "CVCurrency", isRequiredCVCurrency, 0
				pcs_ValidateTextField "CICAmount", false, 0
				pcs_ValidateTextField "CMCAmount", false, 0
				pcs_ValidateTextField "CFCAmount", false, 0
				pcs_ValidateTextField "CCIPurpose", false, 0
				pcs_ValidateTextField "CCIInvoiceNumber", false, 0
				pcs_ValidateTextField "CCIComments", false, 0
				'//Shipper Notification
				pcs_ValidateTextField	"NotificationShipperEmail", false, 0
				pcs_ValidateTextField	"NotificationRecipientEmail", false, 0
				pcs_ValidateTextField	"ShipperNotificationFormat", false, 0
				pcs_ValidateTextField "ShipperShipmentNotification", false, 0 'value="1"
				pcs_ValidateTextField "ShipperDeliveryNotification", false, 0 'value="1"
				pcs_ValidateTextField "ShipperExceptionNotification", false, 0 'value="1"

				'//Recipient Notification
				pcs_ValidateTextField "RecipientShipmentNotification", false, 0 'value="1"
				pcs_ValidateTextField "RecipientDeliveryNotification", false, 0 'value="1"
				pcs_ValidateTextField "RecipientExceptionNotification", false, 0 'value="1"

				'//Other Notification
				pcs_ValidateTextField "OtherShipmentNotification", false, 0 'value="1"
				pcs_ValidateTextField "OtherDeliveryNotification", false, 0 'value="1"
				pcs_ValidateTextField "OtherExceptionNotification", false, 0 'value="1"

				pcs_ValidateTextField "DeliveryType", false, 14
				pcs_ValidateTextField "DeliveryInstructions", false, 74
				pcs_ValidateTextField "DeliveryDate", false, 0
				pcs_ValidatePhoneNumber	"DeliveryPhone", false, 16
				pcs_ValidateTextField "RCIdValue", false, 0
				pcs_ValidateTextField "RCIdType", false, 0

				'Express Freight
				pcs_ValidateTextField "EFPackingListEnclosed", false, 0
				pcs_ValidateTextField "EFShippersLoadAndCount", false, 0
				pcs_ValidateTextField "EFBookingConfirmationNumber", false, 0

				'Dangerous Goods
				pcs_ValidateTextField "DGORMD", false, 0
				pcs_ValidateTextField "DGPackagingCount", false, 0
				pcs_ValidateTextField "DGPackagingUnits", false, 0
				pcs_ValidateTextField "DGEmergencyContactNumber", false, 0

				pcs_ValidateTextField "ContainerType", false, 0
				pcs_ValidateTextField "DocumentsOnly", false, 0

				'SMARTPOST
				pcs_ValidateTextField "SMIndicia", false, 0
				pcs_ValidateTextField "SMAncillaryEndorsement", false, 0
				pcs_ValidateTextField "SMHubID", false, 0

				'// Additional Validation for Numerics
				if isNULL(Session("pcAdminDocumentsOnly")) OR Session("pcAdminDocumentsOnly")<>"1" then
					Session("pcAdminDocumentsOnly")="0"
				end if

				if isNULL(Session("pcAdminDGORMD")) OR Session("pcAdminDGORMD")<>"1" then
					Session("pcAdminDGORMD")="0"
				end if

				if isNULL(Session("pcAdminEFPackingListEnclosed")) OR Session("pcAdminEFPackingListEnclosed")<>"1" then
					Session("pcAdminEFPackingListEnclosed")="0"
				end if

				if isNULL(Session("pcAdminResidentialDelivery")) OR Session("pcAdminResidentialDelivery")<>"true" then
					Session("pcAdminResidentialDelivery")="false"
				end if

				if NOT validNum(Session("pcAdminInsideDelivery")) OR Session("pcAdminInsideDelivery")<>"1" then
					Session("pcAdminInsideDelivery")="0"
				end if
				if NOT validNum(Session("pcAdminbSOCODCollection")) OR Session("pcAdminbSOCODCollection")<>"1" then
					Session("pcAdminbSOCODCollection")="0"
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
				if NOT validNum(Session("pcAdminbSOHAL")) OR Session("pcAdminbSOHAL")<>"1" then
					Session("pcAdminbSOHAL")="0"
				end if
				if NOT validNum(Session("pcAdminbSODryIce")) OR Session("pcAdminbSODryIce")<>"1" then
					Session("pcAdminbSODryIce")="0"
				end if
				if NOT validNum(Session("pcAdminbSODGShip")) OR Session("pcAdminbSODGShip")<>"1" then
					Session("pcAdminbSODGShip")="0"
				end if
				if NOT validNum(Session("pcAdminDGAircraftOnly")) OR Session("pcAdminDGAircraftOnly")<>"1" then
					Session("pcAdminDGAircraftOnly")="0"
				end if
				if NOT validNum(Session("pcAdminbISOBrokerSelect")) OR Session("pcAdminbISOBrokerSelect")<>"1" then
					Session("pcAdminbISOBrokerSelect")="0"
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

				'//Smart Post Flag
				mySP = 0

				if Session("pcAdminService1") = "SMART_POST" then
					mySP = 1
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
						objFedExClass.NewXMLLabelWS "ProcessShipmentRequest", FedExkey, FedExPassword, FedExAccountNumber, FedExMeterNumber, "9", "ship"
							'--------------------
							'// TransactionDetail
							'--------------------
							objFedExClass.WriteParent "TransactionDetail", "9", ""
								objFedExClass.AddNewNode "CustomerTransactionId", "9", Session("pcAdminCustomerInvoiceNumber")
							objFedExClass.WriteParent "TransactionDetail", "9", "/"

							'--------------------
							'// Version
							'--------------------
							objFedExClass.WriteParent "Version", "9", ""
								objFedExClass.AddNewNode "ServiceId", "9", "ship"
								objFedExClass.AddNewNode "Major", "9", "9"
								objFedExClass.AddNewNode "Intermediate", "9", "0"
								objFedExClass.AddNewNode "Minor", "9", "0"
							objFedExClass.WriteParent "Version", "9", "/"

							'--------------------
							'// RequestedShipment
							'--------------------
							objFedExClass.WriteParent "RequestedShipment", "9", ""
								objFedExClass.AddNewNode "ShipTimestamp", "9", Session("pcAdminShipDate") & "T09:00:00-00:00"
								objFedExClass.AddNewNode "DropoffType", "9", Session("pcAdminDropoffType")
								objFedExClass.AddNewNode "ServiceType", "9", Session("pcAdminService1")
								objFedExClass.AddNewNode "PackagingType", "9", Session("pcAdminPackaging1")
								'// Off for smartpost
								If mySP=0 Then
									objFedExClass.WriteParent "TotalWeight", "9", ""
										objFedExClass.AddNewNode "Units", "9", Session("pcAdminShipmentWeightUnitS")
										objFedExClass.AddNewNode "Value", "9", Session("pcAdminTotalShipmentWeight")
									objFedExClass.WriteParent "TotalWeight", "9", "/"
								End If

								'--------------------------------
								'// RequestedShipment/Shipper
								'--------------------------------
								objFedExClass.WriteParent "Shipper", "9", ""
									objFedExClass.AddNewNode "AccountNumber", "9", ""
									If Session("pcAdminSenderTINNumber")&""<>"" Then
										objFedExClass.WriteParent "Tins", "9", ""
											objFedExClass.AddNewNode "TinType", "9", Session("pcAdminSenderTINType")
											objFedExClass.AddNewNode "Number", "9", fnStripPhone(Session("pcAdminSenderTINNumber"))
										objFedExClass.WriteParent "Tins", "9", "/"
									End If
									objFedExClass.WriteParent "Contact", "9", ""
										'objFedExClass.AddNewNode "ContactId", "RBB1057"
										objFedExClass.AddNewNode "PersonName", "9", sanitizeField(Session("pcAdminOriginPersonName"))
										objFedExClass.AddNewNode "CompanyName", "9", sanitizeField(Session("pcAdminOriginCompanyName"))
										objFedExClass.AddNewNode "PhoneNumber", "9", fnStripPhone(Session("pcAdminOriginPhoneNumber"))
										objFedExClass.AddNewNode "PagerNumber", "9", fnStripPhone(Session("pcAdminOriginPagerNumber"))
										objFedExClass.AddNewNode "FaxNumber", "9", fnStripPhone(Session("pcAdminOriginFaxNumber"))
										objFedExClass.AddNewNode "EMailAddress", "9", Session("pcAdminOriginEmailAddress")
									objFedExClass.WriteParent "Contact", "9", "/"
									objFedExClass.WriteParent "Address", "9", ""
										objFedExClass.AddNewNode "StreetLines", "9", Session("pcAdminOriginLine1")
										objFedExClass.AddNewNode "StreetLines", "9", Session("pcAdminOriginLine2")
										objFedExClass.AddNewNode "City", "9", Session("pcAdminOriginCity")
										objFedExClass.AddNewNode "StateOrProvinceCode", "9", Session("pcAdminOriginStateOrProvinceCode")
										objFedExClass.AddNewNode "PostalCode", "9", Session("pcAdminOriginPostalCode")

										objFedExClass.AddNewNode "CountryCode", "9", Session("pcAdminOriginCountryCode")
										objFedExClass.AddNewNode "Residential", "9", "false"
									objFedExClass.WriteParent "Address", "9", "/"
								objFedExClass.WriteParent "Shipper", "9", "/"

								'--------------------------------
								'// RequestedShipment/Recipient
								'--------------------------------
								objFedExClass.WriteParent "Recipient", "9", ""
									objFedExClass.AddNewNode "AccountNumber", "9", ""
									objFedExClass.WriteParent "Contact", "9", ""
										objFedExClass.AddNewNode "PersonName", "9", sanitizeField(Session("pcAdminRecipPersonName"))
										objFedExClass.AddNewNode "CompanyName", "9", sanitizeField(Session("pcAdminRecipCompanyName"))
										objFedExClass.AddNewNode "PhoneNumber", "9", fnStripPhone(Session("pcAdminRecipPhoneNumber"))
										objFedExClass.AddNewNode "PagerNumber", "9", fnStripPhone(Session("pcAdminRecipPagerNumber"))
										objFedExClass.AddNewNode "FaxNumber", "9", fnStripPhone(Session("pcAdminRecipFaxNumber"))
										objFedExClass.AddNewNode "EMailAddress", "9", Session("pcAdminRecipEmailAddress")
									objFedExClass.WriteParent "Contact", "9", "/"
									objFedExClass.WriteParent "Address", "9", ""
										objFedExClass.AddNewNode "StreetLines", "9", Session("pcAdminRecipLine1")
										objFedExClass.AddNewNode "StreetLines", "9", Session("pcAdminRecipLine2")
										objFedExClass.AddNewNode "City", "9", Session("pcAdminRecipCity")
										if Session("pcAdminRecipCountryCode")="US" OR Session("pcAdminRecipCountryCode")="CA" then
											objFedExClass.AddNewNode "StateOrProvinceCode", "9", Session("pcAdminRecipStateOrProvinceCode")
										end if

										objFedExClass.AddNewNode "PostalCode", "9", Session("pcAdminRecipPostalCode")
										objFedExClass.AddNewNode "CountryCode", "9", Session("pcAdminRecipCountryCode") '"US"
										objFedExClass.AddNewNode "Residential", "9", Session("pcAdminResidentialDelivery")
									objFedExClass.WriteParent "Address", "9", "/"
								objFedExClass.WriteParent "Recipient", "9", "/"

								'--------------------------------------------
								'// RequestedShipment/ShippingChargesPayment
								'--------------------------------------------
								objFedExClass.WriteParent "ShippingChargesPayment", "9", ""
									objFedExClass.AddNewNode "PaymentType", "9", Session("pcAdminPayorType")
									objFedExClass.WriteParent "Payor", "9", ""
										objFedExClass.AddNewNode "AccountNumber", "9", Session("pcAdminPayorAccountNumber")
										objFedExClass.AddNewNode "CountryCode", "9", Session("pcAdminPayorCountryCode")
									objFedExClass.WriteParent "Payor", "9", "/"
								objFedExClass.WriteParent "ShippingChargesPayment", "9", "/"

								'---------------------------------------------
								'// RequestedShipment/SpecialServicesRequested
								'---------------------------------------------
								BSO = 0
								if Session("pcAdminbISOBrokerSelect")<>"0" then
									BSO = 1
								end if

								ENS = 0

								If ENS = 1 OR BSO = 1 OR Session("pcAdminReturnShipmentIndicator")="PRINTRETURNLABEL" OR (DateDiff("d", FedExDateFormat(Date()), Session("pcAdminShipDate")) >= 1) OR Session("pcAdminSaturdayDelivery")<>"0" OR Session("pcAdminSaturdayPickup")<>"0" OR Session("pcAdminInsideDelivery")<>"0" OR Session("pcAdminSOHAL")<>"0" OR Session("pcAdminDeliveryType")&""<>"" OR (Session("pcAdminbSOCODCollection")<>"0" AND Session("pcAdminService1")<>"FEDEX_GROUND") Then
									objFedExClass.WriteParent "SpecialServicesRequested", "9", ""
										If BSO = 1 Then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "BROKER_SELECT_OPTION"
										End If

										if Session("pcAdminReturnShipmentIndicator")="PRINTRETURNLABEL" then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "RETURN_SHIPMENT"

											objFedExClass.WriteParent "ReturnShipmentDetail", "9", ""
												objFedExClass.AddNewNode "ReturnType", "9", "PRINT_RETURN_LABEL"
												objFedExClass.WriteParent "Rma", "9", ""
													objFedExClass.AddNewNode "Number", "9", "UATtest123"
													objFedExClass.AddNewNode "Reason", "9", "Inoperable"
												objFedExClass.WriteParent "Rma", "9", "/"
											objFedExClass.WriteParent "ReturnShipmentDetail", "9", "/"
										end if

										'// Future Day Shipment
										if (DateDiff("d", FedExDateFormat(Date()), Session("pcAdminShipDate")) >= 1) then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "FUTURE_DAY_SHIPMENT"
										end if


										'// Saturday Services
										if  Session("pcAdminSaturdayDelivery")<>"0" then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "SATURDAY_DELIVERY"
										end if
										if  Session("pcAdminSaturdayPickup")<>"0" then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "SATURDAY_PICKUP"
										end if
										'// Inside Delivery
										if Session("pcAdminInsideDelivery")<>"0" then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "INSIDE_DELIVERY"
										end if

										pcTempNotificationShipperEmail = 0
										pcTempNotificationRecipientEmail = 0
										pcTempNotificationOtherEmail = 0
										ENS = 0

										'Session("pcAdminShipperNotificationFormat")
										If (Session("pcAdminNotificationShipperEmail")&""<>"" AND (Session("pcAdminShipperShipmentNotification") = "1" OR Session("pcAdminShipperDeliveryNotification") = "1" OR Session("pcAdminShipperExceptionNotification") = "1")) Then
											pcTempNotificationShipperEmail = 1
											ENS = 1
										End If
										If (Session("pcAdminNotificationRecipientEmail")&""<>"" AND (Session("pcAdminRecipientShipmentNotification") = "1" OR Session("pcAdminRecipientDeliveryNotification") = "1" OR Session("pcAdminRecipientExceptionNotification") = "1")) Then
											pcTempNotificationRecipientEmail = 1
											ENS = 1
										End If
										If (Session("pcAdminOtherNotification1")&""<>"" AND (Session("pcAdminOtherShipmentNotification") = "1" OR Session("pcAdminOtherDeliveryNotification") = "1" OR Session("pcAdminOtherExceptionNotification") = "1")) Then
											pcTempNotificationOtherEmail = 1
											ENS = 1
										End If

										if ENS = 1 then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "EMAIL_NOTIFICATION"
											objFedExClass.WriteParent "EMailNotificationDetail", "9", ""
												'objFedExClass.AddNewNode "PersonalMessage", "9", "Personal Message Details"
												If pcTempNotificationShipperEmail = 1 Then
													objFedExClass.WriteParent "Recipients", "9", ""
														objFedExClass.AddNewNode "EMailNotificationRecipientType", "9", "SHIPPER"
														objFedExClass.AddNewNode "EMailAddress", "9", Session("pcAdminNotificationShipperEmail")
														objFedExClass.AddNewNode "NotifyOnShipment", "9", Session("pcAdminShipperShipmentNotification")
														objFedExClass.AddNewNode "NotifyOnException", "9", Session("pcAdminShipperExceptionNotification")
														objFedExClass.AddNewNode "NotifyOnDelivery", "9", Session("pcAdminShipperDeliveryNotification")
														objFedExClass.AddNewNode "Format", "9", Session("pcAdminShipperNotificationFormat") 'HTML/TEXT/WIRELESS
														objFedExClass.WriteParent "Localization", "9", ""
															objFedExClass.AddNewNode "LanguageCode", "9", "EN"
														objFedExClass.WriteParent "Localization", "9", "/"
													objFedExClass.WriteParent "Recipients", "9", "/"
												End If
												If pcTempNotificationShipperEmail = 1 Then
													objFedExClass.WriteParent "Recipients", "9", ""
														objFedExClass.AddNewNode "EMailNotificationRecipientType", "9", "RECIPIENT"
														objFedExClass.AddNewNode "EMailAddress", "9", Session("pcAdminNotificationRecipientEmail")
														objFedExClass.AddNewNode "NotifyOnShipment", "9", Session("pcAdminRecipientShipmentNotification")
														objFedExClass.AddNewNode "NotifyOnException", "9", Session("pcAdminRecipientExceptionNotification")
														objFedExClass.AddNewNode "NotifyOnDelivery", "9", Session("pcAdminRecipientDeliveryNotification")
														objFedExClass.AddNewNode "Format", "9", Session("pcAdminShipperNotificationFormat") 'HTML/TEXT/WIRELESS
														objFedExClass.WriteParent "Localization", "9", ""
															objFedExClass.AddNewNode "LanguageCode", "9", "EN"
														objFedExClass.WriteParent "Localization", "9", "/"
													objFedExClass.WriteParent "Recipients", "9", "/"
												End If
												If pcTempNotificationOtherEmail = 1 Then
													objFedExClass.WriteParent "Recipients", "9", ""
														objFedExClass.AddNewNode "EMailNotificationRecipientType", "9", "OTHER"
														objFedExClass.AddNewNode "EMailAddress", "9", Session("pcAdminOtherNotification1")
														objFedExClass.AddNewNode "NotifyOnShipment", "9", Session("pcAdminOtherShipmentNotification")
														objFedExClass.AddNewNode "NotifyOnException", "9", Session("pcAdminOtherExceptionNotification")
														objFedExClass.AddNewNode "NotifyOnDelivery", "9", Session("pcAdminOtherDeliveryNotification")
														objFedExClass.AddNewNode "Format", "9", Session("pcAdminShipperNotificationFormat") 'HTML/TEXT/WIRELESS
														objFedExClass.WriteParent "Localization", "9", ""
															objFedExClass.AddNewNode "LanguageCode", "9", "EN"
														objFedExClass.WriteParent "Localization", "9", "/"
													objFedExClass.WriteParent "Recipients", "9", "/"
												End If
											objFedExClass.WriteParent "EMailNotificationDetail", "9", "/"
										End If

										'// Hold At Location
										if Session("pcAdminbSOHAL")<>"0" then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "HOLD_AT_LOCATION"

											objFedExClass.WriteParent "HoldAtLocationDetail", "9", ""
												objFedExClass.AddNewNode "PhoneNumber", "9", fnStripPhone(Session("pcAdminHALPhone"))
												objFedExClass.WriteParent "LocationContactAndAddress", "9", ""
													objFedExClass.WriteParent "Contact", "9", ""
														objFedExClass.AddNewNode "PersonName", "9", Session("pcAdminHALPersonName")
														objFedExClass.AddNewNode "CompanyName", "9", Session("pcAdminHALCompanyName")
														objFedExClass.AddNewNode "PhoneNumber", "9", fnStripPhone(Session("pcAdminHALPhone"))
													objFedExClass.WriteParent "Contact", "9", "/"
													objFedExClass.WriteParent "Address", "9", ""
														objFedExClass.AddNewNode "StreetLines", "9", Session("pcAdminHALLine1")
														objFedExClass.AddNewNode "City", "9", Session("pcAdminHALCity")
														objFedExClass.AddNewNode "StateOrProvinceCode", "9", Session("pcAdminHALStateOrProvinceCode")
														objFedExClass.AddNewNode "PostalCode", "9", Session("pcAdminHALPostalCode")
														objFedExClass.AddNewNode "CountryCode", "9", Session("pcAdminHALCountryCode")
													objFedExClass.WriteParent "Address", "9", "/"
												objFedExClass.WriteParent "LocationContactAndAddress", "9", "/"
											objFedExClass.WriteParent "HoldAtLocationDetail", "9", "/"
										end if

										'// Home Delivery Premium
										if Session("pcAdminDeliveryType")&""<>"" Then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "HOME_DELIVERY_PREMIUM"
											objFedExClass.WriteParent "HomeDeliveryPremiumDetail", "9", ""
												objFedExClass.AddNewNode "HomeDeliveryPremiumType", "9", Session("pcAdminDeliveryType")
												objFedExClass.AddNewNode "Date", "9", Session("pcAdminDeliveryDate")
												objFedExClass.AddNewNode "PhoneNumber", "9", Session("pcAdminDeliveryPhone")
											objFedExClass.WriteParent "HomeDeliveryPremiumDetail", "9", "/"
										end if

										'//COD DETAILS
										if Session("pcAdminbSOCODCollection")<>"0" AND Session("pcAdminService1")<>"FEDEX_GROUND" then
											objFedExClass.AddNewNode "SpecialServiceTypes", "9", "COD"


											objFedExClass.WriteParent "CodDetail", "9", ""
												objFedExClass.WriteParent "CodCollectionAmount", "9", ""
													objFedExClass.AddNewNode "Currency", "9", "USD"
													objFedExClass.AddNewNode "Amount", "9", Session("pcAdminCODAmount")
												objFedExClass.WriteParent "CodCollectionAmount", "9", "/"
												objFedExClass.AddNewNode "CollectionType", "9", Session("pcAdminCODType")
												objFedExClass.WriteParent "CodRecipient", "9", ""
													objFedExClass.WriteParent "Tins", "9", ""
														objFedExClass.AddNewNode "TinType", "9", Session("pcAdminCODTinType")
														objFedExClass.AddNewNode "Number", "9", Session("pcAdminCODTinNumber")
													objFedExClass.WriteParent "Tins", "9", "/"
													objFedExClass.WriteParent "Contact", "9", ""
														objFedExClass.AddNewNode "PersonName", "9", Session("pcAdminCODPersonName")
														objFedExClass.AddNewNode "Title", "9", Session("pcAdminCODTitle")
														objFedExClass.AddNewNode "CompanyName", "9", Session("pcAdminCODCompanyName")
														objFedExClass.AddNewNode "PhoneNumber", "9", Session("pcAdminCODPhoneNumber")
													objFedExClass.WriteParent "Contact", "9", "/"
													objFedExClass.WriteParent "Address", "9", ""
														objFedExClass.AddNewNode "StreetLines", "9", Session("pcAdminCODStreetLines")
														objFedExClass.AddNewNode "City", "9", Session("pcAdminCODCity")
														objFedExClass.AddNewNode "StateOrProvinceCode", "9", Session("pcAdminCODState")
														objFedExClass.AddNewNode "PostalCode", "9", Session("pcAdminCODPostalCode")
														objFedExClass.AddNewNode "CountryCode", "9", Session("pcAdminCODCountryCode")
													objFedExClass.WriteParent "Address", "9", "/"
												objFedExClass.WriteParent "CodRecipient", "9", "/"
											objFedExClass.WriteParent "CodDetail", "9", "/"

										end if

									objFedExClass.WriteParent "SpecialServicesRequested", "9", "/"
								End If

								'---------------------------------------------
								'// SMARTPOST
								'---------------------------------------------
								if mySP = 1 Then
									objFedExClass.WriteParent "SmartPostDetail", "9", ""
										objFedExClass.AddNewNode "Indicia", "9", Session("pcAdminSMIndicia")
										objFedExClass.AddNewNode "AncillaryEndorsement", "9", Session("pcAdminSMAncillaryEndorsement")
										objFedExClass.AddNewNode "HubId", "9", Session("pcAdminSMHubID")
									objFedExClass.WriteParent "SmartPostDetail", "9", "/"
								end if
								'---------------------------------------------
								'// RequestedShipment/CustomsClearanceDetail
								'---------------------------------------------
'
								If Session("pcAdminEFShippersLoadAndCount")&"" <>"" Then
									objFedExClass.WriteParent "ExpressFreightDetail", "9", ""
										objFedExClass.AddNewNode "PackingListEnclosed", "9", Session("pcAdminEFPackingListEnclosed")
										objFedExClass.AddNewNode "ShippersLoadAndCount", "9", Session("pcAdminEFShippersLoadAndCount")
										objFedExClass.AddNewNode "BookingConfirmationNumber", "9", Session("pcAdminEFBookingConfirmationNumber")
									objFedExClass.WriteParent "ExpressFreightDetail", "9", "/"
								End If

								CCD = 0

								If BSO = 1 OR Session("pcAdminRCIdValue")&""<>"" OR Session("pcAdminDutiesAccountNumber")&"" <> "" OR Session("pcAdminCVAmount")&"" <> "" OR Session("pcAdminCICAmount")&"" <> "" OR Session("pcAdminCFCAmount")&"" <> "" OR Session("pcAdminNumberOfPieces")>0 OR Session("pcAdminB13AFilingOption")&"" <> "" Then
									CCD = 1
								End if

								if CCD = 1 then
									objFedExClass.WriteParent "CustomsClearanceDetail", "9", ""
										If BSO = 1 Then
											objFedExClass.WriteParent "Broker", "9", ""
												objFedExClass.WriteParent "Contact", "9", ""
													objFedExClass.AddNewNode "CompanyName", "9", Session("pcAdminBSOCompanyName")
													objFedExClass.AddNewNode "PhoneNumber", "9", Session("pcAdminBSOPhoneNumber")
												objFedExClass.WriteParent "Contact", "9", "/"
												objFedExClass.WriteParent "Address", "9", ""
													objFedExClass.AddNewNode "StreetLines", "9", Session("pcAdminBSOStreetLines")
													objFedExClass.AddNewNode "City", "9", Session("pcAdminBSOCity")
													objFedExClass.AddNewNode "StateOrProvinceCode", "9", Session("pcAdminBSOStateOrProvinceCode")
													objFedExClass.AddNewNode "PostalCode", "9", Session("pcAdminBSOPostalCode")
													objFedExClass.AddNewNode "CountryCode", "9", Session("pcAdminBSOCountryCode")
												objFedExClass.WriteParent "Address", "9", "/"
											objFedExClass.WriteParent "Broker", "9", "/"
										End If

										If Session("pcAdminRCIdValue")&""<>"" Then
											objFedExClass.WriteParent "RecipientCustomsId", "9", ""
												objFedExClass.AddNewNode "Type", "9", Session("pcAdminRCIdType")
												objFedExClass.AddNewNode "Value", "9", Session("pcAdminRCIdValue")
											objFedExClass.WriteParent "RecipientCustomsId", "9", "/"
										End If

										If Session("pcAdminDutiesAccountNumber")&"" <> "" Then
											objFedExClass.WriteParent "DutiesPayment", "9", ""
												objFedExClass.AddNewNode "PaymentType", "9", Session("pcAdminDutiesPayorType")
												objFedExClass.WriteParent "Payor", "9", ""
													objFedExClass.AddNewNode "AccountNumber", "9", Session("pcAdminDutiesAccountNumber")
													objFedExClass.AddNewNode "CountryCode", "9", Session("pcAdminDutiesCountryCode")
												objFedExClass.WriteParent "Payor", "9", "/"
											objFedExClass.WriteParent "DutiesPayment", "9", "/"
										End If

										If Session("pcAdminDocumentsOnly") <>"0" Then
											if Session("pcAdminDocumentsOnly")="1" Then
												objFedExClass.AddNewNode "DocumentContent", "9", "NON_DOCUMENTS"
											else
												objFedExClass.AddNewNode "DocumentContent", "9", "DOCUMENTS_ONLY"
											end if
										End If

										If Session("pcAdminCVAmount")&"" <> "" Then
											objFedExClass.WriteParent "CustomsValue", "9", ""
												objFedExClass.AddNewNode "Currency", "9", Session("pcAdminCVCurrency")
												objFedExClass.AddNewNode "Amount", "9", Session("pcAdminCVAmount")
											objFedExClass.WriteParent "CustomsValue", "9", "/"
										End If

										If Session("pcAdminCICAmount")&"" <> "" Then
											objFedExClass.WriteParent "InsuranceCharges", "9", ""
												objFedExClass.AddNewNode "Currency", "9", Session("pcAdminCVCurrency")
												objFedExClass.AddNewNode "Amount", "9", Session("pcAdminCICAmount")
											objFedExClass.WriteParent "InsuranceCharges", "9", "/"
										End If

										If Session("pcAdminCFCAmount")&"" <> "" Then
											objFedExClass.WriteParent "CommercialInvoice", "9", ""
												objFedExClass.AddNewNode "Comments", "9", Session("pcAdminCCIComments")
												objFedExClass.WriteParent "FreightCharge", "9", ""
													objFedExClass.AddNewNode "Currency", "9", Session("pcAdminCVCurrency")
													objFedExClass.AddNewNode "Amount", "9", Session("pcAdminCFCAmount")
												objFedExClass.WriteParent "FreightCharge", "9", "/"
												objFedExClass.WriteParent "TaxesOrMiscellaneousCharge", "9", ""
													objFedExClass.AddNewNode "Currency", "9", Session("pcAdminCVCurrency")
													objFedExClass.AddNewNode "Amount", "9", Session("pcAdminCMCAmount")
												objFedExClass.WriteParent "TaxesOrMiscellaneousCharge", "9", "/"
												objFedExClass.AddNewNode "Purpose", "9", Session("pcAdminCCIPurpose")
												objFedExClass.AddNewNode "CustomerInvoiceNumber", "9", Session("pcAdminCCIInvoiceNumber")
											objFedExClass.WriteParent "CommercialInvoice", "9", "/"
										End If

										If Session("pcAdminNumberOfPieces")>0 Then
											objFedExClass.WriteParent "Commodities", "9", ""
												objFedExClass.AddNewNode "NumberOfPieces", "9", Session("pcAdminNumberOfPieces")
												objFedExClass.AddNewNode "Description", "9", Session("pcAdminDescription")
												objFedExClass.AddNewNode "CountryOfManufacture", "9", Session("pcAdminCountryOfManufacture")
												objFedExClass.WriteParent "Weight", "9", ""
													objFedExClass.AddNewNode "Units", "9", Session("pcAdminShipmentWeightUnitS")
													objFedExClass.AddNewNode "Value", "9", Session("pcAdminCommodityWeight")
												objFedExClass.WriteParent "Weight", "9", "/"
												objFedExClass.AddNewNode "Quantity", "9", Session("pcAdminCommodityQuantity")
												objFedExClass.AddNewNode "QuantityUnits", "9", Session("pcAdminCommodityQuantityUnits")

												objFedExClass.WriteParent "UnitPrice", "9", ""
													objFedExClass.AddNewNode "Currency", "9", "USD"
													objFedExClass.AddNewNode "Amount", "9", Session("pcAdminCommodityUnitPrice")
												objFedExClass.WriteParent "UnitPrice", "9", "/"
												objFedExClass.WriteParent "CustomsValue", "9", ""
													objFedExClass.AddNewNode "Currency", "9", "USD"
													objFedExClass.AddNewNode "Amount", "9", Session("pcAdminCommodityCustomsValue")
												objFedExClass.WriteParent "CustomsValue", "9", "/"
											objFedExClass.WriteParent "Commodities", "9", "/"
										End If


										If Session("pcAdminB13AFilingOption")&"" <> "" Then
											objFedExClass.WriteParent "ExportDetail", "9", ""
												objFedExClass.AddNewNode "B13AFilingOption", "9", Session("pcAdminB13AFilingOption")
												objFedExClass.AddNewNode "ExportComplianceStatement", "9", Session("pcAdminExportComplianceStatement")
											objFedExClass.WriteParent "ExportDetail", "9", "/"
										End If

									objFedExClass.WriteParent "CustomsClearanceDetail", "9", "/"
								end if
								'---------------------------------------------
								'// RequestedShipment/LabelSpecification
								'---------------------------------------------
								objFedExClass.WriteParent "LabelSpecification", "9", ""
									objFedExClass.AddNewNode "LabelFormatType", "9", Session("pcAdminType")
									objFedExClass.AddNewNode "ImageType", "9", "PNG"
									objFedExClass.AddNewNode "LabelStockType", "9", "PAPER_LETTER"
								objFedExClass.WriteParent "LabelSpecification", "9", "/"

								objFedExClass.WriteSingleParent "RateRequestTypes", "9",  Session("pcAdminRateRequestType")

								'---------------------------------------------
								'// RequestedShipment/MasterTrackingId
								'---------------------------------------------
								if cint(pcPackageCount) > 1 then
									objFedExClass.WriteParent "MasterTrackingId", "9", ""
										if pcv_xCounter>1 then
											'// Required for multiple-piece shipping if PackageSequenceNumber value is greater than one.
											objFedExClass.AddNewNode "TrackingNumber", "9", Session("MasterTrackingNumber")

										end if
									objFedExClass.WriteParent "MasterTrackingId", "9", "/"
								end if

								'-------------------------------------------------
								'// RequestedShipment/PackageCount
								'-------------------------------------------------
								objFedExClass.WriteSingleParent "PackageCount", "9", pcPackageCount

								'-------------------------------------------------
								'// RequestedShipment/PackageDetail
								'-------------------------------------------------
								objFedExClass.AddNewNode "PackageDetail", "9", "INDIVIDUAL_PACKAGES"

								'-------------------------------------------------
								'// RequestedShipment/RequestedPackageLineItems
								'-------------------------------------------------
								objFedExClass.WriteParent "RequestedPackageLineItems", "9", ""
									objFedExClass.AddNewNode "SequenceNumber", "9", pcv_xCounter
									IF mySP = 0 THEN
										objFedExClass.WriteParent "InsuredValue", "9", ""
											objFedExClass.AddNewNode "Currency", "9", "USD"
											objFedExClass.AddNewNode "Amount", "9", Session("pcAdmindeclaredvalue"&pcv_xCounter)
										objFedExClass.WriteParent "InsuredValue", "9", "/"
									END IF
									objFedExClass.WriteParent "Weight", "9", ""
										pcvTempWeightUnit = ""
										if Session("pcAdminWeightUnits"&pcv_xCounter) = "LB" then
											pcvTempWeightUnit = "LB"
										else
											pcvTempWeightUnit = "KG"
										end if
										objFedExClass.AddNewNode "Units", "9", pcvTempWeightUnit '"LB"
										IF mySP = 0 THEN
											objFedExClass.AddNewNode "Value", "9", FormatNumber(Session("pcAdminWeight"&pcv_xCounter),1)
										Else
											objFedExClass.AddNewNode "Value", "9", Session("pcAdminWeight"&pcv_xCounter)
										End If
									objFedExClass.WriteParent "Weight", "9", "/"

									If TRIM(Session("pcAdminPackaging1"))="YOUR_PACKAGING" AND TRIM(Session("pcAdminService"&pcv_xCounter))<>"INTERNATIONAL_PRIORITY" then
										objFedExClass.WriteParent "Dimensions", "9", ""
											objFedExClass.AddNewNode "Length", "9", Session("pcAdminLength"&pcv_xCounter) '"12"
											objFedExClass.AddNewNode "Width", "9", Session("pcAdminWidth"&pcv_xCounter) '"13"
											objFedExClass.AddNewNode "Height", "9", Session("pcAdminHeight"&pcv_xCounter) '"14"
											objFedExClass.AddNewNode "Units", "9", Session("pcAdminUnits"&pcv_xCounter) '"IN"
										objFedExClass.WriteParent "Dimensions", "9", "/"
									End If

									' START: CUSTOMER REFERENCES
									objFedExClass.WriteParent "CustomerReferences", "9", ""
										objFedExClass.AddNewNode "CustomerReferenceType", "9", "CUSTOMER_REFERENCE"
										objFedExClass.AddNewNode "Value", "9", Session("pcAdminCustomerReference")
									objFedExClass.WriteParent "CustomerReferences", "9", "/"
									If Session("pcAdminCustomerInvoiceNumber")&""<>"" Then
									objFedExClass.WriteParent "CustomerReferences", "9", ""
										objFedExClass.AddNewNode "CustomerReferenceType", "9", "INVOICE_NUMBER"
										objFedExClass.AddNewNode "Value", "9", Session("pcAdminCustomerInvoiceNumber")
									objFedExClass.WriteParent "CustomerReferences", "9", "/"
									End If
									If Session("pcAdminCustomerPONumber")&""<>"" Then
									objFedExClass.WriteParent "CustomerReferences", "9", ""
										objFedExClass.AddNewNode "CustomerReferenceType", "9", "P_O_NUMBER"
										objFedExClass.AddNewNode "Value", "9", Session("pcAdminCustomerPONumber")
									objFedExClass.WriteParent "CustomerReferences", "9", "/"
									End If
									' START: SPECIAL SERVICES
									DG=0
									if Session("pcAdminbSODGShip")<>"0" then
										DG=1
									end if

									if DG=1 OR (Session("pcAdminbSOCODCollection")<>"0" AND Session("pcAdminService"&pcv_xCounter)="FEDEX_GROUND") OR Session("pcAdminbSODGShip")<>"0" OR Session("pcAdminbSODryIce")="1" OR Session("pcAdminContainerType")="1" OR Session("pcAdminPriorityAlert")&""<>"" OR Session("pcAdminSignatureOption")&""<>"" then
										objFedExClass.WriteParent "SpecialServicesRequested", "9", ""

											if Session("pcAdminbSOCODCollection")<>"0" AND Session("pcAdminService"&pcv_xCounter)="FEDEX_GROUND" then
												objFedExClass.AddNewNode "SpecialServiceTypes", "9", "COD"
											end if

											IF DG=1 THEN
												objFedExClass.AddNewNode "SpecialServiceTypes", "9", "DANGEROUS_GOODS"
											END IF

											if Session("pcAdminbSODryIce")="1" then
												objFedExClass.AddNewNode "SpecialServiceTypes", "9", "DRY_ICE"
											end if

											If Session("pcAdminContainerType")="1" Then
												objFedExClass.AddNewNode "SpecialServiceTypes", "9", "NON_STANDARD_CONTAINER"
											End If

											'//priority alert
											if Session("pcAdminPriorityAlert")&""<>"" then
												objFedExClass.AddNewNode "SpecialServiceTypes", "9", "PRIORITY_ALERT"
											end if
											if Session("pcAdminSignatureOption")&""<>"" then
												objFedExClass.AddNewNode "SpecialServiceTypes", "9", "SIGNATURE_OPTION"
											end if

											'//COD FOR FEDEX GROUND
											if Session("pcAdminbSOCODCollection")<>"0" AND Session("pcAdminService"&pcv_xCounter)="FEDEX_GROUND" then
												objFedExClass.WriteParent "CodDetail", "9", ""
													objFedExClass.WriteParent "CodCollectionAmount", "9", ""
														objFedExClass.AddNewNode "Currency", "9", "USD" 'Session("asdf")
														objFedExClass.AddNewNode "Amount", "9", Session("pcAdminCODAmount")
													objFedExClass.WriteParent "CodCollectionAmount", "9", "/"
													objFedExClass.AddNewNode "CollectionType", "9", Session("pcAdminCODType")
													objFedExClass.WriteParent "CodRecipient", "9", ""
														objFedExClass.WriteParent "Tins", "9", ""
															objFedExClass.AddNewNode "TinType", "9", Session("pcAdminCODTinType")
															objFedExClass.AddNewNode "Number", "9", Session("pcAdminCODTinNumber")
														objFedExClass.WriteParent "Tins", "9", "/"
														objFedExClass.WriteParent "Contact", "9", ""
															objFedExClass.AddNewNode "PersonName", "9", Session("pcAdminCODPersonName")
															objFedExClass.AddNewNode "Title", "9", Session("pcAdminCODTitle")
															objFedExClass.AddNewNode "CompanyName", "9", Session("pcAdminCODCompanyName")
															objFedExClass.AddNewNode "PhoneNumber", "9", Session("pcAdminCODPhoneNumber")
														objFedExClass.WriteParent "Contact", "9", "/"
														objFedExClass.WriteParent "Address", "9", ""
															objFedExClass.AddNewNode "StreetLines", "9", Session("pcAdminCODStreetLines")
															objFedExClass.AddNewNode "City", "9", Session("pcAdminCODCity")
															objFedExClass.AddNewNode "StateOrProvinceCode", "9", Session("pcAdminCODState")
															objFedExClass.AddNewNode "PostalCode", "9", Session("pcAdminCODPostalCode")
															objFedExClass.AddNewNode "CountryCode", "9", Session("pcAdminCODCountryCode")
															objFedExClass.AddNewNode "Residential", "9", "false"
														objFedExClass.WriteParent "Address", "9", "/"
													objFedExClass.WriteParent "CodRecipient", "9", "/"
													objFedExClass.AddNewNode "ReferenceIndicator", "9", "INVOICE"
												objFedExClass.WriteParent "CodDetail", "9", "/"
											end if

											'// Dangerous Goods
											IF dg=1 THEN
												objFedExClass.WriteParent "DangerousGoodsDetail", "9", ""
													If session("pcAdminDGAccessibility")&""<>"" Then
														objFedExClass.AddNewNode "Accessibility", "9", session("pcAdminDGAccessibility")
														objFedExClass.AddNewNode "CargoAircraftOnly", "9", session("pcAdminDGAircraftOnly")
													End If
													If session("pcAdminDGORMD")="1" Then
														objFedExClass.AddNewNode "Options", "9", "ORM_D"
													End If
													If session("pcAdminDGPackagingUnits")&""<>"" Then
													objFedExClass.WriteParent "Packaging", "9", ""
														objFedExClass.AddNewNode "Count", "9", session("pcAdminDGPackagingCount")
														objFedExClass.AddNewNode "Units", "9", session("pcAdminDGPackagingUnits")
													objFedExClass.WriteParent "Packaging", "9", "/"
													End If
													If session("pcAdminDGEmergencyContactNumber")&""<>"" Then
														objFedExClass.AddNewNode "EmergencyContactNumber", "9", session("pcAdminDGEmergencyContactNumber")
													End If
												objFedExClass.WriteParent "DangerousGoodsDetail", "9", "/"
											END IF

											'// Dry Ice Shipment
											if Session("pcAdminbSODryIce")="1" then
												objFedExClass.WriteParent "DryIceWeight", "9", ""
													objFedExClass.AddNewNode "Units", "9", Session("pcAdminSDIUnit")
													objFedExClass.AddNewNode "Value", "9", Session("pcAdminSDIValue")
												objFedExClass.WriteParent "DryIceWeight", "9", "/"
											end if

											'//SignatureOption
											if Session("pcAdminSignatureOption")&""<>"" then
												objFedExClass.WriteParent "SignatureOptionDetail", "9", ""
													objFedExClass.AddNewNode "OptionType", "9", Session("pcAdminSignatureOption")
												objFedExClass.WriteParent "SignatureOptionDetail", "9", "/"

											end if

											'<xs:element name="PriorityAlertDetail" type="ns:PriorityAlertDetail" minOccurs="0">
										objFedExClass.WriteParent "SpecialServicesRequested", "9", "/"
									End If

								objFedExClass.WriteParent "RequestedPackageLineItems", "9", "/"
							objFedExClass.WriteParent "RequestedShipment", "9", "/"
						objFedExClass.EndXMLTransaction "ProcessShipmentRequest", "9"

						strLogID= Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
						'response.write fedex_postdataWS&"<HR>"
						'response.end
						'// Log our Transaction
						call objFedExClass.pcs_LogTransaction(fedex_postdataWS, sanitizeField(Session("pcAdminRecipPersonName"))&"_Req_"& strLogID &".txt", true)

						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Send Our Transaction.
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						Set srvFEDEXWSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
						Set objOutputXMLDocWS = Server.CreateObject("Microsoft.XMLDOM")
						Set objFedExStream = Server.CreateObject("ADODB.Stream")
						Set objFEDEXXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
						objFEDEXXmlDoc.async = False
						objFEDEXXmlDoc.validateOnParse = False
						if err.number>0 then
							err.clear
						end if

						srvFEDEXWSXmlHttp.open "POST", FedExWSURL&"/ship", false


						srvFEDEXWSXmlHttp.send(fedex_postdataWS)
						FEDEXWS_result = srvFEDEXWSXmlHttp.responseText

						'response.write FEDEXWS_result
						'response.end
						'// Log our Response
						call objFedExClass.pcs_LogTransaction(FEDEXWS_result, sanitizeField(Session("pcAdminRecipPersonName"))&"_Res_"& strLogID &".txt", true)
						'// Print out our response

						if trim(FEDEXWS_result)="" then
							response.redirect ErrPageName & "?msg=FedEx was unable to send a response. There may have been a connection error. Please try again."
						end if

						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Load Our Response.
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						call objFedExClass.LoadXMLResults(FEDEXWS_result)
						objOutputXMLDocWS.loadXML FEDEXWS_result


						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Check for errors from FedEx.
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						if pcv_xCounter = 1 then
							'// master package error, no processing done
							pcv_strErrorMsg = objFedExClass.ReadResponseNode("//v9:ProcessShipmentReply", "v9:Notifications/v9:Severity")
							strTmpLabelImage = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:Label/v9:Parts/v9:Image")

							if pcv_strErrorMsg="SUCCESS" OR pcv_strErrorMsg="WARNING" OR (pcv_strErrorMsg="NOTE" AND strTmpLabelImage&""<>"") then
								pcv_strErrorMsg = Cstr("")
							else
								pcv_strErrorMsg = objFedExClass.ReadResponseNode("//v9:ProcessShipmentReply", "v9:Notifications/v9:Message")
							end if

							if pcv_strErrorMsg&""="" then
								pcv_strErrorMsg = objFedExClass.ReadResponseNode("//soapenv:Fault", "faultstring")
								pcv_isFault = "&fault="&sanitizeField(Session("pcAdminRecipPersonName"))&"_Res_"& strLogID &".txt"
							end if

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
								response.redirect ErrPageName & "?msg=Your shipment was not processed for the following reason. " & pcv_strErrorMsg & pcv_isFault
							else
								pcLocalArray(pcv_xCounter-1) = "shipped"
								pcv_strItemsList = join(pcLocalArray, chr(124))
								Session("pcGlobalArray") = pcv_strItemsList
								'/////////////////////////////////////////////////////////////
								'// POSTBACK LOGGING
								'/////////////////////////////////////////////////////////////
								'// Tracking Number for Logs
								pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:TrackingIds/v9:TrackingNumber")

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
								pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:TrackingIds/v9:TrackingNumber")
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
							'//Notifications
							pcv_NotificationSeverity = objFedExClass.ReadResponseNode("//v9:Notifications", "v9:Severity")
							pcv_NotificationSource = objFedExClass.ReadResponseNode("//v9:Notifications", "v9:Source")
							pcv_NotificationCode = objFedExClass.ReadResponseNode("//v9:Notifications", "v9:Code")
							pcv_NotificationMessage = objFedExClass.ReadResponseNode("//v9:Notifications", "v9:Message")
							pcv_NotificationLocalizedMessage = objFedExClass.ReadResponseNode("//v9:Notifications", "v9:LocalizedMessage")

							pcv_CustomerTransactionId = objFedExClass.ReadResponseNode("//v9:TransactionDetail", "v9:CustomerTransactionId")

							pcv_VersionServiceId = objFedExClass.ReadResponseNode("//v9:Version", "v9:ServiceId")
							pcv_VersionMajor = objFedExClass.ReadResponseNode("//v9:Version", "v9:Major")
							pcv_VersionIntermediate = objFedExClass.ReadResponseNode("//v9:Version", "v9:Intermediate")
							pcv_VersionMinor = objFedExClass.ReadResponseNode("//v9:Version", "v9:Minor")

							pcv_UsDomestic = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:UsDomestic")
							pcv_CarrierCode = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:CarrierCode")
							'//if multi-piece shipment get master tracking id
							session("MasterTrackingIdType") = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:MasterTrackingId/v9:TrackingIdType")
							session("MasterFormId") = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:MasterTrackingId/v9:FormId")
							session("MasterTrackingNumber") = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:MasterTrackingId/v9:TrackingNumber")

							pcv_ServiceTypeDescription = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:ServiceTypeDescription")
							pcv_PackagingDescription = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:PackagingDescription")

							pcv_ShipmentOriginLocationNumber = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:OperationalDetail/v9:OriginLocationNumber")
							pcv_ShipmentDestinationLocationNumber = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:OperationalDetail/v9:DestinationLocationNumber")
							pcv_ShipmentTransitTime = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:OperationalDetail/v9:TransitTime")
							pcv_ShipmentCustomTransitTime = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:OperationalDetail/v9:CustomTransitTime")
							pcv_ShipmentIneligibleForMoneyBackGuarantee = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:OperationalDetail/v9:IneligibleForMoneyBackGuarantee")
							pcv_ShipmentDeliveryEligibilities = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:OperationalDetail/v9:DeliveryEligibilities")
							pcv_ShipmentServiceCode = objFedExClass.ReadResponseNode("//v9:CompletedShipmentDetail", "v9:OperationalDetail/v9:ServiceCode")

							pcv_ShipmentRateType = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:RateType")
							pcv_ShipmentRateZone = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:RateZone")
							pcv_ShipmentRatedWeightMethod = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:RatedWeightMethod")
							pcv_ShipmentDimDivisor = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:DimDivisor")
							pcv_ShipmentFuelSurchargePercent = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:FuelSurchargePercent")
							pcv_ShipmentTotalBillingWeightUnits = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalBillingWeight/v9:Units")
							pcv_ShipmentTotalBillingWeightValue = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalBillingWeight/v9:Value")

							pcv_ShipmentTotalBaseChargeCurrency = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalBaseCharge/v9:Currency")
							pcv_ShipmentTotalBaseChargeAmount = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalBaseCharge/v9:Amount")
							pcv_ShipmentTotalFreightDiscountsCurrency = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalFreightDiscounts/v9:Currency")
							pcv_ShipmentTotalFreightDiscountsAmount = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalFreightDiscounts/v9:Amount")
							pcv_ShipmentTotalNetFreightCurrency = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalNetFreight/v9:Currency")
							pcv_ShipmentTotalNetFreightAmount = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalNetFreight/v9:Amount")
							pcv_ShipmentTotalSurchargesCurrency = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalSurcharges/v9:Currency")
							pcv_ShipmentTotalSurchargesAmount = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalSurcharges/v9:Amount")
							pcv_ShipmentTotalNetFedExChargeCurrency = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalNetFedExCharge/v9:Currency")
							pcv_ShipmentTotalNetFedExChargeAmount = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalNetFedExCharge/v9:Amount")
							pcv_ShipmentTotalTaxesCurrency = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalTaxes/v9:Currency")
							pcv_ShipmentTotalTaxesAmount = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalTaxes/v9:Amount")
							pcv_ShipmentTotalNetChargeCurrency = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalNetCharge/v9:Currency")
							pcv_ShipmentTotalNetChargeAmount = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalNetCharge/v9:Amount")
							pcv_ShipmentTotalRebatesCurrency = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalRebates/v9:Currency")
							pcv_ShipmentTotalRebatesAmount = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:TotalRebates/v9:Amount")

							pcv_ShipmentSurchargesType = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:Surcharges/v9:SurchargeType")
							pcv_ShipmentSurchargesLevel = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:Surcharges/v9:Level")
							pcv_ShipmentSurchargesDesc = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:Surcharges/v9:Description")
							pcv_ShipmentSurchargesCurrency = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:Surcharges/v9:Amount/v9:Currency")
							pcv_ShipmentSurchargesAmount = objFedExClass.ReadResponseNode("//v9:ShipmentRateDetails", "v9:Surcharges/v9:Amount/v9:Amount")

							pcv_SequenceNumber = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:SequenceNumber")

							pcv_TrackingIdType = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:TrackingIds/v9:TrackingIdType")
							pcv_TrackingNumber = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:TrackingIds/v9:TrackingNumber")

							pcv_GroupNumber = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:GroupNumber")

							pcv_PackageRateType = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:RateType")
							pcv_PackageRatedWeightMethod = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:RatedWeightMethod")
							pcv_PackageBillingWeightUnit = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:BillingWeight/v9:Units")
							pcv_PackageBillingWeightValue = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:BillingWeight/v9:Value")
							pcv_PackageBaseChargeCurrency = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:BaseCharge/v9:Currency")
							pcv_PackageBaseChargeAmount = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:BaseCharge/v9:Amount")
							pcv_PackageTotalFreightDiscountsCurrency = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:TotalFreightDiscounts/v9:Currency")
							pcv_PackageTotalFreightDiscountsAmount = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:TotalFreightDiscounts/v9:Amount")
							pcv_PackageNetFreightCurrency = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:NetFreight/v9:Currency")
							pcv_PackageNetFreightAmount = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:NetFreight/v9:Amount")
							pcv_PackageTotalSurchargesCurrency = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:TotalSurcharges/v9:Currency")
							pcv_PackageTotalSurcharges = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:TotalSurcharges/v9:Amount")
							pcv_PackageNetFedExChargeCurrency = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:NetFedExCharge/v9:Currency")
							pcv_PackageNetFedExChargeAmount = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:NetFedExCharge/v9:Amount")
							pcv_PackageTotalTaxesCurrency = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:TotalTaxes/v9:Currency")
							pcv_PackageTotalTaxesAmount = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:TotalTaxes/v9:Amount")
							pcv_PackageNetChargeCurrency = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:NetCharge/v9:Currency")
							pcv_PackageNetChargeAmount = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:NetCharge/v9:Amount")
							pcv_PackageTotalRebatesCurrency = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:TotalRebates/v9:Currency")
							pcv_PackageTotalRebatesAmount = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:TotalRebates/v9:Amount")
							pcv_PackageSurchargesType = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:Surcharges/v9:SurchargeType")
							pcv_PackageSurchargesLevel = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:Surcharges/v9:Level")
							pcv_PackageSurchargesDesc = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:Surcharges/v9:Description")
							pcv_PackageSurchargesCurrency = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:Surcharges/v9:Amount/v9:Currency")
							pcv_PackageSurchargesAmount = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:PackageRating/v9:PackageRateDetails/v9:Surcharges/v9:Amount/v9:Amount")

							pcv_BarcodesType = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:OperationalDetail/v9:Barcodes/v9:BinaryBarcodes/v9:Type")
							pcv_BarcodesValue = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:OperationalDetail/v9:Barcodes/v9:BinaryBarcodes/v9:Value")
							pcv_StringBarcodesType = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:OperationalDetail/v9:Barcodes/v9:StringBarcodes/v9:Type")
							pcv_StringBarcodesValue = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:OperationalDetail/v9:Barcodes/v9:StringBarcodes/v9:Value")
							pcv_GroundServiceCode = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:OperationalDetail/v9:GroundServiceCode")

							pcv_LabelType = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:Label/v9:Type")
							pcv_LabelShippingDocumentDisposition = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:Label/v9:ShippingDocumentDisposition")
							pcv_LabelResolution = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:Label/v9:Resolution")
							pcv_LabelCopiesToPrint = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:Label/v9:CopiesToPrint")
							pcv_LabelDocumentPartSequenceNumber = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:Label/v9:Parts/v9:DocumentPartSequenceNumber")
							pcv_LabelImage = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:Label/v9:Parts/v9:Image")

							pcv_SignatureOption = objFedExClass.ReadResponseNode("//v9:CompletedPackageDetails", "v9:SignatureOption")



							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' START: SAVE LABEL
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							if pcv_LabelImage <> "" then
								'// Create XML for Label
								GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName=""Label"&pcv_TrackingNumber&".PNG"">"&pcv_LabelImage&"</Base64Data>"

								'// Load label from the request stream
								objFEDEXXmlDoc.loadXML GraphicXML

								'// Use ADO stream to save the binary data
								objFedExStream.Type = 1
								objFedExStream.Open

								objFedExStream.Write objFEDEXXmlDoc.selectSingleNode("/Base64Data").nodeTypedValue
									err.clear
								strFileName = objFEDEXXmlDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue
								'Save the binary stream to the file and overwrite if it already exists in folder
								objFedExStream.SaveToFile server.MapPath("FedExLabels\"&strFileName),2
								objFedExStream.Close()

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
							if pcv_PackageNetChargeAmount = "" then
								pcv_PackageNetChargeAmount = 0
							end if

							If Session("pcAdminResidentialDelivery")="true" Then
								Session("pcAdminResidentialDelivery") = 1
							End If
							If Session("pcAdminResidentialDelivery")="false" Then
								Session("pcAdminResidentialDelivery")=0
							End if

							query=		"UPDATE pcPackageInfo "
							query=query&"SET pcPackageInfo_FDXSPODFlag=0, "
							query=query&"pcPackageInfo_PackageNumber=1, "
							query=query&"pcPackageInfo_PackageWeight=" & Session("pcAdminWeight"&pcv_xCounter) & ", "
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
							query=query&"pcPackageInfo_FDXRate=" & pcv_PackageNetChargeAmount & " "
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
								%>
								<!--#include file="../includes/GoogleCheckout_OrderManagement.asp"-->
							<% End If
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
							response.redirect "FedExWS_ManageShipmentsResults.asp?id=" & pcv_intOrderID & "&msg=Your transaction has been completed successfully."
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
				'pcPackageCount = 1
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
							<li><a id="tabs5" onclick="setpackagedivs();change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '');change('tabs5', 'current');showTab('tab5')">Package Information</a></li>
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
									<th colspan="2">Billing Detail</th>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Payor:</b></td>
									<td align="left">
										<select name="PayorType" id="PayorType">
										<option value="SENDER" <%=pcf_SelectOption("PayorType","SENDER")%>>Sender</option>
										<option value="RECIPIENT" <%=pcf_SelectOption("PayorType","RECIPIENT")%>>Recipient</option>
										<option value="THIRD_PARTY" <%=pcf_SelectOption("PayorType","THIRDPARTY")%>>3rd Party</option>
										<option value="COLLECT" <%=pcf_SelectOption("PayorType","COLLECT")%>>Collect</option>
										</select>
										<%pcs_RequiredImageTag "PayorType", true%></td>
								</tr>
								<tr>
								<td align="right" valign="top"><b>Payor Account Number:</b></td>
								<td align="left">
								<input name="PayorAccountNumber" type="text" id="PayorAccountNumber" value="<%=FedExAccountNumber%>"><%pcs_RequiredImageTag "PayorAccountNumber", false%>
								  </td>
								</tr>
								<tr>
								<td align="right" valign="top"><b>Payor Country Code:</b></td>
								<td align="left">
								<input name="PayorCountryCode" type="text" id="PayorCountryCode" value="US">
								<%pcs_RequiredImageTag "PayorCountryCode", false%>
								e.g.	US	</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Rate Settings</th>
								</tr>
								<tr>
									<td colspan="2">
									<span class="pcCPnotes"> LIST rates retrieve FedEx's list rates and
									ACCOUNT rates retrieve account-specific rates (including any applicable discounts)                          </span>
									</td>
								</tr>
								<tr>
									<td width="24%" align="right" valign="top"><b>Rate Request Type:</b></td>
									<td width="76%" align="left">
									<select name="RateRequestType" id="RateRequestType">
										<option value="LIST" <%=pcf_SelectOption("RateRequestType","LIST")%>>LIST</option>
										<option value="ACCOUNT" <%=pcf_SelectOption("RateRequestType","ACCOUNT")%>>ACCOUNT</option>
									</select>
									</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>

						<tr>
							<th colspan="2">Service Settings  </th>
						</tr>
						<tr>
									<td width="24%" align="right" valign="top"><b>Type of service:</b></td>
									<td width="76%" align="left">
									<script type="text/javascript">
									function setcarriercodedivs() {
									   var div_num = $("#CarrierCode").val();
									   if (div_num == 1) {
									   $("#groundsettings").hide();
										};
									   if (div_num ==2) {
									   $("#groundsettings").show();
										};
									}
									</script>
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
									<select name="CarrierCode" id="CarrierCode" size="1" onchange="setcarriercodedivs();">
								<option value="1" <%=pcf_SelectOption("CarrierCode","FDXE")%>>FedEx Express</option>
								<option value="2" <%=pcf_SelectOption("CarrierCode","FDXG")%>>FedEx Ground</option>
							</select>
							<%pcs_RequiredImageTag "CarrierCode", true %>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Drop off Type:</b></td>
								<td align="left">
								<select name="DropoffType" id="DropoffType">
									<option value="REGULAR_PICKUP" <%=pcf_SelectOption("DropoffType","REGULAR_PICKUP")%>>Regular Pickup</option>
									<option value="REQUEST_COURIER" <%=pcf_SelectOption("DropoffType","REQUEST_COURIER")%>>Courier Pickup</option>
									<option value="DROP_BOX" <%=pcf_SelectOption("DropoffType","DROP_BOX")%>>FedEx Express Drop Box</option>
									<option value="BUSINESS_SERVICE_CENTER" <%=pcf_SelectOption("DropoffType","BUSINESS_SERVICE_CENTER")%>>Business Service Center</option>
									<option value="STATION" <%=pcf_SelectOption("DropoffType","STATION")%>>FedEx Station</option>
								</select><%pcs_RequiredImageTag "DropoffType", isRequiredDropoffType %>
							</td>
						</tr>
						<tr>
									<td align="right" valign="top"><b>Shipment Type:</b></td>
							<td align="left">
								<select name="ReturnShipmentIndicator" id="ReturnShipmentIndicator">
								<option value="NON_RETURN" <%=pcf_SelectOption("ReturnShipmentIndicator","NON_RETURN")%>>Outgoing Shipment</option>
								<option value="PRINT_RETURN_LABEL" <%=pcf_SelectOption("ReturnShipmentIndicator","PRINT_RETURN_LABEL")%>>Return Shipment</option>
								</select>
								<%pcs_RequiredImageTag "ReturnShipmentIndicator", false%>
								</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Total Shipment Weight  </th>
						</tr>
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
								intMPackageWeight=cLng(intMPackageWeight)+1
							End if
							End if
						if request("test")<>"" then
							response.write intMPackageWeight
							response.End()
						end if

						If session("pcAdminTotalShipmentWeight")&""="" Then
							session("pcAdminTotalShipmentWeight") = intMPackageWeight
						End If
						%>

						<tr>
								  <td colspan="2" align="right" valign="top">&nbsp;</td>
							  </tr>
								<tr>
									<td width="24%" align="right" valign="top"><b>Weight Unit:</b></td>
								  <td width="76%" align="left">
									<select name="ShipmentWeightUnits" id="ShipmentWeightUnits">
									  <option value="LB" <%=pcf_SelectOption("ShipmentWeightUnits","LB")%>>LB</option>
									  <option value="KG" <%=pcf_SelectOption("ShipmentWeightUnits","KG")%>>KG</option>
									</select>&nbsp;&nbsp;&nbsp;<b>Weight:&nbsp;
									<input name="TotalShipmentWeight" type="text" id="TotalShipmentWeight" value="<%=pcf_FillFormField("TotalShipmentWeight", false)%>">
									</b></td>
								</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
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
									&nbsp;&nbsp;&nbsp;<b>Ship Time:&nbsp;
									<input name="ShipTime" type="text" id="ShipTime" value="<%=pcf_FillFormField("ShipTime", true)%>"><%pcs_RequiredImageTag "ShipTime", true%>
									* hh:mm:ss </b></td>
							  </tr>
							</table>
							</div>
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
			<tr>
						<td class="pcCPshipping" colspan="2"><span class="titleShip">Additional Settings</span></td>
					</tr>
					<tr>
						<td colspan="2">
							<table width="100%">
								<tr>
									<th colspan="2">
									<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
									<!--
									function jfSOSaturdayServices(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSOSaturdayServices.checked == true) {
									document.getElementById('SOSaturdayServices').style.display='';
									}else{
									document.getElementById('SOSaturdayServices').style.display='none';
									}
									}
									 //-->
									</SCRIPT>
									<%
									if Session("pcAdminbSOSaturdayServices")="true" then
										pcv_strDisplayStyle="style=""display:visible"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOSaturdayServices();" name="bSOSaturdayServices" id="bSOSaturdayServices" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOSaturdayServices", "1")%>>
									Saturday Services </th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOSaturdayServices" <%=pcv_strDisplayStyle%>>
											<Table>
											  <TR>
												<TD width="51" height="20">&nbsp;</TD>
												<TD width="303"><INPUT tabIndex="25" type="checkbox" value="1" name="SaturdayDelivery" class="clearBorder" <%=pcf_CheckOption("SaturdayDelivery", "1")%>>
												&nbsp;Saturday Delivery</TD>
											  </TR>
											  <TR>
												<TD height="20">&nbsp;</TD>
												<TD height="20"><input tabindex="25" type="checkbox" value="1" name="SaturdayPickup" class="clearBorder" <%=pcf_CheckOption("SaturdayPickup", "1")%>>
&nbsp;Saturday Pickup</TD>
											  </TR>
											</Table>
										</div>
									</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">
									<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
									<!--
									function jfSOSignatureOption(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSOSignatureOption.checked == true) {
									document.getElementById('SOSignatureOption').style.display='';
									}else{
									document.getElementById('SOSignatureOption').style.display='none';
									}
									}
									 //-->
									</SCRIPT>
									<%
									if Session("pcAdminbSOSignatureOption")="true" then
										pcv_strDisplayStyle="style=""display:visible"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOSignatureOption();" name="bSOSignatureOption" id="bSOSignatureOption" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOSignatureOption", "1")%>>
									Signature Options </th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOSignatureOption" <%=pcv_strDisplayStyle%>>
											<Table>
						<tr>
						<td align="right" valign="top"><b>Signature Type:</b></td>
							<td align="left">
								<select name="SignatureOption" id="SignatureOption">
									<option value="" <%=pcf_SelectOption("SignatureOption","")%>>No Signature Options</option>
									<option value="SERVICE_DEFAULT" <%=pcf_SelectOption("SignatureOption","SERVICE_DEFAULT")%>>Service default</option>

									<option value="NO_SIGNATURE_REQUIRED" <%=pcf_SelectOption("SignatureOption","NO_SIGNATURE_REQUIRED")%>>No signature required</option>
									<option value="INDIRECT" <%=pcf_SelectOption("SignatureOption","INDIRECT")%>>Indirect</option>
									<option value="DIRECT" <%=pcf_SelectOption("SignatureOption","DIRECT")%>>Direct</option>
									<option value="ADULT" <%=pcf_SelectOption("SignatureOption","ADULT")%>>Adult Signature Required</option>
								</select>
								<%pcs_RequiredImageTag "SignatureOption", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Signature Release:</b></td>
							<td align="left">
								<INPUT type="text" name="SignatureRelease" id="SignatureRelease" value="<%=pcf_FillFormField("SignatureRelease", false)%>">
								<%pcs_RequiredImageTag "SignatureRelease", false%>
								(Deliver Without Signature Only)
							</td>
						</tr>
											</Table>
										</div>
									</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th align="left"><input type="checkbox" name="PriorityAlert" value="1" class="clearBorder" <%=pcf_CheckOption("PriorityAlert", "1")%>><%pcs_RequiredImageTag "PriorityAlert", false%>&nbsp;FedEx Priority Alert®</th>
									<td align="left">

									</td>
								</tr>
						<tr>
									<td colspan="2" class="pcCPspacer"></td>
							  </tr>
								<tr>
									<th colspan="2">
									<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
									<!--
									function jfSOCODCollection(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSOCODCollection.checked == true) {
									document.getElementById('SOCODCollection').style.display='';
									}else{
									document.getElementById('SOCODCollection').style.display='none';
									}
									}
									 //-->
									</SCRIPT>
									<%
									if Session("pcAdminbSOCODCollection")="true" then
										pcv_strDisplayStyle="style=""display:visible"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOCODCollection();" name="bSOCODCollection" id="bSOCODCollection" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOCODCollection", "1")%>>
									COD Collection </th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOCODCollection" <%=pcv_strDisplayStyle%>>
											<Table>
						<tr>
							<td align="right"><b>Collection Amount:</b></td>
							<td align="left">
								<INPUT type="text" name="CODAmount" id="CODAmount" value="<%=pcf_FillFormField("CODAmount", false)%>">
								<%pcs_RequiredImageTag "CODAmount", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Collection Type:</b></td>
							<td align="left">
								<select name="CODType" id="CODType">
									<option value="CASH" <%=pcf_SelectOption("CODType","CASH")%>>Cash</option>
									<option value="COMPANY_CHECK" <%=pcf_SelectOption("CODType","COMPANY_CHECK")%>>Company Check</option>

									<option value="GUARANTEED_FUNDS" <%=pcf_SelectOption("CODType","GUARANTEED_FUNDS")%>>Guaranteed Funds</option>
									<option value="PERSONAL_CHECK" <%=pcf_SelectOption("CODType","PERSONAL_CHECK")%>>Personal Check</option>
								</select>
								<%pcs_RequiredImageTag "CODType", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Tin Type:</b></td>
							<td align="left">
								<select name="CODTinType" id="CODTinType">
									<option value="BUSINESS_NATIONAL" <%=pcf_SelectOption("CODTinType","BUSINESS_NATIONAL")%>>Business National</option>
									<option value="BUSINESS_STATE" <%=pcf_SelectOption("CODTinType","BUSINESS_STATE")%>>Business State</option>
									<option value="BUSINESS_UNION" <%=pcf_SelectOption("CODTinType","BUSINESS_UNION")%>>Business Union</option>
									<option value="PERSONAL_NATIONAL" <%=pcf_SelectOption("CODTinType","PERSONAL_NATIONAL")%>>Personal National</option>
									<option value="PERSONAL_STATE" <%=pcf_SelectOption("CODTinType","PERSONAL_STATE")%>>Personal State</option>
								</select>
								<%pcs_RequiredImageTag "CODTinType", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Tin Number:</b></td>
							<td align="left">
								<INPUT type="text" name="CODTinNumber" id="CODTinNumber" value="<%=pcf_FillFormField("CODTinNumber", false)%>">
								<%pcs_RequiredImageTag "CODTinNumber", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b> Contact Name:</b></td>
							<td align="left">
								<INPUT type="text" name="CODPersonName" id="CODPersonName" value="<%=pcf_FillFormField("CODPersonName", false)%>">
								<%pcs_RequiredImageTag "CODPersonName", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Company Name:</b></td>
							<td align="left">
								<INPUT type="text" name="CODCompanyName" id="CODCompanyName" value="<%=pcf_FillFormField("CODCompanyName", false)%>">
								<%pcs_RequiredImageTag "CODCompanyName", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Phone Number:</b></td>
							<td align="left">
								<INPUT type="text" name="CODPhoneNumber" id="CODPhoneNumber" value="<%=pcf_FillFormField("CODPhoneNumber", false)%>">
								<%pcs_RequiredImageTag "CODPhoneNumber", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Title:</b></td>
							<td align="left">
								<INPUT type="text" name="CODTitle" id="CODTitle" value="<%=pcf_FillFormField("CODTitle", false)%>">
								<%pcs_RequiredImageTag "CODTitle", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Street Lines:</b></td>
							<td align="left">
								<INPUT type="text" name="CODStreetLines" id="CODStreetLines" value="<%=pcf_FillFormField("CODStreetLines", false)%>">
								<%pcs_RequiredImageTag "CODStreetLines", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b> City:</b></td>
							<td align="left">
								<INPUT type="text" name="CODCity" id="CODCity" value="<%=pcf_FillFormField("CODCity", false)%>">
								<%pcs_RequiredImageTag "CODCity", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>State:</b></td>
							<td align="left">
								<INPUT type="text" name="CODState" id="CODState" value="<%=pcf_FillFormField("CODState", false)%>">
								<%pcs_RequiredImageTag "CODState", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Postal Code:</b></td>
							<td align="left">
								<INPUT type="text" name="CODPostalCode" id="CODPostalCode" value="<%=pcf_FillFormField("CODPostalCode", false)%>">
								<%pcs_RequiredImageTag "CODPostalCode", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Country Code:</b></td>
							<td align="left">
								<INPUT type="text" name="CODCountryCode" id="CODCountryCode" value="<%=pcf_FillFormField("CODCountryCode", false)%>">
								<%pcs_RequiredImageTag "CODCountryCode", false%>
							</td>
						</tr>
											</Table>
										</div>
									</td>
								</tr>

								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
								  <th colspan="2">
									<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
									<!--
									function jfSOHAL(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSOHAL.checked == true) {
									document.getElementById('SOHAL').style.display='';
									}else{
									document.getElementById('SOHAL').style.display='none';
									}
									}
									 //-->
									</SCRIPT>
									<%
									if Session("pcAdminbSOHAL")="1" then
										pcv_strDisplayStyle="style=""display:visible"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSOHAL();" name="bSOHAL" id="bSOHAL" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSOHAL", "1")%>>
									Hold At Location </th>
								</tr>
								<tr>
									<td colspan="2">
										<div id="SOHAL" <%=pcv_strDisplayStyle%>>
											<Table>
						<tr>
							<td align="right"><b>Contact Name:</b></td>
							<td align="left">
								<INPUT type="text" name="HALPersonName" id="HALPersonName" value="<%=pcf_FillFormField("HALPersonName", false)%>">
								<%pcs_RequiredImageTag "HALPersonName", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Company Name:</b></td>
							<td align="left">
								<INPUT type="text" name="HALCompanyName" id="HALCompanyName" value="<%=pcf_FillFormField("HALCompanyName", false)%>">
								<%pcs_RequiredImageTag "HALCompanyName", false%>
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
							<td align="right"><b>Country Code:</b></td>
							<td align="left">
								<INPUT name="HALCountryCode" type="text" id="HALCountryCode" value="<%=pcf_FillFormField("HALCountryCode", false)%>">
								<%pcs_RequiredImageTag "HALCountryCode", false%>
							</td>

						</tr>
						<tr>
							<td align="right"><b>Postal Code:</b></td>
							<td align="left">
								<INPUT name="HALPostalCode" type="text" id="HALPostalCode" value="<%=pcf_FillFormField("HALPostalCode", false)%>">
								<%pcs_RequiredImageTag "HALPostalCode", false%>
							</td>
											  </tr>
											</Table>
										</div>
									</td>
						</tr>
						<tr>
									<td colspan="2" class="pcCPspacer"></td>
							  </tr>
								<tr>
									<th colspan="2">
									<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
									<!--
									function jfSODGShip(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSODGShip.checked == true) {
									document.getElementById('SODGShip').style.display='';
									}else{
									document.getElementById('SODGShip').style.display='none';
									}
									}
									 //-->
									</SCRIPT>
									<%
									if Session("pcAdminbSODGShip")="true" then
										pcv_strDisplayStyle="style=""display:visible"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSODGShip();" name="bSODGShip" id="bSODGShip" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bSODGShip", "1")%>>
									Dangerous Goods Shipment </th>
								</tr>
								<tr>
									<td colspan="2">
									<div id="SODGShip" <%=pcv_strDisplayStyle%>>
										<Table>
						<tr>
							<td align="right"><b>Accessibility:</b></td>
							<td align="left">
								<INPUT type="text" name="DGAccessibility" id="DGAccessibility" value="<%=pcf_FillFormField("DGAccessibility", false)%>">
								<%pcs_RequiredImageTag "DGAccessibility", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Cargo Aircraft Only:</b></td>
							<td align="left">
								<INPUT type="radio" name="DGAircraftOnly" id="DGAircraftOnly" value="1" class="clearBorder">Yes
								<INPUT type="radio" name="DGAircraftOnly" id="DGAircraftOnly" value="0" class="clearBorder">No
							</td>
						</tr>
						<tr>
							<td align="right"><b>ORM-D?</b></td>
							<td align="left">
								<INPUT type="radio" name="DGORMD" id="DGORMD" value="1" class="clearBorder">Yes
								<INPUT type="radio" name="DGORMD" id="DGORMD" value="0" class="clearBorder">No
							</td>
						</tr>
						<tr>
							<td align="right"><b>Package Count:</b></td>
							<td align="left">
								<INPUT type="text" name="DGPackagingCount" id="DGPackagingCount" value="<%=pcf_FillFormField("DGPackagingCount", false)%>">
								<%pcs_RequiredImageTag "DGPackagingCount", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Package Units:</b></td>
							<td align="left">
								<INPUT type="text" name="DGPackagingUnits" id="DGPackagingUnits" value="<%=pcf_FillFormField("DGPackagingUnits", false)%>">
								<%pcs_RequiredImageTag "DGPackagingUnits", false%>
							</td>
						</tr>
						<tr>
							<td align="right"><b>Emergency Contact:</b></td>
							<td align="left">
								<INPUT type="text" name="DGEmergencyContactNumber" id="DGEmergencyContactNumber" value="<%=pcf_FillFormField("DGEmergencyContactNumber", false)%>">
								<%pcs_RequiredImageTag "DGEmergencyContactNumber", false%>
							</td>
						</tr>
										</Table>
									</div>
									</td>
								</tr>

								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">
									<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
									<!--
									function jfSODryIce(){

									var selectValDom = document.forms['form1'];
									if (selectValDom.bSODryIce.checked == true) {
									document.getElementById('SODryIce').style.display='';
									}else{
									document.getElementById('SODryIce').style.display='none';
									}
									}
									 //-->
									</SCRIPT>
									<%
									if Session("pcAdminbSODryIce")="true" then
										pcv_strDisplayStyle="style=""display:visible"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfSODryIce();" name="bSODryIce" id="bSODryIce" type="checkbox" class="clearBorder" value=1 <%=pcf_CheckOption("bSODryIce", "1")%>>
									Dry Ice Shipment </th>
								</tr>
								<tr>
									<td colspan="2">
									<div id="SODryIce" <%=pcv_strDisplayStyle%>>
										<Table>
											<tr>
												<td width="273" align="right"><b>Dry Ice Package Count:</b></td>
												<td width="345" align="left">
								<INPUT type="text" name="SDIPackageCount" id="SDIPackageCount" value="<%=pcf_FillFormField("SDIPackageCount", false)%>">
								<%pcs_RequiredImageTag "SDIPackageCount", false%>
							</td>
						</tr>
						<tr>
												<td align="right"><b>Dry Ice Weight:</b></td>
												<td align="left">
													<INPUT type="text" name="SDIValue" id="SDIValue" value="<%=pcf_FillFormField("SDIValue", false)%>">
													<%pcs_RequiredImageTag "SDIValue", false%>
												</td>
										  </tr>
											<tr>
												<td align="right"><b>Weight Units:</b></td>
												<td align="left">
													<INPUT type="text" name="SDIUnit" id="SDIUnit" value="<%=pcf_FillFormField("SDIUnit", false)%>">
													<%pcs_RequiredImageTag "SDIUnit", false%>
												(KG or LB)</td>
											</tr>
										</Table>
									</div>
									</td>
								</tr>
							</table>
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
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td width="24%" align="right" valign="top"><b>Documents  Shipment:</b></td>
							<td width="76%" align="left">
							<input type="radio" name="DocumentsOnly" value="0" class="clearBorder" checked>Not Applicable
							<input type="radio" name="DocumentsOnly" value="1" class="clearBorder">Non-Documents
							<input type="radio" name="DocumentsOnly" value="2" class="clearBorder">Documents Only
							</td>
						</tr>


						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Customs Clearance Detail</th>
						</tr>
						<tr>
							<td align="right"><b>Customs Amount:</b></td>
							<td align="left">
								<INPUT type="text" name="CVAmount" id="CVAmount" value="<%=pcf_FillFormField("CVAmount", false)%>">
								<%pcs_RequiredImageTag "CVAmount", isRequiredCVAmount %>
							</td>
						</tr>
						<tr>
							<td width="25%" align="right" valign="top"><b>Customs Currency Unit:</b></td>
							<td width="75%" align="left">
								<INPUT type="text" name="CVCurrency" id="CVCurrency" value="<%=pcf_FillFormField("CVCurrency", false)%>">
								<%pcs_RequiredImageTag "CVCurrency", isRequiredCVCurrency%>
							</td>
						</tr>
						<tr>
						  <td align="right"><b>Insurance Charges Amount:</b></td>
						  <td align="left">
							<INPUT type="text" name="CICAmount" id="CICAmount" value="<%=pcf_FillFormField("CICAmount", false)%>">
							<%pcs_RequiredImageTag "CICAmount", false%>
							</td>
						  </tr>
						<tr>
						  <td align="right"><b>Taxes or Miscellaneous Charges Amount:</b></td>
						  <td align="left"><INPUT type="text" name="CMCAmount" id="CMCAmount" value="<%=pcf_FillFormField("CMCAmount", false)%>">
							<%pcs_RequiredImageTag "CMCAmount", false%></td>
						  </tr>
						<tr>
							<td align="right"><b>Freight Charges Amount:</b></td>
							<td align="left"><INPUT type="text" name="CFCAmount" id="CFCAmount" value="<%=pcf_FillFormField("CFCAmount", false)%>">
							<%pcs_RequiredImageTag "CFCAmount", false%></td>
						</tr>
						<tr>
							<td align="right"><b>Purpose:</b></td>
							<td align="left">
							 <select name="CCIPurpose" id="CCIPurpose">
								<option value="NOT_SOLD" <%=pcf_SelectOption("CCIPurpose","NOT_SOLD")%>>Not Sold</option>
								<option value="PERSONAL_EFFECTS" <%=pcf_SelectOption("CCIPurpose","PERSONAL_EFFECTS")%>>Personal Effects</option>
								<option value="REPAIR_AND_RETURN" <%=pcf_SelectOption("CCIPurpose","REPAIR_AND_RETURN")%>>Repair and Return</option>
								<option value="SAMPLE" <%=pcf_SelectOption("CCIPurpose","SAMPLE")%>>Sample</option>
								<option value="SOLD" <%=pcf_SelectOption("CCIPurpose","SOLD")%>>Sold</option>
								</select>
								<%pcs_RequiredImageTag "CCIPurpose", false%>
						</tr>
						<tr>
							<td align="right"><b>Invoice Number:</b></td>
							<td align="left"><INPUT type="text" name="CCIInvoiceNumber" id="CCIInvoiceNumber" value="<%=pcf_FillFormField("CCIInvoiceNumber", false)%>">
							<%pcs_RequiredImageTag "CCIInvoiceNumber", false%></td>
						</tr>
						<tr>
							<td align="right"><b>Comments:</b></td>
							<td align="left"><INPUT type="text" name="CCIComments" id="CCIComments" value="<%=pcf_FillFormField("CCIComments", false)%>">
							<%pcs_RequiredImageTag "CCIComments", false%></td>
						</tr>
						<tr>
						  <td colspan="2" class="pcCPspacer"></td>
						</tr>

						<tr>
							<th colspan="2">Commodity</th>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Number of Pieces:</b></td>
							<td align="left">
								<input name="NumberOfPieces" type="text" id="NumberOfPieces" value="<%=pcf_FillFormField("NumberOfPieces", false)%>">
								<%pcs_RequiredImageTag "NumberOfPieces", isRequiredNumberOfPieces %>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Description:</b></td>
							<td align="left">
								<input name="Description" type="text" id="Description" value="<%=pcf_FillFormField("Description", false)%>">
								<%pcs_RequiredImageTag "Description", isRequiredDescription %>
							</td>
						</tr>



						<tr>
							<td align="right" valign="top"><b>Country Code of Manufacture:</b></td>
							<td align="left">
								<input name="CountryOfManufacture" type="text" id="CountryOfManufacture" value="<%=pcf_FillFormField("CountryOfManufacture", false)%>">
								<%pcs_RequiredImageTag "CountryOfManufacture", isRequiredCountryOfManufacture %>
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
								<%pcs_RequiredImageTag "CommodityWeight", isRequiredCommodityWeight %>
							(e.g. 3.0) </td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Quantity:</b></td>
							<td align="left">
								<input name="CommodityQuantity" type="text" id="CommodityQuantity" value="<%=pcf_FillFormField("CommodityQuantity", false)%>">
								<%pcs_RequiredImageTag "CommodityQuantity", isRequiredCommodityQuantity %>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Quantity Units:</b></td>
							<td align="left">
								<input name="CommodityQuantityUnits" type="text" id="CommodityQuantityUnits" value="<%=pcf_FillFormField("CommodityQuantityUnits", false)%>">
								<%pcs_RequiredImageTag "CommodityQuantityUnits", isRequiredCommodityQuantityUnits %>
							(e.g. EA) </td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Unit Price:</b></td>
							<td align="left">
								<input name="CommodityUnitPrice" type="text" id="CommodityUnitPrice" value="<%=pcf_FillFormField("CommodityUnitPrice", false)%>">
								<%pcs_RequiredImageTag "CommodityUnitPrice", isRequiredCommodityUnitPrice %>
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
							<td align="right" valign="top"><b>B13A Filing Option:</b></td>
							<td align="left">
							 <select name="B13AFilingOption" id="B13AFilingOption">
								<option value="" <%=pcf_SelectOption("B13AFilingOption","")%>>Select Option</option>
								<option value="NOT_REQUIRED" <%=pcf_SelectOption("B13AFilingOption","NOT_REQUIRED")%>>Not Required</option>
								<option value="FILED_ELECTRONICALLY" <%=pcf_SelectOption("B13AFilingOption","FILED_ELECTRONICALLY")%>>Filed Electronically</option>
								<option value="MANUALLY_ATTACHED" <%=pcf_SelectOption("B13AFilingOption","MANUALLY_ATTACHED")%>>Manually Attached</option>
								<option value="SUMMARY_REPORTING" <%=pcf_SelectOption("B13AFilingOption","SUMMARY_REPORTING")%>>Summary Reporting</option>
								</select>
								<%pcs_RequiredImageTag "B13AFilingOption", false%>
							</td>
						</tr>
						<tr>
							<td align="right" valign="top"><b>Export Compliance Statement:</b></td>
							<td align="left">
								<select name="ExportComplianceStatement" id="ExportComplianceStatement">
								<option value="NO EEI 30.36" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.36")%>>NO EEI 30.36</option>
								<option value="NO EEI 30.37(a)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(a)")%>>NO EEI 30.37(a)</option>
								<option value="NO EEI 30.37(b)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(b)")%>>NO EEI 30.37(b)</option>
								<option value="NO EEI 30.37(f)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(f)")%>>NO EEI 30.37(f)</option>
								<option value="NO EEI 30.37(g)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(g)")%>>NO EEI 30.37(g)</option>
								<option value="NO EEI 30.37(h)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(h)")%>>NO EEI 30.37(h)</option>
								<option value="NO EEI 30.37(i)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(i)")%>>NO EEI 30.37(i)</option>
								<option value="NO EEI 30.37(j)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(j)")%>>NO EEI 30.37(j)</option>
								<option value="NO EEI 30.37(k)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(k)")%>>NO EEI 30.37(k)</option>
								<option value="NO EEI 30.37(l)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(l)")%>>NO EEI 30.37(l)</option>
								<option value="NO EEI 30.37(p)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.37(p)")%>>NO EEI 30.37(p)</option>
								<option value="NO EEI 30.39" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.39")%>>NO EEI 30.39</option>
								<option value="NO EEI 30.40(a)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.40(a)")%>>NO EEI 30.40(a)</option>
								<option value="NO EEI 30.40(b)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.40(b)")%>>NO EEI 30.40(b)</option>
								<option value="NO EEI 30.40(c)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.40(c)")%>>NO EEI 30.40(c)</option>
								<option value="NO EEI 30.40(d)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.40(d)")%>>NO EEI 30.40(d)</option>
								<option value="NO EEI 30.02 (d)" <%=pcf_SelectOption("ExportComplianceStatement","NO EEI 30.02(d)")%>>NO EEI 30.02(d)</option>
								</select>
								<%pcs_RequiredImageTag "ExportComplianceStatement", false%>
							</td>
						</tr>
								<tr>
								  <td colspan="2" class="pcCPspacer"></td>
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
										<%pcs_RequiredImageTag "DutiesAccountNumber", isRequiredDutiesAccountNumber %>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Duties Payor Country Code:</b></td>
									<td align="left">
									<input name="DutiesCountryCode" type="text" id="DutiesCountryCode" value="<%=pcf_FillFormField("DutiesCountryCode", false)%>">
									<%pcs_RequiredImageTag "DutiesCountryCode", isRequiredDutiesCountryCode %>
									(e.g. US) </td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">Sender TIN Details</th>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Sender TIN Number:</b></td>
									<td align="left">
										<input name="SenderTINNumber" type="text" id="SenderTINNumber" value="<%=pcf_FillFormField("SenderTINNumber", false)%>">
										<%pcs_RequiredImageTag "SenderTINNumber", false%>
									</td>
								</tr>
								<tr>
									<td align="right" valign="top"><b>Sender TIN or DUNS Type:</b></td>
									<td align="left">
										<select name="SenderTINType" id="SenderTINType">
											<option value="BUSINESS_NATIONAL" <%=pcf_SelectOption("SenderTINType","BUSINESS_NATIONAL")%>>Business National</option>
											<option value="BUSINESS_STATE" <%=pcf_SelectOption("SenderTINType","BUSINESS_STATE")%>>Business State</option>
											<option value="BUSINESS_UNION" <%=pcf_SelectOption("SenderTINType","BUSINESS_UNION")%>>Business Union</option>
											<option value="PERSONAL_NATIONAL" <%=pcf_SelectOption("SenderTINType","PERSONAL_NATIONAL")%>>Personal National</option>
											<option value="PERSONAL_STATE" <%=pcf_SelectOption("SenderTINType","PERSONAL_STATE")%>>Personal State</option>
										</select>
										<%pcs_RequiredImageTag "SenderTINType", false%>
									</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">
									<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
									<!--
									function jfISOBrokerSelect(){
									var selectValDom = document.forms['form1'];
									if (selectValDom.bISOBrokerSelect.checked == true) {
									document.getElementById('ISOBrokerSelect').style.display='';
									}else{
									document.getElementById('ISOBrokerSelect').style.display='none';
									}
									}
									 //-->
									</SCRIPT>
									<%
									if Session("pcAdminbISOBrokerSelect")="true" then
										pcv_strDisplayStyle="style=""display:visible"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfISOBrokerSelect();" name="bISOBrokerSelect" id="bISOBrokerSelect" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bISOBrokerSelect", "1")%>>&nbsp;
									Broker Select Special Services Option<br>
									<br>
									<span class="pcCPnotes">Broker Select Option  should be used for Express shipments only.</span></th>
								</tr>
								<tr>
									<td colspan="2">
									<div id="ISOBrokerSelect" <%=pcv_strDisplayStyle%>>
										<Table>
											<tr>
												<td align="right"><b>CompanyName:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOCompanyName" id="BSOCompanyName" value="<%=pcf_FillFormField("BSOCompanyName", false)%>">
													<%pcs_RequiredImageTag "BSOCompanyName", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>PhoneNumber:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOPhoneNumber" id="BSOPhoneNumber" value="<%=pcf_FillFormField("BSOPhoneNumber", false)%>">
													<%pcs_RequiredImageTag "BSOPhoneNumber", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>Address:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOStreetLines" id="BSOStreetLines" value="<%=pcf_FillFormField("BSOStreetLines", false)%>">
													<%pcs_RequiredImageTag "BSOStreetLines", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>City:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOCity" id="BSOCity" value="<%=pcf_FillFormField("BSOCity", false)%>">
													<%pcs_RequiredImageTag "BSOCity", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>StateOrProvinceCode:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOStateOrProvinceCode" id="BSOStateOrProvinceCode" value="<%=pcf_FillFormField("BSOStateOrProvinceCode", false)%>">
													<%pcs_RequiredImageTag "BSOStateOrProvinceCode", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>PostalCode:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOPostalCode" id="BSOPostalCode" value="<%=pcf_FillFormField("BSOPostalCode", false)%>">
													<%pcs_RequiredImageTag "BSOPostalCode", false%>
												</td>
											</tr>
											<tr>
												<td align="right"><b>CountryCode:</b></td>
												<td align="left">
													<INPUT type="text" name="BSOCountryCode" id="SDIPackageCount" value="<%=pcf_FillFormField("BSOCountryCode", false)%>">
													<%pcs_RequiredImageTag "BSOCountryCode", false%>
												</td>
											</tr>
										</Table>
									</div>
									</td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<th colspan="2">
									<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
									<!--
									function jfISOCustomsID(){
									var selectValDom = document.forms['form1'];
									if (selectValDom.bISOCustomsID.checked == true) {
									document.getElementById('ISOCustomsID').style.display='';
									}else{
									document.getElementById('ISOCustomsID').style.display='none';
									}
									}
									 //-->
									</SCRIPT>
									<%
									if Session("pcAdminbISOCustomsID")="true" then
										pcv_strDisplayStyle="style=""display:visible"""
									else
										pcv_strDisplayStyle="style=""display:none"""
									end if
									%>
									<input onClick="jfISOCustomsID();" name="bISOCustomsID" id="bISOCustomsID" type="checkbox" class="clearBorder" value="1" <%=pcf_CheckOption("bISOCustomsID", "1")%>>
									Recipient Customs ID</th>
								</tr>
								<tr>
									<td colspan="2">
									<div id="ISOCustomsID" <%=pcv_strDisplayStyle%>>
										<Table>
											<tr>
											<td align="right" valign="top"><b>Id Type:</b></td>
											<td align="left">
										<select name="RCIdType" id="RCIdType">
											<option value="COMPANY" <%=pcf_SelectOption("RCIdType","COMPANY")%>>Company</option>
											<option value="INDIVIDUAL" <%=pcf_SelectOption("RCIdType","INDIVIDUAL")%>>Individual</option>
											<option value="PASSPORT" <%=pcf_SelectOption("RCIdType","PASSPORT")%>>Passport</option>
										</select>
										<%pcs_RequiredImageTag "RCIdType", false%>
											</td>
										</tr>
										<tr>
											<td align="right" valign="top"><b>Id Number:</b></td>
											<td align="left">
												<input name="RCIdValue" type="text" id="RCIdValue" value="<%=pcf_FillFormField("RCIdValue", false)%>">
												<%pcs_RequiredImageTag "RCIdValue", false%>


							</td>
						</tr>
					</table>
					</div>
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
			You must enter a valid E-mail Address.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>E-mail Address:</p></td>
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
			You must enter a valid E-mail Address.
			</td>
			</tr>
			<% end if %>
			<tr>
			<td align="right"><p>E-mail Address:</p></td>
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
			<input type="checkbox" name="ResidentialDelivery" value="true" class="clearBorder" <%=pcf_CheckOption("ResidentialDelivery", "true")%>>
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
					<p><strong>Shipment notification</strong> &ndash; Automatically send an e-mail message indicating the shipment is on the way.<br>
					  <strong>Delivery notification</strong> &ndash; receive a delivery notification for an express package. <br>
					  <strong>Exception notification</strong> - receive an e-mail notifcation for delivery exceptions.<br>
					  <strong>E-mail address</strong> &ndash; Enter the e-mail addresses to receive the notifications.                    </p></td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td align="left" colspan="2">Notification Format:&nbsp;&nbsp;<select name="ShipperNotificationFormat" id="select">
				<option value="HTML" <%=pcf_SelectOption("ShipperNotificationFormat","HTML")%>>HTML Format</option>
				<option value="TEXT" <%=pcf_SelectOption("ShipperNotificationFormat","TEXT")%>>Plain Text Format</option>
				<option value="WIRELESS" <%=pcf_SelectOption("ShipperNotificationFormat","WIRELESS")%>>Formatted for Wireless Device</option>
			</select></td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td align="left"><strong>Shipper Notification:</strong></td>
				<td align="left">&nbsp;</td>
			</tr>
			<tr>
				<td align="left" colspan="2">Shipper E-mail:&nbsp;&nbsp;
				  <input name="NotificationShipperEmail" type="text" id="NotificationShipperEmail" value="<%=pcf_FillFormField("NotificationShipperEmail", false)%>"></td>
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
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td align="left"><strong>Recipient Notification:</strong></td>
				<td align="left">&nbsp;</td>
			</tr>
			<tr>
				<td align="left" colspan="2">Recipient E-mail:&nbsp;&nbsp;
				  <input name="NotificationRecipientEmail" type="text" id="NotificationRecipientEmail" value="<%=pcf_FillFormField("NotificationRecipientEmail", false)%>"></td>
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
						<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">You must enter a valid Email Address. <br />
					<% end if %>
				</td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td align="left"><strong>Additional Notification:</strong></td>
				<td align="left">&nbsp;</td>
			</tr>
			<tr>
				<td align="left" colspan="2">E-mail Address:&nbsp;&nbsp;<input name="OtherNotification1" type="text" id="OtherNotification1" value="<%=pcf_FillFormField("OtherNotification1", false)%>">
				<%pcs_RequiredImageTag "OtherNotification1", false%></td>
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
															tOS_width=FEDEXWS_WIDTH
															tOS_height=FEDEXWS_HEIGHT
															tOS_length=FEDEXWS_LENGTH
														end if
													else
														tOS_width=FEDEXWS_WIDTH
														tOS_height=FEDEXWS_HEIGHT
														tOS_length=FEDEXWS_LENGTH
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
					  <script type="text/javascript">
							function setpackagedivs() {
							   var div_num = $("#Service1").val();
							   if (div_num == 1) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num == 2) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==3) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==4) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==5) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==6) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==7) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").show();
								};
							   if (div_num ==8) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==9) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==10) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==11) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==12) {
							   $("#smartpost").hide();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							   if (div_num ==13) {
							   $("#smartpost").hide();
							   $("#expressfreight").show();
							   $("#homedelivery").hide();
								};
							   if (div_num ==14) {
							   $("#smartpost").hide();
							   $("#expressfreight").show();
							   $("#homedelivery").hide();
								};
							   if (div_num ==15) {
							   $("#smartpost").hide();
							   $("#expressfreight").show();
							   $("#homedelivery").hide();
								};
							   if (div_num == 16) {
							   $("#smartpost").show();
							   $("#expressfreight").hide();
							   $("#homedelivery").hide();
								};
							}
							</script>
					<tr>
					<td colspan="2">
					<%
					'// Set Service Type to local
					Select Case Service
						Case "FedEx First Overnight": ServiceSelector="FIRST_OVERNIGHT"
						Case "FedEx Priority Overnight": ServiceSelector="PRIORITY_OVERNIGHT"
						Case "FedEx Standard Overnight": ServiceSelector="STANDARD_OVERNIGHT"
						Case "FedEx 2Day": ServiceSelector="FEDEX_2_DAY"
						Case "FedEx Express Saver": ServiceSelector="FEDEX_EXPRESS_SAVER"
						Case "FedEx Ground": ServiceSelector="FEDEX_GROUND"
						Case "FedEx Home Delivery": ServiceSelector="GROUND_HOME_DELIVERY"
						Case "FedEx International First": ServiceSelector="INTERNATIONAL_FIRST"
						Case "FedEx International Priority": ServiceSelector="INTERNATIONAL_PRIORITY"
						Case "FedEx International Economy": ServiceSelector="INTERNATIONAL_ECONOMY"
						Case "FedEx 1Day Freight": ServiceSelector="FEDEX_1_DAY_FREIGHT"
						Case "FedEx 2Day Freight": ServiceSelector="FEDEX_2_DAY_FREIGHT"
						Case "FedEx 3Day Freight": ServiceSelector="FEDEX_3_DAY_FREIGHT"
						Case "FedEx International Priority Freight": ServiceSelector="INTERNATIONAL_PRIORITY_FREIGHT"
						Case "FedEx International Economy Freight": ServiceSelector="INTERNATIONAL_ECONOMY_FREIGHT"
						Case "FedEx SmartPost": ServiceSelector="SMART_POST"
					End Select

					if Session("pcAdminService"&k)="" then
						Session("pcAdminService"&k)=ServiceSelector
					end if
					%>
					<p>
					<strong>Service Type: </strong>
					<select name="Service<%=k%>" id="Service<%=k%>" onchange="setpackagedivs();">
								<option value="1" <%=pcf_SelectOption("Service"&k,"FIRST_OVERNIGHT")%>>FedEx First Overnight&reg;</option>
								<option value="2" <%=pcf_SelectOption("Service"&k,"PRIORITY_OVERNIGHT")%>>FedEx Priority Overnight&reg;</option>
								<option value="3" <%=pcf_SelectOption("Service"&k,"STANDARD_OVERNIGHT")%>>FedEx Standard Overnight&reg;</option>
								<option value="4" <%=pcf_SelectOption("Service"&k,"FEDEX_2_DAY")%>>FedEx 2Day&reg;</option>
								<option value="5" <%=pcf_SelectOption("Service"&k,"FEDEX_EXPRESS_SAVER")%>>FedEx Express Saver&reg;</option>
								<option value="6" <%=pcf_SelectOption("Service"&k,"FEDEX_GROUND")%>>FedEx Ground&reg;</option>
								<option value="7" <%=pcf_SelectOption("Service"&k,"GROUND_HOME_DELIVERY")%>>FedEx Home Delivery&reg;</option>
								<option value="8" <%=pcf_SelectOption("Service"&k,"INTERNATIONAL_FIRST")%>>FedEx International First&reg;</option>
								<option value="9" <%=pcf_SelectOption("Service"&k,"INTERNATIONAL_PRIORITY")%>>FedEx International Priority&reg;</option>
								<option value="10" <%=pcf_SelectOption("Service"&k,"INTERNATIONAL_ECONOMY")%>>FedEx International Economy&reg; </option>
								<option value="11" <%=pcf_SelectOption("Service"&k,"INTERNATIONAL_PRIORITY_FREIGHT")%>>FedEx International Priority&reg; Freight</option>
								<option value="12" <%=pcf_SelectOption("Service"&k,"INTERNATIONAL_ECONOMY_FREIGHT")%>>FedEx International Economy&reg; Freight</option>
								<option value="13" <%=pcf_SelectOption("Service"&k,"FEDEX_1_DAY_FREIGHT")%>>FedEx 1Day&reg; Freight</option>
								<option value="14" <%=pcf_SelectOption("Service"&k,"FEDEX_2_DAY_FREIGHT")%>>FedEx 2Day&reg; Freight</option>
								<option value="15" <%=pcf_SelectOption("Service"&k,"FEDEX_3_DAY_FREIGHT")%>>FedEx 3Day&reg; Freight</option>
								<option value="16" <%=pcf_SelectOption("Service"&k,"SMART_POST")%>>FedEx SmartPost&reg;</option>

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
					<%

					' FEDEX_10KG_BOX
					' FEDEX_25KG_BOX
					' FEDEX_BOX
					' FEDEX_ENVELOPE
					' FEDEX_PAK
					' FEDEX_TUBE
					' YOUR_PACKAGING
					%>
					<select name="Packaging<%=k%>" id="Packaging<%=k%>">
					<option value="FEDEX_ENVELOPE" <%=pcf_SelectOption("Packaging"&k,"FEDEX_ENVELOPE")%>>FedEx&reg; Envelope</option>
					<option value="FEDEX_PAK" <%=pcf_SelectOption("Packaging"&k,"FEDEX_PAK")%>>FedEx&reg; Pak</option>
					<option value="FEDEX_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_BOX")%>>FedEx&reg; Box</option>
					<option value="FEDEX_TUBE" <%=pcf_SelectOption("Packaging"&k,"FEDEX_TUBE")%>>FedEx&reg; Tube</option>
					<option value="FEDEX_10KG_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_10KG_BOX")%>>FedEx&reg; 10kg Box</option>
					<option value="FEDEX_25KG_BOX" <%=pcf_SelectOption("Packaging"&k,"FEDEX_25KG_BOX")%>>FedEx&reg; 25kg Box</option>
					<option value="YOUR_PACKAGING" <%=pcf_SelectOption("Packaging"&k,"YOUR_PACKAGING")%>>Customer Package</option>
					</select>
					<%pcs_RequiredImageTag "Packaging"&k, true%>
					</p>
					<br />
					<input type="checkbox" name="ContainerType" value="1" class="clearBorder">&nbsp;Non-Standard Container

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
					<option value="LB" <%=pcf_SelectOption("WeightUnits"&k,"LB")%>>LB</option>
					<option value="KG" <%=pcf_SelectOption("WeightUnits"&k,"KG")%>>KG</option>
					</select>
					<%pcs_RequiredImageTag "WeightUnits"&k, true
					%>
					<%
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
							<td colspan="2">
							<div id='smartpost'>
								<Table>
									<tr>
										<th colspan="2">Smart Post Details</th>
									</tr>
									<tr>
										<td align="right" valign="top"><b>Indicia:</b></td>
										<td align="left">
										<select name="SMIndicia" id="SMIndicia">
											<option value="MEDIA_MAIL" <%=pcf_SelectOption("SMIndicia","PARCEL_SELECT")%>>Media Mail</option>
											<option value="PARCEL_SELECT" <%=pcf_SelectOption("SMIndicia","PARCEL_SELECT")%>>Parcel Select</option>
											<option value="PRESORTED_STANDARD" <%=pcf_SelectOption("SMIndicia","PRESORTED_STANDARD")%>>Presorted Standard</option>
											<option value="PRESORTED_BOUND_PRINTED_MATTER" <%=pcf_SelectOption("SMIndicia","PRESORTED_BOUND_PRINTED_MATTER")%>>Presorted Bound Printed Matter</option>
										  </select>
											<%pcs_RequiredImageTag "SMIndicia", true%>
										</td>
									</tr>
									<tr>
										<td align="right" valign="top" nowrap><b>Ancillary Endorsement:</b></td>
										<td align="left">
										<select name="SMAncillaryEndorsement" id="SMAncillaryEndorsement">
											<option value="CARRIER_LEAVE_IF_NO_RESPONSE" <%=pcf_SelectOption("SMAncillaryEndorsement","CARRIER_LEAVE_IF_NO_RESPONSE")%>>Carrier leave if no response</option>
											<option value="ADDRESS_CORRECTION" <%=pcf_SelectOption("SMAncillaryEndorsement","ADDRESS_CORRECTION")%>>Address Correction</option>
											<option value="RETURN_SERVICE" <%=pcf_SelectOption("SMAncillaryEndorsement","RETURN_SERVICE")%>>Return Service</option>
										  </select>
											<%pcs_RequiredImageTag "SMAncillaryEndorsement", true%>
										</td>
									</tr>
									<tr>
										<td align="right" valign="top"><b>HUB ID:</b></td>
										<td align="left">
										<select name="SMHubID" id="SMHubID">
											<option value="5015" <%=pcf_SelectOption("SMHubID","5015")%>>5015</option>Northborough, MA</option>
											<option value="5087" <%=pcf_SelectOption("SMHubID","5087")%>>5087</option>Edison, NJ</option>
											<option value="5150" <%=pcf_SelectOption("SMHubID","5150")%>>5150</option>Pittsburgh, PA</option>
											<option value="5185" <%=pcf_SelectOption("SMHubID","5185")%>>5185</option>Allentown, PA</option>
											<option value="5254" <%=pcf_SelectOption("SMHubID","5254")%>>5254</option>Martinsburg, WV</option>
											<option value="5281" <%=pcf_SelectOption("SMHubID","5281")%>>5281</option>Charlotte, NC</option>
											<option value="5303" <%=pcf_SelectOption("SMHubID","5303")%>>5303</option>Atlanta, GA</option>
											<option value="5327" <%=pcf_SelectOption("SMHubID","5327")%>>5327</option>Orlando, FL</option>
											<option value="5379" <%=pcf_SelectOption("SMHubID","5379")%>>5379</option>Memphis, TN</option>
											<option value="5431" <%=pcf_SelectOption("SMHubID","5431")%>>5431</option>Grove City, OH</option>
											<option value="5465" <%=pcf_SelectOption("SMHubID","5465")%>>5465</option>Indianapolis, IN</option>
											<option value="5481" <%=pcf_SelectOption("SMHubID","5481")%>>5481</option>Detroit, MI</option>
											<option value="5531" <%=pcf_SelectOption("SMHubID","5531")%>>5531</option>New Berlin, WI</option>
											<option value="5552" <%=pcf_SelectOption("SMHubID","5552")%>>5552</option>Minneapolis, MN</option>
											<option value="5631" <%=pcf_SelectOption("SMHubID","5631")%>>5631</option>St. Louis, MO</option>
											<option value="5648" <%=pcf_SelectOption("SMHubID","5648")%>>5648</option>Kansas, KS</option>
											<option value="5751" <%=pcf_SelectOption("SMHubID","5751")%>>5751</option>Dallas, TX</option>
											<option value="5771" <%=pcf_SelectOption("SMHubID","5771")%>>5771</option>Houston, TX</option>
											<option value="5802" <%=pcf_SelectOption("SMHubID","5802")%>>5802</option>Denver, CO</option>
											<option value="5843" <%=pcf_SelectOption("SMHubID","5843")%>>5843</option>Salt Lake City, UT</option>
											<option value="5854" <%=pcf_SelectOption("SMHubID","5854")%>>5854</option>Phoenix, AZ</option>
											<option value="5902" <%=pcf_SelectOption("SMHubID","5902")%>>5902</option>Los Angeles, CA</option>
											<option value="5929" <%=pcf_SelectOption("SMHubID","5929")%>>5929</option>Chino, CA</option>
											<option value="5958" <%=pcf_SelectOption("SMHubID","5958")%>>5958</option>Sacramento, CA</option>
											<option value="5983" <%=pcf_SelectOption("SMHubID","5983")%>>5983</option>Seattle, WA</option>
										  </select>
											<%pcs_RequiredImageTag "SMHubID", true%>
										</td>
									</tr>
								</Table>
							</div>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<div id='expressfreight'>
								<Table>
									<tr>
										<th colspan="2">Express Freight Options</th>
									</tr>
									<tr>
										<td width="54%" align="right" valign="top"><b>Packing List Enclosed?</b></td>
										<td width="46%" align="left">
											<input name="EFPackingListEnclosed" type="radio" id="EFPackingListEnclosed" value="1">Yes
											<input name="EFPackingListEnclosed" type="radio" id="EFPackingListEnclosed" value="1">No
										</td>
									</tr>
									<tr>
										<td align="right" valign="top"><b>ShippersLoadAndCount:</b></td>
										<td align="left">
											<input name="EFShippersLoadAndCount" type="text" id="EFShippersLoadAndCount" value="<%=pcf_FillFormField("EFShippersLoadAndCount", false)%>">
											<%pcs_RequiredImageTag "EFShippersLoadAndCount", false%>
										</td>
									</tr>
									<tr>
										<td align="right" valign="top" nowrap><b>Booking Confirmation Number</b></td>
										<td align="left">
											<input name="EFBookingConfirmationNumber" type="text" id="EFBookingConfirmationNumber" value="<%=pcf_FillFormField("EFBookingConfirmationNumber", false)%>">
											<%pcs_RequiredImageTag "EFBookingConfirmationNumber", false%>
										</td>
									</tr>
								</Table>
								</div>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<div id='homedelivery'>
								<Table>
									<tr>
									  <th colspan="2">Home Delivery Options</th>
								  </tr>
									<tr>
									  <td width="25%" align="right" valign="top"><b>Delivery Type:</b></td>
									  <td align="left">
										<select name="DeliveryType" id="DeliveryType">
										  <option value="">Please make a selection. (optional)</option>
										  <option value="DATE_CERTAIN" <%=pcf_SelectOption("DeliveryType","DATE_CERTAIN")%>>Date Certain</option>
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
										2012-02-29</td>
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
								</Table>
								</div>
							</td>
						</tr>
					<% end if %>
				<% else %>
					<tr>
						<th colspan="2">This package has been shipped.</th>
					</tr>
				<% end if %>
				</table>
			</div>
			<%
		next %>

			<br />
			<br />

			<%
			pcv_strPreviousPage = "Orddetails.asp?id=" & pcv_intOrderID
			pcv_strAddPackagePage = "sds_ShipOrderWizard1.asp?idorder="&pcv_intOrderID&"&PageAction=FedExWs&PackageCount="&pcPackageCount&"&ItemsList="&pcv_strItemsList
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