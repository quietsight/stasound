<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/pcUPSClass.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<!--#include file="../includes/pcShipTestModes.asp" -->
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
<script language="JavaScript"><!--
function newWindow(file,window) {
		PackageWindow=open(file,window,'resizable=no,width=500,height=600,scrollbars=1');
		if (PackageWindow.opener == null) PackageWindow.opener = self;
}
//--></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
<!--
function whatPayorTypeSelected(){

var selectValDom = document.forms['form1'].elements['PayorType'].options;
if (selectValDom.value == 'PrePaid') {
document.getElementById('PayorType_table').style.display='none';
}else{
document.getElementById('PayorType_table').style.display='';
}
}
 //-->
</SCRIPT>
<%
Dim objUPSXmlDoc, objUPSStream, strFileName, GraphicXML
Dim iPageCurrent, varFlagIncomplete, uery, strORD, pcv_intOrderID
Dim pcv_strMethodName, pcv_strMethodReply, CustomerTransactionIdentifier, pcv_strAccountNumber, pcv_strMeterNumber, pcv_strUPSServiceCode
Dim pcv_strTrackingNumber, pcv_strShipmentAccountNumber
Dim pcv_strDestinationCountryCode, pcv_strDestinationPostalCode, pcv_strLanguageCode, pcv_strLocaleCode, pcv_strDetailScans, pcv_strPagingToken
Dim UPS_postdata, objUPSClass, objOutputXMLDoc, srvUPSXmlHttp, UPS_result, UPS_URL, pcv_strErrorMsg, pcv_strAction

function fnStripPhone(PhoneField)
	PhoneField=replace(PhoneField," ","")
	PhoneField=replace(PhoneField,"-","")
	PhoneField=replace(PhoneField,".","")
	PhoneField=replace(PhoneField,"(","")
	PhoneField=replace(PhoneField,")","")
	fnStripPhone = PhoneField
end function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
dim intResetSessions
intResetSessions=0
pcv_strOrderID=request("idorder")
pcv_strSessionOrderID=Session("pcAdminOrderID")
if pcv_strSessionOrderID="" OR len(pcv_strOrderID)>0 then
	pcv_intOrderID=pcv_strOrderID
	'// Reset all sessions
	if pcv_strSessionOrderID<>pcv_strOrderID then
		intResetSessions=1
	end if
else
	pcv_intOrderID=pcv_strSessionOrderID
end if
Session("pcAdminOrderID")=pcv_intOrderID

'// REDIRECT
if pcv_intOrderID="" then
	response.redirect "menu.asp"
end if

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
pcPageName="UPS_ManageShipmentsRequest.asp"
ErrPageName="UPS_ManageShipmentsRequest.asp"

'// ACTION
pcv_strAction = request("Action")

'// OPEN DATABASE
dim conntemp, rs, query
call openDb()

'// SET THE UPS OBJECT
set objUPSClass = New pcUPSClass

'// UPS CREDENTIALS
query="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then
	UPS_Active=rs("active")
	UPS_UserId=trim(rs("userID"))
	UPS_Password=trim(rs("password"))
	UPS_LicenseKey=trim(rs("AccessLicense"))
end if

set rs=nothing

'// DATE FUNCTION
function ShowDateFrmt(x)
	ShowDateFrmt = x
end function

'// UPS Ship Preferences
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
query="SELECT orders.idCustomer, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, "
query = query & "orders.shippingCountryCode, orders.shippingZip, orders.shippingCompany, orders.shippingAddress2, orders.pcOrd_shippingPhone, orders.ShippingFullName, orders.pcOrd_ShippingEmail, orders.ordShipType, orders.pcOrd_ShipWeight, shipmentDetails, SRF "
query = query & "FROM orders "
query = query & "WHERE orders.idOrder=" & pcv_intOrderID &" "

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then
	Dim pidorder, pidcustomer, pcv_ShippingAddress, pcv_ShippingCity, pshippingStateCode, pshippingState, pshippingZip, pshippingPhone, pshippingCountryCode, pshippingCompany, pcv_ShippingAddress2, pShippingEmail

	'// ORDER INFO
	pidorder=scpre+int(pcv_intOrderID)
	pcv_IdCustomer=rs("idcustomer")

	'// DESTINATION ADDRESS
	pcv_ShippingAddress=rs("shippingAddress")
	pcv_ShippingCity=rs("shippingCity")
	pcv_ShippingStateCode=rs("shippingStateCode")
	pcv_ShippingState=rs("shippingState")
	pcv_ShippingZip=rs("shippingZip")
	pcv_ShippingPhone=rs("pcOrd_shippingPhone")
	pcv_ShippingCountryCode=rs("shippingCountryCode")
	pcv_ShippingCompany=rs("shippingCompany")
	pcv_ShippingAddress2=rs("shippingAddress2")
	pcv_ShippingFullName=rs("ShippingFullName")
	pcv_ShippingEmail=rs("pcOrd_ShippingEmail")
	pcv_OrdShipType=rs("ordShipType")
	pcv_ShipWeight=rs("pcOrd_ShipWeight")
	pcv_ShipmentDetails=rs("shipmentDetails")
	pcv_SRF=rs("SRF")
end if
set rs=nothing

If pcv_SRF="1" then
	'
else
	shipping=split(pcv_ShipmentDetails,",")
	if ubound(shipping)>1 then
		if NOT isNumeric(trim(shipping(2))) then
			varShip="0"
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
			if ubound(shipping)>4 then
				ServiceCode=trim(shipping(5))
			end if
		end if
	else
		varShip="0"
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
pcv_OwnerEmail=rs("ownerEmail")
set rs = nothing

if Session("pcAdminShipToDiffLocation") = "" then
	pcv_strShipToDiffLocation = "N"
	Session("pcAdminShipToDiffLocation") = pcv_strShipToDiffLocation
end if

if Session("pcAdminWeightUnitOfMeasurement") = "" then
	pcv_strWeightUnitOfMeasurement = scShipFromWeightUnit
	Session("pcAdminWeightUnitOfMeasurement") = pcv_strWeightUnitOfMeasurement
end if

'// DESTINATION ADDRESS
if Session("pcAdminShipToAttentionName") = "" OR intResetSessions=1 then
	pcv_strShipToAttentionName = pcv_ShippingFullName
	if pcv_strShipToAttentionName="" then
		pcv_strShipToAttentionName=pcv_ShippingCompany
	end if
	Session("pcAdminShipToAttentionName") = pcv_strShipToAttentionName
end if

if Session("pcAdminShipToCompanyName") = "" OR intResetSessions=1 then
	pcv_strShipToCompanyName = pcv_ShippingCompany
	if pcv_strShipToCompanyName="" then
		pcv_strShipToCompanyName=pcv_strShipToAttentionName
	end if
	Session("pcAdminShipToCompanyName") = pcv_strShipToCompanyName
end if

if Session("pcAdminShipToPhoneNumber") = "" OR intResetSessions=1 then
	pcv_strShipToPhoneNumber = pcv_ShippingPhone
	Session("pcAdminShipToPhoneNumber") = pcv_strShipToPhoneNumber
end if

if Session("pcAdminShipToEmailAddress") = "" OR intResetSessions=1 then
	pcv_strShipToEmailAddress = pcv_ShippingEmail
	Session("pcAdminShipToEmailAddress") = pcv_strShipToEmailAddress
end if

'// DESTINATION ADDRESS
if Session("pcAdminShipToAddressLine1") = "" OR IntResetSessions=1 then
	pcv_strShipToAddressLine1 = pcv_ShippingAddress
	Session("pcAdminShipToAddressLine1") = pcv_strShipToAddressLine1
end if
if Session("pcAdminShipToAddressLine2") = "" OR intResetSessions=1 then
	pcv_strShipToAddressLine2 = pcv_ShippingAddress2
	Session("pcAdminShipToAddressLine2") = pcv_strShipToAddressLine2
end if
if Session("pcAdminShipToAddressLine3") = "" OR intResetSessions=1 then
	pcv_strShipToAddressLine3 = pcv_ShippingAddress3
	Session("pcAdminShipToAddressLine3") = pcv_strShipToAddressLine3
end if
if Session("pcAdminShipToCity") = "" OR intResetSessions=1 then
	pcv_strShipToCity = pcv_ShippingCity
	Session("pcAdminShipToCity") = pcv_ShippingCity
end if
if Session("pcAdminShipToStateOrProvinceCode") = "" OR intResetSessions=1 then
	pcv_strShipToStateOrProvinceCode = pcv_ShippingStateCode
	Session("pcAdminShipToStateOrProvinceCode") = pcv_strShipToStateOrProvinceCode
end if
if Session("pcAdminShipToPostalCode") = "" OR intResetSessions=1 then
	pcv_strShipToPostalCode = pcv_ShippingZip
	Session("pcAdminShipToPostalCode") = pcv_strShipToPostalCode
end if
if Session("pcAdminShipToCountryCode") = "" OR intResetSessions=1 then
	pcv_strShipToCountryCode = pcv_ShippingCountryCode
	Session("pcAdminShipToCountryCode") = pcv_strShipToCountryCode
end if

strShipToStateCode=Session("pcAdminShipToStateOrProvinceCode")
strShipToCity=Session("pcAdminShipToCity")
strShipToPostalCode=Session("pcAdminShipToPostalCode")
strShipToCountry=Session("pcAdminShipToCountryCode")

'   >>> Recipient Address Conditionals
'// Use the Request object to toggle State (based of Country selection)
isRequiredRecipState =  true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	isRequiredRecipState=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
isRequiredRecipProvince = true
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	isRequiredRecipProvince=pcv_strProvinceCodeRequired
end if

if Session("pcAdminShipToCountryCode") = "US" OR Session("pcAdminShipToCountryCode") = "CA" then
	isRequiredShipToPostal = true
end if

'// SET REQUIRED VARIABLES
pcv_strMethodName = "FDXShipRequest"
pcv_strMethodReply = "FDXShipReply"
CustomerTransactionIdentifier = "ProductCart_Test"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
call closedb()
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2">Order ID#: <b><%=(scpre+int(pcv_intOrderID))%></b></td>
	</tr>
	<tr>
		<th colspan="2">UPS OnLine&reg; Tools Shipping</th>
	</tr>
	<% if UPS_TESTMODE="1" then %>
		<tr>
			<td colspan="2" class="pcSpacer"></td>
		</tr>
		<tr>
			<td colspan="2">
				<div class="pcCPmessage">
					UPS Shipping Wizard is currently running in Test Mode <br>
Only "SAMPLE" labels will be generated while in Test Mode.
				</div>
			</td>
		</tr>
	<% end if %>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
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
				call opendb()
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
				pcv_strTotalInsuredValue = 0
				pcv_strTotalWeight = 0
				For pcv_xCounter = 1 to pcPackageCount
					' If its shipped the field is no longer required
					if pcLocalArray(pcv_xCounter-1) = "shipped" then
						pcv_strToggle = false
					else
						pcv_strToggle = true
					end if

					pcs_ValidateTextField	"FaxLetter"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"PackageTypeCode"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"DimensionsUnitOfMeasurement"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"Length"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"Width"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"Height"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"WeightUnitOfMeasurement"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"PackageWeight"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"OversizePackage"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"CODAmount"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"CODCurrencyCode"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"CODFundscode"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"InsuredValue"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"DeliveryConfirmation"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"VerbalConfirmation"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"AdditionalHandling"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"CODPackage"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"CODAmount"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"CODCurrencyCode"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"CODFundsCode"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"UPSRefNumber1"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"UPSRefData1"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"UPSRefNumber2"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"UPSRefData2"&pcv_xCounter, false, 0
					if Session("pcAdminPackageWeight"&pcv_xCounter)="" then Session("pcAdminPackageWeight"&pcv_xCounter) = 0
					if Session("pcAdminInsuredValue"&pcv_xCounter)="" then Session("pcAdminInsuredValue"&pcv_xCounter) = 0
				Next

				'..VALIDATE ALL OTHER FIELDS
				pcs_ValidateTextField	"idOrder", true, 0
				pcs_ValidateTextField	"ShipToDiffLocation", true, 1
				pcs_ValidateTextField	"packagecount", false, 0
				pcs_ValidateTextField	"itemsList", false, 0
				pcs_ValidateTextField	"CurrencyCode", false, 0
				pcs_ValidateTextField	"UPSServiceCode", false, 0
				pcs_ValidateTextField	"PayorType", false, 0
				pcs_ValidateTextField	"PayorAccountNumber", false, 0
				pcs_ValidateTextField	"PayorcountryCode", false, 0
				pcs_ValidateTextField	"ImageType", false, 0
				pcs_ValidateTextField	"InvoiceCurrencyCode", false, 0
				pcs_ValidateTextField	"InvoiceAmount", false, 0
				pcs_ValidateTextField	"SaturdayDelivery", false, 0
				pcs_ValidateTextField "OnCallPickup", false, 0
				if session("pcAdminOnCallPickup")="1" then
					pcs_ValidateTextField "OnCallDate", true, 0
					pcs_ValidateTextField "UPSReadyHours", true, 0
					pcs_ValidateTextField "UPSReadyMinutes", true, 0
					pcs_ValidateTextField "UPSReadyAMPM", true, 0
					pcs_ValidateTextField "UPSPUHours", true, 0
					pcs_ValidateTextField "UPSPUMinutes", true, 0
					pcs_ValidateTextField "OnCallContactName", true, 0
					pcs_ValidateTextField "OnCallContactPhone", true, 0
				end if
				if session("pcAdminShipToDiffLocation")="Y" then
					isShipFromCompanyNameRequired=true
					isShipFromCountryCodeRequired=true
					isShipFromAddressLine1Required=true
				else
					isShipFromCompanyNameRequired=false
					isShipFromCountryCodeRequired=false
					isShipFromAddressLine1Required=false
				end if
				pcs_ValidateTextField	"ShipFromCompanyName", isShipFromCompanyNameRequired, 0
				pcs_ValidateTextField	"ShipFromAttentionName", false, 0
				pcs_ValidatePhoneNumber	"ShipFromPhoneNumber", false, 0
				pcs_ValidateTextField	"ShipFromPhoneNumberExt", false, 0
				pcs_ValidateTextField	"ShipFromFaxNumber", false, 0
				pcs_ValidateEmailField	"ShipFromEmailAddress", false, 0
				pcs_ValidateTextField	"ShipFromAddressLine1", isShipFromAddressLine1Required, 0
				pcs_ValidateTextField	"ShipFromAddress2", false, 0
				pcs_ValidateTextField	"ShipFromAddress3", false, 0
				pcs_ValidateTextField	"ShipFromCity", false, 0
				pcs_ValidateTextField	"ShipFromStateCode", false, 0
				pcs_ValidateTextField	"ShipFromProvinceCode", false, 0
				if Session("pcAdminShipFromProvinceCode") <> "" then
					Session("pcAdminShipFromStateOrProvinceCode")=Session("pcAdminShipFromProvinceCode")
				else
					Session("pcAdminShipFromStateOrProvinceCode")=Session("pcAdminShipFromStateCode")
				end if
				pcs_ValidateTextField	"ShipFromPostalCode", false, 0
				pcs_ValidateTextField	"ShipFromCountryCode", isShipFromCountryCodeRequired, 0
				pcs_ValidateTextField	"ShipToCompanyName", false, 0
				pcs_ValidateTextField	"ShipToAttentionName", false, 0
				pcs_ValidatePhoneNumber	"ShipToPhoneNumber", false, 0
				pcs_ValidatePhoneNumber	"ShipToFaxNumber", false, 0
				pcs_ValidateTextField	"ShipToPhoneNumberExt", false, 0
				pcs_ValidateTextField	"shipToFaxNumber", false, 0
				pcs_ValidateEmailField	"ShipToEmailAddress", false, 0
				pcs_ValidateTextField	"ShipToAddressLine1", true, 0
				pcs_ValidateTextField	"ShipToAddressLine2", false, 0
				pcs_ValidateTextField	"ShipToAddressLine3", false, 0
				pcs_ValidateTextField	"ShipToCity", true, 0
				pcs_ValidateTextField	"ShipToStateOrProvinceCode", isRequiredRecipState, 0
				pcs_ValidateTextField	"ShipToProvince", isRequiredRecipProvince, 0
				if Session("pcAdminShipToStateOrProvinceCode") = "" then
					Session("pcAdminShipToStateOrProvinceCode")=Session("pcAdminShipToProvince")
				end if
				pcs_ValidateTextField	"ShipToPostalCode", false, 0
				pcs_ValidateTextField	"ShipToCountryCode", true, 0
				pcs_ValidateTextField	"ResidentialDelivery", false, 0
				pcs_ValidateTextField	"ShipmentDescription", false, 0
				pcs_ValidateTextField	"ShipmentNotification", false, 0
				pcs_ValidateTextField	"NotificationCode1", false, 0
				pcs_ValidateTextField	"NotificationCode2", false, 0
				pcs_ValidateTextField	"NotificationCode3", false, 0
				pcs_ValidateTextField	"NotificationCode4", false, 0
				pcs_ValidateTextField	"NotificationCode5", false, 0
				pcs_ValidateTextField	"NotificationEmail1", false, 0
				pcs_ValidateTextField	"NotificationEmail2", false, 0
				pcs_ValidateTextField	"NotificationEmail3", false, 0
				pcs_ValidateTextField	"NotificationEmail4", false, 0
				pcs_ValidateTextField	"NotificationEmail5", false, 0

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Check for Validation Errors. Do not proceed if there are errors.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				If pcv_intErr>0 Then
					response.redirect pcPageName & "?sub=1&msg=" & pcv_strGenericPageError
				Else

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Build Our Transaction.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					objUPSClass.NewXMLTransaction UPS_LicenseKey, UPS_UserId, UPS_Password

					pcv_xCounter = 1
					pcv_strTotalInsuredValue = 0
					pcv_strTotalWeight = 0
					errnum = 0

					objUPSClass.NewXMLShipmentConfirmRequest "ShipConfirm", "nonvalidate"

					'// LABEL SPECIFICATION
					objUPSClass.WriteParent "LabelSpecification", ""
						objUPSClass.WriteParent "LabelPrintMethod", ""
							objUPSClass.AddNewNode "Code", Session("pcAdminImageType")
						objUPSClass.WriteParent "LabelPrintMethod", "/"
						if Session("pcAdminImageType")="EPL" OR Session("pcAdminImageType")="SPL" then
							objUPSClass.WriteParent "LabelStockSize", ""
								objUPSClass.AddNewNode "Height", "4"
								objUPSClass.AddNewNode "Width", "6"
							objUPSClass.WriteParent "LabelStockSize", "/"
						else
							dim strFullUserAgent,strAbbrUserAgentArry,strUserAgent
							strFullUserAgent=Request.ServerVariables("HTTP_USER_AGENT")
							if instr(strFullUserAgent ,"(") then
								strAbbrUserAgentArry=split(strFullUserAgent, "(")
								strUserAgent=strAbbrUserAgentArry(0)
							else
								strUserAgent="Mozilla/5.0"
							end if
							objUPSClass.AddNewNode "HTTPUserAgent", strUserAgent 'strUserAgent&"</HTTPUserAgent>"&vbcrlf
							objUPSClass.WriteParent "LabelImageFormat", ""
								objUPSClass.AddNewNode "Code", "GIF" '</Code>"&vbcrlf
							objUPSClass.WriteParent "LabelImageFormat", "/"
						end if
					objUPSClass.WriteParent "LabelSpecification", "/"

					'// Shipment
					objUPSClass.WriteParent "Shipment", ""

						'// For Canada
						if Session("pcAdminInvoiceAmount")<>"" then
							objUPSClass.WriteParent "InvoiceLineTotal", ""
								objUPSClass.AddNewNode "CurrencyCode", Session("pcAdminInvoiceCurrencyCode")
								objUPSClass.AddNewNode "MonetaryValue", Session("pcAdminInvoiceAmount")
							objUPSClass.WriteParent "InvoiceLineTotal", "/"
						end if
						'// For International/Canada
						if Session("pcAdminShipmentDescription")<>"" then
							objUPSClass.AddNewNode "Description", Session("pcAdminShipmentDescription")
						end if

						'// SHIPPER
						objUPSClass.WriteParent "Shipper", ""
							objUPSClass.AddNewNode "Name", UPS_COMPANYNAME
							objUPSClass.AddNewNode "AttentionName", UPS_ATTENTION
							objUPSClass.AddNewNode "PhoneNumber", replace(UPS_PHONE,"-","")
							if UPS_FAX<>"" then
								objUPSClass.AddNewNode "FaxNumber", replace(UPS_FAX,"-","")
							end if
							if UPS_EMAIL<>"" then
								objUPSClass.AddNewNode "EMailAddress", UPS_EMAIL
							end if
							objUPSClass.AddNewNode "ShipperNumber", Session("pcAdminUPSAccountNumber") '</ShipperNumber>"&vbcrlf
							objUPSClass.WriteParent "Address", ""
								objUPSClass.AddNewNode "AddressLine1", UPS_ADDRESS1
								objUPSClass.AddNewNode "AddressLine2", UPS_ADDRESS2
								objUPSClass.AddNewNode "AddressLine3", UPS_ADDRESS3
								objUPSClass.AddNewNode "City", UPS_CITY
								objUPSClass.AddNewNode "StateProvinceCode", UPS_STATE
								objUPSClass.AddNewNode "PostalCode", UPS_POSTALCODE
								objUPSClass.AddNewNode "CountryCode", UPS_COUNTRY
							objUPSClass.WriteParent "Address", "/"
						objUPSClass.WriteParent "Shipper", "/"

						'// SHIP TO
						objUPSClass.WriteParent "ShipTo", ""
							objUPSClass.AddNewNode "CompanyName", Session("pcAdminShipToCompanyName")
							objUPSClass.AddNewNode "AttentionName", Session("pcAdminShipToAttentionName")
							objUPSClass.AddNewNode "PhoneNumber", replace(Session("pcAdminShipToPhoneNumber"),"-","")&Session("pcAdminShipToPhoneNumberExt")
							if Session("pcAdminShipToFaxNumber")<>"" then
								objUPSClass.AddNewNode "FaxNumber", Session("pcAdminShipToFaxNumber")
							end if
							if Session("pcAdminShipToEmailAddress")<>"" then
								objUPSClass.AddNewNode "EmailAddress", Session("pcAdminShipToEmailAddress")
							end if
							objUPSClass.WriteParent "Address", ""
								objUPSClass.AddNewNode "AddressLine1", Session("pcAdminShipToAddressLine1")
								objUPSClass.AddNewNode "AddressLine2", Session("pcAdminShipToAddressLine2")
								objUPSClass.AddNewNode "AddressLine3", Session("pcAdminShipToAddressLine3")
								objUPSClass.AddNewNode "City", Session("pcAdminShipToCity")
								objUPSClass.AddNewNode "StateProvinceCode", Session("pcAdminShipToStateOrProvinceCode")
								objUPSClass.AddNewNode "PostalCode", Session("pcAdminShipToPostalCode")
								objUPSClass.AddNewNode "CountryCode", Session("pcAdminShipToCountryCode")
								if session("pcAdminResidentialDelivery")="1" then
									objUPSClass.AddNewNode "ResidentialAddress", Session("pcAdminResidentialDelivery")
								end if
							objUPSClass.WriteParent "Address", "/"
						objUPSClass.WriteParent "ShipTo", "/"

						'// SHIP FROM
						if session("pcAdminShipToDiffLocation")="Y" then
							objUPSClass.WriteParent "ShipFrom", ""
								objUPSClass.AddNewNode "CompanyName", Session("pcAdminShipFromCompanyName")
								objUPSClass.AddNewNode "AttentionName", Session("pcAdminShipFromAttentionName")
								objUPSClass.AddNewNode "PhoneNumber", replace(Session("pcAdminShipToPhoneNumber"),"-","")&Session("pcAdminShipToPhoneNumberExt")
								objUPSClass.WriteParent "Address", ""
									objUPSClass.AddNewNode "AddressLine1", Session("pcAdminShipFromAddressLine1")
									objUPSClass.AddNewNode "AddressLine2", Session("pcAdminShipFromAddressLine2")
									objUPSClass.AddNewNode "AddressLine3", Session("pcAdminShipFromAddressLine3")
									objUPSClass.AddNewNode "City", Session("pcAdminShipFromCity")
									objUPSClass.AddNewNode "StateProvinceCode", Session("pcAdminShipFromStateOrProvinceCode")
									objUPSClass.AddNewNode "PostalCode", Session("pcAdminShipFromPostalCode")
									objUPSClass.AddNewNode "CountryCode", Session("pcAdminShipFromCountryCode")
								objUPSClass.WriteParent "Address", "/"
							objUPSClass.WriteParent "ShipFrom", "/"
						End If

						'// PAYMENT INFORMATION
						objUPSClass.WriteParent "PaymentInformation", ""
							objUPSClass.WriteParent "Prepaid", ""
								objUPSClass.WriteParent "BillShipper", ""
									objUPSClass.AddNewNode "AccountNumber", Session("pcAdminUPSAccountNumber") '</AccountNumber>"&vbcrlf
								objUPSClass.WriteParent "BillShipper", "/"
							objUPSClass.WriteParent "Prepaid", "/"
						objUPSClass.WriteParent "PaymentInformation", "/"

						'// RATE INFORMATION
						if UPS_USENEGOTIATEDRATES="1" then
							objUPSClass.WriteParent "RateInformation", ""
								objUPSClass.AddNewNode "NegotiatedRatesIndicator", "1"
							objUPSClass.WriteParent "RateInformation", "/"
						end if

						'// SERVICE
						objUPSClass.WriteParent "Service", ""
							objUPSClass.AddNewNode "Code", Session("pcAdminUPSServiceCode")
						objUPSClass.WriteParent "Service", "/"

						'// SHIPMENT SERVICE OPTIONS
						if session("pcAdminSaturdayDelivery")="1" OR session("pcAdminShipmentNotification")="1" OR ("pcAdminOnCallPickup")="1" then
							objUPSClass.WriteParent "ShipmentServiceOptions", ""
								if session("pcAdminSaturdayDelivery")="1" then
									objUPSClass.AddNewNode "SaturdayDelivery", session("pcAdminSaturdayDelivery")
								end if
								if session("pcAdminOnCallPickup")="1" then
									objUPSClass.WriteParent "OnCallAir", ""
										objUPSClass.WriteParent "PickupDetails", ""
											objUPSClass.AddNewNode "PickupDate", session("pcAdminOnCallDate")
											'//if UPSReadyAMPM = "PM" then we add 12
											tempReadyHours=Cint(session("pcAdminUPSReadyHours"))
											tempReadyMinutes=Cint(session("pcAdminUPSReadyMinutes"))
											if session("pcAdminUPSReadyAMPM")="PM" then
												tempReadyHours=tempReadyHours+12
											end if
											tempReadyTime=((tempReadyHours*100)+tempReadyMinutes)
											if tempReadyTime>2400 then
												tempReadyTime=tempReadyTime-2400
											end if
											if len(tempReadyTime)=3 then
												tempReadyTime="0"+Cstr(tempReadyTime)
											end if
											objUPSClass.AddNewNode "EarliestTimeReady", tempReadyTime
											tempPUHours=Cint(session("pcAdminUPSPUHours"))
											tempPUMinutes=Cint(session("pcAdminUPSPUMinutes"))
											if tempPUHours<12 then
												tempPUHours=tempPUHours+12
											end if
											tempPUTime=(tempPUHours*100)+tempPUMinutes
											if len(tempPUTime)=3 then
												tempPUTime="0"+Cstr(tempPUTime)
											end if
											objUPSClass.AddNewNode "LatestTimeReady", tempPUTime

											objUPSClass.WriteParent "ContactInfo", ""
												objUPSClass.AddNewNode "Name", session("pcAdminOnCallContactName")
												objUPSClass.AddNewNode "PhoneNumber", session("pcAdminOnCallContactPhone")
											objUPSClass.WriteParent "ContactInfo", "/"
										objUPSClass.WriteParent "PickupDetails", "/"
									objUPSClass.WriteParent "OnCallAir", "/"
								end if

								for iNotifiCnt=1 to 5
									if session("pcAdminShipmentNotification")="1" AND session("pcAdminNotificationCode"&iNotifiCnt)<>"" AND session("pcAdminNotificationEmail"&iNotifiCnt)<>"" then
									objUPSClass.WriteParent "Notification", ""
										objUPSClass.AddNewNode "NotificationCode", session("pcAdminNotificationCode"&iNotifiCnt)
										objUPSClass.WriteParent "EMailMessage", ""
											objUPSClass.AddNewNode "EMailAddress", session("pcAdminNotificationEmail"&iNotifiCnt)
										objUPSClass.WriteParent "EMailMessage", "/"
									objUPSClass.WriteParent "Notification", "/"
									end if
								next
							objUPSClass.WriteParent "ShipmentServiceOptions", "/"
						end if
						'///////////////////////////////////////////////////////////////////////
						'// START LOOP FOR PACKAGE TAG
						'///////////////////////////////////////////////////////////////////////
						For pcv_xCounter = 1 to pcPackageCount
							'// If the package was processed, skip it.
							if pcLocalArray(pcv_xCounter-1) <> "shipped" then
								objUPSClass.WriteParent "Package", ""
									objUPSClass.WriteParent "PackagingType", ""
										objUPSClass.AddNewNode "Code", session("pcAdminPackageTypeCode"&pcv_xCounter)
									objUPSClass.WriteParent "PackagingType", "/"
									if session("pcAdminPackageTypeCode"&pcv_xCounter)="02" then
										objUPSClass.WriteParent "Dimensions", ""
											objUPSClass.AddNewNode "Length", session("pcAdminLength"&pcv_xCounter)
											objUPSClass.AddNewNode "Width", session("pcAdminWidth"&pcv_xCounter)
											objUPSClass.AddNewNode "Height", session("pcAdminHeight"&pcv_xCounter)
										objUPSClass.WriteParent "Dimensions", "/"
									end if
									if session("pcAdminUPSRefNumber1"&pcv_xCounter)<>"NONE" AND session("pcAdminUPSRefData1"&pcv_xCounter)<>"" then
										objUPSClass.WriteParent "ReferenceNumber", ""
											objUPSClass.AddNewNode "Code", session("pcAdminUPSRefNumber1"&pcv_xCounter)
											objUPSClass.AddNewNode "Value", session("pcAdminUPSRefData1"&pcv_xCounter)
										objUPSClass.WriteParent "ReferenceNumber", "/"
									end if
									if session("pcAdminUPSRefNumber2"&pcv_xCounter)<>"NONE" AND session("pcAdminUPSRefData2"&pcv_xCounter)<>"" then
										objUPSClass.WriteParent "ReferenceNumber", ""
											objUPSClass.AddNewNode "Code", session("pcAdminUPSRefNumber2"&pcv_xCounter)
											objUPSClass.AddNewNode "Value", session("pcAdminUPSRefData2"&pcv_xCounter)
										objUPSClass.WriteParent "ReferenceNumber", "/"
									end if
									'//Package Weight
									objUPSClass.WriteParent "PackageWeight", ""
										objUPSClass.WriteParent "UnitOfMeasurement", ""
											objUPSClass.AddNewNode "Code", session("pcAdminWeightUnitOfMeasurement"&pcv_xCounter)
										objUPSClass.WriteParent "UnitOfMeasurement", "/"
										objUPSClass.AddNewNode "Weight", session("pcAdminPackageWeight"&pcv_xCounter)
									objUPSClass.WriteParent "PackageWeight", "/"

									'// Oversized Package Indicator
									if session("pcAdminOversizePackage"&pcv_xCounter)<>"0" then
										objUPSClass.AddNewNode "OverSizePackage", session("pcAdminOversizePackage"&pcv_xCounter)
									End If

									'// Additional Handling Indicator
									if session("pcAdminAdditionalHandling"&pcv_xCounter)="1" then
										objUPSClass.AddNewNode "AdditionalHandling", session("pcAdminAdditionalHandling"&pcv_xCounter)
									End If

									objUPSClass.WriteParent "PackageServiceOptions", ""
										'// Package Description
										objUPSClass.AddNewNode "Description", "gifts"
										'// Delivery Confirmation
										if session("pcAdminDeliveryConfirmation"&pcv_xCounter)<>"NONE" AND session("pcAdminDeliveryConfirmation"&pcv_xCounter)<>"" then
											objUPSClass.WriteParent "DeliveryConfirmation", ""
												objUPSClass.AddNewNode "DCISType", session("pcAdminDeliveryConfirmation"&pcv_xCounter)
											objUPSClass.WriteParent "DeliveryConfirmation", "/"
										end if
										'// Verbal Confirmation
										if session("pcAdminVerbalConfirmation"&pcv_xCounter)="1" then
											objUPSClass.AddNewNode "VerbalConfirmation", session("pcAdminVerbalConfirmation"&pcv_xCounter)
										end if
										'// COD Indicator
										if session("pcAdminCODPackage"&pcv_xCounter)="1" then
											objUPSClass.WriteParent "COD", ""
												objUPSClass.AddNewNode "CODFundsCode", session("pcAdminCODFundsCode"&pcv_xCounter)
												objUPSClass.AddNewNode "CODCode", "3"
												objUPSClass.WriteParent "CODAmount", ""
													objUPSClass.AddNewNode "CurrencyCode", session("pcAdminCODCurrencyCode"&pcv_xCounter)
													objUPSClass.AddNewNode "MonetaryValue", session("pcAdminCODAmount"&pcv_xCounter)
												objUPSClass.WriteParent "CODAmount", "/"
											objUPSClass.WriteParent "COD", "/"
										end if
										'// Insured Value
										if session("pcAdminInsuredValue"&pcv_xCounter)<>"" then
											objUPSClass.WriteParent "InsuredValue", ""
												objUPSClass.AddNewNode "CurrencyCode", session("pcAdminCODCurrencyCode"&pcv_xCounter)
												objUPSClass.AddNewNode "MonetaryValue", session("pcAdminInsuredValue"&pcv_xCounter)
											objUPSClass.WriteParent "InsuredValue", "/"
										end if
									objUPSClass.WriteParent "PackageServiceOptions", "/"
								objUPSClass.WriteParent "Package", "/"
							End if '// end skip shipped packages
						Next
						'///////////////////////////////////////////////////////////////////////
						'// END LOOP
						'///////////////////////////////////////////////////////////////////////
					objUPSClass.WriteParent "Shipment", "/"
					objUPSClass.WriteParent "ShipmentConfirmRequest", "/"

					'//Clear illegal ampersand characters from XML
					UPS_postdata=replace(UPS_postdata, "&", "and")
					UPS_postdata=replace(UPS_postdata, "andamp;", "and")

					'// Print out our newly formed request xml
					'response.Clear()
					'response.ContentType="text/xml"
					'response.Write(UPS_postdata)
					'response.End()

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Send Our Transaction.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if UPS_TESTMODE="1" then
						UPS_URL="https://wwwcie.ups.com/ups.app/xml/ShipConfirm"
					else
						UPS_URL="https://www.ups.com/ups.app/xml/ShipConfirm"
					end if

					call objUPSClass.SendXMLRequest(UPS_postdata, UPS_URL)

					'// Print out our response
					'response.Clear()
					'response.ContentType="text/xml"
					'response.Write(UPS_result)
					'response.End()

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Load Our Response.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					call objUPSClass.LoadXMLResults(UPS_result)


					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Check for errors from UPS.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					'//SOME ERROR CHECKING HERE
					call objUPSClass.XMLResponseVerify(ErrPageName)

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Redirect with a Message OR complete some task.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if NOT len(pcv_strErrorMsg)>0 then


						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Set Our Response Data to Local.
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

						ResponseStatusCode = objUPSClass.ReadResponseNode("//Response", "ResponseStatusCode")
						ResponseStatusDescription = objUPSClass.ReadResponseNode("//Response", "ResponseStatusDescription")
						session("UPSShippingResponseStatusCode")=ResponseStatusCode
						session("UPSShippingResponseStatusDescription")=ResponseStatusDescription

						'Set Nodes = objOutputXMLDoc.selectNodes("//Error")
						'For Each Node In Nodes
							'ErrorSeverity=Node.selectSingleNode("ErrorSeverity").Text
							'ErrorCode=Node.selectSingleNode("ErrorCode").Text
							'ErrorDescription=Node.selectSingleNode("ErrorDescription").Text
							'session("UPSShippingErrorSeverity")=ErrorSeverity
							'session("UPSShippingErrorCode")=ErrorCode
							'session("UPSShippingErrorDescription")=ErrorDescription
						'Next

						ShipmentCustomerContext = objUPSClass.ReadResponseNode("//TransactionReference", "CustomerContext")
						session("UPSShippingShipmentCustomerContext")=ShipmentCustomerContext

						ShipmentIdentificationNumber = objUPSClass.ReadResponseNode("//ShipmentConfirmResponse", "ShipmentIdentificationNumber")
						ShipmentDigest = objUPSClass.ReadResponseNode("//ShipmentConfirmResponse", "ShipmentDigest")
						session("UPSShippingShipmentIdentificationNumber")=ShipmentIdentificationNumber
						session("UPSShippingShipmentDigest")=ShipmentDigest

						CurrencyCode = objUPSClass.ReadResponseNode("//TransportationCharges", "CurrencyCode")
						MonetaryValue = objUPSClass.ReadResponseNode("//TransportationCharges", "MonetaryValue")
						session("UPS_TC_ShippingCurrencyCode")=CurrencyCode
						session("UPS_TC_ShippingMonetaryValue")=MonetaryValue

						CurrencyCode = objUPSClass.ReadResponseNode("//ServiceOptionsCharges", "CurrencyCode")
						MonetaryValue = objUPSClass.ReadResponseNode("//ServiceOptionsCharges", "MonetaryValue")
						session("UPS_SOC_ShippingCurrencyCode")=CurrencyCode
						session("UPS_SOC_ShippingMonetaryValue")=MonetaryValue

						CurrencyCode = objUPSClass.ReadResponseNode("//TotalCharges", "CurrencyCode")
						MonetaryValue = objUPSClass.ReadResponseNode("//TotalCharges", "MonetaryValue")
						session("UPSShippingCurrencyCode")=CurrencyCode
						session("UPSShippingMonetaryValue")=MonetaryValue

						if UPS_USENEGOTIATEDRATES="1" then
							CurrencyCode = objUPSClass.ReadResponseNode("//GrandTotal", "CurrencyCode")
							MonetaryValue = objUPSClass.ReadResponseNode("//GrandTotal", "MonetaryValue")
							session("UPSNegShippingCurrencyCode")=CurrencyCode
							session("UPSNegShippingMonetaryValue")=MonetaryValue
						end if
						pcv_usedNegotiatedRateToDisplay = 0
						If session("UPSNegShippingMonetaryValue")&""<>"" Then
							pcv_usedNegotiatedRateToDisplay = 1
							session("UPSShippingCurrencyCode")=session("UPSNegShippingCurrencyCode")
							session("UPSShippingMonetaryValue")=session("UPSNegShippingMonetaryValue")
						End If

						Weight = objUPSClass.ReadResponseNode("//BillingWeight", "Weight")
						session("UPSShippingWeight")=Weight
						%>

						<form name="form1" action="pcUPSConfirmLabel.asp" method="post" class="pcForms">
							<input type="hidden" name="pcIntPackageInfo_ID" value="<%=pcIntPackageInfo_ID%>">

							<p style="padding-top: 5px;">
							<span style="font-weight: bold">Response Status:</span> <%=session("UPSShippingResponseStatusDescription")%><br>
								<br>
								<span style="font-weight: bold">Retails Shipment Charges</span><br>
								&nbsp;&nbsp;Transportation Charges: <%=session("UPS_TC_ShippingMonetaryValue")%> <br>
								&nbsp;&nbsp;Service Option Charges: <%=session("UPS_SOC_ShippingMonetaryValue")%> <br>
                                <% If pcv_usedNegotiatedRateToDisplay = 1 Then %>
                               		<br><span style="font-weight: bold">Negotiated Shipment Charges</span><br>
								&nbsp;&nbsp;Total Charges: <%=session("UPSShippingMonetaryValue")%> <br>
                                <% Else %>
									&nbsp;&nbsp;Total Charges: <%=session("UPSShippingMonetaryValue")%> <br>
                               	<% End If %>
							</p>
							<p>&nbsp;</p>
							<p>
								<input name='back' type='button' value='Change Package' onClick='javascript:history.go(-1)'>&nbsp;
								<input name="Continue" type="submit" value="Confirm Shipment" class="submit2">
							</p>
							<p>&nbsp;</p>
							<p>&nbsp;</p>
							<div align="center"><table><tr>
					  <td valign="top"><div align="center">
					   <%= pcf_UPSWriteLegalDisclaimers %>
						</div></td>
					  </tr></table></div>
						</form>

					<% End If ' If pcv_intErr>0 Then
				end if
				call closedb()
			else
				call opendb()
				'*******************************************************************************
				' END: ON POSTBACK
				'*******************************************************************************

				'*******************************************************************************
				' START: LOAD HTML FORM
				'*******************************************************************************

				msg=request.querystring("msg")
				if instr(lcase(msg),"invoicelinetotal/monetaryvalue") then
					msg="Invoice Line Total amount is required for this shipment.<br>Click on the &quot;Recipient&quot; tab and enter in a Monetary Value and Currency Code Type."
				end if
				if instr(lcase(msg),"shipment description is required") then
					msg="Shipment Description is required for this shipment.<br>Click on the &quot;Recipient&quot; tab and enter in a Shipment Description."
				end if
				if instr(lcase(msg),"the selected service is not available") then
					msg="The selected service is not available for this shipment.<br>Click on the &quot;Ship Settings&quot; tab and choose different Type of service for this shipment."
				end if
				if msg<>"" then %>
					<div class="pcCPmessage">
						<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
					</div>
				<% end if %>

				<% if request("sub")=1 then %>
					<% if Session("pcAdminPayorType") ="" then
						Session("pcAdminPayorType")=Session("pcAdminUPSPayorType")
					end if %>
					<% if Session("pcAdminSaturdayDelivery")="" then
						Session("pcAdminSaturdayDelivery")=Session("pcAdminUPSSaturdayDelivery")
					end if %>
					<%
					if session("pcAdminShipmentNotification")="" then
						session("pcAdminShipmentNotification")=Session("pcAdminUPSShipmentNotification")
					end if
					if session("pcAdminNotificationCode1")="" then
						session("pcAdminNotificationCode1")=Session("pcAdminUPSNotifiCode1")
					end if
					if session("pcAdminNotificationCode2")="" then
						session("pcAdminNotificationCode2")=Session("pcAdminUPSNotifiCode2")
					end if
					if session("pcAdminNotificationCode3")="" then
						session("pcAdminNotificationCode3")=Session("pcAdminUPSNotifiCode3")
					end if
					if session("pcAdminNotificationCode4")="" then
						session("pcAdminNotificationCode4")=Session("pcAdminUPSNotifiCode4")
					end if
					if session("pcAdminNotificationCode5")="" then
						session("pcAdminNotificationCode5")=Session("pcAdminUPSNotifiCode5")
					end if
					if session("pcAdminNotificationEmail1")="" then
						session("pcAdminNotificationEmail1")=Session("pcAdminUPSNotifiEmail1")
					end if
					if session("pcAdminNotificationEmail2")="" then
						session("pcAdminNotificationEmail2")=Session("pcAdminUPSNotifiEmail2")
					end if
					if session("pcAdminNotificationEmail3")="" then
						session("pcAdminNotificationEmail3")=Session("pcAdminUPSNotifiEmail3")
					end if
					if session("pcAdminNotificationEmail4")="" then
						session("pcAdminNotificationEmail4")=Session("pcAdminUPSNotifiEmail4")
					end if
					if session("pcAdminNotificationEmail5")="" then
						session("pcAdminNotificationEmail5")=Session("pcAdminUPSNotifiEmail5")
					end if
				else
					Session("pcAdminPayorType")=Session("pcAdminUPSPayorType")
					Session("pcAdminSaturdayDelivery")=Session("pcAdminUPSSaturdayDelivery")
					session("pcAdminShipmentNotification")=Session("pcAdminUPSShipmentNotification")
					session("pcAdminNotificationCode1")=Session("pcAdminUPSNotifiCode1")
					session("pcAdminNotificationCode2")=Session("pcAdminUPSNotifiCode2")
					session("pcAdminNotificationCode3")=Session("pcAdminUPSNotifiCode3")
					session("pcAdminNotificationCode4")=Session("pcAdminUPSNotifiCode4")
					session("pcAdminNotificationCode5")=Session("pcAdminUPSNotifiCode5")
					session("pcAdminNotificationEmail1")=Session("pcAdminUPSNotifiEmail1")
					session("pcAdminNotificationEmail2")=Session("pcAdminUPSNotifiEmail2")
					session("pcAdminNotificationEmail3")=Session("pcAdminUPSNotifiEmail3")
					session("pcAdminNotificationEmail4")=Session("pcAdminUPSNotifiEmail4")
					session("pcAdminNotificationEmail5")=Session("pcAdminUPSNotifiEmail5")
				end if %>

				<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">
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
								<td colspan="2"><span class="title">Ship Settings:</span></td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2">Service Settings  </th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<%
							dim pcv_ShowCustSelection, pcv_SelectionStr
							pcv_ShowCustSelection=0
							pcv_SelectionStr=""

							if pcv_SRF="1" then
								pcv_ShowCustSelection=1
								pcv_SelectionStr = pcv_ShipmentDetails
							else
								if varShip<>"0" AND Service<>"" then
									If Service = "UPS Next Day Air" or Service = "UPS 2nd Day Air" or Service = "UPS Ground" or Service = "UPS Worldwide Express " or Service = "UPS Worldwide Expedited " or Service = "UPS Standard To Canada" or Service = "UPS 3-Day Select " or Service="UPS 3 Day Select" or Service = "UPS Next Day Air Saver" or Service = "UPS Next Day Air Early A.M." or Service = "UPS Worldwide Express Plus " or Service = "UPS 2nd Day Air A.M." then
										Session("pcAdminUPSServiceCode") = ServiceCode
									else
										pcv_ShowCustSelection=1
										pcv_SelectionStr = Service
									end if
								else
									pcv_ShowCustSelection=1
									pcv_SelectionStr = pcv_ShipmentDetails
								end if
							end if %>

							<% 'Show customer's selection if not a match for a UPS default service type
							if pcv_ShowCustSelection=1 then %>
								<tr>
									<td width="23%" align="right" valign="top"><b>Customer selected:</b></td>
									<td align="left" valign="top">
									<%=pcv_SelectionStr%>
									</td>
								</tr>
							<% end if %>

							<tr>
								<td width="23%" align="right" valign="top"><b>Type of service:</b></td>
								<td width="77%" align="left">
								<%
								'// Set Carrier Code to local
								pcv_strDropOptions = Session("pcAdminUPSServiceCode")
								%>
								<select name="UPSServiceCode">
									<option value="01" <%=pcf_SelectOption("UPSServiceCode","01")%>>UPS Next Day Air&reg;</option>
									<option value="02" <%=pcf_SelectOption("UPSServiceCode","02")%>>UPS 2nd Day Air&reg;</option>
									<option value="03" <%=pcf_SelectOption("UPSServiceCode","03")%>>UPS Ground</option>
									<option value="07" <%=pcf_SelectOption("UPSServiceCode","07")%>>UPS Worldwide Express <sup>	SM</sup></option>
									<option value="08" <%=pcf_SelectOption("UPSServiceCode","08")%>>UPS Worldwide Expedited <sup>SM</sup></option>
									<option value="11" <%=pcf_SelectOption("UPSServiceCode","11")%>>UPS Standard To Canada</option>
									<option value="12" <%=pcf_SelectOption("UPSServiceCode","12")%>>UPS 3-Day Select <sup>SM</sup></option>
									<option value="13" <%=pcf_SelectOption("UPSServiceCode","13")%>>UPS Next Day Air Saver&reg;</option>
									<option value="14" <%=pcf_SelectOption("UPSServiceCode","14")%>>UPS Next Day Air&reg; Early A.M.&reg;</option>
									<option value="54" <%=pcf_SelectOption("UPSServiceCode","54")%>>UPS Worldwide Express Plus <sup>SM</sup></option>
									<option value="59" <%=pcf_SelectOption("UPSServiceCode","59")%>>UPS 2nd Day Air A.M.&reg;</option>
								</select>
								<%pcs_UPSRequiredImageTag "UPSServiceCode", true %>											</td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
								<th colspan="2">Shipment Billing</th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
								<td align="right" valign="top"><b>BillShipper:</b></td>
								<td align="left">
									<select name="PayorType" id="PayorType" ONCHANGE="whatPayorTypeSelected();">
									<option value="PrePaid" <%=pcf_SelectOption("PayorType","PrePaid")%>>PrePaid</option>
									<option value="BillThirdParty" <%=pcf_SelectOption("PayorType","BillThirdParty")%>>Bill 3rd Party</option>
									<option value="ConsigneeBilled" <%=pcf_SelectOption("PayorType","ConsigneeBilled")%>>Consignee Billing</option>
									<option value="FreightCollect" <%=pcf_SelectOption("PayorType","FreightCollect")%>>Freight Collect</option>
									</select>
									<%pcs_UPSRequiredImageTag "PayorType", true%></td>
							</tr>
							<tr>
							  <td align="left" valign="top" colspan="2">
								<table id="PayorType_table" <% if Session("pcAdminUPSPayorType")="PrePaid" then%>style="display:none"<%end if %>>
									<tr>
									<td width="134" align="right" valign="top"><b>Billing Account Number:</b></td>
									<td width="463" align="left">
									<input name="PayorAccountNumber" type="text" id="PayorAccountNumber" value="<%=pcf_FillFormField("PayorAccountNumber", false)%>">
									<%pcs_UPSRequiredImageTag "PayorAccountNumber", false%>		</td>
									</tr>
									<tr>
									<td align="right" valign="top"><b>Billing Country Code:</b></td>
									<td align="left">
									<input name="PayorCountryCode" type="text" id="PayorCountryCode" value="<%=pcf_FillFormField("PayorCountryCode", false)%>">
									<%pcs_UPSRequiredImageTag "PayorCountryCode", false%>		</td>
									</tr>
								</table>
							  </td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							<th colspan="2">Labels</th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							<td align="right" valign="top"><b>Image Type:</b></td>
							<td align="left">
							<select name="ImageType" id="ImageType">
							<option value="GIF" <%=pcf_SelectOption("ImageType","GIF")%>>GIF (Plain Paper)</option>
							</select>
							<%pcs_UPSRequiredImageTag "ImageType", true%>			</td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							<th colspan="2">Special Services</th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
				<td align="right"><INPUT tabIndex="25" type="checkbox" value="1" name="SaturdayDelivery" class="clearBorder" <%=pcf_CheckOption("SaturdayDelivery", "1")%>></td>
							  <td><strong>Saturday Delivery </strong> </td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
							  <td align="right"><INPUT tabIndex="25" type="checkbox" value="1" name="OnCallPickup" class="clearBorder" <%=pcf_CheckOption("OnCallPickup", "1")%>></td>
							  <td><span style="font-weight: bold">UPS On Call Pickup&reg; </span></td>
						  </tr>
							<tr>
							  <td>&nbsp;</td>
							  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr class="pcCPcontent">
					<td align="left">Date:</td>
					<td align="left">
										<select name="OnCallDate">
											<% dtTodayDate=Date()
											Function UPSDateFormat (UPSDate)
												UPSDay=Day(UPSDate)
												UPSMonth=Month(UPSDate)
												UPSYear= Year(UPSDate)
												UPSDateFormat=UPSYear&Right(Cstr(UPSMonth + 100),2)&Right(Cstr(UPSDay + 100),2)
											End Function %>
											<option value="<%=UPSDateFormat(dtTodayDate)%>" <%=pcf_SelectOption("ShipDate",UPSDateFormat(dtTodayDate))%>>Today</option>
											<% for d=1 to 5
												if DatePart("W", dtTodayDate+d, VBSUNDAY)=1 then
												else %>
												<option value="<%=UPSDateFormat((dtTodayDate+d))%>" <%=pcf_SelectOption("ShipDate",UPSDateFormat(dtTodayDate+d))%>><%=FormatDateTime((dtTodayDate+d), 1)%></option>
												<% end if
											next %>
										</select></td>
				  </tr>
				  <tr class="pcCPcontent">
					<td width="25%" align="left">Shipment Ready At  :</td>
					<td width="75%" align="left"><select name="UPSReadyHours" id="UPSReadyHours">
						<option value="01" <%=pcf_SelectOption("UPSReadyHours","01")%>>01</option>
						<option value="02" <%=pcf_SelectOption("UPSReadyHours","02")%>>02</option>
						<option value="03" <%=pcf_SelectOption("UPSReadyHours","03")%>>03</option>
						<option value="04" <%=pcf_SelectOption("UPSReadyHours","04")%>>04</option>
						<option value="05" <%=pcf_SelectOption("UPSReadyHours","05")%>>05</option>
						<option value="06" <%=pcf_SelectOption("UPSReadyHours","06")%>>06</option>
						<option value="07" <%=pcf_SelectOption("UPSReadyHours","07")%>>07</option>
						<option value="08" <%=pcf_SelectOption("UPSReadyHours","08")%>>08</option>
						<option value="09" <%=pcf_SelectOption("UPSReadyHours","09")%>>09</option>
						<option value="10" <%=pcf_SelectOption("UPSReadyHours","10")%>>10</option>
						<option value="11" <%=pcf_SelectOption("UPSReadyHours","11")%>>11</option>
						<option value="12" <%=pcf_SelectOption("UPSReadyHours","12")%>>12</option>
					  </select>
					  :
					  <select name="UPSReadyMinutes" id="UPSReadyMinutes">
						<option value="00" <%=pcf_SelectOption("UPSReadyMinutes","00")%>>00</option>
						<option value="01" <%=pcf_SelectOption("UPSReadyMinutes","01")%>>01</option>
						<option value="02" <%=pcf_SelectOption("UPSReadyMinutes","02")%>>02</option>
						<option value="03" <%=pcf_SelectOption("UPSReadyMinutes","03")%>>03</option>
						<option value="04" <%=pcf_SelectOption("UPSReadyMinutes","04")%>>04</option>
						<option value="05" <%=pcf_SelectOption("UPSReadyMinutes","05")%>>05</option>
						<option value="06" <%=pcf_SelectOption("UPSReadyMinutes","06")%>>06</option>
						<option value="07" <%=pcf_SelectOption("UPSReadyMinutes","07")%>>07</option>
						<option value="08" <%=pcf_SelectOption("UPSReadyMinutes","08")%>>08</option>
						<option value="09" <%=pcf_SelectOption("UPSReadyMinutes","09")%>>09</option>
						<% for iHHCnt=10 to 59
					response.write "<option value="""&iHHCnt&""" "&pcf_SelectOption("UPSReadyMinutes",""&iHHCnt&"")&">"&iHHCnt&"</option>"
				next %>
					  </select>
					  &nbsp;
					  <input name="UPSReadyAMPM" type="radio" value="AM" <%=pcf_CheckOption("UPSReadyAMPM","AM")%> class="clearBorder">
					  A.M.
					  &nbsp;
					  <input name="UPSReadyAMPM" type="radio" value="PM" <%=pcf_CheckOption("UPSReadyAMPM","PM")%> class="clearBorder">
					  P.M. </td>
				  </tr>
				  <tr class="pcCPcontent">
					<td align="left">Pick Up by :</td>
					<td align="left"><select name="UPSPUHours" id="UPSPUHours">
						<option value="12" <%=pcf_SelectOption("UPSPUHours","12")%>>12</option>
						<option value="01" <%=pcf_SelectOption("UPSPUHours","01")%>>01</option>
						<option value="02" <%=pcf_SelectOption("UPSPUHours","02")%>>02</option>
						<option value="03" <%=pcf_SelectOption("UPSPUHours","03")%>>03</option>
						<option value="04" <%=pcf_SelectOption("UPSPUHours","04")%>>04</option>
						<option value="05" <%=pcf_SelectOption("UPSPUHours","05")%>>05</option>
						<option value="06" <%=pcf_SelectOption("UPSPUHours","06")%>>06</option>
						<option value="07" <%=pcf_SelectOption("UPSPUHours","07")%>>07</option>
						<option value="08" <%=pcf_SelectOption("UPSPUHours","08")%>>08</option>
						<option value="09" <%=pcf_SelectOption("UPSPUHours","09")%>>09</option>
						<option value="10" <%=pcf_SelectOption("UPSPUHours","10")%>>10</option>
						<option value="11" <%=pcf_SelectOption("UPSPUHours","11")%>>11</option>
					  </select>
					  :
					  <select name="UPSPUMinutes" id="UPSPUMinutes">
						<option value="00" <%=pcf_SelectOption("UPSPUMinutes","00")%>>00</option>
						<option value="01" <%=pcf_SelectOption("UPSPUMinutes","01")%>>01</option>
						<option value="02" <%=pcf_SelectOption("UPSPUMinutes","02")%>>02</option>
						<option value="03" <%=pcf_SelectOption("UPSPUMinutes","03")%>>03</option>
						<option value="04" <%=pcf_SelectOption("UPSPUMinutes","04")%>>04</option>
						<option value="05" <%=pcf_SelectOption("UPSPUMinutes","05")%>>05</option>
						<option value="06" <%=pcf_SelectOption("UPSPUMinutes","06")%>>06</option>
						<option value="07" <%=pcf_SelectOption("UPSPUMinutes","07")%>>07</option>
						<option value="08" <%=pcf_SelectOption("UPSPUMinutes","08")%>>08</option>
						<option value="09" <%=pcf_SelectOption("UPSPUMinutes","09")%>>09</option>
						<% for iHHCnt=10 to 59
					response.write "<option value="""&iHHCnt&""" "&pcf_SelectOption("UPSPUMinutes",""&iHHCnt&"")&">"&iHHCnt&"</option>"
				next %>
					  </select>
					  P.M. </td>
				  </tr>

				  <tr>
					<td>Contact Name: </td>
										<td><input name="OnCallContactName" type="text" id="OnCallContactName" value="<%=pcf_FillFormField("OnCallContactName", false)%>">
										<%pcs_UPSRequiredImageTag "OnCallContactName", false%></td>
				  </tr>
				  <tr>
					<td>Contact Phone: </td>
										<td><input name="OnCallContactPhone" type="text" id="OnCallContactPhone" value="<%=pcf_FillFormField("OnCallContactPhone", false)%>">
										<%pcs_UPSRequiredImageTag "OnCallContactPhone", false%></td>
				  </tr>
				  <tr>
					<td>&nbsp;</td>
					<td>Both &quot;Contact Name&quot; and &quot;Contact Phone&quot; <br />
					  are required for UPS OnCall Pickup.</td>
				  </tr>
				</table></td>
						  </tr>
							<tr>
							  <td>&nbsp;</td>
							  <td></td>
						  </tr>
						</table>
						</div>


						<!--
						/////////////////////////////////////////////////////////////////////////////////
						// SHIPPER
						//////////////////////////////////////////////////////////////////////////////////
						-->
						<% if IntResetSessions=1 then
							'session("pcAdminShipToDiffLocation")=""
							session("pcAdminShipFromCompanyName")=""
							session("pcAdminShipFromAttentionName")=""
							session("pcAdminShipFromPhoneNumber")=""
							session("pcAdminShipFromPhoneNumberExt")=""
							session("pcAdminShipFromFaxNumber")=""
							session("pcAdminShipFromEmailAddress")=""
							session("pcAdminShipFromAddressLine1")=""
							session("pcAdminShipFromAddressLine2")=""
							session("pcAdminShipFromAddressLine3")=""
							session("pcAdminShipFromCity")=""
							session("pcAdminShipFromPostalCode")=""
							session("pcAdminShipFromCountryCode")=""
							session("pcAdminShipFromProvinceCode")=""
							session("pcAdminShipFromStateCode")=""
							session("pcAdminShipFromStateOrProvinceCode")=""
							Session("pcAdminShipToFaxNumber")=""
							Session("pcAdminShipToProvince")=""
							Session("pcAdminPackageCount")=""
						end if %>
						<div id="tab2" class="panes">
						<table class="pcCPcontent">
							<tr>
							<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
							<td colspan="2"><span class="title">Shipper:</span></td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							  <td colspan="2">Shipper Details are set via the UPS settings page. </td>
							</tr>
							<tr>
							  <td colspan="2">
								<% response.write "<b>Company Name: </b>"&UPS_COMPANYNAME&"<BR>" %>
								<% if len(UPS_ATTENTION)>0 then
									response.write "<b>Attention Name: </b>"&UPS_ATTENTION&"<BR>"
								end if %>
								<% response.write "<b>Address: </b><BR>&nbsp;&nbsp;"&UPS_ADDRESS1&"<BR>" %>
								<% if len(UPS_ADDRESS2)>0 then
									response.write "&nbsp;&nbsp;"&UPS_ADDRESS2&"<BR>"
								end if %>
								<% if len(UPS_ADDRESS3)>0 then
									response.write "&nbsp;&nbsp;"&UPS_ADDRESS3&"<BR>"
								end if %>
								<% response.write "&nbsp;&nbsp;"&UPS_CITY %><% if len(UPS_STATE)>0 then %>, &nbsp;<% response.write UPS_STATE %><% end if %>
								<br>
								<% if len(UPS_POSTALCODE)>0 then
									response.write "&nbsp;&nbsp;"&UPS_POSTALCODE&"<BR>"
								end if %>
								<% response.write "&nbsp;&nbsp;"&UPS_COUNTRY&"<BR>" %>
								<% if len(UPS_PHONE)>0 then
									response.write "<b>Phone: </b>"& UPS_PHONE&"<BR>"
								end if %>
								<% if len(UPS_FAX)>0 then
									response.write "<b>FAX: </b>"&UPS_FAX&"<BR>"
								end if %>
							  </td>
						  </tr>
							<tr>
							<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
				<td colspan="2"><span class="title">Ship From:</span></td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							  <td colspan="2">The Ship From is only required if the pickup location is different from the Shipper's address.</td>
							</tr>
							<tr>
							  <td colspan="2">Will this package be shipped from a different location then the &quot;Shipper&quot;?
							  <input name="ShipToDiffLocation" type="radio" value="N" class="clearBorder" <%=pcf_CheckOption("ShipToDiffLocation", "N")%>>
No
							&nbsp;
							<input name="ShipToDiffLocation" type="radio" value="Y" class="clearBorder" <%=pcf_CheckOption("ShipToDiffLocation", "Y")%>>
							Yes</td>
						  </tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							<th colspan="2">Contact Details</th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							<td><p>Company Name:</p></td>
							<td>
							<p>
							<input name="ShipFromCompanyName" type="text" id="ShipFromCompanyName" value="<%=pcf_FillFormField("ShipFromCompanyName", false)%>">
							<%pcs_UPSRequiredImageTag "ShipFromCompanyName", false %>
							</p>
							</td>
							</tr>
							<tr>
							<td width="25%"><p>Attention Name:</p></td>
							<td width="75%">
							<p>
							<input name="ShipFromAttentionName" type="text" id="ShipFromAttentionName" value="<%=pcf_FillFormField("ShipFromAttentionName", true)%>">
							<%pcs_UPSRequiredImageTag "ShipFromAttentionName", false %>
							</p>
							</td>
							</tr>
							<% if len(Session("ErrShipFromPhoneNumber"))>0 then %>
							<tr>
							<td colspan="2">
							<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
							You must enter a valid Phone Number.			</td>
							</tr>
							<% end if %>
							<tr>
							<td><p>Phone Number:</p></td>
							<td>
							<p>
							<input name="ShipFromPhoneNumber" type="text" id="ShipFromPhoneNumber" value="<%=pcf_FillFormField("ShipFromPhoneNumber", true)%>">
							<%pcs_UPSRequiredImageTag "ShipFromPhoneNumber", false %>
							&nbsp;Ext:
							<input name="ShipFromPhoneNumberExt" type="text" id="ShipFromPhoneNumberExt" value="<%=pcf_FillFormField("ShipFromPhoneNumberExt", false)%>" size="4" maxlength="4">
							<%pcs_UPSRequiredImageTag "ShipFromPhoneNumberExt", false %>
							</p>
							</td>
							</tr>
							<tr>
							<td><p>Fax Number:</p></td>
							<td>
							<p>
							<input name="ShipFromFaxNumber" type="text" id="ShipFromFaxNumber" value="<%=pcf_FillFormField("ShipFromFaxNumber", false)%>">
							<%pcs_UPSRequiredImageTag "ShipFromFaxNumber", false%>
							</p>
							</td>
							</tr>
							<% if len(Session("ErrShipFromEmailAddress"))>0 then %>
							<tr>
							<td colspan="2">
							<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
							You must enter a valid Email Address.			</td>
							</tr>
							<% end if %>
							<tr>
							<td><p>Email Address:</p></td>
							<td>
							<p>
							<input name="ShipFromEmailAddress" type="text" id="ShipFromEmailAddress" value="<%=pcf_FillFormField("ShipFromEmailAddress", true)%>">
							<%pcs_UPSRequiredImageTag "ShipFromEmailAddress", false %>
							</p>
							</td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							<th colspan="2">Location Details</th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
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
							pcv_isStateCodeRequired = false '// determines if validation is performed (true or false)
							pcv_isProvinceCodeRequired = false '// determines if validation is performed (true or false)
							pcv_isCountryCodeRequired = false '// determines if validation is performed (true or false)

							'// #3 Additional Required Info
							pcv_strTargetForm = "form1" '// Name of Form
							pcv_strCountryBox = "ShipFromCountryCode" '// Name of Country Dropdown
							pcv_strTargetBox = "ShipFromStateCode" '// Name of State Dropdown
							pcv_strProvinceBox =  "ShipFromProvinceCode" '// Name of Province Field
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
							<td><p>Address Line 1:</p></td>
							<td>
							<p>
							<input name="ShipFromAddressLine1" type="text" id="ShipFromAddressLine1" value="<%=pcf_FillFormField("ShipFromAddressLine1", true)%>">
							<%pcs_UPSRequiredImageTag "ShipFromAddressLine1", false %>
							</p>
							</td>
							</tr>
							<tr>
							<td><p>Address Line 2:</p></td>
							<td>
							<p>
							<input name="ShipFromAddressLine2" type="text" id="ShipFromAddressLine2" value="<%=pcf_FillFormField("ShipFromAddressLine2", false)%>">
							<%pcs_UPSRequiredImageTag "ShipFromAddressLine2", false %>
							</p>
							</td>
							</tr>
							<tr>
								<td><p>Address Line 3:</p></td>
								<td>
								<p>
								<input name="ShipFromAddressLine3" type="text" id="ShipFromAddressLine3" value="<%=pcf_FillFormField("ShipFromAddressLine3", false)%>">
								<%pcs_UPSRequiredImageTag "ShipFromAddressLine3", false%>
								</p>
								</td>
							</tr>
							<tr>
							<td><p>City:</p></td>
							<td>
							<p>
							<input name="ShipFromCity" type="text" id="ShipFromCity" value="<%=pcf_FillFormField("ShipFromCity", true)%>">
							<%pcs_UPSRequiredImageTag "ShipFromCity", false %>
							</p>
							</td>
							</tr>

							<%
							'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
							pcs_StateProvince
							%>

							<tr>
							<td><p>Postal Code:</p></td>
							<td>
							<p>
							<input name="ShipFromPostalCode" type="text" id="ShipFromPostalCode" value="<%=pcf_FillFormField("ShipFromPostalCode", true)%>">
							<%pcs_UPSRequiredImageTag "ShipFromPostalCode", false %>
							</p>
							</td>
							</tr>

							<tr>
								<td colspan="2" class="pcCPspacer"></td>
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
							<td colspan="2"><span class="title">Recipient:</span></td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							<th colspan="2">Contact Details</th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							<td width="25%"><p>Company Name:</p></td>
							<td width="75%">
							<p>
							<input name="ShipToCompanyName" type="text" id="ShipToCompanyName" value="<%=pcf_FillFormField("ShipToCompanyName", true)%>">
							<%pcs_UPSRequiredImageTag "ShipToCompanyName", true%>
							</p>
							</td>
							</tr>
							<tr>
							<td><p>Attention Name:</p></td>
							<td>
							<p>
							<input name="ShipToAttentionName" type="text" id="ShipToAttentionName" value="<%=pcf_FillFormField("ShipToAttentionName", false)%>">
							<%pcs_UPSRequiredImageTag "ShipToAttentionName", true%>
							</p>
							</td>
							</tr>
							<% if len(Session("ErrShipToPhoneNumber"))>0 then %>
							<tr>
							<td colspan="2">
							<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
							You must enter a valid Phone Number.			</td>
							</tr>
							<% end if %>
							<tr>
							<td><p>Phone Number:</p></td>
							<td>
							<p>
							<input name="ShipToPhoneNumber" type="text" id="ShipToPhoneNumber" value="<%=pcf_FillFormField("ShipToPhoneNumber", true)%>">
							<%pcs_UPSRequiredImageTag "ShipToPhoneNumber", true%>
							&nbsp;Ext.
							<input name="ShipToPhoneNumberExt" type="text" id="ShipToPhoneNumberExt" value="<%=pcf_FillFormField("ShipToPhoneNumberExt", false)%>" size="4" maxlength="4">
							<%pcs_UPSRequiredImageTag "ShipToPhoneNumberExt", false%>
							</p>
							</td>
							</tr>
							<tr>
							<td><p>Fax Number:</p></td>
							<td>
							<p>
							<input name="ShipToFaxNumber" type="text" id="ShipToFaxNumber" value="<%=pcf_FillFormField("ShipToFaxNumber", false)%>">
							<%pcs_UPSRequiredImageTag "ShipToFaxNumber", false%>
							</p>
							</td>
							</tr>
							<% if len(Session("ErrShipToEmailAddress"))>0 then %>
							<tr>
							<td colspan="2">
							<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
							You must enter a valid Email Address.			</td>
							</tr>
							<% end if %>
							<tr>
							<td><p>Email Address:</p></td>
							<td>
							<p>
							<input name="ShipToEmailAddress" type="text" id="ShipToEmailAddress" value="<%=pcf_FillFormField("ShipToEmailAddress", true)%>">
							<%pcs_UPSRequiredImageTag "ShipToEmailAddress", false %>
							</p>
							</td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
						  </tr>
							<tr>
							<th colspan="2">Location Details</th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
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
							pcv_isStateOrProvinceCodeRequired = isRequiredRecipState '// determines if validation is performed (true or false)
							pcv_isProvinceCodeRequired = isRequiredRecipProvince '// determines if validation is performed (true or false)
							pcv_isCountryCodeRequired = true '// determines if validation is performed (true or false)

							'// #3 Additional Required Info
							pcv_strTargetForm = "form1" '// Name of Form
							pcv_strCountryBox = "ShipToCountryCode" '// Name of Country Dropdown
							pcv_strTargetBox = "ShipToStateOrProvinceCode" '// Name of State Dropdown
							pcv_strProvinceBox =  "ShipToProvince" '// Name of Province Field

							'// Set local Country to Session
							if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
							end if

							'// Set local State to Session
							if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session("pcAdminShipToStateOrProvinceCode")
							end if

							'// Set local Province to Session
							if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session("pcAdminShipToStateOrProvinceCode")
							end if

							'///////////////////////////////////////////////////////////
							'// END: COUNTRY AND STATE/ PROVINCE CONFIG
							'///////////////////////////////////////////////////////////
							%>

							<%
							'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
							pcs_CountryDropdown
							%>
							<tr>
							<td><p>Address Line 1:</p></td>
							<td>
							<p>
							<input name="ShipToAddressLine1" type="text" id="ShipToAddressLine1" value="<%=pcf_FillFormField("ShipToAddressLine1", true)%>">
							<%pcs_UPSRequiredImageTag "ShipToAddressLine1", true%>
							</p>
							</td>
							</tr>
							<tr>
							<td><p>Address Line 2:</p></td>
							<td>
							<p>
							<input name="ShipToAddressLine2" type="text" id="ShipToAddressLine2" value="<%=pcf_FillFormField("ShipToAddressLine2", false)%>">
							<%pcs_UPSRequiredImageTag "ShipToAddressLine2", false%>
							</p>
							</td>
							</tr>
							<tr>
								<td><p>Address Line 3:</p></td>
								<td>
								<p>
								<input name="ShipToAddressLine3" type="text" id="ShipToAddressLine3" value="<%=pcf_FillFormField("ShipToAddressLine3", false)%>">
								<%pcs_UPSRequiredImageTag "ShipToAddressLine3", false%>
								</p>
								</td>
						  </tr>
							<tr>
							<td><p>City:</p></td>
							<td>
							<p>
							<input name="ShipToCity" type="text" id="ShipToCity" value="<%=pcf_FillFormField("ShipToCity", true)%>">
							<%pcs_UPSRequiredImageTag "ShipToCity", true%>
							</p>
							</td>
							</tr>

							<%
							'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
							pcs_StateProvince
							%>

							<tr>
							<td><p>Postal Code:</p></td>
							<td>
							<p>
							<input name="ShipToPostalCode" type="text" id="ShipToPostalCode" value="<%=pcf_FillFormField("ShipToPostalCode", isRequiredRecipPostal)%>">
							<%pcs_UPSRequiredImageTag "ShipToPostalCode", isRequiredRecipPostal %>
							</p>
							</td>
							</tr>


							<tr>
							<td></td>
							<td>
							<p>
							<input type="checkbox" name="ResidentialDelivery" value="1" class="clearBorder" <%=pcf_CheckOption("ResidentialDelivery", "1")%> <% 	if pcv_OrdShipType=0 then 'Residential
%>checked<%end if%>>
							<strong>This is a Residential Delivery</strong>
							</p>
							</td>
							</tr>
							<tr>
							<td></td>
							<td><p>
							<%
							if ucase(strShipToCountry)="US" then %>
							<br>
							<a href="javascript:;" onclick="newWindow('UPS_AVPopup.asp?State=<%=strShipToStateCode%>&City=<%=strShipToCity%>&PC=<%=strShipToPostalCode%>&Country=<%=strShipToCountry%>','ProductWindow')">UPS OnLine&reg; Tools Address Validation</a><br>
							<span style="font-weight: bold">NOTICE</span>: The address validation functionality will validate<br>
P.O. Box addresses, however, UPS does not deliver to P.O.<br>
boxes, attempts by customer to ship to a P.O. Box via UPS<br>
may result in additional charges.<br>
</p>
							<% end if %></p></td>
							</tr>
							<% '// If recipient is Canada or International %>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2">Special Requirements  International Shipments </th>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><span style="font-weight: bold">The Description of Goods for the shipment -  Applies to international shipments only</span>.
								<br>
								Provide a detailed description of items being shipped for documents and non-documents. Provide specific descriptions, such as " annual reports" and " 9 mm steel screws" . Required if all of the listed conditions are true: Ship From and Recipient countries are not the same; The packaging type is not UPS Letter; The Ship From and or Recipient countries are not in the European Union or the Ship From and Recipient countries are both in the European Union and the shipment's service type is not UPS Standard.</td>
							</tr>
							<tr>
								<td align="right"><b>Description:</b></td>
								<td>
								<p>
								<input name="ShipmentDescription" type="text" id="ShipmentDescription" value="<%=pcf_FillFormField("ShipmentDescription", isRequiredShipmentDescription)%>">
								<%pcs_UPSRequiredImageTag "ShipmentDescription", isRequiredShipmentDescription %>
								</p>
								</td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><span style="font-weight: bold">Invoice Line Total amount for the entire shipment and Currency Type</span><br>
								Required for forward shipments (without a return service) that do not have any packages of type UPS Letter and whose origin s the US and destination is Puerto Rico or Canada. </td>
							</tr>
							<tr>
								<td width="23%" align="right"><b>Invoice Line Total Amount:</b></td>
								<td width="77%" align="left">
									<input name="InvoiceAmount" type="text" id="InvoiceAmount" value="<%=pcf_FillFormField("InvoiceAmount", false)%>">
									<%pcs_UPSRequiredImageTag "InvoiceAmount", false%>*This must be a whole number. Do not use a decimals or commas.
								</td>
							</tr>
							<tr>
								<td align="right"><b>Currency Type:</b></td>
								<td align="left">
									<select name="InvoiceCurrencyCode" id="InvoiceCurrencyCode">
										<option value="USD" <%=pcf_SelectOption("InvoiceCurrencyCode","USD")%>>USD</option>
										<option value="AUD" <%=pcf_SelectOption("InvoiceCurrencyCode","AUD")%>>AUD</option>
										<option value="CAD" <%=pcf_SelectOption("InvoiceCurrencyCode","CAD")%>>CAD</option>
										<option value="CHF" <%=pcf_SelectOption("InvoiceCurrencyCode","CHF")%>>CHF</option>
										<option value="CZK" <%=pcf_SelectOption("InvoiceCurrencyCode","CZK")%>>CZK</option>
										<option value="DKK" <%=pcf_SelectOption("InvoiceCurrencyCode","DKK")%>>DKK</option>
										<option value="EUR" <%=pcf_SelectOption("InvoiceCurrencyCode","EUR")%>>EUR</option>
										<option value="GBP" <%=pcf_SelectOption("InvoiceCurrencyCode","GBP")%>>GBP</option>
										<option value="GRD" <%=pcf_SelectOption("InvoiceCurrencyCode","GRD")%>>GRD</option>
										<option value="HKD" <%=pcf_SelectOption("InvoiceCurrencyCode","HKD")%>>HKD</option>
										<option value="HUF" <%=pcf_SelectOption("InvoiceCurrencyCode","HUF")%>>HUF</option>
										<option value="INR" <%=pcf_SelectOption("InvoiceCurrencyCode","INR")%>>INR</option>
										<option value="MXN" <%=pcf_SelectOption("InvoiceCurrencyCode","MXN")%>>MXN</option>
										<option value="MYR" <%=pcf_SelectOption("InvoiceCurrencyCode","MYR")%>>MYR</option>
										<option value="NOK" <%=pcf_SelectOption("InvoiceCurrencyCode","NOK")%>>NOK</option>
										<option value="NZD" <%=pcf_SelectOption("InvoiceCurrencyCode","NZD")%>>NZD</option>
										<option value="PLN" <%=pcf_SelectOption("InvoiceCurrencyCode","PLN")%>>PLN</option>
										<option value="SEK" <%=pcf_SelectOption("InvoiceCurrencyCode","SEK")%>>SEK</option>
										<option value="SGD" <%=pcf_SelectOption("InvoiceCurrencyCode","SGD")%>>SGD</option>
										<option value="THB" <%=pcf_SelectOption("InvoiceCurrencyCode","THB")%>>THB</option>
										<option value="TWD" <%=pcf_SelectOption("InvoiceCurrencyCode","TWD")%>>TWD</option>
								  </select>
									<%pcs_UPSRequiredImageTag "InvoiceCurrencyCode", false%>				</td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
							<td align="right"></td>
							<td align="left">			</td>
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
							<td colspan="2"><span class="title">Quantum View Notify<sup>SM</sup></span></td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>

						<tr>
							<td width="17%" align="right"><input type="checkbox" name="ShipmentNotification" value="1" class="clearBorder" <%=pcf_CheckOption("ShipmentNotification", "1")%>></td>
							<td width="83%"><strong>Shipment Notification</strong>			</td>
						  </tr>
							<tr>
							  <td align="right">Notification Type: </td>
							  <td align="left">
								<select name="NotificationCode1" id="NotificationCode1">
				  <option value="0" <%=pcf_SelectOption("NotificationCode1","0")%>>None Selected</option>
				  <option value="6" <%=pcf_SelectOption("NotificationCode1","6")%>>QVN Ship Notification</option>
				  <option value="7" <%=pcf_SelectOption("NotificationCode1","7")%>>QVN Exception Notification</option>
				  <option value="8" <%=pcf_SelectOption("NotificationCode1","8")%>>QVN Delivery Notification</option>
				</select>
								<%pcs_UPSRequiredImageTag "NotificationCode1", false%>
								&nbsp;E-Mail Address:
								<input name="NotificationEmail1" type="text" id="NotificationEmail1" value="<%=pcf_FillFormField("NotificationEmail1", false)%>">
								<%pcs_UPSRequiredImageTag "NotificationEmail1", false%></td>
						  </tr>
							<tr>
								<td width="17%" align="right">Notification Type:</td>
								<td width="83%" align="left"><select name="NotificationCode2" id="NotificationCode2">
				  <option value="0" <%=pcf_SelectOption("NotificationCode2","0")%>>None Selected</option>
				  <option value="6" <%=pcf_SelectOption("NotificationCode2","6")%>>QVN Ship Notification</option>
				  <option value="7" <%=pcf_SelectOption("NotificationCode2","7")%>>QVN Exception Notification</option>
				  <option value="8" <%=pcf_SelectOption("NotificationCode2","8")%>>QVN Delivery Notification</option>
				</select>
				  <%pcs_UPSRequiredImageTag "NotificationCode2", false %>
								&nbsp;E-Mail Address:
				  <input name="NotificationEmail2" type="text" id="NotificationEmail2" value="<%=pcf_FillFormField("NotificationEmail2", false)%>">
				  <%pcs_UPSRequiredImageTag "NotificationEmail2", false%></td>
							</tr>
							<tr>
								<td align="right">Notification Type:</td>
								<td align="left"><select name="NotificationCode3" id="NotificationCode3">
				  <option value="0" <%=pcf_SelectOption("NotificationCode3","0")%>>None Selected</option>
				  <option value="6" <%=pcf_SelectOption("NotificationCode3","6")%>>QVN Ship Notification</option>
				  <option value="7" <%=pcf_SelectOption("NotificationCode3","7")%>>QVN Exception Notification</option>
				  <option value="8" <%=pcf_SelectOption("NotificationCode3","8")%>>QVN Delivery Notification</option>
				</select>
				  <%pcs_UPSRequiredImageTag "NotificationCode3", false %>
								&nbsp;E-Mail Address:
				  <input name="NotificationEmail3" type="text" id="NotificationEmail3" value="<%=pcf_FillFormField("NotificationEmail3", false)%>">
				  <%pcs_UPSRequiredImageTag "NotificationEmail3", false %></td>
							</tr>
							<tr>
								<td align="right">Notification Type:</td>
								<td align="left"><select name="NotificationCode4" id="NotificationCode4">
				  <option value="0" <%=pcf_SelectOption("NotificationCode4","0")%>>None Selected</option>
				  <option value="6" <%=pcf_SelectOption("NotificationCode4","6")%>>QVN Ship Notification</option>
				  <option value="7" <%=pcf_SelectOption("NotificationCode4","7")%>>QVN Exception Notification</option>
				  <option value="8" <%=pcf_SelectOption("NotificationCode4","8")%>>QVN Delivery Notification</option>
				</select>
				  <%pcs_UPSRequiredImageTag "NotificationCode4", false %>
								&nbsp;E-Mail Address:
				  <input name="NotificationEmail4" type="text" id="NotificationEmail4" value="<%=pcf_FillFormField("NotificationEmail4", false)%>">
				  <%pcs_UPSRequiredImageTag "NotificationEmail4", false %></td>
							</tr>
							<tr>
								<td align="right">Notification Type:</td>
								<td align="left"><select name="NotificationCode5" id="NotificationCode5">
				  <option value="0" <%=pcf_SelectOption("NotificationCode5","0")%>>None Selected</option>
				  <option value="6" <%=pcf_SelectOption("NotificationCode5","6")%>>QVN Ship Notification</option>
				  <option value="7" <%=pcf_SelectOption("NotificationCode5","7")%>>QVN Exception Notification</option>
				  <option value="8" <%=pcf_SelectOption("NotificationCode5","8")%>>QVN Delivery Notification</option>
				</select>
				  <%pcs_UPSRequiredImageTag "NotificationCode5", false %>
								&nbsp;E-Mail Address:
				  <input name="NotificationEmail5" type="text" id="NotificationEmail5" value="<%=pcf_FillFormField("NotificationEmail5", false)%>">
				  <%pcs_UPSRequiredImageTag "NotificationEmail5", false %></td>
							</tr>
							<tr>
								<td align="right">&nbsp;</td>
								<td align="left">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
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
							if request("sub")=1 then
								if Session("pcAdminVerbalConfirmation"&k)="" then
									Session("pcAdminVerbalConfirmation"&k)=Session("pcAdminUPSVerbalConfirmation")
								end if
								if session("pcAdminDimensionsUnitOfMeasurement"&k)="" then
									session("pcAdminDimensionsUnitOfMeasurement"&k)=ucase(UPS_DIM_UNIT)
								end if
								if session("pcAdminHeight"&k)="" then
									session("pcAdminHeight"&k)=UPS_HEIGHT
								end if
								if session("pcAdminWidth"&k)="" then
									session("pcAdminWidth"&k)=UPS_WIDTH
								end if
								if session("pcAdminLength"&k)="" then
									session("pcAdminLength"&k)=UPS_LENGTH
								end if
								if session("pcAdminWeightUnitOfMeasurement"&k)="" then
									session("pcAdminWeightUnitOfMeasurement"&k)=scShipFromWeightUnit
								end if
								if session("pcAdminCODPackage"&k)="" then
									session("pcAdminCODPackage"&k)=Session("pcAdminUPSCODPackage")
								end if
								if session("pcAdminCODAmount"&k)="" then
									Session("pcAdminCODAmount"&k)=Session("pcAdminUPSCODAmount")
								end if
								if session("pcAdminCODCurrencyCode"&k)="" then
									Session("pcAdminCODCurrencyCode"&k)=Session("pcAdminUPSCODCurrency")
								end if
								if session("pcAdminCODFundsCode"&k)="" then
									Session("pcAdminCODFundsCode"&k)=Session("pcAdminUPSCODFunds")
								end if
								if session("pcAdminInsuredValue"&k)="" then
									session("pcAdminInsuredValue"&k)=Session("pcAdminUPSInsuredValue")
								end if
							else
								Session("pcAdminVerbalConfirmation"&k)=Session("pcAdminUPSVerbalConfirmation")
								session("pcAdminDimensionsUnitOfMeasurement"&k)=ucase(UPS_DIM_UNIT)
								session("pcAdminHeight"&k)=UPS_HEIGHT
								session("pcAdminWidth"&k)=UPS_WIDTH
								session("pcAdminLength"&k)=UPS_LENGTH
								session("pcAdminWeightUnitOfMeasurement"&k)=scShipFromWeightUnit
								session("pcAdminCODPackage"&k)=Session("pcAdminUPSCODPackage")
								Session("pcAdminCODAmount"&k)=Session("pcAdminUPSCODAmount")
								Session("pcAdminCODCurrencyCode"&k)=Session("pcAdminUPSCODCurrency")
								Session("pcAdminCODFundsCode"&k)=Session("pcAdminUPSCODFunds")
								session("pcAdminInsuredValue"&k)=Session("pcAdminUPSInsuredValue")
							end if %>
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
														pcv_CumulativeWeight = Cint(0)
														For pcv_xCounter=0 to (ubound(xProductDisplayArray)-1)
															pcv_intPackageInfo_ID = xProductDisplayArray(pcv_xCounter)
															' GET THE PACKAGE CONTENTS
															' >>> Tables: products, ProductsOrdered
															query = "SELECT ProductsOrdered.pcPackageInfo_ID, ProductsOrdered.quantity , products.description, products.idProduct, products.weight, products.OverSizeSpec  "
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
																	pcOverSizedFlag = Cint(0)
																	pcv_strProductQty = rs2("quantity")
																	pcv_strProductDescription = rs2("description")
																	pcv_POWeight = rs2("weight")
																	pcv_strOverSizeSpec = rs2("OverSizeSpec") '40||20||20||0||400
																	if pcv_strOverSizeSpec<>"NO" then
																		pcOverSizedFlag = Cint(1)
																		pOSArray=split(pcv_strOverSizeSpec,"||")
																		if ubound(pOSArray)>2 then
																			tOS_width=pOSArray(0)
																			tOS_height=pOSArray(1)
																			tOS_length=pOSArray(2)
																			'if pcPackageCount  = 1 then
																				session("pcAdminHeight"&k)= tOS_height
																				session("pcAdminWidth"&k)= tOS_width
																				session("pcAdminLength"&k)= tOS_length
																			'end if
																		else
																			tOS_width=0
																			tOS_height=0
																			tOS_length=0
																		end if
																	end if
																	pcv_CumulativeWeight = Cint(pcv_CumulativeWeight) + (Cint(pcv_POWeight)*Cint(pcv_strProductQty))
																	%>
																	<li><%=pcv_strProductQty&"&nbsp;"&pcv_strProductDescription%></li>
																	<%
																	rs2.movenext
																Loop
															end if
														Next
														%>
													</td>
												</tr>
											</table></td>
									</tr>
									<tr>
									<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
									<th colspan="2">Settings <%'=k%></th>
									</tr>
									<tr>
									<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
									<td colspan="2">
									<p>
									Package Type:
									<% Session("pcAdminPackageTypeCode"&k)=Session("pcAdminUPSPackageType") %>
									<select name="PackageTypeCode<%=k%>" id="Service<%=k%>">
									<option value="01" <%=pcf_SelectOption("PackageTypeCode"&k,"01")%>>UPS Letter</option>
									<option value="02" <%=pcf_SelectOption("PackageTypeCode"&k,"02")%>>Your Packaging</option>
									<option value="03" <%=pcf_SelectOption("PackageTypeCode"&k,"03")%>>UPS Tube</option>
									<option value="04" <%=pcf_SelectOption("PackageTypeCode"&k,"04")%>>UPS PAK</option>
									<option value="21" <%=pcf_SelectOption("PackageTypeCode"&k,"21")%>>UPS 25KG Box&reg;</option>
									<option value="24" <%=pcf_SelectOption("PackageTypeCode"&k,"24")%>>UPS 10KG Box&reg;</option>
									</select>
									<%pcs_UPSRequiredImageTag "PackageTypeCode"&k, true%>
									</p>

									<p style="padding-top: 5px;">
									<strong>Packaging:</strong><br>
									When using UPS packaging, select the
									packaging type from the drop-down list.<br>
									When using non-UPS packaging, select &quot;Your
									Packaging&quot;, and then enter
									the dimensions manually.</p>
									<p>&nbsp;</p>
									</td>
									</tr>
									<tr>
									<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
									<th colspan="2">Dimensions and Weight</th>
									</tr>
									<tr>
									<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
									<td colspan="2">
									<p><strong>Package Dimensions:</strong></p>
									<p style="padding-top: 5px;">
									Maximum 274 cm in length (always the longest side)<br>
									Maximum 330 cm in length and girth combined. Girth = (2 x height) + (2 x   width)
									</p>
									<p style="padding-top: 5px;">
									Units:
									<select name="DimensionsUnitOfMeasurement<%=k%>" id="DimensionsUnitOfMeasurement<%=k%>">
									<option value="">Select A Unit</option>
									<option value="IN" <%=pcf_SelectOption("DimensionsUnitOfMeasurement"&k,"IN")%>>Inches</option>
									<option value="CM" <%=pcf_SelectOption("DimensionsUnitOfMeasurement"&k,"CM")%>>Centimeters</option>
									</select>
									<%pcs_UPSRequiredImageTag "DimensionsUnitOfMeasurement"&k, false%>
									</p>
									<p style="padding-top: 5px;">
									Length:
									<input name="Length<%=k%>" type="text" id="Length<%=k%>" value="<%=pcf_FillFormField("Length"&k, false)%>" width="4">
									<%pcs_UPSRequiredImageTag "Length"&k, false%>
									&nbsp;
									Width: <input name="Width<%=k%>" type="text" id="Width<%=k%>" value="<%=pcf_FillFormField("Width"&k, false)%>" width="4">
									<%pcs_UPSRequiredImageTag "Width"&k, false%>
									&nbsp;
									Height: <input name="Height<%=k%>" type="text" id="Height<%=k%>" value="<%=pcf_FillFormField("Height"&k, false)%>" width="4">
									<%pcs_UPSRequiredImageTag "Height"&k, false%>
									</p>
									<p style="padding-top: 5px;"><strong>Package Weight:</strong><br>
									Enter the weight of the package. If there is more than one package in the shipment, enter the weight of the first package or the total shipment weight.</p>
									<p style="padding-top: 5px;">WeightUnits:
									<select name="WeightUnitOfMeasurement<%=k%>" id="WeightUnitOfMeasurement<%=k%>">
									<option value="LBS" <%=pcf_SelectOption("WeightUnitOfMeasurement"&k,"LBS")%>>LBS</option>
									<option value="KGS" <%=pcf_SelectOption("WeightUnitOfMeasurement"&k,"KGS")%>>KGS</option>
									</select>
									<% pcs_UPSRequiredImageTag "WeightUnitOfMeasurement"&k, true %>
									</p>
									<%
									if scShipFromWeightUnit="KGS" then
										intShipWeightPounds=int(pcv_CumulativeWeight/1000)
										intShipWeightOunces=pcv_CumulativeWeight-(intShipWeightPounds*1000)
									else
										intShipWeightPounds=Int(pcv_CumulativeWeight/16) 'intPounds used for USPS
										intShipWeightOunces=pcv_CumulativeWeight-(intShipWeightPounds*16) 'intUniversalOunces used for USPS
									end if
									intMPackageWeight=intShipWeightPounds
									if intMPackageWeight<1 AND intShipWeightOunces<1 then
										intMPackageWeight=0
									end if
									if intMPackageWeight<1 AND intShipWeightOunces>0 then 'if total weight is less then a pound, make UPS/FedEX weight 1 pound
										intMPackageWeight=1
									else  'total weight is not less then a pound and ounces exist, round weight up one more pound.
										If intMPackageWeight>0 AND intShipWeightOunces>0 then
											intMPackageWeight=(intMPackageWeight+1)
										End if
									end if
									'response.write intMPackageWeight&"<BR>"
									'if pcPackageCount=1 then
										'Get weight
										Session("pcAdminPackageWeight"&k) = intMPackageWeight
									'end if %>
									<p style="padding-top: 5px;">Weight: <input name="PackageWeight<%=k%>" type="text" id="PackageWeight<%=k%>" value="<%=pcf_FillFormField("PackageWeight"&k, true) %>">
									<%pcs_UPSRequiredImageTag "PackageWeight"&k, true%>
									<br>
									<br>
									<span style="font-weight: bold">Weight</span><br>
									<br>
									Weight values may be given as whole numbers or as a real numbers with precision<br>
									to the tenth. The fractional separator must be a &quot; .&quot; Examples of valid values are<br>
									25, 25.6 and 25.0. Examples of invalid values are 25.05 and 25,6.</p>

									<%
									strOSSelected = Cstr("")
									if pcOverSizedFlag = 1 then
										strOSSelected = "selected"
									end if
									%>
									<p style="padding-top: 5px;">Oversized Package?
										<select name="OversizePackage<%=k%>" id="select">
											<option value="NO" <%=pcf_SelectOption("OversizePackage"&k,"NO")%>>Not Oversized</option>
											<option value="OS1" <%=pcf_SelectOption("OversizePackage"&k,"OS1")%> <% =strOSSelected%>>Oversize 1</option>
											<option value="OS2" <%=pcf_SelectOption("OversizePackage"&k,"OS2")%>>Oversize 2</option>
											<option value="OS3" <%=pcf_SelectOption("OversizePackage"&k,"OS3")%>>Oversize 3</option>
										</select>
										<br>
									  <span style="font-weight: bold"><br>
									  Oversized and/or Large Packages: </span></p>
									<p style="padding-top: 5px;">Ground shipments will be declared OS1 when the dimensions are supplied and the<br>
									length and girth combined exceeds 84&quot; and is equal to or less than 108&quot; and the<br>
									package weighs less than 30 LBS.<br>
									Ground shipments will be declared OS2 when dimensions are supplied and the length<br>
									and girth combined exceeds 108&quot; and is equal to or less than 130&quot; and the package<br>
									weighs less than 70LBS.<br>
									Ground and Air shipments will be declared OS3/LP when dimensions are<br>
									supplied and the length and girth combined exceeds 130&quot; and is equal to or less<br>
									than 165&quot; .<br>
									For Ground shipments if the actual weight is less than 150 pounds, then the<br>
									billable weight is 150 pounds. Air and 3 Day Select shipments will not be subject<br>
									to a minimum billable weight.
									</p>
									<p style="padding-top: 5px;">
									  <input type="checkbox" name="AdditionalHandling<%=k%>" value="1" class="clearBorder" <%=pcf_CheckOption("AdditionalHandling"&k, "1")%>>
									  <span style="font-weight: bold">Additional Handling </span>
									  </p></td>
									</tr>
									<tr>
									<td colspan="2" class="pcCPspacer"></td>
									</tr>
							<tr>
							<th colspan="2"><b>COD - Collect On Delivery Settings:  <b></th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
								  </tr>
														<tr>
				<td align="right"><input type="checkbox" name="CODPackage<%=k%>" value="1" class="clearBorder" <%=pcf_CheckOption("CODPackage"&k, "1")%>></td>
							  <td><span style="font-weight: bold">C.O.D. is required for this package. </span></td>
							  </tr>
							<tr>
								<td width="23%" align="right"><b>Collection Amount:</b></td>
								<td width="77%" align="left">
									<input name="CODAmount<%=k%>" type="text" id="CODAmount<%=k%>" value="<%=pcf_FillFormField("CODAmount"&k, false)%>">
									<%pcs_UPSRequiredImageTag "CODAmount"&k, false%>				</td>
							</tr>
							<tr>
								<td align="right"><b>Collection Currency:</b></td>
								<td align="left">
									<select name="CODCurrencyCode<%=k%>" id="CODCurrencyCode<%=k%>">
										<option value="USD" <%=pcf_SelectOption("CODCurrencyCode"&k,"USD")%>>USD</option>
										<option value="AUD" <%=pcf_SelectOption("CODCurrencyCode"&k,"AUD")%>>AUD</option>
										<option value="CAD" <%=pcf_SelectOption("CODCurrencyCode"&k,"CAD")%>>CAD</option>
										<option value="CHF" <%=pcf_SelectOption("CODCurrencyCode"&k,"CHF")%>>CHF</option>
										<option value="CZK" <%=pcf_SelectOption("CODCurrencyCode"&k,"CZK")%>>CZK</option>
										<option value="DKK" <%=pcf_SelectOption("CODCurrencyCode"&k,"DKK")%>>DKK</option>
										<option value="EUR" <%=pcf_SelectOption("CODCurrencyCode"&k,"EUR")%>>EUR</option>
										<option value="GBP" <%=pcf_SelectOption("CODCurrencyCode"&k,"GBP")%>>GBP</option>
										<option value="GRD" <%=pcf_SelectOption("CODCurrencyCode"&k,"GRD")%>>GRD</option>
										<option value="HKD" <%=pcf_SelectOption("CODCurrencyCode"&k,"HKD")%>>HKD</option>
										<option value="HUF" <%=pcf_SelectOption("CODCurrencyCode"&k,"HUF")%>>HUF</option>
										<option value="INR" <%=pcf_SelectOption("CODCurrencyCode"&k,"INR")%>>INR</option>
										<option value="MXN" <%=pcf_SelectOption("CODCurrencyCode"&k,"MXN")%>>MXN</option>
										<option value="MYR" <%=pcf_SelectOption("CODCurrencyCode"&k,"MYR")%>>MYR</option>
										<option value="NOK" <%=pcf_SelectOption("CODCurrencyCode"&k,"NOK")%>>NOK</option>
										<option value="NZD" <%=pcf_SelectOption("CODCurrencyCode"&k,"NZD")%>>NZD</option>
										<option value="PLN" <%=pcf_SelectOption("CODCurrencyCode"&k,"PLN")%>>PLN</option>
										<option value="SEK" <%=pcf_SelectOption("CODCurrencyCode"&k,"SEK")%>>SEK</option>
										<option value="SGD" <%=pcf_SelectOption("CODCurrencyCode"&k,"SGD")%>>SGD</option>
										<option value="THB" <%=pcf_SelectOption("CODCurrencyCode"&k,"THB")%>>THB</option>
										<option value="TWD" <%=pcf_SelectOption("CODCurrencyCode"&k,"TWD")%>>TWD</option>
									</select>
									<%pcs_UPSRequiredImageTag "CODCurrencyCode"&k, false%>				</td>
							</tr>
							<tr>
								<td align="right"><b>Collection Fund Type:</b></td>
								<td align="left">
									<select name="CODFundsCode<%=k%>" id="CODFundsCode<%=k%>">
									<option value="0" <%=pcf_SelectOption("CODFundsCode"&k,"0")%>>Cash</option>
									<option value="8" <%=pcf_SelectOption("CODFundsCode"&k,"8")%>>Check, Cashier’s Check or Money Order</option>
									</select>
									<%pcs_UPSRequiredImageTag "CODFundsCode"&k, false%></td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
								  </tr>
									<tr>
									<th colspan="2">Value</th>
									</tr>
									<tr>
									<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
									<td colspan="2">
									<p>Insured Value:
									  <input name="InsuredValue<%=k%>" type="text" id="InsuredValue<%=k%>" value="<%=pcf_FillFormField("InsuredValue"&k, true)%>">
									<%pcs_UPSRequiredImageTag "InsuredValue"&k, true%>
									</p>									</td>
									</tr>
									<tr>
										<td colspan="2">
											<p>Delivery Confirmation :
												<select name="DeliveryConfirmation<%=k%>" id="DeliveryConfirmation<%=k%>">
													<option value="NONE" <%=pcf_SelectOption("DeliveryConfirmation"&k,"NONE")%>>None</option>
													<option value="1" <%=pcf_SelectOption("DeliveryConfirmation"&k,"1")%>>Deliver Without Signature</option>
													<option value="2" <%=pcf_SelectOption("DeliveryConfirmation"&k,"2")%>>Signature Required</option>
													<option value="3" <%=pcf_SelectOption("DeliveryConfirmation"&k,"3")%>>Adult Signature Required</option>
												</select>
												<%pcs_UPSRequiredImageTag "DeliveryConfirmation"&k, false%>
											</p>
											<p style="padding-top: 5px;">Only allowed for shipment with US ShipFrom/destination combination.</p></td>
								  </tr>
										<tr>
											<td align="right"><input type="checkbox" name="VerbalConfirmation<%=k%>" value="1" class="clearBorder" <%=pcf_CheckOption("VerbalConfirmation"&k, "1")%>><%pcs_UPSRequiredImageTag "VerbalConfirmation"&k, false%></td>
											<td><strong>Verbal Confirmation </strong> </td>
										</tr>
										<tr>
									<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<th colspan="2">Package Reference Number(s)</th>
									</tr>
									<tr>
									<td colspan="2" class="pcCPspacer"></td>
									</tr>
										<tr>
										  <td align="right">Reference Code: </td>
										  <td align="left">
											<select name="UPSRefNumber1<%=k%>" id="UPSRefNumber1<%=k%>">
												<option value="NONE">None Selected</option>
												<option value="AJ" <%=pcf_SelectOption("UPSRefNumber1","AJ")%>>Acct. Rec. Customer Acct.</option>
												<option value="AT" <%=pcf_SelectOption("UPSRefNumber1","AT")%>>Appropriation Number</option>
												<option value="BM" <%=pcf_SelectOption("UPSRefNumber1","BM")%>>Bill of Lading Number</option>
												<option value="9V" <%=pcf_SelectOption("UPSRefNumber1","9V")%>>COD Number</option>
												<option value="ON" <%=pcf_SelectOption("UPSRefNumber1","ON")%>>Dealer Order Number</option>
												<option value="DP" <%=pcf_SelectOption("UPSRefNumber1","DP")%>>Department Number</option>
												<option value="EI" <%=pcf_SelectOption("UPSRefNumber1","EI")%>>Employer's ID Number</option>
												<option value="3Q" <%=pcf_SelectOption("UPSRefNumber1","3Q")%>>FDA Product Code
												<option value="TJ" <%=pcf_SelectOption("UPSRefNumber1","TJ")%>>Federal Taxpayer ID No.</option>
												<option value="IK" <%=pcf_SelectOption("UPSRefNumber1","IK")%>>Invoice Number</option>
												<option value="MK" <%=pcf_SelectOption("UPSRefNumber1","MK")%>>Manifest Key Number</option>
												<option value="MJ" <%=pcf_SelectOption("UPSRefNumber1","MJ")%>>Model Number</option>
												<option value="PM" <%=pcf_SelectOption("UPSRefNumber1","PM")%>>Part Number</option>
												<option value="PC" <%=pcf_SelectOption("UPSRefNumber1","PC")%>>Production Code</option>
												<option value="PO" <%=pcf_SelectOption("UPSRefNumber1","PO")%>>Purchase Order Number</option>
												<option value="RQ" <%=pcf_SelectOption("UPSRefNumber1","RQ")%>>Purchase Req. Number</option>
												<option value="RZ" <%=pcf_SelectOption("UPSRefNumber1","RZ")%>>Return Authorization No.</option>
												<option value="SA" <%=pcf_SelectOption("UPSRefNumber1","SA")%>>Salesperson Number</option>
												<option value="SE" <%=pcf_SelectOption("UPSRefNumber1","SE")%>>Serial Number</option>
												<option value="SY" <%=pcf_SelectOption("UPSRefNumber1","SY")%>>Social Security Number</option>
												<option value="ST" <%=pcf_SelectOption("UPSRefNumber1","ST")%>>Store Number</option>
												<option value="TN" <%=pcf_SelectOption("UPSRefNumber1","TN")%>>Transaction Ref. No.</option>
												</option>
											  </select>
											</td>
										</tr>
										<% if session("pcAdminUPSRefNumber1")="IK" then
											if session("pcAdminUPSRefData1")="" then
												session("pcAdminUPSRefData1")=scpre+int(pcv_intOrderID)
												'tmpvalue=
											end if
										end if %>

										<tr>
										  <td align="right">Reference Value: </td>
										  <td align="left"><input name="UPSRefData1<%=k%>" type="text" value="<%=pcf_FillFormField("UPSRefData1", false)%>" size="35"></td>
										</tr>
										<tr>
										  <td align="right">Reference Code: </td>
										  <td align="left"><select name="UPSRefNumber2<%=k%>" id="UPSRefNumber2<%=k%>">
												<option value="NONE">None Selected</option>
												<option value="AJ" <%=pcf_SelectOption("UPSRefNumber2","AJ")%>>Acct. Rec. Customer Acct.</option>
												<option value="AT" <%=pcf_SelectOption("UPSRefNumber2","AT")%>>Appropriation Number</option>
												<option value="BM" <%=pcf_SelectOption("UPSRefNumber2","BM")%>>Bill of Lading Number</option>
												<option value="9V" <%=pcf_SelectOption("UPSRefNumber2","9V")%>>COD Number</option>
												<option value="ON" <%=pcf_SelectOption("UPSRefNumber2","ON")%>>Dealer Order Number</option>
												<option value="DP" <%=pcf_SelectOption("UPSRefNumber2","DP")%>>Department Number</option>
												<option value="EI" <%=pcf_SelectOption("UPSRefNumber2","EI")%>>Employer's ID Number</option>
												<option value="3Q" <%=pcf_SelectOption("UPSRefNumber2","3Q")%>>FDA Product Code
												<option value="TJ" <%=pcf_SelectOption("UPSRefNumber2","TJ")%>>Federal Taxpayer ID No.</option>
												<option value="IK" <%=pcf_SelectOption("UPSRefNumber2","IK")%>>Invoice Number</option>
												<option value="MK" <%=pcf_SelectOption("UPSRefNumber2","MK")%>>Manifest Key Number</option>
												<option value="MJ" <%=pcf_SelectOption("UPSRefNumber2","MJ")%>>Model Number</option>
												<option value="PM" <%=pcf_SelectOption("UPSRefNumber2","PM")%>>Part Number</option>
												<option value="PC" <%=pcf_SelectOption("UPSRefNumber2","PC")%>>Production Code</option>
												<option value="PO" <%=pcf_SelectOption("UPSRefNumber2","PO")%>>Purchase Order Number</option>
												<option value="RQ" <%=pcf_SelectOption("UPSRefNumber2","RQ")%>>Purchase Req. Number</option>
												<option value="RZ" <%=pcf_SelectOption("UPSRefNumber2","RZ")%>>Return Authorization No.</option>
												<option value="SA" <%=pcf_SelectOption("UPSRefNumber2","SA")%>>Salesperson Number</option>
												<option value="SE" <%=pcf_SelectOption("UPSRefNumber2","SE")%>>Serial Number</option>
												<option value="SY" <%=pcf_SelectOption("UPSRefNumber2","EY")%>>Social Security Number</option>
												<option value="ST" <%=pcf_SelectOption("UPSRefNumber2","ST")%>>Store Number</option>
												<option value="TN" <%=pcf_SelectOption("UPSRefNumber2","TN")%>>Transaction Ref. No.</option>
												</option>
											  </select></td>
								  </tr>
										<tr>
										  <td align="right">Reference Value: </td>
										  <td align="left"><input name="UPSRefData2<%=k%>" type="text" value="<%=pcf_FillFormField("UPSRefData2", false)%>" size="35"></td>
								  </tr>
										<tr>
											<td align="right">&nbsp;</td>
											<td align="left">&nbsp;</td>
										</tr>
												<% else %>
												<tr>
												<th colspan="2">This package has been shipped.</th>
												</tr>
												<%
												end if
												%>
								</table>
							</div>
						<% next %>

							<br />
							<br />

							<%
							pcv_strPreviousPage = "Orddetails.asp?id=" & pcv_intOrderID
							pcv_strAddPackagePage = "sds_ShipOrderWizard1.asp?idorder="&pcv_intOrderID&"&PageAction=UPS&PackageCount="&pcPackageCount&"&ItemsList="&pcv_strItemsList
							%>

							<p>
						  <div align="center">
									<input type="button" name="Button" value="Start Over" onclick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
									<% if pcPackageCount<20 then %>
									<input type="button" name="Button" value="Add Another Package" onclick="document.location.href='<%=pcv_strAddPackagePage%>'" class="ibtnGrey">
									<% end if %>
									<input type="submit" name="submit" value="Process Shipment" class="ibtnGrey">
									<br />
									<br />
									<input type="button" name="Button" value="Go Back To Order Details" onclick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
						  </div>
							</p>						</td>
						</tr>
					<tr>
					  <td valign="top"><div align="center">
					   <%= pcf_UPSWriteLegalDisclaimers %>
						</div></td>
					  </tr>
					<!--End -->
					</table>
				</form>
			<%
			end if
			call closedb()

			'*******************************************************************************
			' END: LOAD HTML FORM
			'*******************************************************************************
			%></td>
	</tr>
</table>
<%
'// DESTROY THE UPS OBJECT
set objUPSClass = nothing
%>
<!--#include file="AdminFooter.asp"-->
