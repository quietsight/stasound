<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="USPS Shipping Wizard" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/USPSCountry.asp"-->
<!--#include file="../includes/pcUSPSClass.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Global.asp"-->
<!--#include file="../includes/GoogleCheckout_APIFunctions.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Handler.asp"-->
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
<%
Dim objUSPSXmlDoc, objUSPSStream, strFileName, GraphicXML, iPageCurrent, pcv_intOrderID, pcv_LabelMode, pcv_strPackageCount
Dim pcv_strSessionPackageCount, pcPackageCount, pcArraySize, pcv_strOrderID, pcv_strSessionOrderID
Dim USPS_postdata, objUSPSClass, objOutputXMLDoc, srvUSPSXmlHttp, USPS_result, pcv_USPSLabelServer, pcv_strErrorMsg, pcv_strAction

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

function exFormatDate(strDate, DateTemplate)
  If not IsDate(strDate) Then
	exFormatDate = strDate
	Exit function
  End If
  DateTemplate = replace(DateTemplate,"%mm",right("0" & DatePart("m",strDate),2),1,-1,0)
  DateTemplate = replace(DateTemplate,"%dd",right("0" & DatePart("d",strDate),2),1,-1,0)
  DateTemplate = replace(DateTemplate,"%yyyy",DatePart("yyyy",strDate,2),1,-1,0)
  exFormatDate = DateTemplate
end function

'// MODE
pcv_LabelMode=request("LabelMode")

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
			pcv_strHiddenField=pcv_strHiddenField&"<input type='hidden' name='C"&i&"' value='1'><input type='hidden' name='IDPrd"&i&"' value='"&request("IDPrd" & i)&"'>"
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
pcPageName="USPS_ManageShipmentsRequest.asp"
ErrPageName="USPS_ManageShipmentsRequest.asp"

'// ACTION
pcv_strAction = request("Action")

dim conntemp, query, rs
'// Retrieve current orderstatus of this order

'// SET THE USPS OBJECT
set objUSPSClass = New pcUSPSClass
query="SELECT orderdate, orderstatus FROM orders WHERE idOrder="&Session("pcAdminOrderID")&";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

dim pOrderStatus
pOrderStatus=rs("orderstatus")
pcv_orderDate=rs("orderDate")

'// DATE FUNCTION
function ShowDateFrmt(x)
	ShowDateFrmt = x
end function


'// SELECT DATA SET
' >>> Tables: pcPackageInfo
query="SELECT orders.idCustomer, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.shippingCompany, orders.shippingAddress2, orders.pcOrd_shippingPhone, orders.ShippingFullName, orders.pcOrd_ShippingEmail, orders.ordShipType, orders.pcOrd_ShipWeight, orders.address, orders.address2, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode FROM orders WHERE orders.idOrder=" & pcv_intOrderID &" "

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
	pcv_UseAltAddress = "NO"
	'//If Shipping Variables are empty use billing address
	If pcv_ShippingAddress&""="" Then
		pcv_UseAltAddress = "YES"
		pcv_ShippingAddress=rs("address")
		pcv_ShippingAddress2=rs("address2")
		pcv_ShippingZip=rs("zip")
		pcv_ShippingStateCode=rs("stateCode")
		pcv_ShippingState=rs("state")
		pcv_ShippingCity=rs("city")
		pcv_ShippingCountryCode=rs("countryCode")
		query="SELECT name, lastname, customerCompany, email FROM customers WHERE idcustomer="& pcv_IdCustomer
		Set rsGetCustObj=Server.CreateObject("ADODB.Recordset")
		Set rsGetCustObj=conntemp.execute(query)
		pcv_strAltToFirstName=rsGetCustObj("name")
		pcv_strAltToLastName=rsGetCustObj("lastname")
		pcv_strAltToFullName = pcv_strAltToFirstName&" "&pcv_strAltToLastName
		pcv_strAltToFirm=rsGetCustObj("customerCompany")
		pcv_AltShippingEmail=rsGetCustObj("email")
		Set rsGetCustObj = Nothing
	End If
end if
set rs=nothing

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

'// ORIGIN ADDRESS
if Session("pcAdminFromName") = "" OR intResetSessions=1 then
	pcv_strFromName = scOriginPersonName
	if pcv_strFromName="" then
		pcv_strFromName=scShipFromName
	end if
	Session("pcAdminFromName") = pcv_strFromName
end if

if instr(scOriginPersonName, " ") then
	pcv_FromNameArry=split(scOriginPersonName, " ")
	pcv_strFromFirstName=pcv_FromNameArry(0)
	pcv_strFromLastName=pcv_FromNameArry(1)
end if

if Session("pcAdminFromFirstName") = "" OR intResetSessions=1 then
	pcv_strFromFirstName = pcv_strFromFirstName
	if pcv_strFromFirstName="" then
		pcv_strFromFirstName=pcv_strFromName
	end if
	Session("pcAdminFromFirstName") = pcv_strFromFirstName
end if

if Session("pcAdminFromLastName") = "" OR intResetSessions=1 then
	pcv_strFromLastName = pcv_strFromLastName
	Session("pcAdminFromLastName") = pcv_strFromLastName
end if

if Session("pcAdminFromFirm") = "" OR intResetSessions=1 then
	pcv_strFromFirm = scShipFromName
	if pcv_strFromFirm="" then
		pcv_strFromFirm=scOriginPersonName
	end if
	Session("pcAdminFromFirm") = pcv_strFromFirm
end if

if Session("pcAdminFromPhone") = "" OR intResetSessions=1 then
	pcv_strFromPhone = scOriginPhoneNumber
	Session("pcAdminFromPhone") = pcv_strFromPhone
end if

if Session("pcAdminSenderEMail") = "" OR intResetSessions=1 then
	pcv_strSenderEMail = scFrmEmail
	Session("pcAdminSenderEMail") = pcv_strSenderEMail
end if

'// DESTINATION ADDRESS
if Session("pcAdminFromAddress1") = "" OR IntResetSessions=1 then
	pcv_strFromAddress1 = scShipFromAddress1
	Session("pcAdminFromAddress1") = pcv_strFromAddress1
end if
if Session("pcAdminFromAddress2") = "" OR intResetSessions=1 then
	pcv_strFromAddress2 = scShipFromAddress2
	Session("pcAdminFromAddress2") = pcv_strFromAddress2
end if
if Session("pcAdminFromCity") = "" OR intResetSessions=1 then
	pcv_strFromCity = scShipFromCity
	Session("pcAdminFromCity") = pcv_strFromCity
end if
if Session("pcAdminFromState") = "" OR intResetSessions=1 then
	pcv_strFromState = scShipFromState
	Session("pcAdminFromState") = pcv_strFromState
end if
if Session("pcAdminFromZip5") = "" OR intResetSessions=1 then
	pcv_strFromZip5 = scShipFromPostalCode
	Session("pcAdminFromZip5") = pcv_strFromZip5
end if
if Session("pcAdminFromZip4") = "" OR intResetSessions=1 then
	pcv_strFromZip4 = scShipFromZip4
	Session("pcAdminFromZip4") = pcv_strFromZip4
end if
if Session("pcAdminFromCountryCode") = "" OR intResetSessions=1 then
	pcv_strFromCountryCode = "US"
	Session("pcAdminFromCountryCode") = pcv_strFromCountryCode
end if

if Session("pcAdmincustomerRefNo1") = "" OR intResetSessions=1 then
	Session("pcAdmincustomerRefNo1") = pcv_IdCustomer
end if

strFromState=Session("pcAdminFromState")
strFromCity=Session("pcAdminFromCity")
strFromZip5=Session("pcAdminFromZip5")
strFromZip4=Session("pcAdminFromZip4")
strShipFromCountry=Session("pcAdminFromCountryCode")

if pcv_LabelMode<>"" then
	select case pcv_LabelMode
		case "D"
			pcv_LabelDescription="Delivery Confirmation"
			isPOZipCodeValueReq=0
			isLabelOptionRequired=true
			isLabelOptionValueReq=1
			isFromFirmRequired=false
		case "S"
			pcv_LabelDescription="Signature Confirmation"
			isPOZipCodeValueReq=0
			isLabelOptionRequired=true
			isLabelOptionValueReq=1
			isFromFirmRequired=false
		case "E"
			pcv_LabelDescription="Express Mail"
			isPOZipCodeValueReq=1
			isLabelOptionRequired=false
			isLabelOptionValueReq=1
			isFromFirmRequired=true
	end select
end if

'   >>> Shipper Address Conditionals
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

isRequiredShipFromPostal = true

'// DESTINATION ADDRESS
if Session("pcAdminToName") = "" OR intResetSessions=1 then
	IF pcv_UseAltAddress = "YES" Then
		pcv_strToName = pcv_strAltToFullName
		if pcv_strToName="" then
			pcv_strToName=pcv_strAltToFirm
		end if
	Else
		pcv_strToName = pcv_ShippingFullName
		if pcv_strToName="" then
			pcv_strToName=pcv_ShippingCompany
		end if
	End If
	Session("pcAdminToName") = pcv_strToName
end if

if instr(pcv_ShippingFullName, " ") then
	pcv_ShippingNameArry=split(pcv_ShippingFullName, " ")
	pcv_strToFirstName=pcv_ShippingNameArry(0)
	pcv_strToLastName=pcv_ShippingNameArry(1)
end if

if Session("pcAdminToFirstName") = "" OR intResetSessions=1 then
	IF pcv_UseAltAddress = "YES" Then
		pcv_strToFirstName = pcv_strAltToFirstName
	Else
		pcv_strToFirstName = pcv_strToFirstName
		if pcv_strToFirstName="" then
			pcv_strToFirstName=pcv_strToName
		end if
	End If
	Session("pcAdminToFirstName") = pcv_strToFirstName
end if

if Session("pcAdminToLastName") = "" OR intResetSessions=1 then
	IF pcv_UseAltAddress = "YES" Then
		pcv_strToFirstName = pcv_strAltToLastName
	Else
		pcv_strToLastName = pcv_strToLastName
	End If
	Session("pcAdminToLastName") = pcv_strToLastName
end if

if Session("pcAdminToFirm") = "" OR intResetSessions=1 then
	If pcv_UseAltAddress = "YES" Then
		pcv_strToFirm = pcv_strAltToFirm
		if pcv_strToFirm="" then
			pcv_strToFirm=pcv_strAltToFullName
		end if
	Else
		pcv_strToFirm = pcv_ShippingCompany
		if pcv_strToFirm="" then
			pcv_strToFirm=pcv_strToName
		end if
	End If
	Session("pcAdminToFirm") = pcv_strToFirm
end if

if Session("pcAdminToPhone") = "" OR intResetSessions=1 then
	pcv_strToPhone = pcv_ShippingPhone
	Session("pcAdminToPhone") = pcv_strToPhone
end if

if Session("pcAdminRecipientEMail") = "" OR intResetSessions=1 then
	If pcv_UseAltAddress = "YES" Then
		pcv_strRecipientEMail = pcv_AltShippingEmail
	Else
		pcv_strRecipientEMail = pcv_ShippingEmail
	End If
	Session("pcAdminRecipientEMail") = pcv_strRecipientEMail
end if

'// DESTINATION ADDRESS
if Session("pcAdminToAddress1") = "" OR IntResetSessions=1 then
	pcv_strToAddress1 = pcv_ShippingAddress
	Session("pcAdminToAddress1") = pcv_strToAddress1
end if
if Session("pcAdminToAddress2") = "" OR intResetSessions=1 then
	pcv_strToAddress2 = pcv_ShippingAddress2
	Session("pcAdminToAddress2") = pcv_strToAddress2
end if
if Session("pcAdminToCity") = "" OR intResetSessions=1 then
	pcv_strToCity = pcv_ShippingCity
	Session("pcAdminToCity") = pcv_strToCity
end if
if Session("pcAdminToState") = "" OR intResetSessions=1 then
	pcv_strToState = pcv_ShippingStateCode
	Session("pcAdminToState") = pcv_strToState
end if
if Session("pcAdminToProvince") = "" OR intResetSessions=1 then
	pcv_strToState = pcv_ShippingProvince
	Session("pcAdminToState") = pcv_strToState
end if
if Session("pcAdminToZip5") = "" OR intResetSessions=1 then
	pcv_strToZip5 = pcv_ShippingZip
	Session("pcAdminToZip5") = pcv_strToZip5
end if
if Session("pcAdminToZip4") = "" OR intResetSessions=1 then
	Session("pcAdminToZip4") = ""
end if
if Session("pcAdminToCountryCode") = "" OR intResetSessions=1 then
	pcv_strToCountryCode = pcv_ShippingCountryCode
	Session("pcAdminToCountryCode") = pcv_strToCountryCode
end if
if Session("pcAdminToPostalCode") = "" OR intResetSessions=1 then
	pcv_strToPostalCode = pcv_ShippingZip
	Session("pcAdminToPostalCode") = pcv_strToPostalCode
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2">Order ID#: <b><%=(scpre+int(pcv_intOrderID))%></b></td>
	</tr>
	<% if pcv_LabelMode="" then %>
		<tr>
			<th colspan="2">USPS Shipping - Choose Label Type</th>
		</tr>
	<% else
		if request.form("submit")<>"" then %>
			<tr>
				<th colspan="2">USPS Shipping - <%=pcv_LabelDescription%></th>
			</tr>
		<% else %>
			<tr>
				<th colspan="2">USPS Shipping - <%=pcv_LabelDescription%> - <a href=<%=pcPageName%>>Change Label Type</a></th>
			</tr>
		<% end if %>
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

					pcs_ValidateTextField "ViewPackages"&pcv_xCounter, false, 0
					pcs_ValidateTextField "ServiceType"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"Pounds"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"Ounces"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"CustomerRefNo"&pcv_xCounter, false, 30
					pcs_ValidateTextField	"Description"&pcv_xCounter, false, 30
					pcs_ValidateTextField	"Value"&pcv_xCounter, false, 30
					pcs_ValidateTextField	"InsuredAmount"&pcv_xCounter, false, 30
					pcs_ValidateTextField	"Size"&pcv_xCounter, false, 30
					pcs_ValidateTextField	"Length"&pcv_xCounter, false, 30
					pcs_ValidateTextField	"Width"&pcv_xCounter, false, 30
					pcs_ValidateTextField	"Height"&pcv_xCounter, false, 30
					pcs_ValidateTextField	"Girth"&pcv_xCounter, false, 30
					if pcv_LabelMode="E" then
						pcs_ValidateTextField	"FlatRate"&pcv_xCounter, false, 0
						pcs_ValidateTextField	"WaiverOfSignature"&pcv_xCounter, false, 0
						pcs_ValidateTextField	"NoHoliday"&pcv_xCounter, false, 0
						pcs_ValidateTextField	"NoWeekend"&pcv_xCounter, false, 0
					end if
					if Session("pcAdminPounds"&pcv_xCounter)="" then Session("pcAdminPounds"&pcv_xCounter) = 0
					if Session("pcAdminOunces"&pcv_xCounter)="" then Session("pcAdminOunces"&pcv_xCounter) = 0
					intPounds=Session("pcAdminPounds"&pcv_xCounter)
					intOunces=Session("pcAdminOunces"&pcv_xCounter)
					intWeightInOunces=(intPounds*16)+intOunces
					Session("pcAdminWeightInOunces"&pcv_xCounter)=intWeightInOunces
				Next
				'..VALIDATE ALL OTHER FIELDS
				pcs_ValidateTextField	"idOrder", false, 0
				pcs_ValidateTextField	"packagecount", false, 0
				pcs_ValidateTextField	"itemsList", false, 0
				pcs_ValidateTextField	"ImageType", true, 0					'<ImageType>

				pcs_ValidateTextField	"Container", false, 0					'<ImageType>
				pcs_ValidateTextField	"ContentType", false, 0					'<ImageType>
				pcs_ValidateTextField	"FirstClassMailType", false, 0					'<ImageType>
				pcs_ValidateTextField	"ContentTypeOther", false, 0					'<ImageType>
				pcs_ValidateTextField	"ToPostalCode", false, 0					'<ImageType>
				pcs_ValidateTextField	"ToCountry", false, 0					'<ImageType>
				pcs_ValidateTextField	"ToPOBoxFlag", false, 0					'<ImageType>


				pcs_ValidateTextField	"LabelOption", false, 0	'<Option>
				pcs_ValidateTextField	"LabelDate", false, 0		'<LabelDate>
				if session("pcAdminLabelDate")<>"" then
					session("pcAdminLabelDate")=exFormatDate(session("pcAdminLabelDate"),"%mm/%dd/%yyyy")
				end if
				pcs_ValidateTextField	"SeparateReceiptPage", false, 0			'<SeparateReceiptPage>
				pcs_ValidateTextField	"FromFirm", isFromFirmRequired, 0		'<FromFirm>
				if pcv_LabelMode="E" then
					pcs_ValidateTextField	"FromFirstName", true, 0			'<FromFirstName>
					pcs_ValidateTextField	"FromLastName", true, 0				'<FromLastName>
					pcs_ValidatePhoneNumber	"FromPhone", true, 0
					if session("pcAdminFromFirm")="" then
						session("pcAdminFromFirm")=session("pcAdminFromFirstName")&" "&session("pcAdminFromLastName")
					end if
				else
					pcs_ValidateTextField	"FromName", false, 0					'<FromName>
				end if
				pcs_ValidateEmailField	"SenderEMail", false, 0
				pcs_ValidateTextField	"FromAddress1", false, 26					'<FromAddress2>
				pcs_ValidateTextField	"FromAddress2", false, 26				'<FromAddress1>
				pcs_ValidateTextField	"FromCity", false, 13						'<FromCity>
				pcs_ValidateTextField	"FromState", false, 2					'<FromState>
				pcs_ValidateTextField	"FromZip5", false, 5						'<FromZip5>
				pcs_ValidateTextField	"FromZip4", false, 4					'<FromZip4>
				pcs_ValidateTextField	"ToFirm", false, 26						'<ToFirm>
				if pcv_LabelMode="E" then
					pcs_ValidateTextField	"ToFirstName", false, 26				'<ToFirstName>
					pcs_ValidateTextField	"ToLastName", false, 26				'<ToLastName>
					pcs_ValidatePhoneNumber	"ToPhone", false, 0
				else
					pcs_ValidateTextField	"ToName", false, 0					'<ToName>
				end if
				pcs_ValidateTextField	"POZipcode", false, 0					'<POZipcode>
				pcs_ValidateEmailField	"RecipientEMail", false, 0				'<RecipientEMail>
				pcs_ValidateTextField	"ToAddress1", false, 0					'<ToAddress2>
				pcs_ValidateTextField	"ToAddress2", false, 0					'<ToAddress1>
				pcs_ValidateTextField	"ToCity", false, 24						'<ToCity>
				pcs_ValidateTextField	"ToState", false, 2						'<ToState>
				pcs_ValidateTextField	"ToZip5", false, 5						'<ToZip5>
				pcs_ValidateTextField	"ToPostalCode", false, 5				'<ToPostalCode>
				pcs_ValidateTextField	"ToZip4", false, 4						'<ToZip4>
				if request("MarkedAsShipped")="1" then
					pcs_ValidateTextField "AdmComments", false, 0 '// Admin Comments
				end if

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Check for Validation Errors. Do not proceed if there are errors.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				If pcv_intErr>0 Then
					response.redirect pcPageName & "?sub=1&msg=" & pcv_strGenericPageError
				Else
					if pcv_LabelMode="E" AND request("TandC")<>"1" then
						response.redirect pcPageName & "?LabelMode="&pcv_LabelMode&"&msg=There was an error processing your request.<br>You must agree to the Terms and Conditions listed under the Package Information Tab."
					end if
					'//USPS Variables
					query="SELECT active, AccessLicense, userID FROM ShipmentTypes WHERE idshipment=4;"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					pcv_USPSUserID=trim(rs("userID"))
					pcv_USPSLabelServer=trim(rs("AccessLicense"))
					pcv_USPSActive=rs("active")


					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Build Our Transaction.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					select case pcv_LabelMode
						case "D"
							if USPS_TESTMODE="1" then
								objUSPSClass.NewXMLTransaction "DelivConfirmCertifyV3", "DelivConfirmCertifyV3.0Request", pcv_USPSUserID
								strXMLClosingTag="DelivConfirmCertifyV3.0Request"
							else
								objUSPSClass.NewXMLTransaction "DeliveryConfirmationV3", "DeliveryConfirmationV3.0Request", pcv_USPSUserID
								strXMLClosingTag="DeliveryConfirmationV3.0Request"
							end if
						case "S"
							if USPS_TESTMODE="1" then
								objUSPSClass.NewXMLTransaction "SignatureConfirmationCertifyV3", "SigConfirmCertifyV3.0Request", pcv_USPSUserID
								strXMLClosingTag="SigConfirmCertifyV3.0Request"
							else
								objUSPSClass.NewXMLTransaction "SignatureConfirmationV3", "SignatureConfirmationV3.0Request", pcv_USPSUserID
								strXMLClosingTag="SignatureConfirmationV3.0Request"
							end if
						case "E"
							if USPS_TESTMODE="1" then
								objUSPSClass.NewXMLTransaction "ExpressMailLabelCertify", "ExpressMailLabelCertifyRequest", pcv_USPSUserID
								strXMLClosingTag="ExpressMailLabelCertifyRequest"
							else
								objUSPSClass.NewXMLTransaction "ExpressMailLabel", "ExpressMailLabelRequest", pcv_USPSUserID
								strXMLClosingTag="ExpressMailLabelRequest"
							end if
						case "FCINT"
							if USPS_TESTMODE="1" then
								objUSPSClass.NewXMLTransaction "FirstClassMailIntlCertify", "FirstClassMailIntlCertifyRequest", pcv_USPSUserID
								strXMLClosingTag="FirstClassMailIntlCertifyRequest"
							else
								objUSPSClass.NewXMLTransaction "FirstClassMailIntl", "FirstClassMailIntlRequest", pcv_USPSUserID
								strXMLClosingTag="FirstClassMailIntlRequest"
							end if
						case "PMINT"
							if USPS_TESTMODE="1" then
								objUSPSClass.NewXMLTransaction "PriorityMailIntlCertify", "PriorityMailIntlCertifyRequest", pcv_USPSUserID
								strXMLClosingTag="PriorityMailIntlCertifyRequest"
							else
								objUSPSClass.NewXMLTransaction "PriorityMailIntl", "PriorityMailIntlRequest", pcv_USPSUserID
								strXMLClosingTag="PriorityMailIntlRequest"
							end if
						case "EMINT"
							if USPS_TESTMODE="1" then
								objUSPSClass.NewXMLTransaction "ExpressMailIntlCertify", "ExpressMailIntlCertifyRequest", pcv_USPSUserID
								strXMLClosingTag="ExpressMailIntlCertifyRequest"
							else
								objUSPSClass.NewXMLTransaction "ExpressMailIntl", "ExpressMailIntlRequest", pcv_USPSUserID
								strXMLClosingTag="ExpressMailIntlRequest"
							end if
					end select

					pcv_xCounter = 1
					pcv_strTotalWeight = 0
					errnum = 0

					'///////////////////////////////////////////////////////////////////////
					'// START LOOP FOR PACKAGE TAG
					'///////////////////////////////////////////////////////////////////////
					For pcv_xCounter = 1 to pcPackageCount
						'// If the package was processed, skip it.
						if pcLocalArray(pcv_xCounter-1) <> "shipped" then
							'// LABEL SPECIFICATION
							If pcv_LabelMode = "FCINT" OR pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" Then
								'Option - for future use
								objUSPSClass.AddNewNode "Revision", "2", 1
							Else
								if pcv_LabelMode="E" then
									objUSPSClass.WriteEmptyParent "Option", "/"
								else
									objUSPSClass.AddNewNode "Option", Session("pcAdminLabelOption"), isLabelOptionValueReq
								end if
							End If
							If pcv_LabelMode = "FCINT" OR pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" Then
								objUSPSClass.AddNewNode "FromFirstName", replace(Session("pcAdminFromFirstName"),"''","'"), 1
								objUSPSClass.AddNewNode "FromLastName", replace(Session("pcAdminFromLastName"),"''","'"), 1
							Else
								if pcv_LabelMode="E" then
									objUSPSClass.WriteEmptyParent "EMCAAccount", "/"
									objUSPSClass.WriteEmptyParent "EMCAPassword", "/"
									objUSPSClass.WriteEmptyParent "ImageParameters", "/"
									objUSPSClass.AddNewNode "FromFirstName", replace(Session("pcAdminFromFirstName"),"''","'"), 1
									objUSPSClass.AddNewNode "FromLastName", replace(Session("pcAdminFromLastName"),"''","'"), 1
								else
									objUSPSClass.WriteEmptyParent "ImageParameters", "/"
									objUSPSClass.AddNewNode "FromName", replace(Session("pcAdminFromName"),"''","'"), 1
								end if
							End If
							objUSPSClass.AddNewNode "FromFirm", replace(Session("pcAdminFromFirm"),"''","'"), 1
							objUSPSClass.AddNewNode "FromAddress1", replace(Session("pcAdminFromAddress1"),"''","'"), 1
							objUSPSClass.AddNewNode "FromAddress2", replace(Session("pcAdminFromAddress1"),"''","'"), 1
							objUSPSClass.AddNewNode "FromCity", replace(Session("pcAdminFromCity"),"''","'"), 1
							objUSPSClass.AddNewNode "FromState", Session("pcAdminFromState"), 1
							objUSPSClass.AddNewNode "FromZip5", Session("pcAdminFromZip5"), 1
							If Session("pcAdminFromZip4")&""<>"" Then
								objUSPSClass.AddNewNode "FromZip4", Session("pcAdminFromZip4"), 1
							End If
							If pcv_LabelMode = "FCINT" OR pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" Then
								objUSPSClass.AddNewNode "FromPhone", fnStripPhone(Session("pcAdminFromPhone")), 1
								If pcv_LabelMode = "EMINT" OR pcv_LabelMode = "PMINT" Then
									objUSPSClass.AddNewNode "FromCustomsReference", Session("pcAdminFromCustomsReference"), 1
								End If
								objUSPSClass.AddNewNode "ToName", replace(Session("pcAdminToName"),"''","'"), 1
							Else
								if pcv_LabelMode="E" then
									objUSPSClass.AddNewNode "FromPhone", fnStripPhone(Session("pcAdminFromPhone")), 1
									objUSPSClass.AddNewNode "ToFirstName", replace(Session("pcAdminToFirstName"),"''","'"), 1
									objUSPSClass.AddNewNode "ToLastName", replace(Session("pcAdminToLastName"),"''","'"), 1
								else
									objUSPSClass.AddNewNode "ToName", replace(Session("pcAdminToName"),"''","'"), 1
								end if
							End If
							objUSPSClass.AddNewNode "ToFirm", replace(Session("pcAdminToFirm"),"''","'"), 1
							objUSPSClass.AddNewNode "ToAddress1", replace(Session("pcAdminToAddress2"),"''","'"), 1
							objUSPSClass.AddNewNode "ToAddress2", replace(Session("pcAdminToAddress1"),"''","'"), 1
							objUSPSClass.AddNewNode "ToCity", replace(Session("pcAdminToCity"),"''","'"), 1
							If pcv_LabelMode = "FCINT" OR pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" Then
								objUSPSClass.AddNewNode "ToProvince", Session("pcAdminToState"), 0
								objUSPSClass.AddNewNode "ToCountry", USPSCountry(Session("pcAdminToCountry")), 1


								objUSPSClass.AddNewNode "ToPostalCode", Session("pcAdminToPostalCode"), 1
								If Session("pcAdminToPOBoxFlag")&""="" Then
									objUSPSClass.AddNewNode "ToPOBoxFlag", "N", 1
								Else
									objUSPSClass.AddNewNode "ToPOBoxFlag", "Y", 1
								End If
								objUSPSClass.AddNewNode "ToPhone", fnStripPhone(Session("pcAdminToPhone")), 0
								objUSPSClass.AddNewNode "ToEmail", Session("pcAdminRecipientEmail"), 0
							Else
								objUSPSClass.AddNewNode "ToState", Session("pcAdminToState"), 1
								objUSPSClass.AddNewNode "ToZip5", Session("pcAdminToZip5"), 1
								objUSPSClass.AddNewNode "ToZip4", Session("pcAdminToZip4"), 1
							End If

							if pcv_LabelMode="E" then
								objUSPSClass.AddNewNode "ToPhone", fnStripPhone(Session("pcAdminToPhone")), 1
							end if
							if pcv_LabelMode = "FCINT" then
								objUSPSClass.AddNewNode "FirstClassMailType", Session("pcAdminFirstClassMailType"), 1
							end if
							if pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" then
								objUSPSClass.AddNewNode "Container", Session("pcAdminContainer"), 1
							end if
							if pcv_LabelMode = "PMINT" then
								'objUSPSClass.AddNewNode "FirstClassMailType", Session("pcAdminFirstClassMailType"), 1
							end if
							if pcv_LabelMode = "EMINT" then
								'objUSPSClass.AddNewNode "FirstClassMailType", Session("pcAdminFirstClassMailType"), 1
							end if


							if pcv_LabelMode="E" OR pcv_LabelMode="S" OR pcv_LabelMode="D" then
								if UCASE(Session("pcAdminFlatRate"&pcv_xCounter))="TRUE" then
									objUSPSClass.WriteEmptyParent "WeightInOunces", "/"
								else
									objUSPSClass.AddNewNode "WeightInOunces", Session("pcAdminWeightInOunces"&pcv_xCounter), 1
								end if

								if Session("pcAdminFlatRate"&pcv_xCounter)<>"TRUE" then
									Session("pcAdminFlatRate"&pcv_xCounter)=cstr("")
								end if

								if Session("pcAdminStandardizeAddressFalse"&pcv_xCounter)<>"TRUE" then
									Session("pcAdminStandardizeAddressFalse"&pcv_xCounter)=cstr("")
								end if

								if Session("pcAdminWaiverOfSignature"&pcv_xCounter)<>"TRUE" then
									Session("pcAdminWaiverOfSignature"&pcv_xCounter)=cstr("")
								end if

								if Session("pcAdminNoHoliday"&pcv_xCounter)<>"TRUE" then
									Session("pcAdminNoHoliday"&pcv_xCounter)=cstr("")
								end if

								if Session("pcAdminNoWeekend"&pcv_xCounter)<>"TRUE" then
									Session("pcAdminNoWeekend"&pcv_xCounter)=cstr("")
								end if

								if pcv_LabelMode="E" then
									objUSPSClass.AddNewNode "FlatRate",	Session("pcAdminFlatRate"&pcv_xCounter), 1
									objUSPSClass.AddNewNode "StandardizeAddress",	Session("pcAdminStandardizeAddressFalse"), 1
									objUSPSClass.AddNewNode "WaiverOfSignature",	Session("pcAdminWaiverOfSignature"&pcv_xCounter), 1
									objUSPSClass.AddNewNode "NoHoliday",	Session("pcAdminNoHoliday"&pcv_xCounter), 1
									objUSPSClass.AddNewNode "NoWeekend",	Session("pcAdminNoWeekend"&pcv_xCounter), 1
								end if
								if pcv_LabelMode="D" OR pcv_LabelMode="S" then
									objUSPSClass.AddNewNode "ServiceType", session("pcAdminServiceType"&pcv_xCounter), 1
								end if
								if Session("pcAdminSeparateReceiptPage")="1" then
									objUSPSClass.AddNewNode "SeparateReceiptPage", "True", 0
								end if
							else
								objUSPSClass.WriteParent "ShippingContents", ""
								objUSPSClass.WriteParent "ItemDetail", ""
									objUSPSClass.AddNewNode "Description", Session("pcAdminDescription"&pcv_xCounter), 1
									objUSPSClass.AddNewNode "Quantity", "1", 1
									objUSPSClass.AddNewNode "Value", Session("pcAdminValue"&pcv_xCounter), 1
									objUSPSClass.AddNewNode "NetPounds", Session("pcAdminPounds"&pcv_xCounter), 1
									objUSPSClass.AddNewNode "NetOunces", Session("pcAdminOunces"&pcv_xCounter), 1
									objUSPSClass.AddNewNode "HSTariffNumber", "", 1
									objUSPSClass.AddNewNode "CountryOfOrigin", "United States", 1
								objUSPSClass.WriteParent "ItemDetail", "/"
								objUSPSClass.WriteParent "ShippingContents", "/"
							end if

							If pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" Then
								objUSPSClass.AddNewNode "InsuredAmount", Session("pcAdminInsuredAmount"&pcv_xCounter), 1
							End If
							If pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" OR pcv_LabelMode = "FCINT" Then
								objUSPSClass.AddNewNode "Postage", "", 1
								objUSPSClass.AddNewNode "GrossPounds", Session("pcAdminPounds"&pcv_xCounter), 1
								objUSPSClass.AddNewNode "GrossOunces", Session("pcAdminOunces"&pcv_xCounter), 1
							End If

							If pcv_LabelMode = "FCINT" Then
								objUSPSClass.AddNewNode "Machinable", Session("pcAdminMachinable"&pcv_xCounter), 1
							End If

							If pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" OR pcv_LabelMode = "FCINT" Then
								objUSPSClass.AddNewNode "ContentType", Session("pcAdminContentType"), 1
								objUSPSClass.AddNewNode "ContentTypeOther", Session("pcAdminContentTypeOther"), 1
								objUSPSClass.AddNewNode "Agreement", "Y", 1
								objUSPSClass.AddNewNode "ImageType", Session("pcAdminImageType"), 1
								objUSPSClass.AddNewNode "CustomerRefNo", Session("pcAdminCustomerRefNo"), 0
								If pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" Then
									objUSPSClass.AddNewNode "POZipCode", Session("pcAdminPOZipCode"), isPOZipCodeValueReq
								End If
								objUSPSClass.AddNewNode "LabelDate", Session("pcAdminLabelDate"), 1
								if pcv_LabelMode = "FCINT" then
									objUSPSClass.AddNewNode "Container", Session("pcAdminContainer"), 1
								end if
								If Session("pcAdminLength"&pcv_xCounter)>"12" OR Session("pcAdminWidth"&pcv_xCounter)>"12" OR Session("pcAdminHeight"&pcv_xCounter)>"12" Then
									objUSPSClass.AddNewNode "Size", "LARGE", 0
								Else
									objUSPSClass.AddNewNode "Size", "REGULAR", 0
								End If

								objUSPSClass.AddNewNode "Length", Session("pcAdminLength"&pcv_xCounter), 0
								objUSPSClass.AddNewNode "Width", Session("pcAdminWidth"&pcv_xCounter), 0
								objUSPSClass.AddNewNode "Height", Session("pcAdminHeight"&pcv_xCounter), 0
								objUSPSClass.AddNewNode "Girth", Session("pcAdminGirth"&pcv_xCounter), 0
							Else
								objUSPSClass.AddNewNode "POZipCode", Session("pcAdminPOZipCode"), isPOZipCodeValueReq
								objUSPSClass.AddNewNode "ImageType", Session("pcAdminImageType"), 1
								objUSPSClass.AddNewNode "LabelDate", Session("pcAdminLabelDate"), 0
								objUSPSClass.AddNewNode "CustomerRefNo", Session("pcAdminCustomerRefNo"), 0
								objUSPSClass.AddNewNode "AddressServiceRequested", "", 0
							End If

							If pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" OR pcv_LabelMode = "FCINT" Then
							Else
								if pcv_LabelMode="E" then
									objUSPSClass.AddNewNode "SenderName", Session("pcAdminFromFirstName")&" "&Session("pcAdminFromLastName"), 0
								else
									objUSPSClass.AddNewNode "SenderName", Session("pcAdminFromName"), 0
								end if
								objUSPSClass.AddNewNode "SenderEMail", Session("pcAdminSenderEMail"), 0
								if pcv_LabelMode="E" then
									objUSPSClass.AddNewNode "RecipientName", Session("pcAdminToFirstName")&" "&Session("pcAdminToLastName"), 0
								else
									objUSPSClass.AddNewNode "RecipientName", Session("pcAdminToName"), 0
								end if
								objUSPSClass.AddNewNode "RecipientEMail", Session("pcAdminRecipientEMail"), 0
							End If
						End if '// end skip shipped packages
					Next

					ObjUSPSClass.WriteParent strXMLClosingTag, "/"

					'///////////////////////////////////////////////////////////////////////
					'// END LOOP
					'///////////////////////////////////////////////////////////////////////

					'//Clear illegal ampersand characters from XML
					USPS_postdata=replace(USPS_postdata, "&XML", "andXML")
					USPS_postdata=replace(USPS_postdata, "&", "and")
					USPS_postdata=replace(USPS_postdata, "andamp;", "and")
					USPS_postdata=replace(USPS_postdata, "andXML", "&XML")

					'// Print out our newly formed request xml
					'response.write USPS_postdata&"<HR>"
					'response.end

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Send Our Transaction.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					err.clear

					call objUSPSClass.SendXMLRequest(USPS_postdata, pcv_USPSLabelServer)

					'// Print out our response
					'response.write "<HR>"& USPS_result
					'response.end

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Load Our Response.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					call objUSPSClass.LoadXMLResults(USPS_result)


					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Check for errors from USPS.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					'//SOME ERROR CHECKING HERE
					call objUSPSClass.XMLResponseVerify(ErrPageName)

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Redirect with a Message OR complete some task.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					if NOT len(pcv_strErrorMsg)>0 then
						'// -----------------------------------
						if strErrorNumber="" AND strErrorDescription="" then
							select case pcv_LabelMode
								case "D"
									if USPS_TESTMODE="1" then
										Set LNodes = objOutputXMLDoc.selectNodes("//DelivConfirmCertifyV3.0Response")
									else
										Set LNodes = objOutputXMLDoc.selectNodes("//DeliveryConfirmationV3.0Response")
									end if
									intLNode=0
									For Each LNode In LNodes
										strTrackingNumber=LNode.selectSingleNode("DeliveryConfirmationNumber").Text
										strGraphicImage=LNode.selectSingleNode("DeliveryConfirmationLabel").Text
										if Session("pcAdminSeparateReceiptPage")="1" then
											strGraphicReceipt=LNode.selectSingleNode("DeliveryConfirmationReceipt").Text
										end if
									Next
								case "S"
									if USPS_TESTMODE="1" then
										Set LNodes = objOutputXMLDoc.selectNodes("//SigConfirmCertifyV3.0Response")
									else
										Set LNodes = objOutputXMLDoc.selectNodes("//SignatureConfirmationV3.0Response")
									end if
									intLNode=0
									For Each LNode In LNodes
										strTrackingNumber=LNode.selectSingleNode("SignatureConfirmationNumber").Text
										strGraphicImage=LNode.selectSingleNode("SignatureConfirmationLabel").Text
										if Session("pcAdminSeparateReceiptPage")="1" then
											strGraphicReceipt=LNode.selectSingleNode("SignatureConfirmationReceipt").Text
										end if
									Next
								case "E"
									if USPS_TESTMODE="1" then
										Set LNodes = objOutputXMLDoc.selectNodes("//ExpressMailLabelCertifyResponse")
									else
										Set LNodes = objOutputXMLDoc.selectNodes("//ExpressMailLabelResponse")
									end if
									intLNode=0
									For Each LNode In LNodes
										strToFirstName=LNode.selectSingleNode("ToFirstName").Text
										strToLastName=LNode.selectSingleNode("ToLastName").Text
										strToFirm=LNode.selectSingleNode("ToFirm").Text
										strToAddress1=LNode.selectSingleNode("ToAddress1").Text
										strToAddress2=LNode.selectSingleNode("ToAddress2").Text
										strToCity=LNode.selectSingleNode("ToCity").Text
										strToState=LNode.selectSingleNode("ToState").Text
										strToZip5=LNode.selectSingleNode("ToZip5").Text
										strToZip4=LNode.selectSingleNode("ToZip4").Text
										strPostage=LNode.selectSingleNode("Postage").Text
										strTrackingNumber=LNode.selectSingleNode("EMConfirmationNumber").Text
										strGraphicImage=LNode.selectSingleNode("EMLabel").Text
										if Session("pcAdminSeparateReceiptPage")="1" then
											strGraphicReceipt=LNode.selectSingleNode("EMReceipt").Text
										end if
									Next

								case "FCINT"
									if USPS_TESTMODE="1" then
										Set LNodes = objOutputXMLDoc.selectNodes("//FirstClassMailIntlCertifyResponse")
									else
										Set LNodes = objOutputXMLDoc.selectNodes("//FirstClassMailIntlResponse")
									end if
									intLNode=0
									For Each LNode In LNodes
										strTrackingNumber=LNode.selectSingleNode("BarcodeNumber").Text
										strGraphicImage=LNode.selectSingleNode("LabelImage").Text
									Next

								case "PMINT"
									if USPS_TESTMODE="1" then
										Set LNodes = objOutputXMLDoc.selectNodes("//PriorityMailIntlCertifyResponse")
									else
										Set LNodes = objOutputXMLDoc.selectNodes("//PriorityMailIntlResponse")
									end if
									intLNode=0
									For Each LNode In LNodes
										strTrackingNumber=LNode.selectSingleNode("BarcodeNumber").Text
										strGraphicImage=LNode.selectSingleNode("LabelImage").Text
									Next
								case "EMINT"
									if USPS_TESTMODE="1" then
										Set LNodes = objOutputXMLDoc.selectNodes("//ExpressMailIntlCertifyResponse")
									else
										Set LNodes = objOutputXMLDoc.selectNodes("//ExpressMailIntlResponse")
									end if
									intLNode=0
									For Each LNode In LNodes
										strTrackingNumber=LNode.selectSingleNode("BarcodeNumber").Text
										strGraphicImage=LNode.selectSingleNode("LabelImage").Text
									Next

							end select
						end if

						if strTrackingNumber<>"" AND strGraphicImage<>"" then
							GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName=""label"&strTrackingNumber&"."&session("pcAdminImageType")&""">"&strGraphicImage&"</Base64Data>"
							if strGraphicReceipt<>"" then
								ReceiptXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName=""receipt"&strTrackingNumber&"."&session("pcAdminImageType")&""">"&strGraphicReceipt&"</Base64Data>"
							else
								ReceiptXML=""
							end if

							'Create MSXML DOMDocument Object
							Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument"&scXML)
							objXMLDoc.async = False
							objXMLDoc.validateOnParse = False

							'And load it from the request stream
							If objXMLDoc.loadXML (GraphicXML) Then

								'Use ADO stream to save the binary data
								Set objStream = Server.CreateObject("ADODB.Stream")
								objStream.Type = 1
								objStream.Open

								'The nodeTypedValue automagically converts Base64 data to binary data
								'Write that binary data to the stream
								objStream.Write objXMLDoc.selectSingleNode("/Base64Data").nodeTypedValue

								'Get the FileName attribute's value
								strFileName = objXMLDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue

								'on error resume next
								err.clear
								'Save the binary stream to the file
								'response.write strFileName
								objStream.SaveToFile server.mappath("USPSLabels\" & strFileName), 2
								if err.number<>0 then
									'response.write "This label has already been saved..."
								end if
								objStream.Close()

								Set objStream = Nothing

								'// LINK TO FILE
								strLabelLink = "<a href='USPSLabels\" & strFileName&"' target='_blank'>View/Print Label</a>"
								if ReceiptXML<>"" then
									If objXMLDoc.loadXML (ReceiptXML) Then
										'Use ADO stream to save the binary data
										Set objStream = Server.CreateObject("ADODB.Stream")
										objStream.Type = 1
										objStream.Open

										'The nodeTypedValue automagically converts Base64 data to binary data
										'Write that binary data to the stream
										objStream.Write objXMLDoc.selectSingleNode("/Base64Data").nodeTypedValue

										'Get the FileName attribute's value
										strFileName = objXMLDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue

										'on error resume next
										err.clear
										'Save the binary stream to the file
										'response.write strFileName
										objStream.SaveToFile server.mappath("USPSLabels\" & strFileName), 2
										if err.number<>0 then
											'response.write "This label has already been saved..."
										end if
										objStream.Close()

										Set objStream = Nothing

										'// LINK TO FILE
										if Session("pcAdminSeparateReceiptPage")="1" then
											strReceiptLink="<br><a href='USPSLabels\" & strFileName&"' target='_blank'>View/Print Receipt</a><br>"
										end if
									Else
										strReceiptLink=""
									End if
								End If

								'// SAVE all data to the database
								pcv_intOrderID=Session("pcAdminOrderID")

								dim dtShippedDate
								dtShippedDate=Date()
								if pcv_shippedDate<>"" then
									'dtShippedDate=objFedExClass.pcf_FedExDateFormat(dtShippedDate)
									if SQL_Format="1" then
										dtShippedDate=(day(dtShippedDate)&"/"&month(dtShippedDate)&"/"&year(dtShippedDate))
									else
										dtShippedDate=(month(dtShippedDate)&"/"&day(dtShippedDate)&"/"&year(dtShippedDate))
									end if
								end if

								if scDB="Access" then
									pcInsertDate="#"
								else
									pcInsertDate="'"
								end if

								err.clear
								query="INSERT INTO pcPackageInfo (idOrder, pcPackageInfo_PackageNumber, pcPackageInfo_PackageWeight, pcPackageInfo_ShipToName, pcPackageInfo_ShipToAddress1, pcPackageInfo_ShipToAddress2, pcPackageInfo_ShipToCity, pcPackageInfo_ShipToStateCode, pcPackageInfo_ShipToZip, pcPackageInfo_ShipToCountry, pcPackageInfo_ShipToEmail, pcPackageInfo_ShipFromCompanyName, pcPackageInfo_ShipFromAttentionName, pcPackageInfo_ShipFromAddress1, pcPackageInfo_ShipFromAddress2, pcPackageInfo_ShipFromCity, pcPackageInfo_ShipFromStateProvinceCode, pcPackageInfo_ShipFromPostalCode, pcPackageInfo_ShipFromCountryCode, pcPackageInfo_PackageLength, pcPackageInfo_PackageWidth, pcPackageInfo_PackageHeight, pcPackageInfo_Status, pcPackageInfo_UPSServiceCode, pcPackageInfo_TrackingNumber, pcPackageInfo_ShipMethod, pcPackageInfo_ShippedDate, pcPackageInfo_Comments, pcPackageInfo_ShipToContactName, pcPackageInfo_UPSLabelFormat, pcPackageInfo_MethodFlag) "
								query=query&"VALUES ("&pcv_intOrderID&", 1, 0, '"&pcPackageInfo_ShipToName&"', '"&Session("pcAdminToAddress1")&"', '"&Session("pcAdminToAddress2")&"', '"&Session("pcAdminToCity")&"', '"&Session("pcAdminToState")&"', '"&Session("pcAdminToZip5")&"', 'US', '"&pcPackageInfo_ShipToEmail&"', '"&Session("pcAdminFromFirm")&"', '"&pcPackageInfo_ShipFromAttentionName&"', '"&Session("pcAdminFromAddress1")&"', '"&Session("pcAdminFromAddress2")&"', '"&Session("pcAdminFromCity")&"', '"&Session("pcAdminFromState")&"', '"&Session("pcAdminFromZip5")&"', 'US', '"&pcPackageInfo_PackageLength&"', '"&pcPackageInfo_PackageWidth&"', '"&pcPackageInfo_PackageHeight&"', 0, '"&pcv_LabelMode&"', '"&strTrackingNumber&"', '"&session("pcAdminServiceType1")&"', "&pcInsertDate&dtShippedDate&pcInsertDate&", '"&pcPackageInfo_Comments&"', '"&pcPackageInfo_ShipToContactName&"', '"&Session("pcAdminImageType")&"', 4);"
								set rs=connTemp.execute(query)
								set rs=nothing
								if err.number<>0 then
									response.write err.description
									response.end
								end if

								query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & pcv_intOrderID & " AND pcPackageInfo_TrackingNumber='"&strTrackingNumber&"' ORDER by pcPackageInfo_ID DESC;"
								set rs=connTemp.execute(query)
								if err.number<>0 then
									response.write err.description
									response.end
								else
									pcv_PackageID=rs("pcPackageInfo_ID")
								end if
								set rs=nothing

								pcA=split(Session("pcGlobalArray"),",")
								For i=lbound(pcA) to ubound(pcA)
									if trim(pcA(i)<>"") then
										query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=2, pcPackageInfo_ID=" & pcv_PackageID & " WHERE (idorder=" & pcv_intOrderID & " AND idProductOrdered=" & pcA(i) & ");"
										'response.write query
										'response.end
										set rs=connTemp.execute(query)
										set rs=nothing
									end if
								Next

								if request("MarkedAsShipped")="1" then
									if scDB="SQL" then
										query="UPDATE pcPackageInfo SET pcPackageInfo_ShippedDate='" & dtShippedDate & "',, pcPackageInfo_Comments='" & Session("pcAdminAdmComments") & "' WHERE idOrder="&pcv_intOrderID&" AND pcPackageInfo_ID="&pcv_PackageID&";"
									else
										query="UPDATE pcPackageInfo SET pcPackageInfo_ShippedDate=#" & dtShippedDate & "#,  pcPackageInfo_Comments='" & Session("pcAdminAdmComments") & "' WHERE idOrder="&pcv_intOrderID&" AND pcPackageInfo_ID="&pcv_PackageID&";"
									end if
									set rs=connTemp.execute(query)
									set rs=nothing

									query="DELETE FROM pcAdminComments WHERE idorder=" & pcv_intOrderID & " AND pcACom_ComType=2 AND pcPackageInfo_ID=" & pcv_PackageID & ";"
									set rstemp=connTemp.execute(query)
									query="INSERT INTO pcAdminComments (idorder,pcACom_ComType,pcACom_Comments,pcDropShipper_ID,pcACom_IsSupplier,pcPackageInfo_ID) VALUES (" & pcv_intOrderID & ",2,'" & Session("pcAdminAdmComments") & "',0,0," & pcv_PackageID & ");"
									set rstemp=connTemp.execute(query)

									query="UPDATE ProductsOrdered SET pcPrdOrd_Shipped=1 WHERE idOrder="&pcv_intOrderID&" AND pcPackageInfo_ID="&pcv_PackageID&";"
									set rsQ=connTemp.execute(query)
									set rsQ=nothing

									pcv_SendCust="1"
									pcv_SendAdmin="0"

									pcv_LastShip="0"
									query="SELECT ProductsOrdered.pcPrdOrd_Shipped FROM ProductsOrdered INNER JOIN Orders ON (ProductsOrdered.idorder=Orders.idorder AND ProductsOrdered.pcPrdOrd_Shipped=0) WHERE Orders.idorder=" & pcv_intOrderID & " AND Orders.orderstatus<>4;"
									set rs=connTemp.execute(query)
									if not rs.eof then
										pcv_LastShip="0"
									else
										pcv_LastShip="1"
									end if
									set rs=nothing

									query="SELECT * FROM productsOrdered WHERE pcPrdOrd_Shipped<>1 AND idorder=" & pcv_intOrderID & ";"
									set rsQ=connTemp.execute(query)
									if rsQ.eof then
										query="UPDATE Orders SET orderStatus=4 WHERE idorder=" & pcv_intOrderID & ";"
									else
										query="UPDATE Orders SET orderStatus=7 WHERE idorder=" & pcv_intOrderID & ";"
									end if
									set rsQ=nothing
									set rs=connTemp.execute(query)
									set rs=nothing

									If pcv_LastShip="1" Then
										'// Perform a Google Action
										pcv_strGoogleMethod = "mark" ' // Marks the order shipped at Google
										%> <!--#include file="../includes/GoogleCheckout_OrderManagement.asp"--> <%
									End If
									%>
									<!--#include file="../pc/inc_PartShipEmail.asp"-->

									<% '// End
									Session("pcAdminAdmComments")=""
								end if

							Else
								'Failed to load the document
								response.write "<br><br>Failed to load doc..."
							End If
						Else
							response.write "NO LABEL WAS GENERATED<hr>"
							response.write strErrorNumber &": "&	pcv_strErrorMsg
							response.write "<HR>"& USPS_result
							'response.end
						End If
						%>
						<table class="pcCPcontent">
							<tr>
							  <td colspan="2" class="pcCPspacer"><p>Your label has been created and saved. You can click on the link below to print your label. You 																will be able to access this label from the order details shipping area and also have access to tracking information.<br>
								<br>
								<%=strLabelLink%></p>
								<% if strReceiptLink<>"" then %>
									<p>
									<br>
									Your request includes a separate receipt. You can access and print the receipt from the link below.<br>
									<%= strReceiptLink %></p>
								<% end if %>
								<p>
								<br>
								<a href="OrdDetails.asp?id=<%=pcv_intOrderID%>">Back to order details &gt;&gt;</a><br><br></p>
								</td>
							</tr>
						</table>
						<%
						response.End()
					End If ' If pcv_intErr>0 Then
				end if
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
'//qualify message
If instr(msg, "Gross Pounds must be a positive numeric value 0 to 4") Then
	msg = "First Class weight limit is 4 pounds. Please choose another label type by clicking the link above."
End If
If instr(msg, "Each Shipping item must have description, quantity, and value") Then
	msg = "A package &quot;Description&quot; as well as the package &quot;Value&quot; are both required and can be found under each &quot;Package Information&quot; tab."
End If
				if msg<>"" then %>
					<div class="pcCPmessage">
						<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
					</div>
				<% end if %>
				<% if pcv_LabelMode="" then
					call opendb()
					query="SELECT AccessLicense FROM ShipmentTypes WHERE idshipment=4;"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					pcv_TempUSPSLabelServer=trim(rs("AccessLicense"))
					set rs=nothing
					call closedb()

					if pcv_TempUSPSLabelServer="" then %>
						<table class="pcCPcontent">
							<tr>
								<td width="1100" colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
							  <th colspan="2">USPS Labels API</th>
							</tr>
							<tr>
								<td colspan="2"><p>It appears that your  USPS Secured Server URL has not been set. Please go to your USPS<strong> </strong><a href="USPS_EditLicense.asp">Production Server Settings</a> and populate the &quot;Secured Server URL&quot; with the URL that USPS provided you when you activated your web tools account. </p>
							  <br /><p>Prior to using the Labels API, you must submit a separate request to icustomercare@usps.com to provide you permissions to use the Labels API. </p></td>
							</tr>
						</table>

					<% else %>
					<form name="form1" method="get" action="<%=pcPageName%>" class="pcForms">
						<table class="pcCPcontent">
							<% If Session("pcAdminToCountryCode") = "US" Then %>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
								  <th colspan="2"><input type="radio" name="LabelMode" id="LabelMode" value="D" class="clearBorder" checked>
									&nbsp;Delivery Confirmation</th>
								</tr>
								<tr>
									<td colspan="2"><p>With Delivery Confirmation you and your customers can access information on   the Internet about the delivery status of a package shipped via USPS. From your   ProductCart store  or from   the USPS   website, you can   check the delivery status of Delivery Confirmation packages shipped via Priority   Mail, First-Class Mail parcel, and Package Services (Standard Post, Bound Printed   Matter, Media Mail, and Library Mail). The information returned will include the   date, time, and ZIP Code of delivery, as well as attempted deliveries,   forwarding, and returns. Delivery Confirmation service is not available to   APO/FPO addresses, foreign countries, or most U.S.   territories.<br />
									<br />
									Postage is required on these labels, as well as the   Confirmation Services charge (known as the “electronic option rate”) for   Delivery Confirmation. This discounted “electronic option rate” for Confirmation   Services must be added into the total postage amount affixed to these labels (by   using stamps, meter strips, or other indicia). The Delivery Confirmation fee   varies by different service and is significantly   discounted.
									</ol>
									</p></td>
								</tr>
							<% End If %>
							<% If Session("pcAdminToCountryCode") = "US" Then %>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
								  <th colspan="2"><input type="radio" name="LabelMode" id="LabelMode" value="S" class="clearBorder">&nbsp;Signature Confirmation</th>
								</tr>
								<tr>
									<td colspan="2">With the USPS’s Signature   Confirmation, you or your customers can access information on the Internet about   the delivery status of First-Class Mail parcels, Priority Mail and Package   Services (Standard Post, Bound Printed Matter, Media Mail, and Library Mail),   including the date, time, and ZIP Code of delivery, as well as attempted   deliveries, forwarding, and returns. Signature Confirmation service is not   available to APO/FPO addresses, foreign countries, or most U.S.   territories. <br />
									<br />
									The charge (known as the “electronic option rate”) for   Signature Confirmation is $1.30 for Priority Mail, First-Class mail parcels, and   Package Services parcels. From ProductCart or from   the USPS   website, you can   check the delivery status of Signature Confirmation labels (barcodes) generated   by these Web Tools.</td>
								</tr>
							<% End If %>
							<% If Session("pcAdminToCountryCode") = "US" Then %>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
								  <th colspan="2"><input type="radio" name="LabelMode" id="LabelMode" value="E" class="clearBorder">&nbsp;Online Express Mail Label</th>
								</tr>
								<tr>
									<td colspan="2"> Express Mail is the U.S.   Postal Service’s fastest service, with next day delivery to most destinations.   Express Mail is delivered 365 days a year, with no extra charge for Saturday,   Sunday, or holiday delivery. Features include merchandise and document   reconstruction, tracking and tracing, delivery to post office boxes and rural   addresses, money-back guarantee, COD, return Customer Online Record service, and   waiver of signature. Insurance is provided at no additional cost up to $500.   Additional merchandise insurance is available up to $5,000. Pickup service is   available for one low fee per stop, regardless of the number of pieces. All   packages must use an Express Mail label. This Web Tool allows you to generate a   USPS Express Mail shipping label complete with addresses, barcode, and Customer   Online Record.</td>
								</tr>
							<% End If %>


							<% If Session("pcAdminToCountryCode") <> "US" Then %>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
								  <th colspan="2"><input type="radio" name="LabelMode" id="LabelMode" value="FCINT" class="clearBorder" checked>
									&nbsp;First Class International</th>
								</tr>
								<tr>
									<td colspan="2"><p>The most economical way to send letters, large envelopes,   small packages, postal cards, printed matter, and small packets 4 pounds and   under, worldwide. Postage is required on these labels. <BR>
									  <br />
									Please refer to the <a href="http://pe.usps.com/text/imm/immc2_016.htm" title="IMM Conditions for Mailing First Class International" target="_blank">International Mail Manual</a> for &quot;Conditions for Mailing&quot; First Class International Mail. You will find all the information regarding size and weight limits.</p></td>
								</tr>
							<% End If %>
							<% If Session("pcAdminToCountryCode") <> "US" Then %>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
								  <th colspan="2"><input type="radio" name="LabelMode" id="LabelMode" value="PMINT" class="clearBorder">
									&nbsp;Priority Mail International</th>
								</tr>
								<tr>
									<td colspan="2"><p>Priority Mail International (PMI) provides customers with a   reliable and economical means of sending correspondence and merchandise up to 70   pounds                                      to over 190 countries and territories worldwide.                                      Postage is required on these labels. <BR>
										<br />
Please refer to the <a href="http://pe.usps.com/text/imm/immc2_011.htm" title="IMM Conditions for Mailing Priority Mail International" target="_blank">International Mail Manual</a> for &quot;Conditions for Mailing&quot; Priority Mail International Mail. You will find all the information regarding size and weight limits.</p></td>
								</tr>
							<% End If %>
							<% If Session("pcAdminToCountryCode") <> "US" Then %>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
								  <th colspan="2"><input type="radio" name="LabelMode" id="LabelMode" value="EMINT" class="clearBorder">
									&nbsp;Express Mail International</th>
								</tr>
								<tr>
									<td colspan="2"><p>Express Mail&reg; is USPS's fastest service for time-sensitive   letters, documents or merchandise. Guaranteed overnight delivery to most   locations or your money back.                                      Postage is required on these labels.<BR>
									  <br />
									Please refer to the <a href="http://pe.usps.com/text/imm/immc2_006.htm" title="IMM Conditions for Mailing Express Mail International" target="_blank">International Mail Manual</a> for &quot;Conditions for Mailing&quot; Express Mail International Mail. You will find all the information regarding size and weight limits.
									</ol>
									</p></td>
								</tr>
							<% End If %>
							<tr>
							  <td colspan="2">&nbsp;</td>
							</tr>
							<tr>
							  <td colspan="2"><input type="submit" name="Submit2" value="Submit" class="ibtnGrey"></td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
						</table>
						<input type="hidden" name="idOrder" value="<%=request("idOrder")%>">
						<input type="hidden" name="count" value="<%=pcv_count%>">
						<%=pcv_strHiddenField%>
					</form>
					<% end if
				else %>
					<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">
						<input type="hidden" name="LabelMode" value="<%=pcv_LabelMode%>">
						<table class="pcCPcontent">
							<tr>
							<%
							dim strJSOnChangeTabCnt, k, intTempJSChangeCnt
							strTabCnt=""
							for k=1 to pcPackageCount
								if k=1 then
									strTabCnt="""tab4"""
								else
									iCnt=3+int(k)
									strTabCnt=strTabCnt&",""tab"&iCnt&""""
								end if
							next

							strJSOnChangeTabCnt=""
							for k=1 to pcPackageCount
								intTempJSChangeCnt=3+int(k)
								strJSOnChangeTabCnt=strJSOnChangeTabCnt&";change('tabs"&intTempJSChangeCnt&"', '')"
							next %>
							<!--#include file="../includes/javascripts/pcFedExLabelTabs.asp"-->
							<td valign="top">
								<div class="menu">
									<ul>
										<li><a id="tabs1" class="current" onclick="change('tabs1', 'current');change('tabs2', '');change('tabs3', '');<%=strJSOnChangeTabCnt%>;showTab('tab1')">Ship Settings</a></li>
										<li><a id="tabs2" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '')<%=strJSOnChangeTabCnt%>;showTab('tab2')">Ship From</a></li>
										<li><a id="tabs3" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', 'current');<%=strJSOnChangeTabCnt%>;showTab('tab3')">Recipient</a></li>
										<% strOnclickTabCnt=""
										if pcPackageCount=1 then %>
										<li><a id="tabs4" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current');showTab('tab4')">Package Information</a></li>
										<% else %>
											<% for k=1 to pcPackageCount
												intTempPackageCnt=3+int(k)
												strOnclickTabCnt=""
												for l=1 to pcPackageCount
													intCPC=3+int(l)
													if intCPC=intTempPackageCnt then
														strOnclickTabCnt=strOnclickTabCnt&";change('tabs"&intCPC&"', 'current')"
													else
														strOnclickTabCnt=strOnclickTabCnt&";change('tabs"&intCPC&"', '')"
													end if
												next
												%>
												<li><a id="tabs<%=3+int(k)%>" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');<%=strOnclickTabCnt%>;showTab('tab<%=intTempPackageCnt%>')">Package <%=k%></a></li>
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
								<% For xArrayCount = LBound(pcLocalArray) TO UBound(pcLocalArray) %>
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
										  <th colspan="2">Package Status</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2">This package will not be automatically flagged as shipped after the label is generated. You will need to manually update the package status after you have placed postage on the package and have  shipped the package. If you wish to flag the package as &quot;Shipped&quot; when the label is generated, you can check the box below. </td>
								  </tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<script>
									<!--
									function HideCommentRow(){
									if(document["form1"]["MarkedAsShipped"].checked){
									document.getElementById("AdmCommentsRow").style.display=''
									}
									else{
									document.getElementById("AdmCommentsRow").style.display='none'
									}
									}
									//-->
									</script>
									<tr>
										<td align="right"><INPUT tabIndex="25" type="checkbox" value="1" name="MarkedAsShipped" class="clearBorder" <%=pcf_CheckOption("MarkedAsShipped", "1")%> onclick = "HideCommentRow()"></td>
										<td><strong>Flag this package as shipped</strong></td>
									</tr>
									<tr>
									  <td colspan="2" class="pcCPspacer"></td>
									</tr>
									<%
									' Look up today's date
									Dim varMonth, varDay, varYear
									varMonth=Month(Date)
									varDay=Day(Date)
									varYear=Year(Date)
									dim dtInputStr
									dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
									if scDateFrmt="DD/MM/YY" then
										dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
									end if
									' Setup default Order Shipped message

									' Get customer information

									query="SELECT name,lastname FROM customers WHERE idcustomer="& pcv_IdCustomer
									Set rs=Server.CreateObject("ADODB.Recordset")
									Set rs=conntemp.execute(query)
									pcv_CustomerName = rs("name")&" "&rs("lastname")

									' Prepare message
									customerShippedEmail=""
									personalmessage=replace(scShippedEmail,"<br>", vbCrlf)
									personalmessage=replace(personalmessage,"<COMPANY>",scCompanyName)
									personalmessage=replace(personalmessage,"<COMPANY_URL>",scStoreURL)
									personalmessage=replace(personalmessage,"<TODAY_DATE>",dtInputStr)
									personalmessage=replace(personalmessage,"<CUSTOMER_NAME>",pcv_CustomerName)
									personalmessage=replace(personalmessage,"<ORDER_ID>",pidorder)
									personalmessage=replace(personalmessage,"<ORDER_DATE>",ShowDateFrmt(pcv_orderDate))
									If scShippedEmail<>"" Then
										customerShippedEmail=customerShippedEmail & vbCrLf & personalmessage & vbCrLf & vbCrLf
									end if
									CustomerShippedEmail=replace(CustomerShippedEmail,"//","/")
									CustomerShippedEmail=replace(CustomerShippedEmail,"http:/","http://")
									CustomerShippedEmail=replace(CustomerShippedEmail,"https:/","https://")
									CustomerShippedEmail=replace(CustomerShippedEmail,"''",chr(39))

									If Session("pcAdminAdmComments")="" Then
										pcv_AdmComments=CustomerShippedEmail
									Else
										pcv_AdmComments=replace(Session("pcAdminAdmComments"),"''","'") '// reverse db apostrophe santization
									End If
									%>

									<tr id="AdmCommentsRow" style="display:none">
										<td valign="top" align="right"><b>Comments:</b></td>
										<td valign="top">
										<textarea name="AdmComments" size="40" rows="10" cols="65"><%=pcv_AdmComments%></textarea>
										<div style="margin: 10px 15px 15px 0;" class="pcCPnotes">Please note that additional text will appear in the message that is emailed to the customer depending on whether this is a partial or final shipment, and depending on which shipping provider was used for the shipment, if any. The additional text can be edited by editing the file &quot;includes/languages_ship.asp". We recommend that you ship a few test orders in different scenarios to become familiar with the way the final message appears.</div>
										<%
										Session("pcAdminAdmComments")=""
										%>
										</td>
									</tr>
									<% if pcv_LabelMode="PMINT" then %>
										<tr>
											<th colspan="2">Priority Mail</th>
										</tr>
										<tr>
											<td colspan="2" class="pcCPspacer"></td>
										</tr>
										<tr>
										<td width="23%" align="right" valign="top"><b>Shipping Container Type:</b></td>
										<td width="77%" align="left">
										<select name="Container" id="Container">
										<option value="RECTANGULAR" <%=pcf_SelectOption("Container","RECTANGULAR")%>>Rectangular</option>
										<option value="NONRECTANGULAR" <%=pcf_SelectOption("Container","NONRECTANGULAR")%>>Non-Rectangular</option>
										<option value="MDFLATRATEBOX" <%=pcf_SelectOption("Container","MDFLATRATEBOX")%>>Medium Flat Rate Box</option>
										<option value="LGFLATRATEBOX" <%=pcf_SelectOption("Container","LGFLATRATEBOX")%>>Large Flat Rate Box</option>
										<option value="SMFLATRATEBOX" <%=pcf_SelectOption("Container","LGFLATRATEBOX")%>>Small Flat Rate Box</option>
										<option value="LGVIDEOBOX" <%=pcf_SelectOption("Container","LGVIDEOBOX")%>>Large Video Box</option>
										<option value="DVDBOX" <%=pcf_SelectOption("Container","DVDBOX")%>>DVD Box</option>
										<option value="FLATRATEENV" <%=pcf_SelectOption("Container","FLATRATEENV")%>>Flat Rate Envelope</option>
										<option value="LEGALFLATRATEENV" <%=pcf_SelectOption("Container","LEGALFLATRATEENV")%>>Legal Flat Rate Envelope</option>
										<option value="PADDEDFLATRATEENV" <%=pcf_SelectOption("Container","PADDEDFLATRATEENV")%>>Padded Flat Rate Envelope</option>
										<option value="WINDOWFLATRATEENV" <%=pcf_SelectOption("Container","WINDOWFLATRATEENV")%>>Window Flat Rate Envelope</option>
										<option value="SMFLATRATEENV" <%=pcf_SelectOption("Container","SMFLATRATEENV")%>>Small Flat Rate Envelope</option>
										<option value="GIFTCARDFLATRATEENV" <%=pcf_SelectOption("Container","GIFTCARDFLATRATEENV")%>>Gift Card Flat Rate Envelope</option>

										</select>
										<%pcs_UPSRequiredImageTag "Container", true%></td>
										</tr>
										<tr>
											<td width="23%" align="right" valign="top"><b>Content Type:</b></td>
											<td width="77%" align="left">
											<select name="ContentType" id="ContentType">
											<option value="MERCHANDISE" <%=pcf_SelectOption("ContentType","MERCHANDISE")%>>Merchandise</option>
											<option value="SAMPLE" <%=pcf_SelectOption("ContentType","SAMPLE")%>>Sample</option>
											<option value="GIFT" <%=pcf_SelectOption("ContentType","GIFT")%>>Gift</option>
											<option value="DOCUMENTS" <%=pcf_SelectOption("ContentType","DOCUMENTS")%>>Documents</option>
											<option value="RETURN" <%=pcf_SelectOption("ContentType","RETURN")%>>Return</option>
											<option value="OTHER" <%=pcf_SelectOption("ContentType","OTHER")%>>Other</option>
											</select>
											<%pcs_UPSRequiredImageTag "ContentType", true%></td>
										</tr>

										<tr>
											<td></td>
											<td>If &quot;Content Type&quot; of &quot;Other&quot; was selected please give a short description (15 characters max)</td>
										</tr>

										<tr>
											<td width="23%" align="right">&nbsp;</td>
											<td width="77%" align="left"><input name="ContentTypeOther3" type="text" value="<%=pcf_FillFormField("ContentTypeOther"&k, false)%>" size="30" maxlength="15"></td>
													</tr>
										<tr>
										  <td colspan="2" class="pcCPspacer"></td>
										</tr>
									<% end if %>
									<% if pcv_LabelMode="EMINT" then %>
										<tr>
											<th colspan="2">Priority Mail</th>
										</tr>
										<tr>
											<td colspan="2" class="pcCPspacer"></td>
										</tr>
										<tr>
										<td width="23%" align="right" valign="top"><b>Shipping Container Type:</b></td>
										<td width="77%" align="left">
										<select name="Container" id="Container">
										<option value="RECTANGULAR" <%=pcf_SelectOption("Container","RECTANGULAR")%>>Rectangular</option>
										<option value="NONRECTANGULAR" <%=pcf_SelectOption("Container","NONRECTANGULAR")%>>Non-Rectangular</option>
										<option value="FLATRATEENV" <%=pcf_SelectOption("Container","FLATRATEENV")%>>Flat Rate Envelope</option>
										<option value="LEGALFLATRATEENV" <%=pcf_SelectOption("Container","LEGALFLATRATEENV")%>>Legal Flat Rate Envelope</option>

										</select>
										<%pcs_UPSRequiredImageTag "Container", true%></td>
										</tr>
										<tr>
											<td width="23%" align="right" valign="top"><b>Content Type:</b></td>
											<td width="77%" align="left">
											<select name="ContentType" id="ContentType">
											<option value="MERCHANDISE" <%=pcf_SelectOption("ContentType","MERCHANDISE")%>>Merchandise</option>
											<option value="SAMPLE" <%=pcf_SelectOption("ContentType","SAMPLE")%>>Sample</option>
											<option value="GIFT" <%=pcf_SelectOption("ContentType","GIFT")%>>Gift</option>
											<option value="DOCUMENTS" <%=pcf_SelectOption("ContentType","DOCUMENTS")%>>Documents</option>
											<option value="RETURN" <%=pcf_SelectOption("ContentType","RETURN")%>>Return</option>
											<option value="OTHER" <%=pcf_SelectOption("ContentType","OTHER")%>>Other</option>
											</select>
											<%pcs_UPSRequiredImageTag "ContentType", true%></td>
										</tr>
										<tr>
											<td></td>
											<td>If &quot;Content Type&quot; of &quot;Other&quot; was selected please give a short description (15 characters max)</td>
										</tr>

										<tr>
											<td width="23%" align="right">&nbsp;</td>
											<td width="77%" align="left"><input name="ContentTypeOther2" type="text" value="<%=pcf_FillFormField("ContentTypeOther"&k, false)%>" size="30" maxlength="15"></td>
							</tr>
										<tr>
										  <td colspan="2" class="pcCPspacer"></td>
										</tr>
									<% end if %>
									<% if pcv_LabelMode="FCINT" then %>
										<tr>
											<th colspan="2">First Class Mail Type</th>
										</tr>
										<tr>
											<td colspan="2" class="pcCPspacer"></td>
										</tr>
										<tr>
										<td width="23%" align="right" valign="top"><b>Mail Type:</b></td>
										<td width="77%" align="left">
										<select name="FirstClassMailtype" id="FirstClassMailtype">
										<option value="PARCEL" <%=pcf_SelectOption("FirstClassMailtype","PARCEL")%>>Parcel</option>
										<option value="FLAT" <%=pcf_SelectOption("FirstClassMailtype","FLAT")%>>Flat</option>
										 <option value="LETTER" <%=pcf_SelectOption("FirstClassMailtype","LETTER")%>>Letter</option>
									   </select>
										<%pcs_UPSRequiredImageTag "FirstClassMailtype", true%></td>
										</tr>
										<tr>
											<td width="23%" align="right" valign="top"><b>Content Type:</b></td>
											<td width="77%" align="left">
											<select name="ContentType" id="ContentType">
											<option value="MERCHANDISE" <%=pcf_SelectOption("ContentType","MERCHANDISE")%>>Merchandise</option>
											<option value="SAMPLE" <%=pcf_SelectOption("ContentType","SAMPLE")%>>Sample</option>
											<option value="GIFT" <%=pcf_SelectOption("ContentType","GIFT")%>>Gift</option>
											<option value="DOCUMENTS" <%=pcf_SelectOption("ContentType","DOCUMENTS")%>>Documents</option>
											<option value="OTHER" <%=pcf_SelectOption("ContentType","OTHER")%>>Other</option>
											</select>
											<%pcs_UPSRequiredImageTag "ContentType", true%></td>
										</tr>
										<tr>
											<td></td>
											<td>If &quot;Content Type&quot; of &quot;Other&quot; was selected please give a short description (15 characters max)</td>
										</tr>

										<tr>
											<td width="23%" align="right">&nbsp;</td>
											<td width="77%" align="left"><input name="ContentTypeOther" type="text" value="<%=pcf_FillFormField("ContentTypeOther"&k, false)%>" size="30" maxlength="15"></td>
							</tr>
										<tr>
										  <td colspan="2" class="pcCPspacer"></td>
										</tr>
									<% End If %>


									<tr>
										<th colspan="2">Ship Date</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
									  <td colspan="2">Date Package Will Be Mailed. Ship date may be up to 3 days in advance. If you are not sure of the exact date of shipment, select &quot;No Date Selected&quot;.</td>
									</tr>
									<tr>
									  <td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
									  <td align="right" valign="top"><b>Ship Date:</b></td>
									  <td><p>
										<select name="LabelDate" id="LabelDate">
											<% pcv_TodayDate=Date() %>
											<option value="" <%=pcf_SelectOption("LabelDate","")%>>No Date Selected</option>
											<option value="<%=pcv_TodayDate%>" <%=pcf_SelectOption("LabelDate", (exFormatDate(pcv_TodayDate,"%mm/%dd/%yyyy")))%>><%=exFormatDate(pcv_TodayDate,"%mm/%dd/%yyyy")%></option>
											<option value="<%=pcv_TodayDate+1%>" <%=pcf_SelectOption("LabelDate",exFormatDate(pcv_TodayDate+1,"%mm/%dd/%yyyy"))%>><%=exFormatDate(pcv_TodayDate+1,"%mm/%dd/%yyyy")%></option>
											<option value="<%=pcv_TodayDate+2%>" <%=pcf_SelectOption("LabelDate",exFormatDate(pcv_TodayDate+2,"%mm/%dd/%yyyy"))%>><%=exFormatDate(pcv_TodayDate+2,"%mm/%dd/%yyyy")%></option>
										</select>
										  <%pcs_UPSRequiredImageTag "LabelDate", false%>
									  </p></td>
									</tr>
									<tr>
									  <td colspan="2" class="pcCPspacer"></td>
									</tr>

										  <th colspan="2">Label Options</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td width="23%" align="right" valign="top"><b>Label Output Format:</b></td>
										<td width="77%" align="left">
										<select name="ImageType" id="ImageType">
										<% if pcv_LabelMode="E" then %>
											<option value="GIF" <%=pcf_SelectOption("ImageType","GIF")%>>GIF</option>
										<% end if %>
										<option value="PDF" <%=pcf_SelectOption("ImageType","PDF")%>>PDF</option>
										<% if pcv_LabelMode<>"E" then %>
											<option value="TIF" <%=pcf_SelectOption("ImageType","TIF")%>>TIF</option>
										<% end if %>
										</select>
										<%pcs_UPSRequiredImageTag "ImageType", true%>			</td>
									</tr>
									<% '//<Option>
									if pcv_LabelMode="E" OR pcv_LabelMode="FCINT" OR pcv_LabelMode="PMINT" OR pcv_LabelMode="EMINT" then %>
										<input type="hidden" name="LabelOption" value="">
									<% else %>
										<tr>
											<td align="right" valign="top"><b>Label Type Option:</b></td>
											<td align="left">
											<select name="LabelOption" id="LabelOption">
											<option value="1" <%=pcf_SelectOption("LabelOption","1")%>>Full Label</option>
											<option value="2" <%=pcf_SelectOption("LabelOption","2")%>>Bar Code Only</option>
											</select>
											<%pcs_UPSRequiredImageTag "LabelOption", true%></td>
										</tr>
									<% end if %>
									<% If pcv_LabelMode="FCINT" OR pcv_LabelMode="PMINT" OR pcv_LabelMode="EMINT" then %>
									<% Else %>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td align="right"><INPUT tabIndex="25" type="checkbox" value="1" name="SeparateReceiptPage" class="clearBorder" <%=pcf_CheckOption("SeparateReceiptPage", "1")%>></td>
										<td><strong>Print label on separate page.</strong> </td>
									</tr>
									<tr>
										<td></td>
										<td>If you would like the Label and the Customer Online Record Printed on 2 pages   instead of one, check the box above. If you want them printed on the same single   page, leave the box unchecked.</td>
									</tr>
									<% End If %>

									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
								</table>
								</div>

								<!--
								/////////////////////////////////////////////////////////////////////////////////
								// SHIPPER
								//////////////////////////////////////////////////////////////////////////////////
								-->
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
										<th colspan="2">Contact Details</th>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td width="25%"><p>Company Name:</p></td>
										<td width="77%">
											<p>
											<input name="FromFirm" type="text" id="FromFirm" value="<%=pcf_FillFormField("FromFirm", true)%>" size="50">
											<%pcs_UPSRequiredImageTag "FromFirm", true%>
											</p>                                </td>
									</tr>
									<% if pcv_LabelMode="E" OR pcv_LabelMode="FCINT" OR pcv_LabelMode="PMINT" OR pcv_LabelMode="EMINT" then %>
										<tr>
											<td><p>First Name:</p></td>
											<td>
											<p>
											<input name="FromFirstName" type="text" id="FromFirstName" value="<%=pcf_FillFormField("FromFirstName", false)%>">
											<%pcs_UPSRequiredImageTag "FromFirstName", true%>
											</p></td>
										</tr>
										<tr>
											<td><p>Last Name:</p></td>
											<td>
											<p>
											<input name="FromLastName" type="text" id="FromLastName" value="<%=pcf_FillFormField("FromLastName", false)%>">
											<%pcs_UPSRequiredImageTag "FromLastName", true%>
											</p></td>
										</tr>

									<% else %>
										<tr>
											<td><p>Attention  Name:</p></td>
											<td>
											<p>
											<input name="FromName" type="text" id="FromName" value="<%=pcf_FillFormField("FromName", false)%>">
											<%pcs_UPSRequiredImageTag "FromName", true%>
											</p></td>
										</tr>
									<% end if %>
									<% if pcv_LabelMode="E" OR pcv_LabelMode="PMINT" OR pcv_LabelMode="EMINT" then
										if len(Session("ErrFromPhone"))>0 then %>
											<tr>
												<td colspan="2">
												<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
												You must enter a valid Phone Number.</td>
											</tr>
										<% end if %>
										<tr>
											<td><p>Phone Number:</p></td>
											<td>
											<p><input name="FromPhone" type="text" id="FromPhone" value="<%=pcf_FillFormField("FromPhone", true)%>">
											<%pcs_UPSRequiredImageTag "FromPhone", true%></p></td>
										</tr>
									<% end if %>
									<% if len(Session("ErrSenderEMail"))>0 then %>
										<tr>
											<td colspan="2">
											<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
											You must enter a valid Email Address.</td>
										</tr>
									<% end if %>
									<tr>
										<td><p>Email Address:</p></td>
										<td>
										<p>
										<input name="SenderEMail" type="text" id="SenderEMail" value="<%=pcf_FillFormField("SenderEMail", true)%>" size="50">
										<%pcs_UPSRequiredImageTag "SenderEMail", false %>
										</p></td>
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
									<tr>
										<td><p>Address Line 1:</p></td>
										<td>
										<p>
										<input name="FromAddress1" type="text" id="FromAddress1" value="<%=pcf_FillFormField("FromAddress1", true)%>" size="50">
										<%pcs_UPSRequiredImageTag "FromAddress1", true%>
										</p></td>
										</tr>
										<tr>
										<td><p>Address Line 2:</p></td>
										<td>
										<p>
										<input name="FromAddress2" type="text" id="FromAddress2" value="<%=pcf_FillFormField("FromAddress2", false)%>" size="50">
										<%pcs_UPSRequiredImageTag "FromAddress2", false%>
										</p></td>
									</tr>
									<tr>
										<td><p>City:</p></td>
										<td>
										<p>
										<input name="FromCity" type="text" id="FromCity" value="<%=pcf_FillFormField("FromCity", true)%>" size="20" maxlength="13">
										<%pcs_UPSRequiredImageTag "FromCity", true%>
										</p></td>
									</tr>
									<% err.clear
									err.number=0
									dim rsStateObj, FromStateOptAry, ToStateOptAry
									FromStateOptAry=""
									ToStateOptAry=""
									call opendb()
									query="SELECT stateCode, stateName FROM states WHERE pcCountryCode='US' ORDER BY stateName;"
									set rsStateObj=server.CreateObject("ADODB.RecordSet")
									set rsStateObj=conntemp.execute(query)
									if err.number<>0 then
										call LogErrorToDatabase()
										set rs=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									if NOT rsStateObj.eof then
										do until rsStateObj.eof
											strTempStateCode=rsStateObj("stateCode")
											strSelectedValue=pcf_SelectOption("FromState",""&strTempStateCode&"")
											FromStateOptAry=FromStateOptAry&"<option value="""&strTempStateCode&""" "&strSelectedValue&">"&rsStateObj("stateName")&"</option>"
											strSelectedValue=pcf_SelectOption("ToState",""&strTempStateCode&"")
											ToStateOptAry=ToStateOptAry&"<option value="""&strTempStateCode&""" "&strSelectedValue&">"&rsStateObj("stateName")&"</option>"
											rsStateObj.moveNext
										loop
									end if

									set rsStateObj=nothing
									call closedb()
									%>
									<tr>
										<td><p>State:</p></td>
										<td><p>
										<select name="FromState" id="FromState">
										<option value="">-Select State-</option>
										<%=FromStateOptAry%>
										</select></p></td>
									</tr>
									<tr>
										<td><p>Postal Code:</p></td>
										<td>
										<p>
										<input name="FromZip5" type="text" id="FromZip5" value="<%=pcf_FillFormField("FromZip5", true)%>" size="5" maxlength="5">
										<%pcs_UPSRequiredImageTag "FromZip5", true %> - <input name="FromZip4" type="text" id="FromZip4" value="<%=pcf_FillFormField("FromZip4", false)%>" size="4" maxlength="4">
										<%pcs_UPSRequiredImageTag "FromZip4", false %>
										</p>
										</td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<% If pcv_LabelMode="FCINT" then %>
									<% Else %>
			  <tr>
										<th colspan="2">Collection Point</th>
									</tr>

									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2">When the ZIP Code of a collection point for a given package is different from the Zip Code posted in the &quot;Location Details&quot; of the &quot;Shipper&quot; or the person mailing the package, this value must be used to convey this difference to the USPS. Enter the Zip Code of the post office or collection box where the item is mailed. May  be different than from zip code.</td>
									</tr>
									<tr>
										<td width="25%"><p>Collection Point Postal Code:</p></td>
										<td width="77%">
										<p>
										<input name="POZipCode" type="text" id="POZipCode" value="<%=pcf_FillFormField("POZipCode", false)%>">
										<%pcs_UPSRequiredImageTag "POZipCode", false %>
										</p></td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
								<% End If %>
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
									<td width="77%">
									<p>
									<input name="ToFirm" type="text" id="ToFirm" value="<%=pcf_FillFormField("ToFirm", true)%>" size="50">
									<%pcs_UPSRequiredImageTag "ToFirm", true%>
									</p></td>
									</tr>
									<% if pcv_LabelMode="E" then %>
										<tr>
											<td><p>First Name:</p></td>
											<td>
											<p>
											<input name="ToFirstName" type="text" id="ToFirstName" value="<%=pcf_FillFormField("ToFirstName", false)%>">
											<%pcs_UPSRequiredImageTag "ToFirstName", true%>
											</p></td>
										</tr>
										<tr>
											<td><p>Last Name:</p></td>
											<td>
											<p>
											<input name="ToLastName" type="text" id="ToLastName" value="<%=pcf_FillFormField("ToLastName", false)%>">
											<%pcs_UPSRequiredImageTag "ToLastName", true%>
											</p></td>
										</tr>
									<% else %>
										<tr>
											<td><p>Attention Name:</p></td>
											<td>
											<p>
											<input name="ToName" type="text" id="ToName" value="<%=pcf_FillFormField("ToName", false)%>">
											<%pcs_UPSRequiredImageTag "ToName", true%>
											</p></td>
										</tr>
									<% end if %>
									<% if pcv_LabelMode="E" then %>
										<% if len(Session("ErrToPhone"))>0 then %>
											<tr>
												<td colspan="2">
												<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
												You must enter a valid Phone Number.</td>
											</tr>
										<% end if %>
										<tr>
											<td><p>Phone Number:</p></td>
											<td>
											<p>
											<input name="ToPhone" type="text" id="ToPhone" value="<%=pcf_FillFormField("ToPhone", true)%>">
											<%pcs_UPSRequiredImageTag "ToPhone", true%>
											</p></td>
										</tr>
									<% end if %>
									<% if len(Session("ErrRecipientEMail"))>0 then %>
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
									<input name="RecipientEMail" type="text" id="RecipientEMail" value="<%=pcf_FillFormField("RecipientEMail", true)%>" size="50">
									<%pcs_UPSRequiredImageTag "RecipientEMail", false %>
									</p></td>
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

									<tr>
									<td><p>Address Line 1:</p></td>
									<td>
									<p>
									<input name="ToAddress1" type="text" id="ToAddress1" value="<%=pcf_FillFormField("ToAddress1", true)%>" size="50">
									<%pcs_UPSRequiredImageTag "ToAddress1", true%>
									</p></td>
									</tr>
									<tr>
									<td><p>Address Line 2:</p></td>
									<td>
									<p>
									<input name="ToAddress2" type="text" id="ToAddress2" value="<%=pcf_FillFormField("ToAddress2", false)%>" size="50">
									<%pcs_UPSRequiredImageTag "ToAddress2", false%>
									</p></td>
									</tr>
									<tr>
									<td><p>City:</p></td>
									<td>
									<p>
									<input name="ToCity" type="text" id="ToCity" value="<%=pcf_FillFormField("ToCity", true)%>">
									<%pcs_UPSRequiredImageTag "ToCity", true%>
									</p></td>
									</tr>

									<% If pcv_LabelMode="FCINT" OR pcv_LabelMode="PMINT" OR pcv_LabelMode="EMINT" Then %>
										<% call opendb()
										'///////////////////////////////////////////////////////////
										'// START: COUNTRY AND STATE/ PROVINCE CONFIG
										'///////////////////////////////////////////////////////////
										'
										' 1) Place this section ABOVE the Country field
										' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
										' 3) Additional Required Info

										'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
										pcv_isStateCodeRequired =  False '// determines if validation is performed (true or false)
										pcv_isProvinceCodeRequired =  False '// determines if validation is performed (true or false)
										pcv_isCountryCodeRequired =  False '// determines if validation is performed (true or false)

										'// #3 Additional Required Info
										pcv_strTargetForm = "form1" '// Name of Form
										pcv_strCountryBox = "ToCountry" '// Name of Country Dropdown
										pcv_strTargetBox = "ToState" '// Name of State Dropdown
										pcv_strProvinceBox =  "ToProvince" '// Name of Province Field

										'// Set local Country to Session
										if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
											Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session("pcAdminToCountryCode")
										end if

										'// Set local State to Session
										if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
											Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session("pcAdminToState")
										end if

										'// Set local Province to Session
										if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
											Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  Session("pcAdminToProvince")
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
										<%
										'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
										pcs_StateProvince
										call closedb()%>
										<tr>
											<td><p>Postal Code:</p></td>
											<td>
											<p>
											<input name="ToPostalCode" type="text" id="ToPostalCode" value="<%=pcf_FillFormField("ToPostalCode", true)%>" size="10" maxlength="9">
											<%pcs_UPSRequiredImageTag "ToPostalCode", true %>
											</p></td>
										</tr>
									<% Else %>
									<tr>
									<td><p>State:</p></td>
									<td><p>
									<select name="ToState" id="ToState">
									<option value="">-Select State-</option>
									<%=ToStateOptAry%>
									</select></p></td>
									</tr>
									<% End If %>
									<% If pcv_LabelMode="D" OR pcv_LabelMode="S" OR pcv_LabelMode="E" Then %>
										<tr>
											<td><p>Postal Code:</p></td>
											<td>
											<p>
											<input name="ToZip5" type="text" id="ToZip5" value="<%=pcf_FillFormField("ToZip5", true)%>" size="5" maxlength="5">
											<%pcs_UPSRequiredImageTag "ToZip5", true %> - <input name="ToZip4" type="text" id="ToZip4" value="<%=pcf_FillFormField("ToZip4", false)%>" size="4" maxlength="4">
											<%pcs_UPSRequiredImageTag "ToZip4", false %>
											</p></td>
										</tr>
									<% End If %>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<% If pcv_LabelMode="FCINT" Then %>
										<tr>
											<td align="right"><INPUT tabIndex="25" type="checkbox" value="False" name="ToPOBoxFlag" class="clearBorder" <%=pcf_CheckOption("ToPOBoxFlag", "False")%>></td>
											<td><strong>Is this a Post Office Box?</strong></td>
										</tr>
									<% End If %>
									<% if pcv_LabelMode="E" then %>
										<tr>
											<td align="right"><INPUT tabIndex="25" type="checkbox" value="False" name="StandardizeAddressFalse" class="clearBorder" <%=pcf_CheckOption("StandardizeAddressFalse", "False")%>></td>
											<td><strong>Do Not Standardize Delivery Address</strong></td>
								  </tr>
											<tr>
												<td></td>
												<td><p>USPS will check the recipient's address for accuracy and standardize the address using the USPS databases. If you do NOT want the address to be standardize, check the box above and the recipient's address will not be checked for accuracy. </p>
												</td>
										</tr>
									<% end if %>

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
								<% for k=1 to pcPackageCount  %>
									<div id="tab<%=3+int(k)%>" class="panes">
										<table class="pcCPcontent">
											<tr>
												<td colspan="2" class="pcCPspacer"></td>
											</tr>
											<% '// If the tab was processed, skip it.
											if pcLocalArray(k-1) <> "shipped" then %>
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
															function ViewPackageSelected<%=k%>(){

															var selectValDom = document.forms['form1'];
															if (selectValDom.ViewPackages<%=k%>.checked == true) {
															document.getElementById('FaxTable<%=k%>').style.display='';
															}else{
															document.getElementById('FaxTable<%=k%>').style.display='none';
															}
															}
															 //-->
														</SCRIPT>
														<%
														if Session("pcAdminViewPackages"&k)="true" then
															pcv_strDisplayStyle="style=""display:visible"""
														else
															pcv_strDisplayStyle="style=""display:none"""
														end if
														%>
														<input onClick="ViewPackageSelected<%=k%>();" name="ViewPackages<%=k%>" id="ViewPackages<%=k%>" type="checkbox" class="clearBorder" value=true <%=pcf_CheckOption("ViewPackages"&k, "true")%>>
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
																	call opendb()
																	query = 		"SELECT ProductsOrdered.pcPackageInfo_ID , products.description, products.idProduct  "
																	query = query & "FROM ProductsOrdered "
																	query = query & "INNER JOIN products "
																	query = query & "ON ProductsOrdered.idProduct = products.idProduct "
																	query = query & "WHERE ProductsOrdered.idProductOrdered=" & pcv_intPackageInfo_ID &" "

																	set rs2=server.CreateObject("ADODB.RecordSet")
																	set rs2=conntemp.execute(query)

																	if err.number<>0 then
																		'// handle admin error
																		response.write err.description
																		response.end
																	end if

																	if NOT rs2.eof then
																		Do until rs2.eof
																			pcv_strProductDescription = rs2("description")
																			%>
																			<li><%=pcv_strProductDescription%></li>
																			<%
																			rs2.movenext
																		Loop
																	end if
																	call closedb()
																Next
																%>                                                                </td>
															</tr>
														</table>                                                    </td>
												</tr>
												<% if pcv_LabelMode="D" OR pcv_LabelMode="S" then %>
												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<tr>
													<th colspan="2">Service Type
													  <%'=k%></th>
												</tr>
												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<tr>
													<td colspan="2">
														<p>
														Mail Service Type:
														<% Session("pcAdminServiceType"&k)=Session("pcAdminUSPSPackageType") %>
														<select name="ServiceType<%=k%>" id="Service<%=k%>">
															<option value="Priority" <%=pcf_SelectOption("ServiceType"&k,"Priority")%>>Priority</option>
															<option value="First Class" <%=pcf_SelectOption("ServiceType"&k,"First Class")%>>First Class</option>
															<option value="Standard Post" <%=pcf_SelectOption("ServiceType"&k,"Standard Post")%>>Standard Post</option>
															<option value="Bound Printed Matter" <%=pcf_SelectOption("ServiceType"&k,"Bound Printed Matter")%>>Bound Printed Matter</option>
															<option value="Media Mail" <%=pcf_SelectOption("ServiceType"&k,"Media Mail")%>>Media Mail</option>
															<option value="Library Mail" <%=pcf_SelectOption("ServiceType"&k,"Library Mail")%>>Library Mail</option>
														</select>
														<%pcs_UPSRequiredImageTag "ServiceType"&k, true%>
														</p>														</td>
												</tr>
												<% end if %>
												<% if pcv_LabelMode="E" then %>
													<input type="hidden" name="ServiceType<%=k%>" id="Service<%=k%>" value="Express Mail">
												<% end if %>
												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<tr>
													<th colspan="2">Package Weight</th>
												</tr>
												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<% if pcv_LabelMode="E" then %>
													<tr>
													  <td colspan="2"> <p><INPUT tabIndex="25" type="checkbox" value="TRUE" name="FlatRate<%=k%>" class="clearBorder" <%=pcf_CheckOption("FlatRate"&k, "TRUE")%>><strong>&nbsp;Flat Rate Request</strong><br>
														Using an USPS Express Mail Flat Rate envelope should select this option. This option allows mailers to place as much material, regardless of weight, into the envelope providing it can be properly sealed.</p></td>
													</tr>
													<tr>
														<td colspan="2" class="pcCPspacer"></td>
													</tr>
												<% end if %>
												<tr>
													<td colspan="2"><p><strong>Package Weight:</strong><br>
														<% if pcv_LabelMode="E" then %>
															When &quot;Flat Rate&quot; is not selected, weight is required.
														<% end if %>
														Enter the weight of the package in ounces. If there is more than one package in the shipment, enter the weight of the first package or the total shipment weight.</p>
														<p style="padding-top: 5px;">&nbsp;</p>
														<%
														intShipWeightPounds=Int(pcv_ShipWeight/16) 'intPounds used for USPS
														intShipWeightOunces=pcv_ShipWeight-(intShipWeightPounds*16) 'intUniversalOunces used for USPS

														if pcPackageCount=1 then
															'Get weight
															Session("pcAdminPounds"&k) = intShipWeightPounds
															Session("pcAdminOunces"&k) = intShipWeightOunces
														end if %>
														<p style="padding-top: 5px;">Weight: <input name="Pounds<%=k%>" type="text" id="Pounds<%=k%>" value="<%=pcf_FillFormField("Pounds"&k, true)%>" size="4">
														lbs.
														  <%pcs_UPSRequiredImageTag "Pounds"&k, true%>&nbsp;<input name="Ounces<%=k%>" type="text" id="Ounces<%=k%>" value="<%=pcf_FillFormField("Ounces"&k, true)%>" size="4">
														  ozs.
														  <%pcs_UPSRequiredImageTag "Ounces"&k, true%>
														</p>                                                    </td>
												</tr>
												<% If pcv_LabelMode = "FCINT" OR pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" Then %>


												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<tr>
													<th colspan="2">Package Size</th>
												</tr>
												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<% if pcv_LabelMode = "PMINT" OR pcv_LabelMode = "EMINT" Then %>
												<% else %>
												<tr>
													<td width="24%" align="right" valign="top"><b>Container:</b></td>
													<td width="76%" align="left">
													<select name="Container" id="Container">
													<option value="RECTANGULAR" <%=pcf_SelectOption("Container","RECTANGULAR")%>>Rectangular</option>
													<option value="NONRECTANGULAR" <%=pcf_SelectOption("Container","NONRECTANGULAR")%>>Non-Rectangular</option>
													</select>
													<%pcs_UPSRequiredImageTag "Container", true%></td>
												</tr>
												<% end if %>

												<tr>
													<td colspan="2"><p><strong>Deminsions for &quot;Large&quot; package size:</strong> <br>
													  If any one side of your package is over 12" your package is considered &quot;Large&quot; and you must supply the measurments for each side of your package. The unit must be inches and any fraction of an inch should be put in decimal format.
													  </p>
														<p style="padding-top: 5px;">
														Length: <input name="Length<%=k%>" type="text" id="Length<%=k%>" value="<%=pcf_FillFormField("Length"&k, true)%>" size="10">
														<%pcs_UPSRequiredImageTag "Length"&k, true%>
														Width: <input name="Width<%=k%>" type="text" id="Width<%=k%>" value="<%=pcf_FillFormField("Width"&k, true)%>" size="10">
														<%pcs_UPSRequiredImageTag "Width"&k, true%>
														Height: <input name="Height<%=k%>" type="text" id="Height<%=k%>" value="<%=pcf_FillFormField("Height"&k, true)%>" size="10">
														<%pcs_UPSRequiredImageTag "Height"&k, true%>
													  </p>                                                    </td>
												</tr>
												<tr>
													<td colspan="2"><p><strong><br>
													Girth is required when the package is considered &quot;Large&quot; and also a &quot;Non-Rectangular&quot; package size:</strong> Measurments must be calculated in inches. Fractions of an inch should be posted in decimal format. Girth = 2(W+H)<br>
													<br>
Girth:
<input name="Girth<%=k%>" type="text" id="Girth<%=k%>" value="<%=pcf_FillFormField("Girth"&k, true)%>" size="10">
													  <%pcs_UPSRequiredImageTag "Girth"&k, false%>

													  </p></td>
												</tr>
												<% End IF %>
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
													  <td colspan="2">If you need to cross-reference information about a shipment using your own tracking or inventory systems, you can use the &quot;Customer Reference Number&quot; for this. </td>
												</tr>
												<tr>
													<td width="24%" align="right">Customer Reference Number: </td>
													<td width="76%" align="left"><input name="CustomerRefNo<%=k%>" type="text" value="<%=pcf_FillFormField("CustomerRefNo"&k, false)%>" size="15"></td>
												</tr>
												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<% If pcv_LabelMode="FCINT" OR pcv_LabelMode="PMINT" OR pcv_LabelMode = "EMINT" Then %>
													<tr>
														<th colspan="2">Package Contents</th>
													</tr>
													<tr>
														<td colspan="2" class="pcCPspacer"></td>
													</tr>
													<tr>
														  <td colspan="2">Enter a brief description of your item(s) within the package. Non-descriptive wording such as 'Gift' will result in an error.</td>
													</tr>
													<tr>
														<td width="24%" align="right">Description: </td>
														<td width="76%" align="left"><input name="Description<%=k%>" type="text" value="<%=pcf_FillFormField("Description"&k, false)%>" size="56"></td>
													</tr>
													<tr>
														<td colspan="2">Enter the value of the contents within this package.</td>
													</tr>
													<tr>
														<td width="24%" align="right">Value: </td>
														<td width="76%" align="left"><input name="Value<%=k%>" type="text" value="<%=pcf_FillFormField("Value"&k, false)%>" size="56"></td>
													</tr>
												<% End If %>
												<% if pcv_LabelMode="PMINT" OR pcv_LabelMode = "EMINT" then %>
													<tr>
														<td colspan="2">Enter the insured amount of this package.</td>
													</tr>
													<tr>
														<td width="24%" align="right">Insured Amount: </td>
														<td width="76%" align="left"><input name="InsuredAmount<%=k%>" type="text" value="<%=pcf_FillFormField("InsuredAmount"&k, false)%>" size="56"></td>
													</tr>

												<% end if %>

												<% if pcv_LabelMode="E" then %>
													<tr>
														<td colspan="2" class="pcCPspacer"></td>
													</tr>
													<tr>
														<th colspan="2">Miscellaneous Settings</th>
													</tr>
													<tr>
														<td colspan="2" class="pcCPspacer"></td>
													</tr>

													<tr>
														<td align="right"><INPUT tabIndex="25" type="checkbox" value="TRUE" name="WaiverOfSignature<%=k%>" class="clearBorder" <%=pcf_CheckOption("WaiverOfSignature"&k, "TRUE")%>></td>
														<td><strong>No Signature Required for Delivery</strong></td>
													</tr>
													<tr>
														<td></td>
														<td>This feature allows the user to waive the signature requirement and authorize the delivery employee signature as proof of delivery. This option is not available if additional insurance is requested.</td>
													</tr>

													<tr>
													  <td align="right">&nbsp;</td>
													  <td>&nbsp;</td>
													</tr>
													<tr>
														<td align="right"><INPUT tabIndex="25" type="checkbox" value="TRUE" name="NoHoliday<%=k%>" class="clearBorder" <%=pcf_CheckOption("NoHoliday"&k, "TRUE")%>></td>
														<td><strong>No Holiday Delivery</strong></td>
													</tr>

													<tr>
														<td align="right"><INPUT tabIndex="25" type="checkbox" value="TRUE" name="NoWeekend<%=k%>" class="clearBorder" <%=pcf_CheckOption("NoWeekend"&k, "TRUE")%>></td>
														<td><strong>No Weekend Delivery</strong></td>
												   </tr>
													<tr>
														<td></td>
														<td>The above selections of &quot;No Holiday Delivery&quot; and &quot;No Weekend Delivery&quot; allows the user to defer
														  delivery on a weekend and/or holiday.
														  This may be selected if the sender
														  knows the recipient is not available on
														  the weekend or holiday (e.g., business
														addresses).</td>
													</tr>

													<tr>
													  <td align="right">&nbsp;</td>
													  <td align="left">&nbsp;</td>
													</tr>
													<tr>
													  <td colspan="2" align="left"><p><strong>Note on Additional Special Services</strong><br>
														Return Receipt, Collect-On-Delivery (COD), and additional Insurance are available for an additional fee.<br>
													  To request any of these additional services, you must bring your item to a post office for acceptance.</p></td>
													</tr>
													<tr>
													  <td align="right">&nbsp;</td>
													  <td align="left">&nbsp;</td>
													</tr>
													<tr>
														<td colspan="2" class="pcCPspacer"></td>
													</tr>
													<tr>
														<th colspan="2">Terms and Conditions</th>
													</tr>
													<tr>
														<td colspan="2" class="pcCPspacer"></td>
													</tr>
													<tr>
													  <td colspan="2" align="left"><p><strong>Service Guarantee:
</strong>Online Express Mail&reg; labels are restricted to Post Office-to-Addressee domestic shipments to
														U.S. destinations, including Alaska, Hawaii, Puerto Rico, and the U.S. Virgin Islands. They cannot be used for
														international shipments, shipments to APO/FPOs, or to the remaining U.S. Territories, Possessions, and Freely
														Associated States.</p>
														<p>&nbsp;</p>
														<p> If the shipment is mailed at a designated USPS&reg; Express Mail facility on or before the specified deposit time for
														overnight delivery to the addressee, delivery to the addressee or agent will be attempted before the guaranteed time
														the next delivery day. Signature of the addressee, addressee's agent, or delivery employee is required upon delivery,
														unless this requirement is expressly waived by the mailer. If a delivery attempt is not made by the guaranteed time
														and the mailer files a claim for a refund, the USPS will refund the postage, unless: 1) delivery was attempted but
														could not be made, or the article was available for pickup at the destination, 2) the shipment was delayed by strike or work stoppage, 3) detention was properly made for a law enforcement purpose, or 4) delay in delivery was due to certain other factors specified in the Domestic Mail Manual.</p>
														<p>&nbsp;</p>
														<p> A notice is left for the addressee when an item cannot be delivered on a first attempt. If the item cannot be delivered on the second attempt and is not claimed by the addressee within five days of the second attempt, it will be returned to sender at no additional postage.</p>
														<p>&nbsp;</p>
														<p> Please consult your local Express Mail directory for noon and 3:00 p.m. delivery areas. See the Domestic Mail Manual for details.</p>
														<p>&nbsp;</p>
														<p>                                                          <strong> Insurance Coverage:</strong> Insurance is provided only in accordance with postal regulations in the Domestic Mail Manual
														(DMM&reg;). The DMM sets forth the specific types of losses that are covered, the limitations on coverage, terms of
														insurance, conditions of payment, and adjudication procedures. Copies of the DMM are available for inspection at
														any Post Office&trade;. If copies are not available and information on Express Mail insurance is requested, please contact your Postmaster prior to mailing. The DMM consists of federal regulations, and USPS personnel are NOT authorized to change or waive these regulations or grant exceptions. Limitations prescribed in the DMM provide, in part, that:<br>
														</p>
														<ul>
														<li> The contents of Express Mail shipments defined by postal regulations as merchandise are insured against
														loss, damage, or rifling. Coverage up to $100 per shipment is included at no additional charge. Additional
														merchandise insurance up to $5,000 per shipment may be purchased for an additional fee; however,
														additional insurance is void if waiver of the addressee's signature is requested.</li>
														<li> Coverage extends to the actual value of the contents at the time of mailing or the cost of repairs, not to
														exceed the limit fixed for the insurance coverage obtained.</li>
														<li> Items defined by postal regulations as &quot;negotiable items&quot; (items that can be converted to cash without resort
														to forgery), currency, or bullion are insured up to a maximum of $15 per shipment.</li>
														<li> Items defined by postal indemnity regulations as nonnegotiable documents are insured against loss,
														damage, or rifling up to $100 per shipment for document reconstruction, subject to additional limitations for
														multiple pieces lost or damaged in a single catastrophic occurrence. Document reconstruction insurance
														provides reimbursement for the reasonable costs incurred in reconstructing duplicates of nonnegotiable
														documents mailed. Document reconstruction insurance coverage above $100 per shipment is NOT
														available, and attempts to purchase additional document insurance are void.</li>
														<li> No coverage is provided for consequential losses due to loss, damage, or delay of Express Mail, or for
														concealed damage, spoilage of perishable items, and articles improperly packaged or too fragile to
														withstand normal handling in the mail.</li>
														</ul>
														<p><strong>COVERAGE, TERMS, AND LIMITATIONS ARE SUBJECT TO CHANGE. </strong>Please consult Domestic Mail Manual for
														additional limitations and terms of coverage.</p>
														<p>&nbsp;</p>
														<p>                                                          <strong>Claims: </strong>Online Customer Receipt of the Express Mail label must be presented when filing an indemnity claim and/or<br>
														  for a postage refund.</p>
														<p>&nbsp;</p>
														<p> All claims for delay, loss, damage, or rifling must be made within 90 days of the date of mailing.
														Claim forms may be obtained and filed at any Post Office.</p>
														<p>&nbsp;</p>
														<p> To file a claim for damage, the article, container, and packaging must be presented to the USPS for inspection. To file
														  a claim for loss of contents, the container and packaging must be presented to the USPS for inspection. PLEASE DO
														NOT REMAIL.</p>
														<p><br>
														  <strong>THANK YOU FOR CHOOSING EXPRESS MAIL SERVICE.</strong></p>
														<p>&nbsp;</p>
														<p>
														  <INPUT tabIndex="25" type="checkbox" value="1" name="TandC" class="clearBorder" <%=pcf_CheckOption("TandC", "1")%>>
														<strong>&nbsp;I agree to these Terms and Conditions</strong></p>
													  </p></td>
													</tr>
													<tr>
														<td align="right">&nbsp;</td>
														<td align="left">&nbsp;</td>
													</tr>
												<% end if %>
											<% else %>
												<tr>
													<th colspan="2">This package has been shipped.</th>
												</tr>
											<% end if %>
										</table>
									</div>
								<% next %>
								<br />
								<br />
								<%
								pcv_strPreviousPage = "sds_ShipOrderWizard1.asp?idorder="&pcv_intOrderID&"&PageAction=USPS"
								pcv_strAddPackagePage = "sds_ShipOrderWizard1.asp?idorder="&pcv_intOrderID&"&PageAction=USPS&PackageCount="&pcPackageCount&"&ItemsList="&pcv_strItemsList
								%>
								<p>
							  <div align="center">
										<input type="button" name="Button" value="Start Over" onclick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
									   &nbsp;                                       <input type="submit" name="submit" value="Process Shipment" class="ibtnGrey">
										<br />
										<br />
										<input type="button" name="Button" value="Go Back To Order Details" onclick="document.location.href='<%=pcv_strPreviousPage%>'" class="ibtnGrey">
							  </div>
								</p>
							</td>
						</tr>
						<tr>
							<td valign="top"><div align="center">
							</div></td>
						</tr>
						<!--End -->
					</table>
				</form>
				<% end if
			end if
			'*******************************************************************************
			' END: LOAD HTML FORM
			'*******************************************************************************
			%>
		</td>
	</tr>
</table>
<%call closedb()

'// DESTROY THE USPS OBJECT
set objUSPSClass = nothing
%>
<!--#include file="AdminFooter.asp"-->