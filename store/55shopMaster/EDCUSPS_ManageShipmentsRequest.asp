<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="USPS Shipping - Endicia Postage Services Wizard" %>
<% response.Buffer=true %>
<% section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/pcUSPSClass.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Global.asp"-->
<!--#include file="../includes/GoogleCheckout_APIFunctions.asp"-->
<!--#include file="../pc/pcPay_GoogleCheckout_Handler.asp"-->
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<!--#include file="../includes/pcShipTestModes.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/EndiciaFunctions.asp"-->
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

Function ReplaceN(TestStr)
if (TestStr<>"") OR (Not IsNull(TestStr)) then
	ReplaceN=replace(TestStr,"'","''")
else
	ReplaceN=TestStr
end if
End Function

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
pcPageName="EDCUSPS_ManageShipmentsRequest.asp"
ErrPageName="EDCUSPS_ManageShipmentsRequest.asp"

'// ACTION
pcv_strAction = request("Action")

dim conntemp, query, rs
'// Retrieve current orderstatus of this order

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
query="SELECT orders.idCustomer, orders.total, orders.shipmentDetails, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.shippingCompany, orders.shippingAddress2, orders.pcOrd_shippingPhone, orders.ShippingFullName, orders.pcOrd_ShippingEmail, orders.ordShipType, orders.pcOrd_ShipWeight, orders.Address, orders.Address2, orders.city, orders.state, orders.stateCode, orders.zip, orders.CountryCode,customers.[name],customers.[lastname] FROM customers INNER JOIN orders ON customers.idcustomer=orders.idcustomer WHERE orders.idOrder=" & pcv_intOrderID &" "

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then		
	Dim pidorder, pidcustomer, pcv_ShippingAddress, pcv_ShippingCity, pshippingStateCode, pshippingState, pshippingZip, pshippingPhone, pcv_shippingCountryCode, pshippingCompany, pcv_ShippingAddress2, pShippingEmail,pcv_OrderTotal
	
	'// ORDER INFO
	pidorder=scpre+int(pcv_intOrderID)
	pcv_IdCustomer=rs("idcustomer")
	pcv_OrderTotal = rs("total")
	'//Find default USPS Ship Service
	pcv_shipmentDetails=rs("shipmentDetails")
	tmpEDCShipMethod=""
	tmpEDCPackType=""
	if pcv_shipmentDetails<>"" then
		on error resume next
		tmpArr=split(pcv_shipmentDetails,",")
		if UCase(tmpArr(0))="USPS" then
			if tmpArr(5)<>"" then
				Select Case tmpArr(5)
				Case "9901": tmpEDCShipMethod="Priority"
				Case "9902": tmpEDCShipMethod="Express"
				Case "9903": tmpEDCShipMethod="ParcelPost"
				Case "9904": tmpEDCShipMethod="First"
				Case "9905": tmpEDCShipMethod="ExpressMailInternational"
				Case "9906": tmpEDCShipMethod="ExpressMailInternational"
				Case "9907": tmpEDCShipMethod="PriorityMailInternational"
				Case "9908": 
							tmpEDCShipMethod="PriorityMailInternational"
							tmpEDCPackType="FlatRateEnvelope"
				Case "9909":
							tmpEDCShipMethod="PriorityMailInternational"
							tmpEDCPackType="SmallFlatRateBox"
				Case "9910": tmpEDCShipMethod="ExpressMailInternational"
				Case "9911":
							tmpEDCShipMethod="ExpressMailInternational"
							tmpEDCPackType="FlatRateEnvelope"
				Case "9912": tmpEDCShipMethod="FirstClassMailInternational"
				Case "9913": tmpEDCShipMethod="ParcelPost"
				Case "9914": tmpEDCShipMethod="ExpressMailInternational"
				Case "9915": tmpEDCShipMethod="StandardMail"
				Case "9916": tmpEDCShipMethod="MediaMail"
				Case "9917": tmpEDCShipMethod="LibraryMail"
				End Select
			end if
		end if
		err.clear
		on error goto 0
	end if
	
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
	intShipWeightPounds=Int(pcv_ShipWeight/16) 'intPounds used for USPS
	intShipWeightOunces=pcv_ShipWeight-(intShipWeightPounds*16) 'intUniversalOunces used for USPS

	pAddress=rs("Address")
	pAddress2=rs("address2")
	pcity=rs("city")
	pstate=rs("state")
	pstateCode=rs("stateCode")
	pzip=rs("zip")
	pCountryCode=rs("CountryCode")
	pcv_FName=rs("name")
	pcv_LName=rs("lastName")
	
	if pcv_ShippingAddress="" then
		pcv_ShippingAddress=pAddress
		pcv_ShippingCity=pcity
		pcv_ShippingStateCode=pstateCode
		pcv_ShippingState=pstate
		pcv_ShippingZip=pzip
		pcv_ShippingCountryCode=pCountryCode	
		pcv_ShippingAddress2=pAddress2
	end if
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

If intResetSessions=1 Then
	Session("pcAdminAdmComments")=""
End If

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

'// FROM ADDRESS	
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
	Session("pcAdminFromZip4") = ""
end if
if Session("pcAdminFromCountryCode") = "" OR intResetSessions=1 then
	pcv_strFromCountryCode = "US"
	Session("pcAdminFromCountryCode") = pcv_strFromCountryCode
end if

if Session("pcAdmincustomerRefNo1") = "" OR intResetSessions=1 then
	Session("pcAdmincustomerRefNo1") = pcv_IdCustomer
end if
k=1
if Session("pcAdminPounds"&k)="" OR intResetSessions=1 then
	'Get weight
	Session("pcAdminPounds"&k) = intShipWeightPounds
	Session("pcAdminOunces"&k) = intShipWeightOunces
end if

strFromState=Session("pcAdminFromState")
strFromCity=Session("pcAdminFromCity")
strFromZip5=Session("pcAdminFromZip5")
strFromZip4=Session("pcAdminFromZip4")
strShipFromCountry=Session("pcAdminFromCountryCode")

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
	pcv_strToName = trim(pcv_ShippingFullName)
	if pcv_strToName="" then
		pcv_strToName=pcv_FName & " " & pcv_LName
	end if
	Session("pcAdminToName") = pcv_strToName
end if

if instr(pcv_ShippingFullName, " ") then
	pcv_ShippingNameArry=split(pcv_ShippingFullName, " ")
	pcv_strToFirstName=pcv_ShippingNameArry(0)
	pcv_strToLastName=pcv_ShippingNameArry(1)
else
	pcv_strToFirstName=pcv_FName
	pcv_strToLastName=pcv_LName
end if

if Session("pcAdminToFirstName") = "" OR intResetSessions=1 then
	if pcv_strToFirstName="" then
		pcv_strToFirstName = pcv_FName
	end if
	Session("pcAdminToFirstName") = pcv_strToFirstName
end if

if Session("pcAdminToLastName") = "" OR intResetSessions=1 then
	if pcv_strToLastName="" then
		pcv_strToLastName = pcv_LName
	end if
	Session("pcAdminToLastName") = pcv_strToLastName
end if

if Session("pcAdminToFirm") = "" OR intResetSessions=1 then
	pcv_strToFirm = pcv_ShippingCompany
	Session("pcAdminToFirm") = pcv_strToFirm
end if

if Session("pcAdminToPhone") = "" OR intResetSessions=1 then
	pcv_strToPhone = pcv_ShippingPhone
	Session("pcAdminToPhone") = pcv_strToPhone
end if

if Session("pcAdminRecipientEMail") = "" OR intResetSessions=1 then
	pcv_strRecipientEMail = pcv_ShippingEmail
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

if Session("pcAdminToStateCode") = "" OR intResetSessions=1 then
	pcv_strToStateCode = pcv_ShippingStateCode
	Session("pcAdminToStateCode") = pcv_ShippingStateCode
end if
if Session("pcAdminToState") = "" OR intResetSessions=1 then
	pcv_strToState = pcv_ShippingState
	Session("pcAdminToState") = pcv_ShippingState
end if
if Session("pcAdminToZip5") = "" OR intResetSessions=1 then
	if pcv_ShippingZip<>"" then
		pcv_strToZip5 = Left(pcv_ShippingZip,9)
	else
		pcv_strToZip5 = ""
	end if
	Session("pcAdminToZip5") = pcv_strToZip5
end if
if Session("pcAdminToZip4") = "" OR intResetSessions=1 then
	pcv_strToZip4=""
	if pcv_ShippingZip<>"" then
		if Instr(pcv_ShippingZip,"-")>0 then
			pcv_ShippingZip=replace(pcv_ShippingZip,"-","")
		end if
		if Instr(pcv_ShippingZip," ")>0 then
			pcv_ShippingZip=replace(pcv_ShippingZip," ","")
		end if
		if Len(pcv_ShippingZip)>=9 then
			pcv_strToZip4 = Right(pcv_ShippingZip,4)
		end if
	end if
	Session("pcAdminToZip4") = pcv_strToZip4
end if
if Session("pcAdminToCountryCode") = "" OR intResetSessions=1 then
	pcv_strToCountryCode = pcv_ShippingCountryCode
	Session("pcAdminToCountryCode") = pcv_ShippingCountryCode
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Order #<b><%=(scpre+int(pcv_intOrderID))%></b> - Print USPS Label via Endicia Postage Label Services</th>
	</tr>
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
			if (request.form("submit")<>"") OR (request("endicia")="buypostage") then
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' ServerSide Validate the Required Fields and Formatting.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				
			IF (request.form("submit")<>"") THEN
				'// Generic error for page
				pcv_strGenericPageError = "At least one required field was empty."
				
				'// Clear error string
				pcv_strSecondaryErrors = ""
				pcv_strErrorMsg = ""
				
				session("pcEDCMailClass")=request("edcMailClass")
				if (session("pcEDCMailClass")="") then
					session("pcEDCMailClass")="Express"
				end if
				If session("pcEDCMailClass") = "FirstEnvelope" Or session("pcEDCMailClass") = "FirstLetter" then
					'session("pcEDCLabelType") = "DestinationConfirm" - Currently producing "invalid mail type" error
					session("pcEDCLabelType")="Default"
				Else
					session("pcEDCLabelType")="Default"
				End If				

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
					pcs_ValidateTextField	"Pounds"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"Ounces"&pcv_xCounter, false, 0
					pcs_ValidateTextField	"CustomerRefNo"&pcv_xCounter, false, 30
					if Session("pcAdminPounds"&pcv_xCounter)="" then Session("pcAdminPounds"&pcv_xCounter) = 0
					if Session("pcAdminOunces"&pcv_xCounter)="" then Session("pcAdminOunces"&pcv_xCounter) = 0
					intPounds=Session("pcAdminPounds"&pcv_xCounter)
					intOunces=Session("pcAdminOunces"&pcv_xCounter)
					intWeightInOunces=(intPounds*16)+intOunces
					Session("pcAdminWeightInOunces"&pcv_xCounter)=intWeightInOunces
					
					session("pcEDCPakValue"&pcv_xCounter)=request("edcPakValue"&pcv_xCounter)
					if session("pcEDCPakValue"&pcv_xCounter)="" then
						session("pcEDCPakValue"&pcv_xCounter)="0"
					end if
					session("pcEDCPakDesc"&pcv_xCounter)=request("edcPakDesc"&pcv_xCounter)
					if session("pcEDCPakDesc"&pcv_xCounter)="" then
						session("pcEDCPakDesc"&pcv_xCounter)="Package for the Order #" & (scpre+int(pcv_intOrderID))
					end if
					session("pcEDCShape"&pcv_xCounter)=request("edcShape"&pcv_xCounter)
					if (session("pcEDCShape"&pcv_xCounter)="") then
						session("pcEDCShape"&pcv_xCounter)="Parcel"
					end if
					session("pcEDCLength"&pcv_xCounter)=request("edcLength"&pcv_xCounter)
					if session("pcEDCLength"&pcv_xCounter)="" then
						session("pcEDCLength"&pcv_xCounter)="0"
					end if
					session("pcEDCWidth"&pcv_xCounter)=request("edcWidth"&pcv_xCounter)
					if session("pcEDCWidth"&pcv_xCounter)="" then
						session("pcEDCWidth"&pcv_xCounter)="0"
					end if
					session("pcEDCHeight"&pcv_xCounter)=request("edcHeight"&pcv_xCounter)
					if session("pcEDCHeight"&pcv_xCounter)="" then
						session("pcEDCHeight"&pcv_xCounter)="0"
					end if
					session("pcEDCIM"&pcv_xCounter)=request("edcIM"&pcv_xCounter)
					if (session("pcEDCIM"&pcv_xCounter)="") then
						session("pcEDCIM"&pcv_xCounter)="OFF"
					end if
					session("pcEDCInValue"&pcv_xCounter)=request("edcInValue"&pcv_xCounter)
					if session("pcEDCInValue"&pcv_xCounter)="" then
						session("pcEDCInValue"&pcv_xCounter)="0"
					end if
					'//Calculate Endicia Insurance Total
					pcTmpInValue = Ccur(session("pcEDCInValue"&pcv_xCounter))
					If Session("pcAdminFromCountryCode") = "US" then
						pcTmpInValue = ((pcTmpInValue + (100-1)) \ 100 )*0.75 'US
					Else
						pcTmpInValue = ((pcTmpInValue + (100-1)) \ 100 )*1.35 'US
					End If
					session("pcEDCExInsValue"&pcv_xCounter)=pcTmpInValue
					session("pcEDCSC"&pcv_xCounter)=request("edcSC"&pcv_xCounter)
					if (session("pcEDCSC"&pcv_xCounter)="") then
						session("pcEDCSC"&pcv_xCounter)="OFF"
					end if
					if session("pcEDCLabelType")<>"Default" OR session("pcEDCMailClass")="Express" then
						session("pcEDCSC"&pcv_xCounter)=""
					end if
					session("pcEDCNWD"&pcv_xCounter)=request("edcNWD"&pcv_xCounter)
					if (session("pcEDCNWD"&pcv_xCounter)<>"TRUE") then
						session("pcEDCNWD"&pcv_xCounter)=""
					end if
					session("pcEDCSHD"&pcv_xCounter)=request("edcSHD"&pcv_xCounter)
					if (session("pcEDCSHD"&pcv_xCounter)<>"TRUE") then
						session("pcEDCSHD"&pcv_xCounter)=""
					end if
					
				Next

				'..VALIDATE ALL OTHER FIELDS				
				pcs_ValidateTextField	"idOrder", false, 0	
				pcs_ValidateTextField	"packagecount", false, 0	
				pcs_ValidateTextField	"itemsList", false, 0	
				pcs_ValidateTextField	"LabelDate", false, 0		'<LabelDate>
				if session("pcAdminLabelDate")<>"" then
					session("pcAdminLabelDate")=exFormatDate(session("pcAdminLabelDate"),"%mm/%dd/%yyyy")
				end if
				pcs_ValidateTextField	"FromFirm", isFromFirmRequired, 0		'<FromFirm>	
				pcs_ValidatePhoneNumber	"FromPhone", false, 0
				if session("pcAdminFromFirm")="" then
					session("pcAdminFromFirm")=session("pcAdminFromFirstName")&" "&session("pcAdminFromLastName")
				end if	
				pcs_ValidateTextField	"FromName", true, 0					'<FromName>
				pcs_ValidateEmailField	"SenderEMail", false, 0	
				pcs_ValidateTextField	"FromAddress1", true, 26					'<FromAddress2>	
				pcs_ValidateTextField	"FromAddress2", false, 26				'<FromAddress1>
				pcs_ValidateTextField	"FromCity", true, 13						'<FromCity>
				pcs_ValidateTextField	"FromState", true, 2					'<FromState>
				pcs_ValidateTextField	"FromZip5", true, 5						'<FromZip5>
				pcs_ValidateTextField	"FromZip4", false, 4					'<FromZip4>
				pcs_ValidateTextField	"ToFirm", false, 26						'<ToFirm>
				pcs_ValidatePhoneNumber	"ToPhone", false, 0					
				pcs_ValidateTextField	"ToName", true, 0					'<ToName>
				pcs_ValidateEmailField	"RecipientEMail", false, 0				'<RecipientEMail>
				pcs_ValidateTextField	"ToAddress1", true, 0					'<ToAddress2>
				pcs_ValidateTextField	"ToAddress2", false, 0					'<ToAddress1>
				pcs_ValidateTextField	"ToCity", true, 24						'<ToCity>
				pcs_ValidateTextField	"ToStateCode", false, 10						'<ToState>
				pcs_ValidateTextField	"ToState", false, 0
				pcs_ValidateTextField	"ToCountryCode", true, 0
				pcs_ValidateTextField	"ToZip5", true, 9						'<ToZip5>
				pcs_ValidateTextField	"ToZip4", false, 4						'<ToZip4>
				session("pcAdminMarkedAsShipped")=request("MarkedAsShipped")
				if session("pcAdminMarkedAsShipped")="" then
					session("pcAdminMarkedAsShipped")="0"
				end if
				if request("MarkedAsShipped")="1" then
					pcs_ValidateTextField "AdmComments", false, 0 '// Admin Comments
				end if
				
				If session("pcEDCLabelType") = "DestinationConfirm" Then
					session("pcEDCLabelSize")= "7x4"
				Else
					session("pcEDCLabelSize")= "4x6"
				End If
				
				session("pcEDCLabelFormat")=request("edcLabelFormat")
				if (session("pcEDCLabelFormat")="") then
					session("pcEDCLabelFormat")="GIF"
				end if
				
			END IF 'Get POSTBACK	
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Check for Validation Errors. Do not proceed if there are errors.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				If pcv_intErr>0 Then
					response.redirect pcPageName & "?sub=1&msg=" & pcv_strGenericPageError
				Else
					pcv_xCounter = 1
					pcv_strTotalWeight = 0
					errnum = 0
					
					EDC_ErrMsg=""
					EDC_SuccessMsg=""
					
					call GetEDCSettings()
					tmpEDCWeb=EDCURL & "GetPostageLabelXML"
					
					tmpEDCCal=EDCURL & "CalculatePostageRateXML"
					
					IF (request.form("submit")<>"") THEN
					
						if EDCTestMode<>"1" then
							EDCCalMode="yes"
						else
							EDCCalMode="no"
						end if
										
					tmpXMLOrg="<LabelRequest"
					
					tmpXMLCalOrg="<PostageRateRequest>"
					
					if (EDCTestMode="1") OR (request.form("submit")<>"") then
						tmpXMLOrg=tmpXMLOrg & " Test=""YES"""
					else
						tmpXMLOrg=tmpXMLOrg & " Test=""NO"""
					end if
					tmpXMLOrg=tmpXMLOrg & " LabelType=""" & session("pcEDCLabelType") & """"
					if session("pcEDCLabelSize")<>"" then
						tmpXMLOrg=tmpXMLOrg & " LabelSize=""" & session("pcEDCLabelSize") & """"
					end if
					tmpXMLOrg=tmpXMLOrg & " ImageFormat=""" & session("pcEDCLabelFormat") & """>"
					tmpXMLOrg=tmpXMLOrg & "<RequesterID>" & XMLReplace(EDCPartnerID) & "</RequesterID>"
					
					tmpXMLCalOrg=tmpXMLCalOrg & "<RequesterID>" & XMLReplace(EDCPartnerID) & "</RequesterID>"
					
					tmpXMLOrg=tmpXMLOrg & "<AccountID>" & XMLReplace(EDCUserID) & "</AccountID>"
					tmpXMLOrg=tmpXMLOrg & "<PassPhrase>" & XMLReplace(EDCPassP) & "</PassPhrase>"
					
					tmpXMLCalOrg=tmpXMLCalOrg & "<CertifiedIntermediary>"
					tmpXMLCalOrg=tmpXMLCalOrg & "<AccountID>" & XMLReplace(EDCUserID) & "</AccountID>"
					tmpXMLCalOrg=tmpXMLCalOrg & "<PassPhrase>" & XMLReplace(EDCPassP) & "</PassPhrase>"
					tmpXMLCalOrg=tmpXMLCalOrg & "</CertifiedIntermediary>"					
					
					if session("pcEDCMailClass") = "FirstLetter" OR session("pcEDCMailClass") = "FirstEnvelope" then
						tmpXMLOrg=tmpXMLOrg & "<MailClass>First</MailClass>"
						tmpXMLCalOrg=tmpXMLCalOrg & "<MailClass>First</MailClass>"
					else
						tmpXMLOrg=tmpXMLOrg & "<MailClass>" & XMLReplace(session("pcEDCMailClass")) & "</MailClass>"
						tmpXMLCalOrg=tmpXMLCalOrg & "<MailClass>" & XMLReplace(session("pcEDCMailClass")) & "</MailClass>"
					end if
					if session("pcEDCMailClass") = "Parcel" or session("pcEDCMailClass") = "ParcelSelect" then
						tmpXMLOrg=tmpXMLOrg & "<SortType>Nonpresorted</SortType>"
						tmpXMLOrg=tmpXMLOrg & "<EntryFacility>Other</EntryFacility>"
						
						tmpXMLCalOrg=tmpXMLCalOrg & "<SortType>Nonpresorted</SortType>"
						tmpXMLCalOrg=tmpXMLCalOrg & "<EntryFacility>Other</EntryFacility>"
					end if
					tmpXMLOrg=tmpXMLOrg & "<DateAdvance>0</DateAdvance>"
					
					tmpXMLCalOrg=tmpXMLCalOrg & "<DateAdvance>0</DateAdvance>"
					
					tmpXMLOrg=tmpXMLOrg & "<Stealth>FALSE</Stealth>"
					tmpXMLOrg=tmpXMLOrg & "<PartnerCustomerID>" & XMLReplace(Session("pcAdmincustomerRefNo1")) & "</PartnerCustomerID>"
					tmpXMLOrg=tmpXMLOrg & "<PartnerTransactionID>" & XMLReplace(Session("pcAdminOrderID")) & "</PartnerTransactionID>"
					tmpXMLOrg=tmpXMLOrg & "<ToName>" & XMLReplace(replace(Session("pcAdminToName"),"''","'")) & "</ToName>"
					if Session("pcAdminToFirm")<>"" then
						tmpXMLOrg=tmpXMLOrg & "<ToCompany>" & XMLReplace(replace(Session("pcAdminToFirm"),"''","'")) & "</ToCompany>"
					end if
					tmpXMLOrg=tmpXMLOrg & "<ToAddress1>" & XMLReplace(replace(Session("pcAdminToAddress1"),"''","'")) & "</ToAddress1>"
					if Session("pcAdminToAddress2")<>"" then
						tmpXMLOrg=tmpXMLOrg & "<ToAddress2>" & XMLReplace(replace(Session("pcAdminToAddress2"),"''","'")) & "</ToAddress2>"
					end if
					tmpXMLOrg=tmpXMLOrg & "<ToCity>" & XMLReplace(replace(Session("pcAdminToCity"),"''","'")) & "</ToCity>"
					if Session("pcAdminToState") & Session("pcAdminToStateCode")<>"" then
						tmpXMLOrg=tmpXMLOrg & "<ToState>" & XMLReplace(Session("pcAdminToState") & Session("pcAdminToStateCode")) & "</ToState>"
					end if
					if Session("pcAdminToZip5")<>"" then
					tmpXMLOrg=tmpXMLOrg & "<ToPostalCode>" & XMLReplace(Session("pcAdminToZip5")) & "</ToPostalCode>"
					tmpXMLCalOrg=tmpXMLCalOrg & "<ToPostalCode>" & XMLReplace(Session("pcAdminToZip5")) & "</ToPostalCode>"
					end if
					if Session("pcAdminToZip4")<>"" then
					tmpXMLOrg=tmpXMLOrg & "<ToZIP4>" & XMLReplace(Session("pcAdminToZip4")) & "</ToZIP4>"
					end if
					if Session("pcAdminToPhone")<>"" then
					tmpXMLOrg=tmpXMLOrg & "<ToPhone>" & XMLReplace(fnStripPhone(Session("pcAdminToPhone"))) & "</ToPhone>"
					end if
					if Session("pcAdminRecipientEMail")<>"" then
					tmpXMLOrg=tmpXMLOrg & "<ToEMail>" & XMLReplace(Session("pcAdminRecipientEMail")) & "</ToEMail>"
					end if
					if Session("pcAdminToCountryCode")<>"" then
						call opendb()
						query="SELECT countryName FROM countries WHERE countryCode like '" & Session("pcAdminToCountryCode") & "';"
						set rsQ=connTemp.execute(query)
						if not rsQ.eof then
							tmpXMLOrg=tmpXMLOrg & "<ToCountry>" & XMLReplace(rsQ("countryName")) & "</ToCountry>"
							tmpXMLCalOrg=tmpXMLCalOrg & "<ToCountry>" & XMLReplace(rsQ("countryName")) & "</ToCountry>"
						end if
						set rsQ=nothing
					end if
					tmpXMLOrg=tmpXMLOrg & "<FromName>" & XMLReplace(replace(Session("pcAdminFromName"),"''","'")) & "</FromName>"
					tmpXMLOrg=tmpXMLOrg & "<FromCompany>" & XMLReplace(replace(Session("pcAdminFromFirm"),"''","'")) & "</FromCompany>"
					tmpXMLOrg=tmpXMLOrg & "<ReturnAddress1>" & XMLReplace(replace(Session("pcAdminFromAddress1"),"''","'")) & "</ReturnAddress1>"
					if Session("pcAdminFromAddress2")<>"" then
						tmpXMLOrg=tmpXMLOrg & "<ReturnAddress2>" & XMLReplace(replace(Session("pcAdminFromAddress2"),"''","'")) & "</ReturnAddress2>"
					end if
					tmpXMLOrg=tmpXMLOrg & "<FromCity>" & XMLReplace(replace(Session("pcAdminFromCity"),"''","'")) & "</FromCity>"
					tmpXMLOrg=tmpXMLOrg & "<FromState>" & XMLReplace(Session("pcAdminFromState")) & "</FromState>"
					tmpXMLOrg=tmpXMLOrg & "<FromPostalCode>" & XMLReplace(Session("pcAdminFromZip5")) & "</FromPostalCode>"
					
					tmpXMLCalOrg=tmpXMLCalOrg & "<FromPostalCode>" & XMLReplace(Session("pcAdminFromZip5")) & "</FromPostalCode>"
					
					if Session("pcAdminFromZip4")<>"" then
						tmpXMLOrg=tmpXMLOrg & "<FromZIP4>" & XMLReplace(Session("pcAdminFromZip4")) & "</FromZIP4>"
					end if
					if Session("pcAdminFromPhone")<>"" then
						 tmpXMLOrg=tmpXMLOrg & "<FromPhone>" & XMLReplace(fnStripPhone(Session("pcAdminFromPhone"))) & "</FromPhone>"
					end if
					if session("pcAdminLabelDate")<>"" then
						tmpXMLOrg=tmpXMLOrg & "<ShipDate>" & XMLReplace(session("pcAdminLabelDate")) & "</ShipDate>"
						tmpXMLCalOrg=tmpXMLCalOrg & "<ShipDate>" & XMLReplace(session("pcAdminLabelDate")) & "</ShipDate>"
					end if
					tmpXMLOrg=tmpXMLOrg & "<ResponseOptions PostagePrice=""TRUE""/>"
					
					session("pcEDCpcPackageCount")=pcPackageCount
					
					ELSE
						pcPackageCount=session("pcEDCpcPackageCount")
					END IF
					
					'///////////////////////////////////////////////////////////////////////
					'// START LOOP FOR PACKAGE TAG
					'///////////////////////////////////////////////////////////////////////
					For pcv_xCounter = 1 to pcPackageCount
						'// If the package was processed, skip it.
						if pcLocalArray(pcv_xCounter-1) <> "shipped" then
							IF (request.form("submit")<>"") THEN
							tmpXMLPak=""
							tmpXMLCal=""
							if Session("pcAdminWeightInOunces"&pcv_xCounter)<>"" then
								tmpXMLPak=tmpXMLPak & "<WeightOz>" & XMLReplace(Session("pcAdminWeightInOunces"&pcv_xCounter)) & "</WeightOz>"
								tmpXMLCal=tmpXMLCal & "<WeightOz>" & XMLReplace(Session("pcAdminWeightInOunces"&pcv_xCounter)) & "</WeightOz>"
							end if
							if session("pcEDCPakValue"&pcv_xCounter)<>"" then
								tmpXMLPak=tmpXMLPak & "<Value>" & XMLNumber(session("pcEDCPakValue"&pcv_xCounter)) & "</Value>"
								tmpXMLCal=tmpXMLCal & "<Value>" & XMLNumber(session("pcEDCPakValue"&pcv_xCounter)) & "</Value>"
							end if
					
							if session("pcEDCPakDesc"&pcv_xCounter)<>"" then
								tmpXMLPak=tmpXMLPak & "<Description>" & XMLReplace(session("pcEDCPakDesc"&pcv_xCounter)) & "</Description>"
							end if
					
							if session("pcEDCShape"&pcv_xCounter)<>"" then
								if session("pcEDCMailClass") = "FirstLetter" OR session("pcEDCMailClass") = "FirstEnvelope" then
									if session("pcEDCMailClass") = "FirstLetter" then
										if Session("pcAdminWeightInOunces"&pcv_xCounter)>"13" then
											tmpXMLPak=tmpXMLPak & "<MailpieceShape>Flat</MailpieceShape>"
											tmpXMLCal=tmpXMLCal & "<MailpieceShape>Flat</MailpieceShape>"
										else
											tmpXMLPak=tmpXMLPak & "<MailpieceShape>Letter</MailpieceShape>"
											tmpXMLCal=tmpXMLCal & "<MailpieceShape>Letter</MailpieceShape>"
										end if
									else
										tmpXMLPak=tmpXMLPak & "<MailpieceShape>Flat</MailpieceShape>"
										tmpXMLCal=tmpXMLCal & "<MailpieceShape>Flat</MailpieceShape>"
									end if
								else
									tmpXMLPak=tmpXMLPak & "<MailpieceShape>" & XMLReplace(session("pcEDCShape"&pcv_xCounter)) & "</MailpieceShape>"
									tmpXMLCal=tmpXMLCal & "<MailpieceShape>" & XMLReplace(session("pcEDCShape"&pcv_xCounter)) & "</MailpieceShape>"
								end if
							end if
					
							if ((session("pcEDCShape"&pcv_xCounter)="Parcel") OR (session("pcEDCShape"&pcv_xCounter)="IrregularParcel") OR (session("pcEDCShape"&pcv_xCounter)="LargeParcel") OR (session("pcEDCShape"&pcv_xCounter)="OversizedParcel")) AND (session("pcEDCLength"&pcv_xCounter)>"0") AND (session("pcEDCWidth"&pcv_xCounter)>"0") AND (session("pcEDCHeight"&pcv_xCounter)>"0") then
								tmpXMLPak=tmpXMLPak & "<MailpieceDimensions>"
								tmpXMLPak=tmpXMLPak & "<Length>" & XMLNumber(session("pcEDCLength"&pcv_xCounter)) & "</Length>"
								tmpXMLPak=tmpXMLPak & "<Width>" & XMLNumber(session("pcEDCWidth"&pcv_xCounter)) & "</Width>"
								tmpXMLPak=tmpXMLPak & "<Height>" & XMLNumber(session("pcEDCHeight"&pcv_xCounter)) & "</Height>"
								tmpXMLPak=tmpXMLPak & "</MailpieceDimensions>"
								
								tmpXMLCal=tmpXMLCal & "<MailpieceDimensions>"
								tmpXMLCal=tmpXMLCal & "<Length>" & XMLNumber(session("pcEDCLength"&pcv_xCounter)) & "</Length>"
								tmpXMLCal=tmpXMLCal & "<Width>" & XMLNumber(session("pcEDCWidth"&pcv_xCounter)) & "</Width>"
								tmpXMLCal=tmpXMLCal & "<Height>" & XMLNumber(session("pcEDCHeight"&pcv_xCounter)) & "</Height>"
								tmpXMLCal=tmpXMLCal & "</MailpieceDimensions>"
							end if
					
							if session("pcEDCIM"&pcv_xCounter)<>"" OR session("pcEDCSC"&pcv_xCounter)<>"" then
								tmpXMLPak=tmpXMLPak & "<Services"
								tmpXMLCal=tmpXMLCal & "<Services"
								if session("pcEDCIM"&pcv_xCounter)<>"" then
									tmpXMLPak=tmpXMLPak & " InsuredMail=""" & session("pcEDCIM"&pcv_xCounter) &""""
									tmpXMLCal=tmpXMLCal & " InsuredMail=""" & session("pcEDCIM"&pcv_xCounter) &""""
								end if
								if session("pcEDCSC"&pcv_xCounter)<>"" then
									tmpXMLPak=tmpXMLPak & " SignatureConfirmation=""" & session("pcEDCSC"&pcv_xCounter) & """"
									tmpXMLCal=tmpXMLCal & " SignatureConfirmation=""" & session("pcEDCSC"&pcv_xCounter) & """"
								end if
								tmpXMLPak=tmpXMLPak & " />"
								tmpXMLCal=tmpXMLCal & " />"
							end if
					
							if (session("pcEDCIM"&pcv_xCounter)<>"") AND (session("pcEDCIM"&pcv_xCounter)<>"OFF") AND (session("pcEDCInValue"&pcv_xCounter)>"0") then
								tmpXMLPak=tmpXMLPak & "<InsuredValue>" & XMLNumber(session("pcEDCInValue"&pcv_xCounter)) & "</InsuredValue>"
								tmpXMLCal=tmpXMLCal & "<InsuredValue>" & XMLNumber(session("pcEDCInValue"&pcv_xCounter)) & "</InsuredValue>"
							end if
					
							if session("pcEDCNWD"&pcv_xCounter)<>"" then
								tmpXMLPak=tmpXMLPak & "<NoWeekendDelivery>" & XMLReplace(session("pcEDCNWD"&pcv_xCounter)) & "</NoWeekendDelivery>"
							end if
					
							if session("pcEDCSHD"&pcv_xCounter)<>"" then
								tmpXMLPak=tmpXMLPak & "<SundayHolidayDelivery>" & XMLReplace(session("pcEDCSHD"&pcv_xCounter)) & "</SundayHolidayDelivery>"
								tmpXMLCal=tmpXMLCal & "<SundayHolidayDelivery>" & XMLReplace(session("pcEDCSHD"&pcv_xCounter)) & "</SundayHolidayDelivery>"
							end if
							
							tmpXMLCal=tmpXMLCal & "<ResponseOptions PostagePrice=""TRUE""/>"
							
							tmpXMLCalWeb=tmpXMLCalOrg & tmpXMLCal & "</PostageRateRequest>"

							tmpXMLCalWeb="postageRateRequestXML=" & Server.URLEncode(tmpXMLCalWeb)
					
							tmpXML=tmpXMLOrg & tmpXMLPak & "</LabelRequest>"
							tmpXML="labelRequestXML=" & Server.URLEncode(tmpXML)
							
							session("pcEDCRequestXML" & pcv_xCounter)=tmpXML
							
							ELSE
							
							tmpXML=session("pcEDCRequestXML" & pcv_xCounter)
							tmpXML=replace(tmpXML,"Test%3D%22YES%22","Test%3D%22NO%22")
							
							END IF
							
							if (EDCTestMode="0") AND (request.form("submit")<>"") then
								result=ConnectServer(tmpEDCCal,"POST","","",tmpXMLCalWeb)
							else
								result=ConnectServer(tmpEDCWeb,"POST","","",tmpXML)
							end if

							IF result="ERROR" or result="TIMEOUT" THEN
								EDC_ErrMsg="Cannot connect to Endicia Label Server"
								tmpCode=0
							ELSE
								tmpCode=FindStatusCode(result)
								if tmpCode="0" then
									tmpCode=1
									EDC_SuccessMsg="Got Postage Label successfully!"
								else
									tmpCode=0
									EDC_ErrMsg=FindErrMsg(result)
								end if
							END IF
							Call SaveTrans(tmpXML,result,tmpCode,1)
						End if '// end skip shipped packages 						 
					Next
					
					IF (tmpCode=1) AND (EDC_ErrMsg="") THEN
						IF (request("endicia")="buypostage") OR (EDCTestMode="1") then
						'// SAVE all data to the database
								pcv_intOrderID=Session("pcAdminOrderID")
								pcv_xCounter=1
		
								dtShippedDate=Date()
								if pcv_shippedDate<>"" then									
									'dtShippedDate=objFedExClass.pcf_FedExDateFormat(dtShippedDate)
									if SQL_Format="1" then
										dtShippedDate=(day(dtShippedDate)&"/"&month(dtShippedDate)&"/"&year(dtShippedDate))
									else
										dtShippedDate=(month(dtShippedDate)&"/"&day(dtShippedDate)&"/"&year(dtShippedDate))
									end if
								end if
								tmpEndiciaExp=Date()
								if SQL_Format="1" then
									tmpEndiciaExp=(day(tmpEndiciaExp)&"/"&month(tmpEndiciaExp)&"/"&year(tmpEndiciaExp))
								else
									tmpEndiciaExp=(month(tmpEndiciaExp)&"/"&day(tmpEndiciaExp)&"/"&year(tmpEndiciaExp))
								end if
								tmpEndiciaExp=tmpEndiciaExp & " " & Time()
			
								if scDB="Access" then
									pcInsertDate="#"
								else
									pcInsertDate="'"
								end if	
								
								pcPackNum=0
								query="SELECT Count(*) As TotalPaks FROM pcPackageInfo WHERE IDOrder=" & pcv_intOrderID & ";"
								set rs=connTemp.execute(query)
								
								if not rs.eof then
									pcPackNum=rs("TotalPaks")
								end if
								set rs=nothing
								
								pcPackNum=pcPackNum+1
								
								err.clear
								query="INSERT INTO pcPackageInfo (idOrder, pcPackageInfo_PackageNumber, pcPackageInfo_PackageWeight, pcPackageInfo_ShipToName, pcPackageInfo_ShipToAddress1, pcPackageInfo_ShipToAddress2, pcPackageInfo_ShipToCity, pcPackageInfo_ShipToStateCode, pcPackageInfo_ShipToZip, pcPackageInfo_ShipToCountry, pcPackageInfo_ShipToEmail, pcPackageInfo_ShipFromCompanyName, pcPackageInfo_ShipFromAttentionName, pcPackageInfo_ShipFromAddress1, pcPackageInfo_ShipFromAddress2, pcPackageInfo_ShipFromCity, pcPackageInfo_ShipFromStateProvinceCode, pcPackageInfo_ShipFromPostalCode, pcPackageInfo_ShipFromCountryCode, pcPackageInfo_PackageLength, pcPackageInfo_PackageWidth, pcPackageInfo_PackageHeight, pcPackageInfo_Status, pcPackageInfo_UPSServiceCode, pcPackageInfo_TrackingNumber, pcPackageInfo_ShipMethod, pcPackageInfo_ShippedDate, pcPackageInfo_Comments, pcPackageInfo_ShipToContactName, pcPackageInfo_UPSLabelFormat, pcPackageInfo_MethodFlag,pcPackageInfo_Endicia,pcPackageInfo_EndiciaLabelFile,pcPackageInfo_EndiciaIsPIC,pcPackageInfo_EndiciaExp) " 
								query=query&"VALUES ("&pcv_intOrderID&","&pcPackNum&","&Session("pcAdminWeightInOunces"&pcv_xCounter)&", '"&ReplaceN(Session("pcAdminToName"))&"', '"&ReplaceN(Session("pcAdminToAddress1"))&"', '"&ReplaceN(Session("pcAdminToAddress2"))&"', '"&ReplaceN(Session("pcAdminToCity"))&"', '"&ReplaceN(Session("pcAdminToStateCode"))&ReplaceN(Session("pcAdminToState"))&"', '"&ReplaceN(Session("pcAdminToZip5"))&"', '" & ReplaceN(Session("pcAdminToCountryCode")) & "', '"&ReplaceN(Session("pcAdminRecipientEMail"))&"', '"&ReplaceN(Session("pcAdminFromFirm"))&"', '"&ReplaceN(Session("pcAdminFromName"))&"', '"&ReplaceN(Session("pcAdminFromAddress1"))&"', '"&ReplaceN(Session("pcAdminFromAddress2"))&"', '"&ReplaceN(Session("pcAdminFromCity"))&"', '"&ReplaceN(Session("pcAdminFromState"))&"', '"&Session("pcAdminFromZip5")&"', 'US', '"&session("pcEDCLength"&pcv_xCounter)&"', '"&session("pcEDCWidth"&pcv_xCounter)&"', '"&session("pcEDCHeight"&pcv_xCounter)&"', 0, '"&session("pcEDCMailClass")&"', '"&EDCTrackingNum&"', '"&session("pcEDCMailClass")&"', "&pcInsertDate&dtShippedDate&pcInsertDate&", '"&Session("pcAdminAdmComments")&"', '"&replace(Session("pcAdminToName"),"'","''")&"', '"&session("pcEDCLabelFormat")&"', 4, 1,'" & EDCLabelFile & "'," & EDCIsPIC & ","&pcInsertDate&tmpEndiciaExp&pcInsertDate&");"
								set rs=connTemp.execute(query)
								set rs=nothing
								if err.number<>0 then
									response.write err.description
									response.end
								end if
								
								query="SELECT pcPackageInfo_ID FROM pcPackageInfo WHERE idorder=" & pcv_intOrderID & " AND pcPackageInfo_TrackingNumber='"&EDCTrackingNum&"' ORDER by pcPackageInfo_ID DESC;"
								set rs=connTemp.execute(query)
								if err.number<>0 then
									response.write err.description
									response.end
								else
									pcv_PackageID=rs("pcPackageInfo_ID")
								end if
								set rs=nothing
								
								query="UPDATE pcEDCTrans SET pcPackageInfo_ID=" & pcv_PackageID & " WHERE idorder=" & pcv_intOrderID & " AND pcET_ID=" & EDCID & ";"
								set rs=connTemp.execute(query)
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
								
								if session("pcAdminMarkedAsShipped")="1" then
									if scDB="SQL" then
										query="UPDATE pcPackageInfo SET pcPackageInfo_ShippedDate='" & dtShippedDate & "', pcPackageInfo_Comments='" & Session("pcAdminAdmComments") & "' WHERE idOrder="&pcv_intOrderID&" AND pcPackageInfo_ID="&pcv_PackageID&";"
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
						END IF
					ELSE
						Session("EDC_ErrMsg") = "A Postage Label cannot be generated for this combination of Mail Class and Package Type. The error returned from Endicia is shown below. Please select a different Mail Class or Package Type and try again.<br /><br /> " & EDC_ErrMsg 
						response.redirect pcPageName & "?sub=1&savesession=1&msg=" & "A Postage Label cannot be generated for this combination of Mail Class and Package Type. The error returned from Endicia is shown below. Please select a different Mail Class or Package Type and try again.<br /><br /> " & EDC_ErrMsg 
					END IF
					
					IF (request("endicia")="buypostage") OR (EDCTestMode="1") then
					%>
                        <table class="pcCPcontent">
                            <tr>
                              <td colspan="2" class="pcCPspacer"><p>Your label has been created and saved. You can click on the link below to print your label. You will be able to access this label from the order details shipping area and also have access to tracking information.<br>
							  	<%if EDCTestMode="1" then%>
								<br>
								<b><u>Note:</u> You are in "TEST" mode and this is a SAMPLE label so do NOT use it to ship your package! You can go to <a href="EDC_manage.asp">Endicia's account settings</a> to switch to "LIVE" mode.</b>
								<br>
								<%else%>
								<br>
								<b><u>Note:</u> You are in "LIVE" mode and the label is genuine. If you have any issues, you can create a refund request for this label within 48 hours on the 'Order Details' page or on the website: <a href="https://www.endicia.com/Account/LogIn/" target="_blank">www.endicia.com</a></b>
								<br>
								<%end if%>
                                <br>
                                <%strLabelLink="USPSLabels/" & strFileName%>
								<a href="<%=strLabelLink%>" target="_blank">View/Print Postage Label</a></p>
                                <p>
								<table class="pcCPcontent">
                            	<%if CDbl(EDCFPostage)<>Cdbl(EDCSPostage) then%>
								<tr>
                                	<td colspan="2" class="pcCPspacer"><hr></td>
                                </tr>
								<tr>
									<td>Postage Amount:</td><td><%=scCurSign & money(EDCSPostage)%></td>
								</tr>
								<%if EDCFeesDetails<>"" then%>
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"><hr></td>
                                    </tr>
									<tr>
										<td colspan="2"><strong>Additional Fees</strong></td>
									</tr>
									<%tmp1=split(EDCFeesDetails,"|~|")
									intC1=ubound(tmp1)
									For k=0 to intC1
										if tmp1(k)<>"" then
											tmp2=split(tmp1(k),"|!|")%>
											<tr>
												<td>
                                                <%
												Dim pcvStrAdditionalFees
												pcvStrAdditionalFees=tmp2(0)
												pcvStrAdditionalFees=replace(pcvStrAdditionalFees,"SignatureConfirmation","Signature Confirmation")
												pcvStrAdditionalFees=replace(pcvStrAdditionalFees,"DeliveryConfirmation","Delivery Confirmation (required by Endicia)")
												response.write pcvStrAdditionalFees
												%>
                                                </td>
												<td><%=scCurSign & money(tmp2(1))%></td>
											</tr>
										<%end if
									Next
								end if
								end if%>
								<tr>
                                    <td><font size="3"><b>Postage Price:</b></font></td>
									<td><font size="3"><b><%=scCurSign & money(EDCFPostage)%></b></font></td>
                                </tr>
								<tr>
                                    <td><b>Remaining Balance:</b></td>
									<td><b><%=scCurSign & money(EDCRBalance)%></b></td>
                                </tr>
								</table>	
                                <br>
                                <a href="OrdDetails.asp?id=<%=pcv_intOrderID%>">Back to order details &gt;&gt;</a><br><br></p>
                                </td>
                            </tr>
                        </table>
                <%ELSE%>
					<table class="pcCPcontent">
                            <tr>
                              <td colspan="2" class="pcCPspacer"><p>Postage for this shipment was calculated successfully and is shown below. Click the &quot;Process Shipment&quot; button to purchase this postage from Endicia and print the shipping label. If there is a problem with the shipment, you can request a refund.
                                <p>
                                <table class="pcCPcontent">
                                <% dim pcTotalExInsValue
								pcTotalExInsValue = Ccur(0)
								For pcv_xCounter = 1 to pcPackageCount
									'// If the package was processed, skip it.
									If Session("pcAdminPounds"&pcv_xCounter)="" then Session("pcAdminPounds"&pcv_xCounter) = 0
									If Session("pcAdminOunces"&pcv_xCounter)="" then Session("pcAdminOunces"&pcv_xCounter) = 0
								 	%>
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                    <% 
									tmpEDCShape1 = session("pcEDCShape"&pcv_xCounter)
									select case tmpEDCShape1
										case "Card"
												tmpEDCShape1 = "Card"
										case "Letter"
												tmpEDCShape1 = "Letter"
										case "Flat"
												tmpEDCShape1 = "Flat"
										case "Parcel"
												tmpEDCShape1 = "Parcel"
											case "LargeParcel"
												tmpEDCShape1 = "Large Parcel"
										case "IrregularParcel"
												tmpEDCShape1 = "Irregular Parcel"
										case "OversizedParcel"
												tmpEDCShape1 = "Oversized Parcel"
										case "FlatRateEnvelope"
												tmpEDCShape1 = "Flat Rate Envelope"
										case "FlatRatePaddedEnvelope"
												tmpEDCShape1 = "Flat Rate Padded Envelope (Commercial Plus customers only)"
										case "SmallFlatRateBox"
												tmpEDCShape1 = "Small Flat Rate Box"
										case "MediumFlatRateBox"
												tmpEDCShape1 = "Medium Flat Rate Box"
										case "LargeFlatRateBox"
												tmpEDCShape1 = "Large Flat Rate Box"
									end select
									%>
                                    <tr>
                                    	<td colspan="2"><span style="font-size: 16px; font-weight: bold;">Shipment Summary:</span></td>
                                    </tr>
                                    <tr>
                                        <td width="20%" nowrap>Mail Class:</td>
                                        <td width="80%"><strong><%=session("pcEDCMailClass")%></strong></td>
                                    </tr>
                                    <tr>
                                        <td nowrap>Package Type</td>
                                        <td><strong><%=tmpEDCShape1%></strong></td>
                                    </tr>
                                    <tr>
                                      <td>Weight:</td>
                                      <td nowrap><strong><%=Session("pcAdminPounds"&pcv_xCounter) %> Pounds <%= Session("pcAdminOunces"&pcv_xCounter) %> Ounces</strong></td>
                                    </tr>
									<% if session("pcEDCPakValue"&pcv_xCounter)<>"" then %>
                                        <tr>
                                          <td nowrap>Package Value:</td>
                                          <td><strong><%=scCurSign & money(session("pcEDCPakValue"&pcv_xCounter))%></strong></td>
                                        </tr>
                                    <% end if %>
                                    <tr>
                                      <td>Insurance:</td>
                                      <td><strong><%=session("pcEDCIM"&pcv_xCounter)%></strong></td>
                                    </tr>
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                	<% If session("pcEDCIM"&pcv_xCounter) = "Endicia" Then
										pcTotalExInsValue = ccur(pcTotalExInsValue) + Ccur(session("pcEDCExInsValue"&pcv_xCounter))
									End If
								Next %>
                                </table>
								<table class="pcCPcontent">
                            	<%if CDbl(EDCFPostage)<>Cdbl(EDCSPostage) OR pcTotalExInsValue > "0" then%>
								<tr>
                                	<td colspan="2" class="pcCPspacer"><hr></td>
                                </tr>
                                <% 'If pcTotalExInsValue > "0" then
									'EDCSPostage = EDCSPostage + pcTotalExInsValue
								'End If %>
								<tr>
									<td nowrap>Postage Amount:</td><td><strong><%=scCurSign & money(EDCSPostage)%></strong></td>
								</tr>
								<% if EDCFeesDetails<>"" OR pcTotalExInsValue > "0" then %>
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"><hr></td>
                                    </tr>
									<tr>
										<td colspan="2"><strong>Additional Fees</strong></td>
									</tr>
									<%tmp1=split(EDCFeesDetails,"|~|")
									intC1=ubound(tmp1)
									For k=0 to intC1
										if tmp1(k)<>"" then
											tmp2=split(tmp1(k),"|!|")%>
											<tr>
												<td>
												<%
												pcvStrAdditionalFees=tmp2(0)
												pcvStrAdditionalFees=replace(pcvStrAdditionalFees,"SignatureConfirmation","Signature Confirmation")
												pcvStrAdditionalFees=replace(pcvStrAdditionalFees,"DeliveryConfirmation","Delivery Confirmation (required by Endicia)")
												response.write pcvStrAdditionalFees
												%>
                                                </td>
												<td><strong><%=scCurSign & money(tmp2(1))%></strong></td>
											</tr>
										<%end if
									Next
									If pcTotalExInsValue > 0 then %>
                                    	<tr>
                                            <td>Endicia Insurance</td>
                                            <td><strong><%=scCurSign & money(pcTotalExInsValue)%></strong></td>
                                        </tr>
                                        <tr>
                                        	<td colspan="2"><em>Please Note: Endicia Insurance is <u>not</u> included in the total postage amount shown on this page. You will be billed for all insurance charges at the end of the month, directly by Endicia.</em></td>
                                        </tr>
									<% End If
								end if
								end if%>
								<tr>
                                	<td colspan="2" class="pcCPspacer"><hr></td>
                                </tr>
                                <% 'If pcTotalExInsValue > "0" then
									'EDCFPostage = EDCFPostage + pcTotalExInsValue
								'End If %>
								<tr>
                                    <td colspan="2"><span style="font-size: 16px; font-weight: bold;">Estimated Postage Price: <span style="color:#06F;"><%=scCurSign & money(EDCFPostage)%></span></span></td>
                                </tr>
								</table>
                                <hr>
								<script>
                                   function DisableButton(b)
                                   {
                                      b.disabled = true;
                                      b.value = 'Processing';
                                   }
                                </script>
                                <div style="text-align: center; margin-top: 15px;">
                                	<input type="button" name="submit" value="Process Shipment" class="submit2" onclick="javascript:DisableButton(this);pcf_Open_EndiciaPop();document.location.href='<%=pcPageName%>?endicia=buypostage'">
                                    &nbsp;
                                    <input type="button" name="Button" value="Change Shipment Settings" onclick="document.location.href='<%=pcPageName%>'" class="ibtnGrey">&nbsp;
                                    <input type="button" name="Button" value="Go Back To Order Details" onclick="document.location.href='OrdDetails.asp?id=<%=Session("pcAdminOrderID")%>'" class="ibtnGrey">
                                </div>
                                </p>
                              </td>
                            </tr>
                        </table>
				<%END IF
				End if
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
				if request.QueryString("savesession")="1" then
					msg = session("EDC_ErrMsg")
					Session("EDC_ErrMsg") = ""
				end if
				
				if (request("edcaction")="refill") AND (request("amount")<>"") then
					tmpAmount=request("amount")
					if not IsNumeric(tmpAmount) then
						msg="Refill your Endicia Account error: Amount is not a numeric value"
					else
						if (Cdbl(tmpAmount)<10) OR (Cdbl(tmpAmount)>99999.99) then
							msg="Refill your Endicia Account error: The Refill Amount must be between $10.00 and $99,999.99 in US dollars, rounded to the nearest cent."
						end if
					end if
					if msg="" then
						call GetEDCSettings()
						FuncResult=BuyPostage(tmpAmount)
						if FuncResult<>"1" then
							msg="Refill your Endicia Account error: " & EDC_ErrMsg
						else
							msg=EDC_SuccessMsg
							tmpSuccess=1
						end if
					end if
				end if
					

				if msg<>"" then %>
					<div <%if tmpSuccess=1 then%>class="pcCPmessageSuccess"<%else%>class="pcCPmessage"<%end if%>>
					 <%=msg%>
                     <% if instr(lcase(msg),"address is too ambiguous") then %>
                     <br><br>
                     <a href="http://zip4.usps.com/zip4/welcome.jsp" target="_blank">USPS address lookup</a>
                     <% end if %>
					</div>
				<% end if %>
<!--MAIN-->
					<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">
                        <input type="hidden" name="LabelMode" value="<%=pcv_LabelMode%>">
                        <table class="pcCPcontent">	
                            <tr>
							<% 
                            dim strJSOnChangeTabCnt, k, intTempJSChangeCnt					
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
							<script>
							function change(id, newClass) {
								identity=document.getElementById(id);
								identity.className=newClass;
							}

							function popwin(fileName)
							{
								pcInfoWin = window.open('','InfoWindow','scrollbars=auto,status=no,width=400,height=300')
								pcInfoWin.location.href = fileName;
							}
    
							var tabs = ["tab1","tab2","tab3","tab4",<%=strTabCnt%>];

							function showTab( tab ){

							// first make sure all the tabs are hidden
							for(i=0; i < tabs.length; i++){
								var obj = document.getElementById(tabs[i]);
								obj.style.display = "none";
							}
			
							// show the tab we're interested in
							var obj = document.getElementById(tab);
							obj.style.display = "block";

							}
							</script>
							<td valign="top">
                                <div class="menu">
                                    <ul>
										<li><a id="tabs1" class="current" onclick="change('tabs1', 'current');change('tabs2', '');change('tabs3', '');change('tabs4', '')<%=strJSOnChangeTabCnt%>;showTab('tab1')">Account Balance</a></li>
										<li><a id="tabs2" onclick="change('tabs1', '');change('tabs2', 'current');change('tabs3', '');change('tabs4', '')<%=strJSOnChangeTabCnt%>;showTab('tab2')">Label Settings</a></li>
                                        <li><a id="tabs3" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', 'current');change('tabs4', '')<%=strJSOnChangeTabCnt%>;showTab('tab3')">Ship Settings</a></li>
                                        <li><a id="tabs4" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current')<%=strJSOnChangeTabCnt%>;showTab('tab4')">Ship From/Recipient</a></li>
                                        <% strOnclickTabCnt=""
                                        if pcPackageCount=1 then %>
											<li><a id="tabs5" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', '');change('tabs5', 'current');showTab('tab5')">Package Information</a></li>
                                        <% else %>
                                            <% for k=1 to pcPackageCount
                                                intTempPackageCnt=3+int(k)
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
                                // ENDICIA ACOUNT
                                //////////////////////////////////////////////////////////////////////////////////////////////
                                -->
								<div id="tab1" class="panes" style="display:block">
								<table class="pcCPcontent">
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                    <tr valign="top">
                                        <td ><span class="title">Your Endicia Account</span></td>
										<td align="right"><img src="images/PoweredByEndicia_small.jpg" border="0"></td>
                                    </tr>
									<%
									call GetEDCSettings()
									If (EDCReg="1") AND (EDCUserID>"0") then
										FuncResult=AutoRefill()
										if FuncResult="1" then%>
											<tr>
											<td colspan="2">
											<div class="pcCPmessageSuccess"> 
												<%=EDC_SuccessMsg%>
											</div>
											</td>
											</tr>
										<%end if
									End If
									tmpEDC=GetAccountStatus()
									if tmpEDC="1" then%>
									<tr>
                                        <td><font size="3"><b>Current Balance:</b></font></td>
										<td><font size="3"><b><%=scCurSign & money(EDCABalance)%></b></font></td>
                                    </tr>
									<%if FindXMLValue(EDCReXML,"AscendingBalance")<>"" then%>
									<tr>
                                        <td>Printed Postage: </td>
										<td><%=scCurSign & money(FindXMLValue(EDCReXML,"AscendingBalance"))%></td>
                                    </tr>
									<%end if%>
									<%if FindXMLValue(EDCReXML,"AccountStatus")<>"" then%>
									<tr>
                                        <td>AccountStatus: </td>
										<td nowrap>
											<b>
											<%tmpEDCStatus=FindXMLValue(EDCReXML,"AccountStatus")
											if Ucase(tmpEDCStatus)<>"A" then
												tmpEDC=0
											end if
											Select Case tmpEDCStatus
											Case "A": response.write "Active"

											Case "C": response.write "Closed account"

											Case "P": response.write "Pending customer-initiated closeout"

											Case "S": response.write "Suspended due to multiple bad login attempts"

											Case "X": response.write "Suspended pending new account review"
											
											Case Else: response.write "Unknown"
											
											End Select
											%>
											</b>
										</td>
                                    </tr>
									<%end if%>
									<%if tmpEDC="1" then%>
									<tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
									<tr valign="top">
                                        <th colspan="2">Refill your account</th>
                                    </tr>
									<tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
									<tr valign="top">
                                        <td>Amount:</td>
										<td><%=scCurSign%><input name="edcrefill" id="edcrefill" type="text" size="10" value="0"><br>
										<i>Enter an amount between $10.00 and $99,999.99 US dollars, rounded to the nearest cent.</i></td>
                                    </tr>
									<tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
									<tr>
										<td>&nbsp;</td>
										<td>
										<script>
										function isDigit(s)
										{
											var test=""+s;
											if(test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
											{
												return(true) ;
											}
											return(false);
										}
	
										function allDigit(s)
										{
											var test=""+s ;
											for (var k=0; k <test.length; k++)
											{
												var c=test.substring(k,k+1);
												if (isDigit(c)==false)
												{
													return (false);
												}
											}
											return (true);
										}
										function checkNumber(tmpField)
										{
											if (tmpField.value=="")
											{
												alert("Please enter a value for 'Amount' field");
												tmpField.focus();
												return(false);
											}
											if (allDigit(tmpField.value) == false)
											{
												alert("Please enter a numeric value for 'Amount' field");
												tmpField.focus();
												return(false);
											}
											if ((parseFloat(tmpField.value)<10) || (parseFloat(tmpField.value)>99999.99))
											{
												alert("Please enter a numeric value between: 10.00 - 99999.99 into the 'Amount' field.");
												tmpField.focus();
												return(false);
											}
											return(true);
										}
										</script>
										<input type="button" name="refillbtn" class="submit2" value=" Refill Account " onclick="if (checkNumber(document.form1.edcrefill)) location='<%=pcPageName%>?edcaction=refill&amount='+document.form1.edcrefill.value;"></td>
									</tr>
									<%end if%>
									<%else%>
									<tr>
										<td colspan="2">
										<div class="pcCPmessage"> 
											<%=EDC_ErrMsg%>
										</div>
										</td>
									</tr>
									<%end if%>
								</table>
								</div>
								
								 <!--							
                                //////////////////////////////////////////////////////////////////////////////////////////////
                                // LABEL SETTINGS
                                //////////////////////////////////////////////////////////////////////////////////////////////
                                -->
								
								<div id="tab2" class="panes">
								<table class="pcCPcontent">
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                    <tr>
                                        <td colspan="2"><span class="title">Label Settings:</span></td>
                                    </tr>
									<tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
									<%if (session("pcEDCMailClass")="") then
										if tmpEDCShipMethod<>"" then
											session("pcEDCMailClass")=tmpEDCShipMethod
										else
											session("pcEDCMailClass")="Express"
										end if
									end if%>									
									<tr>
										<th colspan="2">Mail Class</th>
									</tr>
									<tr>
                                        <td style="padding: 15px;" valign="top">
                                        <b>Domestic:</b>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="Express" <%if (session("pcEDCMailClass")="Express") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/shipping/expressmail.htm" target="_blank">Express Mail</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('DestinationConfirm',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="FirstLetter" <%if (session("pcEDCMailClass")="FirstLetter") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/send/waystosendmail/senditwithintheus/firstclassletters.htm" target="_blank">First-Class, Letter</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('DestinationConfirm',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="FirstEnvelope" <%if (session("pcEDCMailClass")="FirstEnvelope") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/send/waystosendmail/senditwithintheus/firstclassflats.htm" target="_blank">First-Class, Large Envelope</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="First" <%if (session("pcEDCMailClass")="First") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/send/waystosendmail/senditwithintheus/firstclassparcels.htm" target="_blank">First-Class, Package</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="LibraryMail" <%if (session("pcEDCMailClass")="LibraryMail") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/send/waystosendmail/senditwithintheus/libraryrate.htm" target="_blank">Library Mail</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="MediaMail" <%if (session("pcEDCMailClass")="MediaMail") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/send/waystosendmail/senditwithintheus/mediamail.htm" target="_blank">Media Mail</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="ParcelPost" <%if (session("pcEDCMailClass")="ParcelPost") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/send/waystosendmail/senditwithintheus/parcelpost.htm" target="_blank">Standard Post</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="ParcelSelect" <%if (session("pcEDCMailClass")="ParcelSelect") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/send/waystosendmail/senditwithintheus/parcelselect.htm" target="_blank">Parcel Select</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="Priority" <%if (session("pcEDCMailClass")="Priority") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/shipping/prioritymail.htm" target="_blank">Priority Mail</a>
                                        </td>
                                        <td style="padding: 15px;" valign="top">
                                        <b>International:</b>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="ExpressMailInternational" <%if (session("pcEDCMailClass")="ExpressMailInternational") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/international/expressmailinternational.htm" target="_blank">Express Mail International</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="FirstClassMailInternational" <%if (session("pcEDCMailClass")="FirstClassMailInternational") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/international/airmailinternational.htm" target="_blank">First-Class Mail International</a>
                                        <br />
                                        <input type="radio" name="edcMailClass" onclick="javascript:setMSOptions1('Default',this.value);ShowHideExServices(this.value);ShowHideSCServices(this.value);" value="PriorityMailInternational" <%if (session("pcEDCMailClass")="PriorityMailInternational") then%>checked<%end if%> class="clearBorder"> <a href="http://www.usps.com/international/prioritymailinternational.htm" target="_blank">Priority Mail International</a>
                                        </td>
									</tr>
									<tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
									<tr>
										<th colspan="2">Label Attributes</th>
									</tr>
									<tr>
                                        <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
									<tr valign="top">
										<td nowrap>Label Ouput Format:</td>
										<%if (session("pcEDCLabelFormat")="") then
											session("pcEDCLabelFormat")="GIF"
										end if%>
										<td>
											<select id="edcLabelFormat" name="edcLabelFormat">
												<option value="GIF" <%if (session("pcEDCLabelFormat")="GIF") then%>selected<%end if%>>GIF</option>
												<option value="JPEG" <%if (session("pcEDCLabelFormat")="JPEG") then%>selected<%end if%>>JPEG</option>
												<option value="PDF" <%if (session("pcEDCLabelFormat")="PDF") then%>selected<%end if%>>PDF</option>
												<option value="PNG" <%if (session("pcEDCLabelFormat")="PNG") then%>selected<%end if%>>PNG</option>
											</select>
                                            &nbsp;Note: Only GIF is available for some international labels.</i>
										</td>
									</tr>
								</table>
								</div>
								
								 <!--							
                                //////////////////////////////////////////////////////////////////////////////////////////////
                                // SHIPPING SETTINGS
                                //////////////////////////////////////////////////////////////////////////////////////////////
                                -->
								
                                <div id="tab3" class="panes">		
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
										<td colspan="2">If you want ProductCart to automatically flag the package as &quot;Shipped&quot; at the time the label is generated, check the checkbox below. Otherwise, you can manually update the package status from the Order Details page.</td>
								  	</tr>
                                    <tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<%if session("pcAdminMarkedAsShipped")="" then
										session("pcAdminMarkedAsShipped")="1"
									end if%>
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
                                        <td align="right"><INPUT tabIndex="25" type="checkbox" value="1" name="MarkedAsShipped" class="clearBorder" <%if session("pcAdminMarkedAsShipped")="1" then%>checked<%end if%> onclick = "HideCommentRow()"></td>
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

                                    <tr id="AdmCommentsRow" <%if session("pcAdminMarkedAsShipped")<>"1" then%>style="display:none"<%end if%>>
                                        <td valign="top" align="right"><b>Comments:</b></td>
                                        <td valign="top">
                                        <textarea name="AdmComments" size="40" rows="10" cols="65"><%=pcv_AdmComments%></textarea>
                                        <div style="margin: 10px 15px 15px 0;" class="pcCPnotes">NOTE: additional text will be added to the message that is e-mailed to the customer depending on whether this is a partial or final shipment, and depending on which shipping provider was used for the shipment, if any. The additional text can be edited by editing the file &quot;<strong>includes/languages_ship.asp</strong>&quot;. Ship a few test orders in different scenarios to see how the e-mail sent to the customer looks like.</div>  
										<%
										Session("pcAdminAdmComments")=""
										%>                                     
                                        </td>
                                    </tr>
                                    <tr>
                                    	<th colspan="2">Ship Date</th>
                                    </tr>	
                                    <tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
                                    <tr>
                                      <td colspan="2">Date Package Will Be Mailed. The Ship Date may be up to 5 days in advance.</td>
                                    </tr>
                                    <tr>
                                      <td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                    <tr>
                                      <td align="right" valign="top"><b>Ship Date:</b></td>
                                      <td><p>
                                        <select name="LabelDate" id="LabelDate">
                                            <% pcv_TodayDate=Date() %>
                                            <option value="<%=pcv_TodayDate%>" <%=pcf_SelectOption("LabelDate", (exFormatDate(pcv_TodayDate,"%mm/%dd/%yyyy")))%>><%=exFormatDate(pcv_TodayDate,"%mm/%dd/%yyyy")%></option>
                                            <option value="<%=pcv_TodayDate+1%>" <%=pcf_SelectOption("LabelDate",exFormatDate(pcv_TodayDate+1,"%mm/%dd/%yyyy"))%>><%=exFormatDate(pcv_TodayDate+1,"%mm/%dd/%yyyy")%></option>
                                            <option value="<%=pcv_TodayDate+2%>" <%=pcf_SelectOption("LabelDate",exFormatDate(pcv_TodayDate+2,"%mm/%dd/%yyyy"))%>><%=exFormatDate(pcv_TodayDate+2,"%mm/%dd/%yyyy")%></option>
											<option value="<%=pcv_TodayDate+3%>" <%=pcf_SelectOption("LabelDate",exFormatDate(pcv_TodayDate+2,"%mm/%dd/%yyyy"))%>><%=exFormatDate(pcv_TodayDate+3,"%mm/%dd/%yyyy")%></option>
											<option value="<%=pcv_TodayDate+4%>" <%=pcf_SelectOption("LabelDate",exFormatDate(pcv_TodayDate+2,"%mm/%dd/%yyyy"))%>><%=exFormatDate(pcv_TodayDate+4,"%mm/%dd/%yyyy")%></option>
                                        </select>
										  <%pcs_UPSRequiredImageTag "LabelDate", false%>
                                      </p></td>
                                    </tr>
                                    <tr>
                                        <td colspan="2" class="pcCPspacer"></td>
									</tr>
                                </table>
                            	</div>

                                <!--							
                                /////////////////////////////////////////////////////////////////////////////////
                                // SHIPPER/RECIPIENT
                                //////////////////////////////////////////////////////////////////////////////////
                                -->
                                <div id="tab4" class="panes">
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
                                        <tr>
                                            <td><p>Attention  Name:</p></td>
                                            <td>
                                            <p>
                                            <input name="FromName" type="text" id="FromName" value="<%=pcf_FillFormField("FromName", false)%>">
                                            <%pcs_UPSRequiredImageTag "FromName", true%>			
                                            </p></td>
                                        </tr>
										<%if len(Session("ErrFromPhone"))>0 then %>
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
                                            </p></td>
                                        </tr>
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
                                            <td><p>Attention Name:</p></td>
                                            <td>
                                            <p>
                                            <input name="ToName" type="text" id="ToName" value="<%=pcf_FillFormField("ToName", false)%>">
                                            <%pcs_UPSRequiredImageTag "ToName", true%>			
                                            </p></td>
                                    </tr>
                                    <tr>
                                    <td width="25%"><p>Company Name:</p></td>
                                    <td width="77%">
                                    <p>
                                    <input name="ToFirm" type="text" id="ToFirm" value="<%=pcf_FillFormField("ToFirm", true)%>" size="50">
                                    </p></td>
                                    </tr>
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
                                            </p></td>
                                        </tr>
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
                                    </p>
									</td>
                                    </tr>
									<%
			                        pcv_isStateCodeRequired = true '// determines if validation is performed (true or false)
			                        pcv_isProvinceCodeRequired = false '// determines if validation is performed (true or false)
			                        pcv_isCountryCodeRequired = true '// determines if validation is performed (true or false)					
                        
		    	                    '// #3 Additional Required Info
			                        pcv_strTargetForm = "form1" '// Name of Form
			                        pcv_strCountryBox = "ToCountryCode" '// Name of Country Dropdown
			                        pcv_strTargetBox = "ToStateCode" '// Name of State Dropdown
			                        pcv_strProvinceBox =  "ToState" '// Name of Province Field
                        
			                        '// Set local Country to Session
			                        if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
			                            Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session("pcAdminToCountryCode")
			                        end if
                            
			                        '// Set local State to Session
			                        if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
			                            Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session("pcAdminToStateCode")
			                        end if
                            
			                        '// Set local Province to Session
			                         if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
			                             Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  Session("pcAdminToState")
			                         end if
			                        call opendb()%>					
			                        <!--#include file="../includes/javascripts/pcStateAndProvince.asp"-->
			                        <%
			                        pcs_CountryDropdown
			                        pcs_StateProvince
					                %>
                                    <tr>
                                        <td><p>Postal Code:</p></td>
                                        <td>
                                        <p>
                                        <input name="ToZip5" type="text" id="ToZip5" value="<%=pcf_FillFormField("ToZip5", true)%>" size="7" maxlength="9">
                                        <%pcs_UPSRequiredImageTag "ToZip5", true %> - <input name="ToZip4" type="text" id="ToZip4" value="<%=pcf_FillFormField("ToZip4", false)%>" size="4" maxlength="4">
                                        <%pcs_UPSRequiredImageTag "ToZip4", false %>			
                                        </p></td>
                                    </tr>
                                    <tr>
                                    	<td colspan="2" class="pcCPspacer"></td>
                                    </tr>
                                    
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
									<div id="tab<%=4+int(k)%>" class="panes">				
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
                                                    Click Here to view <b>package content</b>.
                                            
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
																			- <%=pcv_strProductDescription%><br>
																			<%
																			rs2.movenext
																		Loop								
																	end if	
																	call closedb()
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
												<tr valign="top">
													<%if session("pcEDCPakValue" & k)="" then
														session("pcEDCPakValue" & k)="0"
													end if%>
													<td width="20%" nowrap><b>Package Value:</b></td>
                                                    <% if session("pcEDCPakValue" & k) = "" OR session("pcEDCPakValue" & k) = "0" then
														if pcPackageCount = 1 then 
														else
															pcv_OrderTotal = 0
														end if
													else
														pcv_OrderTotal = session("pcEDCPakValue" & k)
													end if %>
													<td width="80%"><%=scCurSign%><input type="text" size="10" name="edcPakValue<%=k%>" value="<%=Money(pcv_OrderTotal)%>"><br>
													<i>Required for International Mail</i>
													</td>
												</tr>
												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<tr valign="top">
													<%if session("pcEDCPakDesc" & k)="" then
														session("pcEDCPakDesc" & k)="Package for the Order #" & (scpre+int(pcv_intOrderID))
													end if%>
													<td nowrap><b>Package Description:</b></td>
													<td><input type="text" size="50" name="edcPakDesc<%=k%>" value="<%=session("pcEDCPakDesc" & k)%>"><br>
													<i>Description of the item(s) shipped.</i>
													</td>
												</tr>
												
												<tr>
                                                    <td colspan="2" class="pcCPspacer"></td>
                                                </tr>
                                                <tr>
                                                    <th colspan="2">Package Weight</th>
                                                </tr>
                                                <tr>
                                                    <td colspan="2" class="pcCPspacer"></td>
                                                </tr>
												<tr>
													<td colspan="2">
														Enter the weight of the package in ounces. If there is more than one package in the shipment, enter the weight of the first package or the total shipment weight.</p>
                                                        <p style="padding-top: 5px;">&nbsp;</p>
                                                        <p style="padding-top: 5px;">Weight: <input name="Pounds<%=k%>" type="text" id="Pounds<%=k%>" value="<%=pcf_FillFormField("Pounds"&k, true)%>" size="4">
                                                        lbs.
														  <%pcs_UPSRequiredImageTag "Pounds"&k, true%>&nbsp;<input name="Ounces<%=k%>" type="text" id="Ounces<%=k%>" value="<%=pcf_FillFormField("Ounces"&k, true)%>" size="4">
														  ozs.
														  <%pcs_UPSRequiredImageTag "Ounces"&k, true%>
                                                        </p>
                                                    </td>
                                                </tr>
												<tr>
                                                    <td colspan="2" class="pcCPspacer"></td>
                                                </tr>
                                                <tr>
                                                    <th colspan="2">Package Type</th>
                                                </tr>
												<tr>
                                                    <td colspan="2" class="pcCPspacer"></td>
                                                </tr>
												<tr>
													<%
													
													if (session("pcEDCShape" & k)="") then
														if tmpEDCPackType="" then
															if session("pcEDCMailClass")="Express" then
																if USPS_EM_PACKAGE="Flat Rate Envelope" then
																	tmpEDCPackType="FlatRateEnvelope"
																end if
															end if
															if session("pcEDCMailClass")="Priority" then
																Select Case USPS_PM_PACKAGE
																	Case "Flat Rate Envelope": tmpEDCPackType="FlatRateEnvelope"
																	Case "Flat Rate Box": tmpEDCPackType="SmallFlatRateBox"
																	Case "Flat Rate Box1": tmpEDCPackType="MediumFlatRateBox"
																	Case "Flat Rate Box2": tmpEDCPackType="LargeFlatRateBox"
																End Select
																if (tmpEDCPackType="FlatRateEnvelope") AND (Clng(pcv_ShipWeight)>Clng(USPS_PM_FREWeightLimit)) AND (Clng(USPS_PM_FREWeightLimit)>0) then
																	Select Case USPS_PM_FREOption
																		Case "Flat Rate Box": tmpEDCPackType="SmallFlatRateBox"
																		Case "Flat Rate Box1": tmpEDCPackType="MediumFlatRateBox"
																		Case "Flat Rate Box2": tmpEDCPackType="LargeFlatRateBox"
																	End Select
																end if
															end if
														end if
														if tmpEDCPackType<>"" then
															session("pcEDCShape" & k)=tmpEDCPackType
														else
															session("pcEDCShape" & k)="Parcel"
														end if
													end if%>
													<td colspan="2">
														<Select id="edcShape<%=k%>" name="edcShape<%=k%>" onchange="javascript:ShowHide3D(this.value);">
															<option value="Card" <%if (session("pcEDCShape" & k)="Card") then%>selected<%end if%>>Card</option>
															<option value="Letter" <%if (session("pcEDCShape" & k)="Letter") then%>selected<%end if%>>Letter</option>
															<option value="Flat" <%if (session("pcEDCShape" & k)="Flat") then%>selected<%end if%>>Flat</option>
															<option value="Parcel" <%if (session("pcEDCShape" & k)="Parcel") then%>selected<%end if%>>Parcel</option>
															<option value="LargeParcel" <%if (session("pcEDCShape" & k)="LargeParcel") then%>selected<%end if%>>Large Parcel</option>
															<option value="IrregularParcel" <%if (session("pcEDCShape" & k)="IrregularParcel") then%>selected<%end if%>>Irregular Parcel</option>
															<option value="OversizedParcel" <%if (session("pcEDCShape" & k)="OversizedParcel") then%>selected<%end if%>>Oversized Parcel</option>
															<option value="FlatRateEnvelope" <%if (session("pcEDCShape" & k)="FlatRateEnvelope") then%>selected<%end if%>>Flat Rate Envelope</option>
															<option value="FlatRatePaddedEnvelope" <%if (session("pcEDCShape" & k)="FlatRatePaddedEnvelope") then%>selected<%end if%>>Flat Rate Padded Envelope (Commercial Plus customers only)</option>
															<option value="SmallFlatRateBox" <%if (session("pcEDCShape" & k)="SmallFlatRateBox") then%>selected<%end if%>>Small Flat Rate Box</option>
															<option value="MediumFlatRateBox" <%if (session("pcEDCShape" & k)="MediumFlatRateBox") then%>selected<%end if%>>Medium Flat Rate Box</option>
															<option value="LargeFlatRateBox" <%if (session("pcEDCShape" & k)="LargeFlatRateBox") then%>selected<%end if%>>Large Flat Rate Box</option>
														</select>
													  <script>
															function setMSOptions1(lt,mc) 
															{ 
																var select2 = document.form1.edcShape<%=k%>; 
																select2.options.length = 0;
																if (lt == "Default")
																{
																	if (mc == "Express")
																	{ 
																		select2.options[select2.options.length] = new Option("Letter","Letter");
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		select2.options[select2.options.length] = new Option("Flat Rate Envelope","FlatRateEnvelope");
																		select2.options[select2.options.length] = new Option("Flat Rate Padded Envelope (Commercial Plus customers only)","FlatRatePaddedEnvelope");
																	}
																	if (mc == "First")
																	{ 
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																	}
																	if (mc == "LibraryMail")
																	{ 
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																	}
																	if (mc == "MediaMail")
																	{ 
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																	}
																	if (mc == "ParcelPost")
																	{ 
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		select2.options[select2.options.length] = new Option("Large Parcel","LargeParcel");
																		select2.options[select2.options.length] = new Option("Oversized Parcel","OversizedParcel");
																	}
																	if (mc == "ParcelSelect")
																	{ 
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		select2.options[select2.options.length] = new Option("Large Parcel","LargeParcel");
																		select2.options[select2.options.length] = new Option("Oversized Parcel","OversizedParcel");
																	}
																	if (mc == "Priority")
																	{
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		select2.options[select2.options.length] = new Option("Large Parcel","LargeParcel");
																		select2.options[select2.options.length] = new Option("Irregular Parcel","IrregularParcel");
																		select2.options[select2.options.length] = new Option("Flat Rate Envelope","FlatRateEnvelope");
																		select2.options[select2.options.length] = new Option("Flat Rate Padded Envelope (Commercial Plus customers only)","FlatRatePaddedEnvelope");
																		select2.options[select2.options.length] = new Option("Small Flat Rate Box","SmallFlatRateBox");
																		select2.options[select2.options.length] = new Option("Medium Flat Rate Box","MediumFlatRateBox");
																		select2.options[select2.options.length] = new Option("Large Flat Rate Box","LargeFlatRateBox");
																	}
																	if (mc == "ExpressMailInternational")
																	{
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		select2.options[select2.options.length] = new Option("Flat Rate Envelope","FlatRateEnvelope");
																	}
																	if (mc == "FirstClassMailInternational")
																	{
																		select2.options[select2.options.length] = new Option("Letter","Letter");
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																	}
																	if (mc == "PriorityMailInternational")
																	{
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		select2.options[select2.options.length] = new Option("Flat Rate Envelope","FlatRateEnvelope");
																		select2.options[select2.options.length] = new Option("Small Flat Rate Box","SmallFlatRateBox");
																		select2.options[select2.options.length] = new Option("Medium Flat Rate Box","MediumFlatRateBox");
																		select2.options[select2.options.length] = new Option("Large Flat Rate Box","LargeFlatRateBox");
																	}
																}
												
																if (lt == "DestinationConfirm")
																{
																	if (mc == "FirstLetter")
																	{ 
																		select2.options[select2.options.length] = new Option("Letter","Letter");
																	}
																	else
																	{
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																	}
																}
																if (lt == "CertifiedMail")
																{ 
																	if (mc == "First")
																	{ 
																		select2.options[select2.options.length] = new Option("Card","Card");
																		select2.options[select2.options.length] = new Option("Letter","Letter");
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																	}
																	else
																	{
																		if (mc == "Priority")
																		{ 
																			select2.options[select2.options.length] = new Option("Card","Card");
																			select2.options[select2.options.length] = new Option("Letter","Letter");
																			select2.options[select2.options.length] = new Option("Flat","Flat");
																			select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		}
																		else
																		{
																			select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		}
																	}
																}
																if (lt == "International")
																{
																	if (mc == "ExpressMailInternational")
																	{
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		select2.options[select2.options.length] = new Option("Flat Rate Envelope","FlatRateEnvelope");
																	}
																	if (mc == "FirstClassMailInternational")
																	{
																		select2.options[select2.options.length] = new Option("Letter","Letter");
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																	}
																	if (mc == "PriorityMailInternational")
																	{
																		select2.options[select2.options.length] = new Option("Flat","Flat");
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																		select2.options[select2.options.length] = new Option("Flat Rate Envelope","FlatRateEnvelope");
																		select2.options[select2.options.length] = new Option("Small Flat Rate Box","SmallFlatRateBox");
																		select2.options[select2.options.length] = new Option("Medium Flat Rate Box","MediumFlatRateBox");
																		select2.options[select2.options.length] = new Option("Large Flat Rate Box","LargeFlatRateBox");
																	}
																	if ((mc != "ExpressMailInternational") && (mc != "FirstClassMailInternational") && (mc != "PriorityMailInternational"))
																	{
																		select2.options[select2.options.length] = new Option("Parcel","Parcel");
																	}
																}
																ShowHide3D(document.form1.edcShape<%=k%>.value);
															}
															
															function ShowHide3D(tmpValue)
															{
																if ((tmpValue=="Parcel") || (tmpValue=="IrregularParcel"))
																{
																	document.getElementById("3DArea").style.display=''
																}
																else
																{
																	document.getElementById("3DArea").style.display='none'	
																}
															}
															
															function ShowHideExServices(tmpValue)
															{
																if (tmpValue == "Express")
																{
																	document.getElementById("ExServices").style.display=''
																}
																else
																{
																	document.getElementById("ExServices").style.display='none'	
																}
															}
															function ShowHideSCServices(tmpValue)
															{
																if ((tmpValue=="Express") || (tmpValue=="FirstLetter") || (tmpValue=="FirstEnvelope"))
																{
																	document.getElementById("signCon").style.display='none'
																}
																else
																{
																	document.getElementById("signCon").style.display=''	
																}
															}
															
														</script>
													</td>
												</tr>
												<tr>
                                                    <td colspan="2" class="pcCPspacer"></td>
                                                </tr>
												<tr>
                                                    <td colspan="2">
													<table id="3DArea" <%if session("pcEDCShape" & k)<>"Parcel" AND session("pcEDCShape" & k)<>"IrregularParcel" then%>style="display:none"<%end if%> class="pcCPcontent">
													<tr>
    	                                                <td colspan="2"><strong>Package Dimensions</strong></td>
        	                                        </tr>
													<tr>
														<%
														if (session("pcEDCLength" & k)="0" AND session("pcEDCWidth" & k)="0" AND session("pcEDCHeight" & k)="0") then
															session("pcEDCLength" & k)=USPS_LENGTH
															session("pcEDCWidth" & k)=USPS_WIDTH
															session("pcEDCHeight" & k)=USPS_HEIGHT
														end if
														if (session("pcEDCLength" & k)="") then
															session("pcEDCLength" & k)=USPS_LENGTH
														end if%>
														<td width="15%">Length:</td>
														<td width="85%"><input type="text" id="edcLength<%=k%>" size="10" value="<%=session("pcEDCLength" & k)%>"></td>
													</tr>
													<tr>
														<%if (session("pcEDCWidth" & k)="") then
															session("pcEDCWidth" & k)=USPS_WIDTH
														end if%>
														<td>Width:</td>
														<td><input type="text" id="edcWidth<%=k%>" size="10" value="<%=session("pcEDCWidth" & k)%>"></td>
													</tr>
													<tr>
														<%if (session("pcEDCHeight" & k)="") then
															session("pcEDCHeight" & k)=USPS_HEIGHT
														end if%>
														<td>Height:</td>
														<td><input type="text" id="edcHeight<%=k%>" size="10" value="<%=session("pcEDCHeight" & k)%>"></td>
													</tr>
													</table>
												</td>
												</tr>
                           	                    <tr>
                                                    <td colspan="2" class="pcCPspacer">
													<script>
                                                    function GetRValue(tmpField)
                                                    {
                                                        for (var i=0; i < tmpField.length; i++)
                                                        {
                                                            if (tmpField[i].checked)
                                                            {
                                                                return(tmpField[i].value);
                                                            }
                                                        }
                                                        return("");
                                                    }

													setMSOptions<%=k%>('Default',GetRValue(document.form1.edcMailClass));
													document.form1.edcShape<%=k%>.value="<%=session("pcEDCShape" & k)%>";
													</script>
													</td>
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
                                                    <td align="right" nowrap>Customer Reference Number:</td>
													<td align="left"><input name="CustomerRefNo<%=k%>" type="text" value="<%=pcf_FillFormField("CustomerRefNo"&k, false)%>" size="15"></td>
												</tr>
												
												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<tr>
                    			                	  <th colspan="2">Additional Services</th>
                                			    </tr>	
			                                    <tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
												<tr>
													<td colspan="2">
                                                    
                                                    <table class="pcCPcontent" style="width: 100%; margin: 0;">
                                                    	<tr>
                                                        	<td colspan="2"><b><u>Insured Mail</u></b><br>If insurance is requested, please memember to enter insured value for the field below
                                                            </td>
                                                        </tr>
														<%if (session("pcEDCIM" & k)="") then
                                                            session("pcEDCIM" & k)="OFF"
                                                        end if%>
                                                        <tr valign="top">
                                                            <td align="right" width="5%"><input type="radio" name="edcIM<%=k%>" value="OFF" <%if (session("pcEDCIM" & k)="OFF") then%>checked<%end if%> class="clearBorder"></td>
                                                            <td width="95%">OFF - <i>No insurance requested.</i>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" valign="top"><input type="radio" name="edcIM<%=k%>" value="Endicia" <%if (session("pcEDCIM" & k)="Endicia") then%>checked<%end if%> class="clearBorder"></td>
                                                            <td valign="top">ON - Endicia Insurance - <i>Endicia Insurance requested (Maximum insurable value: $10,000). Endicia insurance fee is not included in the postage price. It is billed to your account.</i>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <% if session("pcEDCInValue" & k) = "" OR session("pcEDCInValue" & k) = "0" then
                                                                if pcPackageCount = 1 then 
                                                                    pcv_InsureTotal = pcv_OrderTotal
                                                                else
                                                                    pcv_InsureTotal = 0
                                                                end if
                                                            else
                                                                pcv_InsureTotal = session("pcEDCInValue" & k)
                                                            end if %>
                                                            <td colspan="2">Insured Value: <%=scCurSign%><input type="text" size="10" name="edcInValue<%=k%>" value="<%=money(pcv_InsureTotal)%>">
                                                            </td>
                                                        </tr>
                                                     </table>
                                                   </td>
                                                </tr>
												<tr>
													<td colspan="2">
													<table id="signCon" <%if session("pcEDCLabelType")<>"Default" OR session("pcEDCMailClass")="Express" then%>style="display:none"<%end if%> class="pcCPcontent" style="width: 100%; margin: 0;">
													<tr>
														<td colspan="2" class="pcCPspacer"></td>
													</tr>
													<tr>
														<td colspan="2"><b><u>Signature Confirmation</u></b></td>
													</tr>
													<%if (session("pcEDCSC" & k)="") then
														session("pcEDCSC" & k)="OFF"
													end if%>
													<tr valign="top">
														<td align="right" width="5%"><input type="radio" name="edcSC<%=k%>" value="OFF" <%if (session("pcEDCSC" & k)="OFF") then%>checked<%end if%> class="clearBorder"></td>
														<td width="95%" nowrap>OFF - <i>Signature Confirmation not requested.</i>
														</td>
													</tr>
													<tr valign="top">
														<td align="right"><input type="radio" name="edcSC<%=k%>" value="ON" <%if (session("pcEDCSC" & k)="ON") then%>checked<%end if%> class="clearBorder"></td>
														<td>ON - <i>Signature Confirmation requested.</i>
														</td>
													</tr>
													</table>
													</td>
												</tr>
												<tr>
													<td colspan="2">
													<table id="ExServices" <%if (session("pcEDCMailClass")<>"Express") then%>style="display:none"<%end if%> class="pcCPcontent" style="width: 100%; margin: 0;">
													<tr>
														<td colspan="2" class="pcCPspacer"></td>
													</tr>
													<tr valign="top">
														<td colspan="2"><b><u>Services For Express Mail Only:</u></b></td>
													</tr>
													<%if (session("pcEDCNWD" & k)="") then
														session("pcEDCNWD" & k)="FALSE"
													end if%>
													<tr valign="top">
														<td align="right" width="5%"><input type="checkbox" name="edcNWD<%=k%>" value="TRUE" <%if (session("pcEDCNWD" & k)="TRUE") then%>checked<%end if%> class="clearBorder"></td>
														<td width="95%" nowrap>No Weekend Delivery - <i>Mailpiece should NOT be delivered on a Saturday</i>
														</td>
													</tr>
													<%if (session("pcEDCSHD" & k)="") then
														session("pcEDCSHD" & k)="FALSE"
													end if%>
													<tr valign="top">
														<td align="right"><input type="checkbox" name="edcSHD<%=k%>" value="TRUE" <%if (session("pcEDCSHD" & k)="TRUE") then%>checked<%end if%> class="clearBorder"></td>
														<td>Sunday/Holiday Delivery - <i>Deliver on Sunday or holiday.</i>
														</td>
													</tr>
													</table>
													</td>
												</tr>
												<tr>
													<td colspan="2" class="pcCPspacer"></td>
												</tr>
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
                                <div align="center">
                                   <input type="submit" name="submit" value="<%if (EDCTestMode="1") then%>Process Shipment<%else%>Calculate Postage Price<%end if%>" onclick="javascript:pcf_Open_EndiciaPop();" <%if tmpEDC<>"1" then%>disabled<%end if%> class="submit2">
                                   &nbsp;
                                   <input type="button" name="Button" value="Start Over" onclick="document.location.href='<%=pcv_strPreviousPage%>'">
                                   &nbsp;
                                   <input type="button" name="Button" value="Go Back To Order Details" onclick="document.location.href='OrdDetails.asp?id=<%=Session("pcAdminOrderID")%>'">
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top"><div align="center">
                            </div></td>
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
<%call closedb() 

'// DESTROY THE USPS OBJECT
set objUSPSClass = nothing
%>
<%Response.write(pcf_ModalWindow("Connecting to Endicia Label Server... ","EndiciaPop", 300))%>
<!--#include file="AdminFooter.asp"-->