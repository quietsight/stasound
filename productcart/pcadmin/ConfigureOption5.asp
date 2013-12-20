<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEx Web Services Shipping Configuration"
pageIcon="pcv4_icon_settings.png"
%>
<% Section="shipOpt" %>
<% pcPageName = "ConfigureOption5.asp" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->

<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<%
Dim query, rs, conntemp
Dim pcv_strMethodName, pcv_strMethodReply, fedex_postdataWS, objFedExWSClass, objOutputXMLDocWS
Dim srvFEDEXWSXmlHttp, FEDEXWS_result, FEDEX_URL, pcv_strErrorMsg, objFEDEXXmlDoc, objFedExStream, strFileName, GraphicXML

'// Validate phone
function fnStripPhone(PhoneField)
	PhoneField=replace(PhoneField," ","")
	PhoneField=replace(PhoneField,"-","")
	PhoneField=replace(PhoneField,".","")
	PhoneField=replace(PhoneField,"(","")
	PhoneField=replace(PhoneField,")","")
	fnStripPhone = PhoneField
end function

'**************************************************************************
' START: If registration request was submitted, process request
'**************************************************************************
Dim pcv_strAccountName, pcv_strMeterNumber, pcv_strCarrierCode

pcv_strMethodName = "SubscriptionRequest"
pcv_strMethodReply = "v2:SubscriptionReply"
pcv_strVersion = "2"
CustomerTransactionIdentifier = "Subscription_Request"

if request.form("submit")<>"" then

	'// Set error count
	pcv_intErr=0

	'// generic error for page
	pcv_strGenericPageError = "At least one required field was empty."

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ValidateTextField	"FedExWSMode", true, 4
	pcs_ValidateTextField	"FedExWS_AccountNumber", true, 20

	pcs_ValidateTextField	"FedExWS_BillingAddress", true, 20
	pcs_ValidateTextField	"FedExWS_BillingCity", true, 20
	pcs_ValidateTextField	"FedExWS_BillingCountryCode", true, 2
	pcs_ValidateTextField	"FedExWS_BillingStateOrProvinceCode", true, 20
	pcs_ValidateTextField	"FedExWS_BillingPostalCode", true, 20

	pcs_ValidateTextField	"FedExWS_FirstName", true, 20
	pcs_ValidateTextField	"FedExWS_LastName", true, 20
	pcs_ValidateTextField	"FedExWS_CompanyName", false, 20
	pcs_ValidatePhoneNumber	"FedExWS_PhoneNumber", true, 16
	pcs_ValidateEmailField	"FedExWS_eMailAddress", false, 250
	pcs_ValidateTextField	"FedExWS_Line1", true, 20
	pcs_ValidateTextField	"FedExWS_Line2", false, 20
	pcs_ValidateTextField	"FedExWS_City", true, 20
	pcs_ValidateTextField	"FedExWS_CountryCode", true, 2
	pcs_ValidateTextField	"FedExWS_StateOrProvinceCode", true, 20
	pcs_ValidateTextField	"FedExWS_PostalCode", true, 20

	if len(Session("pcAdminFedExWS_Line1"))<1 AND (Session("pcAdminFedExWS_CountryCode")="US" OR Session("pcAdminFedExWS_CountryCode")="CA") then
		pcv_intErr=pcv_intErr+1
	end if

	pcs_ValidateTextField	"FedExWS_PostalCode", true, 20

	if len(Session("pcAdminFedExWS_PostalCode"))<1 AND (Session("pcAdminFedExWS_CountryCode")="US" OR Session("pcAdminFedExWS_CountryCode")="CA") then
		pcv_intErr=pcv_intErr+1
	end if

	Session("pcAdminFedExWS_AccountNumber") = replace(Session("pcAdminFedExWS_AccountNumber"),"-","")
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Check for Validation Errors. Do not proceed if there are errors.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If pcv_intErr>0 Then
		response.redirect pcPageName & "?msg=" & pcv_strGenericPageError
	Else

		'// Save collected data in database
		FedExWSAPI_ID=getUserInput(request("FedExWSAPI_ID"),4)

		'// Open the DB
		call opendb()

		'// Generate the Query (Save form data)
		if FedExWSAPI_ID=0 then
			query="INSERT INTO FedExWSAPI (FedExAPI_PersonName, FedExAPI_CompanyName, FedExAPI_Department, FedExAPI_PhoneNumber, FedExAPI_FaxNumber, FedExAPI_EmailAddress, FedExAPI_Line1, FedExAPI_Line2, FedExAPI_city, FedExAPI_State, FedExAPI_PostalCode, FedExAPI_Country) VALUES ('"& Session("pcAdminFedExWS_FirstName") & " " & Session("pcAdminFedExWS_LastName") &"', '"&Session("pcAdminFedExWS_CompanyName")&"', '"&Session("pcAdminFedExWS_Department")&"', '"&Session("pcAdminFedExWS_PhoneNumber")&"', '"&Session("pcAdminFedExWS_FaxNumber")&"', '"&Session("pcAdminFedExWS_EmailAddress")&"', '"&Session("pcAdminFedExWS_Line1")&"', '"&Session("pcAdminFedExWS_Line2")&"', '"&Session("pcAdminFedExWS_City")&"', '"&Session("pcAdminFedExWS_StateOrProvinceCode")&"', '"&Session("pcAdminFedExWS_PostalCode")&"', '"&Session("pcAdminFedExWS_CountryCode")&"');"
		else
			query="UPDATE FedExWSAPI SET FedExAPI_PersonName='"& Session("pcAdminFedExWS_FirstName") & " " & Session("pcAdminFedExWS_LastName") &"', FedExAPI_CompanyName='"&Session("pcAdminFedExWS_CompanyName")&"', FedExAPI_Department='"&Session("pcAdminFedExWS_Department")&"', FedExAPI_PhoneNumber='"&Session("pcAdminFedExWS_PhoneNumber")&"', FedExAPI_FaxNumber='"&Session("pcAdminFedExWS_FaxNumber")&"', FedExAPI_EmailAddress='"&Session("pcAdminFedExWS_EmailAddress")&"', FedExAPI_Line1='"&Session("pcAdminFedExWS_Line1")&"', FedExAPI_Line2='"&Session("pcAdminFedExWS_Line2")&"', FedExAPI_city='"&Session("pcAdminFedExWS_City")&"', FedExAPI_State='"&Session("pcAdminFedExWS_StateOrProvinceCode")&"', FedExAPI_PostalCode='"&Session("pcAdminFedExWS_PostalCode")&"', FedExAPI_Country='"&Session("pcAdminFedExWS_CountryCode")&"' WHERE FedExAPI_ID=1;"
		end if

		'// Execute the Query
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing

		'// Close the DB
		call closedb()

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Set our Object.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		set objFedExClass = New pcFedExWSClass

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Build Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		NameOfMethod = "RegisterWebCspUserRequest"
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v2=""http://fedex.com/ws/registration/v2"">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<v2:"&NameOfMethod&">"&vbcrlf

		fedex_postdataWS=fedex_postdataWS&"<v2:WebAuthenticationDetail>"&vbcrlf
		'If CSPTurnOn = 1 Then
			fedex_postdataWS=fedex_postdataWS&"<v2:CspCredential>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v2:Key>CPTi545ATGa1CD89</v2:Key>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v2:Password>8BB07q2XIIOFyNJeJQHMLv094</v2:Password>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</v2:CspCredential>"&vbcrlf
		'End IF
		fedex_postdataWS=fedex_postdataWS&"</v2:WebAuthenticationDetail>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<v2:ClientDetail>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v2:AccountNumber>"&Session("pcAdminFedExWS_AccountNumber")&"</v2:AccountNumber>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v2:ClientProductId>EIPC</v2:ClientProductId>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v2:ClientProductVersion>3424</v2:ClientProductVersion>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v2:Region>US</v2:Region>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</v2:ClientDetail>"&vbcrlf

		'--------------------
		'// TransactionDetail
		'--------------------
		objFedExClass.WriteParent "TransactionDetail", "2", ""
			objFedExClass.AddNewNode "CustomerTransactionId", "2", "Registration Request"
		objFedExClass.WriteParent "TransactionDetail", "2", "/"

		'--------------------
		'// Version
		'--------------------
		objFedExClass.WriteParent "Version", "2", ""
			objFedExClass.AddNewNode "ServiceId", "2", "fcas"
			objFedExClass.AddNewNode "Major", "2", "2"
			objFedExClass.AddNewNode "Intermediate", "2", "1"
			objFedExClass.AddNewNode "Minor", "2", "0"
		objFedExClass.WriteParent "Version", "2", "/"

		objFedExClass.WriteSingleParent "Categories", pcv_strVersion, "SHIPPING"
			objFedExClass.WriteParent "BillingAddress", pcv_strVersion, ""
				objFedExClass.AddNewNode "StreetLines", pcv_strVersion, Session("pcAdminFedExWS_BillingAddress")
				objFedExClass.AddNewNode "City", pcv_strVersion, Session("pcAdminFedExWS_BillingCity")
				objFedExClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, Session("pcAdminFedExWS_BillingStateOrProvinceCode")
				objFedExClass.AddNewNode "PostalCode", pcv_strVersion, Session("pcAdminFedExWS_BillingPostalCode")
				objFedExClass.AddNewNode "CountryCode", pcv_strVersion, Session("pcAdminFedExWS_BillingCountryCode")
			objFedExClass.WriteParent "BillingAddress", pcv_strVersion, "/"

			objFedExClass.WriteParent "UserContactAndAddress", pcv_strVersion, ""
				objFedExClass.WriteParent "Contact", pcv_strVersion, ""
					objFedExClass.WriteParent "PersonName", pcv_strVersion, ""
						objFedExClass.AddNewNode "FirstName", pcv_strVersion, Session("pcAdminFedExWS_FirstName")
						objFedExClass.AddNewNode "LastName", pcv_strVersion, Session("pcAdminFedExWS_LastName")
					objFedExClass.WriteParent "PersonName", pcv_strVersion, "/"
					objFedExClass.AddNewNode "CompanyName", pcv_strVersion, Session("pcAdminFedExWS_CompanyName")
					objFedExClass.AddNewNode "PhoneNumber", pcv_strVersion, fnStripPhone(Session("pcAdminFedExWS_PhoneNumber"))
					objFedExClass.AddNewNode "EMailAddress", pcv_strVersion, Session("pcAdminFedExWS_eMailAddress")
				objFedExClass.WriteParent "Contact", pcv_strVersion, "/"
				objFedExClass.WriteParent "Address", pcv_strVersion, ""
					objFedExClass.AddNewNode "StreetLines", pcv_strVersion, Session("pcAdminFedExWS_Line1") & " " & Session("pcAdminFedExWS_Line2")
					objFedExClass.AddNewNode "City", pcv_strVersion, Session("pcAdminFedExWS_City")
					objFedExClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, Session("pcAdminFedExWS_StateOrProvinceCode")
					objFedExClass.AddNewNode "PostalCode", pcv_strVersion, Session("pcAdminFedExWS_PostalCode")
					objFedExClass.AddNewNode "CountryCode", pcv_strVersion, Session("pcAdminFedExWS_CountryCode")
				objFedExClass.WriteParent "Address", pcv_strVersion, "/"
			objFedExClass.WriteParent "UserContactAndAddress", pcv_strVersion, "/"

		objFedExClass.EndXMLTransaction NameOfMethod, "2"
		'// Print out our newly formed request xml
		'response.Clear()
		'response.ContentType="text/xml"
		'response.Write(fedex_postdataWS)
		'response.End()

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Send Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'call objFedExClass.SendXMLRequest(fedex_postdataWS, Session("pcAdminFedExWSMode"))
		'//srvFEDEXWSXmlHttp.open "POST", "https://wsbeta.fedex.com:443/web-services", false
		'//SMART POST
		srvFEDEXWSXmlHttp.open "POST", FedExWSURL, false
		srvFEDEXWSXmlHttp.send(fedex_postdataWS)
		FEDEXWS_result = srvFEDEXWSXmlHttp.responseText
		'// Print out our response
		'response.Clear()
		'response.ContentType="text/xml"
		'response.write FEDEXWS_result
		'response.end

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Load Our Response.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		call objFedExClass.LoadXMLResults(FEDEXWS_result)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Baseline Logging
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Log our Transaction
		call objFedExClass.pcs_LogTransaction(fedex_postdataWS, pcv_strMethodName&"_in"&q&".in", true)
		'// Log our Response
		call objFedExClass.pcs_LogTransaction(FEDEXWS_result, pcv_strMethodName&"_out"&q&".out", true)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Redirect with a Message OR complete some task.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// ERROR
		pcv_strNotificationCode = objFedExClass.ReadResponseNode("//v2:RegisterWebCspUserReply", "v2:Notifications/v2:Severity")
		If pcv_strNotificationCode <> "SUCCESS" Then
			pcv_strErrorMessage = objFedExClass.ReadResponseNode("//v2:RegisterWebCspUserReply", "v2:Notifications/v2:Message")
			response.redirect "ConfigureOption5.asp?msg="&pcv_strErrorMessage
			response.end
		End If

		'// Web User Credentials
		pcv_strWUKey = objFedExClass.ReadResponseNode("//v2:RegisterWebCspUserReply", "v2:Credential/v2:Key")
		pcv_strWUPassword = objFedExClass.ReadResponseNode("//v2:RegisterWebCspUserReply", "v2:Credential/v2:Password")

			'// Ensure that the MeterNumber exists
		if pcv_strWUKey&""="" OR pcv_strWUPassword&""="" then
			response.redirect pcPageName & "?msg=There was an error activating your FedEx account. The FedEx servers may be down temporarily. Please try again later."
		else
			'// Process Subscribe Request!
			NameOfMethod = "SubscriptionRequest"
			fedex_postdataWS=""
			fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v2=""http://fedex.com/ws/registration/v2"">"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v2:"&NameOfMethod&">"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v2:WebAuthenticationDetail>"&vbcrlf
			'If CSPTurnOn = 1 Then
				fedex_postdataWS=fedex_postdataWS&"<v2:CspCredential>"&vbcrlf
					fedex_postdataWS=fedex_postdataWS&"<v2:Key>CPTi545ATGa1CD89</v2:Key>"&vbcrlf
					fedex_postdataWS=fedex_postdataWS&"<v2:Password>8BB07q2XIIOFyNJeJQHMLv094</v2:Password>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"</v2:CspCredential>"&vbcrlf
			'End IF
			fedex_postdataWS=fedex_postdataWS&"<v2:UserCredential>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v2:Key>"&pcv_strWUKey&"</v2:Key>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v2:Password>"&pcv_strWUPassword&"</v2:Password>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</v2:UserCredential>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</v2:WebAuthenticationDetail>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<v2:ClientDetail>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v2:AccountNumber>"&Session("pcAdminFedExWS_AccountNumber")&"</v2:AccountNumber>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v2:ClientProductId>EIPC</v2:ClientProductId>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v2:ClientProductVersion>3424</v2:ClientProductVersion>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v2:Region>US</v2:Region>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</v2:ClientDetail>"&vbcrlf
			objFedExClass.WriteParent "Version", "2", ""
				objFedExClass.AddNewNode "ServiceId", "2", "fcas"
				objFedExClass.AddNewNode "Major", "2", "2"
				objFedExClass.AddNewNode "Intermediate", "2", "1"
				objFedExClass.AddNewNode "Minor", "2", "0"
			objFedExClass.WriteParent "Version", "2", "/"

			objFedExClass.AddNewNode "CspSolutionId", "2", "120"
			objFedExClass.AddNewNode "CspType", "2", "CERTIFIED_SOLUTION_PROVIDER"

			objFedExClass.WriteParent "Subscriber", pcv_strVersion, ""
				objFedExClass.AddNewNode "AccountNumber", pcv_strVersion, Session("pcAdminFedExWS_AccountNumber")
				objFedExClass.WriteParent "Contact", pcv_strVersion, ""
					objFedExClass.AddNewNode "PersonName", pcv_strVersion, Session("pcAdminFedExWS_LastName")
					objFedExClass.AddNewNode "CompanyName", pcv_strVersion, Session("pcAdminFedExWS_CompanyName")
					objFedExClass.AddNewNode "PhoneNumber", pcv_strVersion, fnStripPhone(Session("pcAdminFedExWS_PhoneNumber"))
					objFedExClass.AddNewNode "EMailAddress", pcv_strVersion, Session("pcAdminFedExWS_eMailAddress")
				objFedExClass.WriteParent "Contact", pcv_strVersion, "/"
				objFedExClass.WriteParent "Address", pcv_strVersion, ""
					objFedExClass.AddNewNode "StreetLines", pcv_strVersion, Session("pcAdminFedExWS_Line1") & " " & Session("pcAdminFedExWS_Line2")
					objFedExClass.AddNewNode "City", pcv_strVersion, Session("pcAdminFedExWS_City")
					objFedExClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, Session("pcAdminFedExWS_StateOrProvinceCode")
					objFedExClass.AddNewNode "PostalCode", pcv_strVersion, Session("pcAdminFedExWS_PostalCode")
					objFedExClass.AddNewNode "CountryCode", pcv_strVersion, Session("pcAdminFedExWS_CountryCode")
				objFedExClass.WriteParent "Address", pcv_strVersion, "/"
			objFedExClass.WriteParent "Subscriber", pcv_strVersion, "/"

			objFedExClass.WriteParent "AccountShippingAddress", pcv_strVersion, ""
				objFedExClass.AddNewNode "StreetLines", pcv_strVersion, Session("pcAdminFedExWS_BillingAddress")
				objFedExClass.AddNewNode "City", pcv_strVersion, Session("pcAdminFedExWS_BillingCity")
				objFedExClass.AddNewNode "StateOrProvinceCode", pcv_strVersion, Session("pcAdminFedExWS_BillingStateOrProvinceCode")
				objFedExClass.AddNewNode "PostalCode", pcv_strVersion, Session("pcAdminFedExWS_BillingPostalCode")
				objFedExClass.AddNewNode "CountryCode", pcv_strVersion, Session("pcAdminFedExWS_BillingCountryCode")
			objFedExClass.WriteParent "AccountShippingAddress", pcv_strVersion, "/"

			objFedExClass.EndXMLTransaction NameOfMethod, "2"
			'// Print out our newly formed request xml
			'response.Clear()
			'response.ContentType="text/xml"
			'response.Write(fedex_postdataWS)
			'response.End()

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Send Our Transaction.
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'call objFedExClass.SendXMLRequest(fedex_postdataWS, Session("pcAdminFedExWSMode"))
			'//srvFEDEXWSXmlHttp.open "POST", "https://wsbeta.fedex.com:443/web-services", false
			'//SMART POST
			srvFEDEXWSXmlHttp.open "POST", FedExWSURL, false
			srvFEDEXWSXmlHttp.send(fedex_postdataWS)
			FEDEXWS_result = srvFEDEXWSXmlHttp.responseText
			'// Print out our response
			'response.Clear()
			'response.ContentType="text/xml"
			'response.write FEDEXWS_result
			'response.end

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Load Our Response.
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			call objFedExClass.LoadXMLResults(FEDEXWS_result)

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Baseline Logging
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Log our Transaction
			call objFedExClass.pcs_LogTransaction(fedex_postdataWS, pcv_strMethodName&"_in"&q&".in", true)
			'// Log our Response
			call objFedExClass.pcs_LogTransaction(FEDEXWS_result, pcv_strMethodName&"_out"&q&".out", true)

			'// ERROR
			pcv_strNotificationCode = objFedExClass.ReadResponseNode("//v2:SubscriptionReply", "v2:Notifications/v2:Severity")
			If pcv_strNotificationCode="SUCCESS" Then
			Else
				pcv_strErrorMessage = objFedExClass.ReadResponseNode("//v2:SubscriptionReply", "v2:Notifications/v2:Message")
				response.redirect pcPageName & "?msg="&pcv_strErrorMessage
			End If

			'// Web User Credentials
			pcv_strWUMeterNumber = objFedExClass.ReadResponseNode("//v2:SubscriptionReply", "v2:MeterNumber")

			'//////////////////////////////
			call opendb()
			
			query="UPDATE ShipmentTypes SET [password]='"&pcv_strWUMeterNumber&"', userID='"&Session("pcAdminFedExWS_AccountNumber")&"', AccessLicense='LIVE', FedExKey='"&pcv_strWUKey&"', FedExPwd='"&pcv_strWUPassword&"' WHERE (((ShipmentTypes.idShipment)=9));"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing

				if err.number<>0 then
					call closedb()
					response.redirect pcPageName & "?msg=There was an error activating your FedEx account. Please submit your registration request again."
				else
					'// Generate the Query (Update shipment types)
					query="UPDATE ShipmentTypes SET AccessLicense='LIVE' WHERE (((ShipmentTypes.idShipment)=9));"
					set rs=server.CreateObject("ADODB.RecordSet")
					'// Execute the Query
					set rs=conntemp.execute(query)
					set rs=nothing
					call closedb()
				'// No errors, redirect to next step
					session("FedExWSSetUP")="YES"
					pcs_ClearAllSessions()
					response.redirect "FEDEXWS_EditShipOptions.asp"
					response.end

				end if
			end if


	end if
end if
'**************************************************************************
' END: If registration request was submitted, process request
'**************************************************************************




'**************************************************************************
' START: Was FedExWS was previously registered by querying the database ?
'**************************************************************************
if request("changeMode")="" then
	call opendb()
	query="SELECT userID FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=9));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	If rs.eof then
		'FedExWS was previously activated - redirect
		session("FedExWSSetUP")="YES"
		response.redirect "FEDEXWS_EditShipOptions.asp"
		response.end
	end if
	set rs=nothing
	call closedb()
end if
'**************************************************************************
' END: Was FedExWS was previously registered by querying the database ?
'**************************************************************************
%>

<%
'**************************************************************************
' START: Get Fed Ex credentials
'**************************************************************************
call opendb()

'// Get Access License
query="SELECT ShipmentTypes.AccessLicense, ShipmentTypes.userID FROM ShipmentTypes WHERE idShipment=9;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

Session("pcAdminFedExWS_AccountNumber") = rs("userID")
strAccessLicense=rs("AccessLicense")

if len(strAccessLicense)<1 then
	strAccessLicense="TEST"
end if

'// Get Form Data
query="SELECT FedExAPI_ID, FedExAPI_PersonName, FedExAPI_CompanyName, FedExAPI_Department, FedExAPI_PhoneNumber, FedExAPI_FaxNumber, FedExAPI_EmailAddress, FedExAPI_Line1, FedExAPI_Line2, FedExAPI_city, FedExAPI_State, FedExAPI_PostalCode, FedExAPI_Country FROM FedExWSAPI;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if NOT rs.eof then
	FedExWSAPI_ID=rs("FedExAPI_ID")
	if request("changeMode")="Y" then
		Session("pcAdminFedExWS_PersonName")=rs("FedExAPI_PersonName")
		Session("pcAdminFedExWS_CompanyName")=rs("FedExAPI_CompanyName")
		Session("pcAdminFedExWS_Department")=rs("FedExAPI_Department")
		Session("pcAdminFedExWS_PhoneNumber")=rs("FedExAPI_PhoneNumber")
		Session("pcAdminFedExWS_PagerNumber")=rs("FedExAPI_PagerNumber")
		Session("pcAdminFedExWS_FaxNumber")=rs("FedExAPI_FaxNumber")
		Session("pcAdminFedExWS_EmailAddress")=rs("FedExAPI_EmailAddress")
		Session("pcAdminFedExWS_Line1")=rs("FedExAPI_Line1")
		Session("pcAdminFedExWS_Line2")=rs("FedExAPI_Line2")
		Session("pcAdminFedExWS_city")=rs("FedExAPI_city")
		Session("pcAdminFedExWS_StateOrProvinceCode")=rs("FedExAPI_State")
		Session("pcAdminFedExWS_PostalCode")=rs("FedExAPI_PostalCode")
		Session("pcAdminFedExWS_Country")=rs("FedExAPI_Country")
	end if
else
	FedExWSAPI_ID=0
end if
set rs=nothing
call closedb()
'**************************************************************************
' END: Get Fed Ex credentials
'**************************************************************************
%>
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">
	<input type="hidden" name="FedExWSAPI_ID" value="<%=FedExWSAPI_ID%>">
	<input type="hidden" name="changeMode" value="<%=request("changeMode")%>">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2">
			If you have any problems with the registration/subscribe process, contact FedEx&reg; Technical Support at 1-800-810-9073 or via e-mail at <a href="mailto:websupport@fedex.com">websupport@fedex.com</a>.
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<% if intErrCnt>0 then %>
			<tr>
			<td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="4">
				<tr>
					<td width="4%" valign="top"><img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"></td>
					<td width="96%" valign="top" class="message"><font color="#FF9900"><b>
					  <% response.write intErrCnt&" error(s) were located. <ul>"&strErrMsg&"</ul>"%></b></font></td>
				</tr>
		</table>
			</td>
			</tr>
		<% end if %>

		<tr>
			<td colspan="2">
			<span class="pcCPnotes">Click &quot;Continue&quot; below to submit your FedEx&reg; subscription request.</span>
			<input name="FedExWSMode" type="hidden" value="LIVE">
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Account Details</th>
		<tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="23%">FedEx Account Number:</td>
			<td width="77%"><input name="FedExWS_AccountNumber" type="text" value="<%=pcf_FillFormField("FedExWS_AccountNumber", true)%>" size="15" maxlength="25">
			<%pcs_RequiredImageTag "FedExWS_AccountNumber", true%></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Shipping Address</th>
		<tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="23%">Address: </td>
			<td width="77%">
			  <input name="FedExWS_BillingAddress" type="text" value="<%=pcf_FillFormField("FedExWS_BillingAddress", true)%>" size="30" maxlength="100">
			  <%pcs_RequiredImageTag "FedExWS_BillingAddress", true%></td>
		</tr>
		<tr>
			<td width="23%">City: </td>
			<td width="77%">
			  <input name="FedExWS_BillingCity" type="text" value="<%=pcf_FillFormField("FedExWS_BillingCity", true)%>" size="30" maxlength="100">
			  <%pcs_RequiredImageTag "FedExWS_BillingCity", true%></td>
		</tr>
		<tr>
			<td width="23%">State Code:  </td>
			<td width="77%">
				<%
				call opendb()
				query="SELECT stateCode,stateName FROM states WHERE pcCountryCode = 'US' ORDER BY stateName ASC;"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				%>
				<select name="FedExWS_BillingStateOrProvinceCode">
				<option value=""></option>
				<% do while not rs.eof
					strStateCode=rs("stateCode")
					strStateName=rs("stateName")
					%>
					<option value="<%=strStateCode%>" <%=pcf_SelectOption("FedExWS_StateOrProvinceCode",strStateCode)%>><%=strStateName%></option>
					<%rs.movenext
				loop
				set rs=nothing
				call closedb() %>
				</select>
				<%pcs_RequiredImageTag "FedExWS_BillingStateOrProvinceCode", true%>
				</td>
			</tr>
		<tr>
			<td width="23%">Postal Code:  </td>
			<td width="77%">
			  <input name="FedExWS_BillingPostalCode" type="text" value="<%=pcf_FillFormField("FedExWS_BillingPostalCode", true)%>" size="15" maxlength="20">
			  <%pcs_RequiredImageTag "FedExWS_BillingPostalCode", true%></td>
			</tr>
		<tr>
			<td width="23%">Country Code:  </td>
			<td width="77%">
			  US<input type="hidden" name="FedExWS_BillingCountryCode" value="US"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Contact Address</th>
		<tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="23%">First Name: </td>
			<td width="77%">
			  <input name="FedExWS_FirstName" type="text" value="<%=pcf_FillFormField("FedExWS_FirstName", true)%>" size="30" maxlength="100">
			  <%pcs_RequiredImageTag "FedExWS_FirstName", true%></td>
		</tr>
		<tr>
			<td width="23%">Last Name: </td>
			<td width="77%">
			  <input name="FedExWS_LastName" type="text" value="<%=pcf_FillFormField("FedExWS_LastName", true)%>" size="30" maxlength="100">
			  <%pcs_RequiredImageTag "FedExWS_LastName", true%></td>
		</tr>
		<tr>
			<td width="23%">Company Name: </td>
			<td width="77%">
			  <input name="FedExWS_CompanyName" type="text" value="<%=pcf_FillFormField("FedExWS_CompanyName", false)%>" size="30" maxlength="100"></td>
		</tr>
		<!--
		<tr>
			<td width="23%">Department: </td>
			<td width="77%">
			  <input name="FedExWS_Department" type="text" value="<%=pcf_FillFormField("FedExWS_Department", false)%>" size="30" maxlength="100"></td>
		</tr>
		-->
		<% if len(Session("ErrFedExWS_PhoneNumber"))>0 then %>
		<tr>
			<td width="23%"></td>
			<td width="77%">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
			You must enter a valid Phone Number.</td>
		</tr>
		<% end if %>
		<tr>
			<td width="23%">Phone Number:  </td>
			<td width="77%">
			  <input name="FedExWS_PhoneNumber" type="text" value="<%=pcf_FillFormField("FedExWS_PhoneNumber", true)%>" size="16" maxlength="16">
			  <%pcs_RequiredImageTag "FedExWS_PhoneNumber", true%></td>
		</tr>
		<!--
		<tr>
			<td width="23%">Pager Number:  </td>
			<td width="77%">
			  <input name="FedExWS_PagerNumber" type="text" value="<%=pcf_FillFormField("FedExWS_PagerNumber", false)%>" size="16" maxlength="16">
			  </td>
		</tr>
		<tr>
			<td width="23%">Fax Number:  </td>
			<td width="77%">
			  <input name="FedExWS_FaxNumber" type="text" value="<%=pcf_FillFormField("FedExWS_FaxNumber", false)%>" size="16" maxlength="16">
			  </td>
		</tr>
		-->
		<% if len(Session("ErrFedExWS_eMailAddress"))>0 then %>
		<tr>
			<td width="23%"></td>
			<td width="77%">
			<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">
			You must enter a valid Email Address.</td>
		</tr>
		<% end if %>
		<tr>
			<td width="23%">E-mail Address:  </td>
			<td width="77%">
			  <input name="FedExWS_eMailAddress" type="text" value="<%=pcf_FillFormField("FedExWS_eMailAddress", true)%>" size="40" maxlength="250">
			  <%pcs_RequiredImageTag "FedExWS_eMailAddress", true%></td>
			</tr>
		<tr>
			<td width="23%">Address:  </td>
			<td width="77%">
			  <input name="FedExWS_Line1" type="text" value="<%=pcf_FillFormField("FedExWS_Line1", true)%>" size="40" maxlength="250">
			  <%pcs_RequiredImageTag "FedExWS_Line1", true%></td>
			</tr>
		<tr>
			<td width="23%">&nbsp;</td>
			<td width="77%">
			<input name="FedExWS_Line2" type="text" value="<%=pcf_FillFormField("FedExWS_Line2", false)%>" size="40" maxlength="250"></td>
			</tr>
		<tr>
			<td width="23%">City:  </td>
			<td width="77%">
			  <input name="FedExWS_City" type="text" value="<%=pcf_FillFormField("FedExWS_City", true)%>" size="30" maxlength="250">
			  <%pcs_RequiredImageTag "FedExWS_City", true%>
			</td>
		</tr>
		<tr>
			<td width="23%">State Code:  </td>
			<td width="77%">
				<%
				call opendb()
				query="SELECT stateCode,stateName FROM states WHERE pcCountryCode = 'US' ORDER BY stateName ASC;"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				%>
				<select name="FedExWS_StateOrProvinceCode">
				<option value=""></option>
				<% do while not rs.eof
					strStateCode=rs("stateCode")
					strStateName=rs("stateName")
					%>
					<option value="<%=strStateCode%>" <%=pcf_SelectOption("FedExWS_StateOrProvinceCode",strStateCode)%>><%=strStateName%></option>
					<%rs.movenext
				loop
				set rs=nothing
				call closedb() %>
				</select>
				<%pcs_RequiredImageTag "FedExWS_StateOrProvinceCode", true%>
				</td>
			</tr>
		<tr>
			<td width="23%">Postal Code:  </td>
			<td width="77%">
			  <input name="FedExWS_PostalCode" type="text" value="<%=pcf_FillFormField("FedExWS_PostalCode", true)%>" size="15" maxlength="20">
			  <%pcs_RequiredImageTag "FedExWS_PostalCode", true%></td>
			</tr>
		<tr>
			<td width="23%">Country Code:  </td>
			<td width="77%">
			  US<input type="hidden" name="FedExWS_CountryCode" value="US"></td>
			</tr>
		<tr>
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2">
			<input type="submit" name="Submit" value="Continue" class="submit2">
			&nbsp;
			<input type="button" name="back" value="Back" onClick="JavaScript:history.back();">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->