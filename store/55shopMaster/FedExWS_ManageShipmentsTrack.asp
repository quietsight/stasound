<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard - Track Packages" %>
<% Section="mngAcc" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->

<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<%
Dim query, rs, conntemp
Dim iPageCurrent, varFlagIncomplete, uery, strORD, pcv_intOrderID
Dim pcv_strMethodName, pcv_strMethodReply, CustomerTransactionIdentifier, pcv_strAccountNumber, pcv_strMeterNumber, pcv_strCarrierCode
Dim pcv_strValue, pcv_strType, pcv_strTrackingNumberUniqueIdentifier, pcv_strShipDateRangeBegin, pcv_strShipDateRangeEnd, pcv_strShipmentAccountNumber
Dim pcv_strDestinationCountryCode, pcv_strDestinationPostalCode, pcv_strLanguageCode, pcv_strLocaleCode, pcv_strDetailScans, pcv_strPagingToken
Dim fedex_postdata, objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEX_result, FEDEX_URL, pcv_strErrorMsg, pcv_strAction



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// GET ORDER ID
pcv_strOrderID=Request("id")
pcv_strSessionOrderID=Session("pcAdminOrderID")
if pcv_strSessionOrderID="" OR len(pcv_strOrderID)>0 then
	pcv_intOrderID=pcv_strOrderID
	Session("pcAdminOrderID")=pcv_intOrderID
else
	pcv_intOrderID=pcv_strSessionOrderID
end if

'// PAGE NAME
pcPageName="FedExWS_ManageShipmentsTrack.asp"
ErrPageName="FedExWS_ManageShipmentsResults.asp"

'// ACTION
pcv_strAction = request("Action")

'// OPEN DATABASE
call openDb()

'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExWSClass

'// REQUEST ARRAY OF PACKAGES TO TRACK "PackageInfo_ID"
if pcv_strAction="batch" then
	pcv_strTrackingNumbers=""
	Count=request("count")
	Dim k
	For k=1 to Count
		if (request("check" & k)<>"") then
			pcv_strTrackingNumbers=pcv_strTrackingNumbers & request("check" & k) & ","
		end if
	Next
	xStringLength = len(pcv_strTrackingNumbers)
	if xStringLength>0 then
		pcv_strTrackingNumbers = left(pcv_strTrackingNumbers,(xStringLength-1))
	end if
else
	pcv_strTrackingNumbers = Request("PackageInfo_ID")
end if

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


'// CREATE ARRAY OF PACKAGES
Dim xIdOptCounter, pcArrayTrackingNumbers
if NOT instr(pcv_strTrackingNumbers,",") then
	pcv_strTrackingNumbers = pcv_strTrackingNumbers&","
end if
pcArrayTrackingNumbers = split(pcv_strTrackingNumbers,",")

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Tracking FedEx<sup>&reg;</sup> Shipments for Order Number <%=(scpre+int(Session("pcAdminOrderID")))%></th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">
			<span class="pcCPnotes">
			<strong>ATTENTION SHIPPERS:</strong> If your package has not yet been scanned by FedEx then the information on this page may not be accurate.
			FedEx sometimes reuses Tracking Numbers, so a Tracking Number may show data from a previous shipment until its scanned again.
			</span>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
</table>

<table class="pcCPcontent">

	<form name="form1" action="<%=pcPageName%>" method="post" class="pcForms">
	<input name="PackageInfo_ID" type="hidden" value="<%=pcv_strTrackingNumbers%>">
	<input name="id" type="hidden" value="<%=pcv_intOrderID%>">
		<tr>
			<td>
			<%
			'***************************************************************************
			' START LOOP THROUGH TRACKING
			'***************************************************************************


			for xIdOptCounter = 0 to Ubound(pcArrayTrackingNumbers)-1


				pcv_strTmpNumber = pcArrayTrackingNumbers(xIdOptCounter)

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Set Required Data
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' SELECT DATA SET
				' >>> Tables: pcPackageInfo
				query = 		"SELECT pcPackageInfo.pcPackageInfo_ID, pcPackageInfo.pcPackageInfo_TrackingNumber, pcPackageInfo.pcPackageInfo_ShipMethod, pcPackageInfo.pcPackageInfo_FDXCarrierCode "
				query = query & "FROM pcPackageInfo "
				query = query & "WHERE pcPackageInfo.pcPackageInfo_ID=" & pcv_strTmpNumber &" "

				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)

				if NOT rs.eof then
					pcv_strValue=rs("pcPackageInfo_TrackingNumber")
					pcv_strType= "" 'rs("")
					pcv_strShipDateRangeBegin= "" 'rs("")
					pcv_strShipDateRangeEnd= "" 'rs("")
					pcv_strDestinationCountryCode= "" 'rs("")
					pcv_strDestinationPostalCode= "" 'rs("")
					pcv_strLanguageCode= "" 'rs("")
					pcv_strLocaleCode= "" 'rs("")
					pcv_strDetailScans= "" 'rs("")
					pcv_strPagingToken= "" 'rs("")
					pcv_strTrackingNumberUniqueIdentifier= "" 'rs("")
					pcv_strShipMethod=rs("pcPackageInfo_ShipMethod")
					pcv_strCarrierCode=rs("pcPackageInfo_FDXCarrierCode")
				end if
				set rs=nothing

				pcv_strShipmentAccountNumber=pcv_strAccountNumber '// Owner's Account Number

				fedex_postdataWS=""
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: Build Transaction
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				if instr(pcv_strShipMethod, "FIRST_OVERNIGHT") then
					pcv_strCarrierCode = "FDXE"
				end if
				if instr(pcv_strShipMethod, "PRIORITY_OVERNIGHT") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "STANDARD_OVERNIGHT") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "FEDEX_2_DAY") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "FEDEX_EXPRESS_SAVER") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "FEDEX_GROUND") then
					pcv_strCarrierCode ="FDXG"
				end if
				if instr(pcv_strShipMethod, "GROUND_HOME_DELIVERY") then
					pcv_strCarrierCode ="FDXG"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_FIRST") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_PRIORITY") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_ECONOMY") then
					pcv_strCarrierCode ="FDXE"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_PRIORITY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "INTERNATIONAL_ECONOMY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "FEDEX_1_DAY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "FEDEX_2_DAY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "FEDEX_3_DAY_FREIGHT") then
					pcv_strCarrierCode ="FXFR"
				end if
				if instr(pcv_strShipMethod, "SMART_POST") then
					pcv_strCarrierCode = "FXSP"
				end if
							
				NameOfMethod = "TrackRequest"
				fedex_postdataWS=""
				fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v6=""http://fedex.com/ws/track/v6"">"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v6:"&NameOfMethod&">"&vbcrlf

				fedex_postdataWS=fedex_postdataWS&"<v6:WebAuthenticationDetail>"&vbcrlf
				If CSPTurnOn = 1 Then
					fedex_postdataWS=fedex_postdataWS&"<v6:CspCredential>"&vbcrlf
						fedex_postdataWS=fedex_postdataWS&"<v6:Key>CPTi545ATGa1CD89</v6:Key>"&vbcrlf
						fedex_postdataWS=fedex_postdataWS&"<v6:Password>8BB07q2XIIOFyNJeJQHMLv094</v6:Password>"&vbcrlf
					fedex_postdataWS=fedex_postdataWS&"</v6:CspCredential>"&vbcrlf
				End If
				fedex_postdataWS=fedex_postdataWS&"<v6:UserCredential>"&vbcrlf
					fedex_postdataWS=fedex_postdataWS&"<v6:Key>" & FedExkey & "</v6:Key>"&vbcrlf
					fedex_postdataWS=fedex_postdataWS&"<v6:Password>" & FedExPassword & "</v6:Password>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"</v6:UserCredential>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"</v6:WebAuthenticationDetail>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<v6:ClientDetail>"&vbcrlf
					fedex_postdataWS=fedex_postdataWS&"<v6:AccountNumber>"&FedExAccountNumber&"</v6:AccountNumber>"&vbcrlf
					fedex_postdataWS=fedex_postdataWS&"<v6:MeterNumber>"&FedExMeterNumber&"</v6:MeterNumber>"&vbcrlf
					fedex_postdataWS=fedex_postdataWS&"<v6:ClientProductId>EIPC</v6:ClientProductId>"&vbcrlf
					fedex_postdataWS=fedex_postdataWS&"<v6:ClientProductVersion>3424</v6:ClientProductVersion>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"</v6:ClientDetail>"&vbcrlf

				'--------------------
				'// TransactionDetail
				'--------------------
				objFedExClass.WriteParent "TransactionDetail", "6", ""
					objFedExClass.AddNewNode "CustomerTransactionId", "6", "Track Ground Shipment"
					objFedExClass.WriteParent "Localization", "6", ""
						objFedExClass.AddNewNode "LanguageCode", "6", "EN"
					objFedExClass.WriteParent "Localization", "6", "/"
				objFedExClass.WriteParent "TransactionDetail", "6", "/"

				'--------------------
				'// Version
				'--------------------
				objFedExClass.WriteParent "Version", "6", ""
					objFedExClass.AddNewNode "ServiceId", "6", "trck"
					objFedExClass.AddNewNode "Major", "6", "6"
					objFedExClass.AddNewNode "Intermediate", "6", "0"
					objFedExClass.AddNewNode "Minor", "6", "0"
				objFedExClass.WriteParent "Version", "6", "/"

				objFedExClass.AddNewNode "CarrierCode", "6", pcv_strCarrierCode
				objFedExClass.WriteParent "PackageIdentifier", "6", ""
					objFedExClass.AddNewNode "Value", "6", pcv_strValue
					objFedExClass.AddNewNode "Type", "6", "TRACKING_NUMBER_OR_DOORTAG"
							objFedExClass.AddNewNode "TrackingNumber", "6", pcv_strTrackingNumber
				objFedExClass.WriteParent "PackageIdentifier", "6", "/"

				objFedExClass.EndXMLTransaction NameOfMethod, "6"

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END: Build Transaction
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

				'// Print out our newly formed request xml
				'response.write fedex_postdataWS&"<hr>"
				'response.end

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Send Our Transaction.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'call objFedExClass.SendXMLRequest(fedex_postdata, pcv_strEnvironment)
				Set srvFEDEXWSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
				Set objOutputXMLDocWS = Server.CreateObject("Microsoft.XMLDOM")
				Set objFedExStream = Server.CreateObject("ADODB.Stream")
				Set objFEDEXXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
				objFEDEXXmlDoc.async = False
				objFEDEXXmlDoc.validateOnParse = False
				if err.number>0 then
					err.clear
				end if

				srvFEDEXWSXmlHttp.open "POST", FedExWSURL&"/track", false


				srvFEDEXWSXmlHttp.send(fedex_postdataWS)
				FEDEXWS_result = srvFEDEXWSXmlHttp.responseText
				'// Print out our response
						'response.write FEDEXWS_result
				'response.end

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Load Our Response.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				call objFedExClass.LoadXMLResults(FEDEXWS_result)

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Check for errors from FedEx.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// master package error, no processing done
		pcv_strErrorMsg = Cstr("")

		pcv_strErrorMsg = objFedExClass.ReadResponseNode("//v6:TrackReply", "v6:Notifications/v6:Severity")

		if pcv_strErrorMsg="SUCCESS" OR pcv_strErrorMsg="NOTE" then
			pcv_strErrorMsg = Cstr("")
		else
			pcv_strErrorMsg = objFedExClass.ReadResponseNode("//v6:TrackReply", "v6:Notifications/v6:Message")
		end if

		if pcv_strErrorMsg&""="" then
			pcv_strErrorMsg = objFedExClass.ReadResponseNode("//soapenv:Fault", "faultstring")
		end if

		If pcv_strErrorMsg&"" <> "" Then
			response.redirect ErrPageName&"?msg="&pcv_strErrorMsg
		End IF

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Redirect with a Message OR complete some task.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if NOT len(pcv_strErrorMsg)>0 then
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Set Our Response Data to Local.
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Available Methods will search unlimited levels of Nodes by separating nodes with a "/".
			' 1) ReadResponseParent
			' 2) ReadResponseNode

			'///////////////////////////////////////////////////////////////////////////////////////////////////
			' Note: these are the primary values, but there are many more possible return values
			'///////////////////////////////////////////////////////////////////////////////////////////////////
					pcv_strTrackingNumber = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:TrackingNumber")
					pcv_strService = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:ServiceType")
					pcv_strShipTimeStamp = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:ShipTimestamp")
					pcv_strActualDeliveryTimeStamp = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:ActualDeliveryTimestamp")
					pcv_strStatusDescription = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:StatusDescription")
					pcv_strWeight = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:PackageWeight")
					pcv_strEstDeliveryTimestamp = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:EstimatedDeliveryTimestamp")

					pcv_strEventDate = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:EventType")
					pcv_strEventTime = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:Timestamp")
					pcv_strEventType = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:EventType")
					pcv_strEventDescription = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:EventDescription")
					pcv_strEventStatusExceptionCode = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:StatusExceptionCode")
					pcv_strEventStatusExceptionDescription = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:StatusExceptionDescription")
					pcv_strEventAddressCity = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:Address/v6:City")
					pcv_strEventAddressStateOrProvinceCode = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:Address/v6:StateOrProvinceCode")
					pcv_strEventAddressPostalCode = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:Address/v6:PostalCode")
					pcv_strEventAddressCountryCode = objFedExClass.ReadResponseNode("//v6:TrackDetails", "v6:Events/v6:Address/v6:CountryCode")
					'//
					tmpShipTimestamp = pcv_strShipTimeStamp
					if instr(tmpShipTimestamp,"T") then
						arrShipTimestamp = split(tmpShipTimestamp, "T")
						tmpShipTime = arrShipTimestamp(1)
						arrShipTimeFormat = split(tmpShipTime,":")
						tmpShipTimeHour = Cint(arrShipTimeFormat(0))
						tmpShipTimeMinutes = arrShipTimeFormat(1)
						tmpShipTimeSeconds = arrShipTimeFormat(2)
						'//Format hour and check for AM/PM
						if tmpShipTimeHour < 12 then
							tmpShipAMPM = "AM"
							tmpShipHour = Cint(tmpShipTimeHour)
						else
							tmpShipAMPM = "PM"
							tmpShipHour = Cint(tmpShipTimeHour) - Cint(12)
						end if
						tmpShipDate = arrShipTimestamp(0)
						arrShipDate = split(tmpShipDate,"-")
						tmpShipDay = arrShipDate(2)
						tmpShipMonth = arrShipDate(1)
						select case tmpShipMonth
							case "01"
								tmpShipMonth = "January"
							case "02"
								tmpShipMonth = "February"
							case "03"
								tmpShipMonth = "March"
							case "04"
								tmpShipMonth = "April"
							case "05"
								tmpShipMonth = "May"
							case "06"
								tmpShipMonth = "June"
							case "07"
								tmpShipMonth = "July"
							case "08"
								tmpShipMonth = "August"
							case "09"
								tmpShipMonth = "September"
							case "10"
								tmpShipMonth = "October"
							case "11"
								tmpShipMonth = "November"
							case "12"
								tmpShipMonth = "December"
						end select
						tmpShipYear = arrShipDate(0)
						FedExShipTimeDateStampF = tmpShipMonth&", "&tmpShipDay&" "&tmpShipYear&" "&tmpShipTimeHour&":"&tmpShipTimeMinutes&" "&tmpShipAMPM
					else
						FedExShipTimeDateStampF = "N/A"
					end if
					
					'//
					tmpActualDeliveryTimestamp = pcv_strActualDeliveryTimestamp
					if instr(tmpActualDeliveryTimestamp,"T") then
						arrActualDeliveryTimestamp = split(tmpActualDeliveryTimestamp, "T")
						tmpActualDeliveryTime = arrActualDeliveryTimestamp(1)
						arrActualTimeFormat = split(tmpActualDeliveryTime,":")
						tmpActualTimeHour = Cint(arrActualTimeFormat(0))
						tmpActualTimeMinutes = arrActualTimeFormat(1)
						tmpActualTimeSeconds = arrActualTimeFormat(2)
						'//Format hour and check for AM/PM
						if tmpActualTimeHour < 12 then
							tmpActualAMPM = "AM"
							tmpActualHour = Cint(tmpActualTimeHour)
						else
							tmpActualAMPM = "PM"
							tmpActualHour = Cint(tmpActualTimeHour) - Cint(12)
						end if
						tmpActualDeliveryDate = arrActualDeliveryTimestamp(0)
						arrActualDeliveryDate = split(tmpActualDeliveryDate,"-")
						tmpActualDeliveryDay = arrActualDeliveryDate(2)
						tmpActualDeliveryMonth = arrActualDeliveryDate(1)
						select case tmpActualDeliveryMonth
							case "01"
								tmpActualDeliveryMonth = "January"
							case "02"
								tmpActualDeliveryMonth = "February"
							case "03"
								tmpActualDeliveryMonth = "March"
							case "04"
								tmpActualDeliveryMonth = "April"
							case "05"
								tmpActualDeliveryMonth = "May"
							case "06"
								tmpActualDeliveryMonth = "June"
							case "07"
								tmpActualDeliveryMonth = "July"
							case "08"
								tmpActualDeliveryMonth = "August"
							case "09"
								tmpActualDeliveryMonth = "September"
							case "10"
								tmpActualDeliveryMonth = "October"
							case "11"
								tmpActualDeliveryMonth = "November"
							case "12"
								tmpActualDeliveryMonth = "December"
						end select
						tmpActualDeliveryYear = arrActualDeliveryDate(0)
						
						FedExActualTimeDateStampF = tmpActualDeliveryMonth&", "&tmpActualDeliveryDay&" "&tmpActualDeliveryYear&" "&tmpActualTimeHour&":"&tmpActualTimeMinutes&" "&tmpActualAMPM
					else
						FedExActualTimeDateStampF = "N/A"
					end if
					
					'//
					tmpEstDeliveryTimestamp = pcv_strEstDeliveryTimestamp
					if instr(tmpEstDeliveryTimestamp,"T") then
						arrEstDeliveryTimestamp = split(tmpEstDeliveryTimestamp, "T")
						tmpEstDeliveryTime = arrEstDeliveryTimestamp(1)
						arrEstTimeFormat = split(tmpEstDeliveryTime,":")
						tmpEstTimeHour = Cint(arrEstTimeFormat(0))
						tmpEstTimeMinutes = arrEstTimeFormat(1)
						tmpEstTimeSeconds = arrEstTimeFormat(2)
						'//Format hour and check for AM/PM
						if tmpEstTimeHour < 12 then
							tmpEstAMPM = "AM"
							tmpEstHour = Cint(tmpEstTimeHour)
						else
							tmpEstAMPM = "PM"
							tmpEstHour = Cint(tmpEstTimeHour) - Cint(12)
						end if
						tmpEstDeliveryDate = arrEstDeliveryTimestamp(0)
						arrEstDeliveryDate = split(tmpEstDeliveryDate,"-")
								tmpEstDeliveryDay = arrEstDeliveryDate(2)
								tmpEstDeliveryMonth = arrEstDeliveryDate(1)
								select case tmpEstDeliveryMonth
									case "01"
										tmpEstDeliveryMonth = "January"
									case "02"
										tmpEstDeliveryMonth = "February"
									case "03"
										tmpEstDeliveryMonth = "March"
									case "04"
										tmpEstDeliveryMonth = "April"
									case "05"
										tmpEstDeliveryMonth = "May"
									case "06"
										tmpEstDeliveryMonth = "June"
									case "07"
										tmpEstDeliveryMonth = "July"
									case "08"
										tmpEstDeliveryMonth = "August"
									case "09"
										tmpEstDeliveryMonth = "September"
									case "10"
										tmpEstDeliveryMonth = "October"
									case "11"
										tmpEstDeliveryMonth = "November"
							case "12"
								tmpEstDeliveryMonth = "December"
						end select
						tmpEstDeliveryYear = arrEstDeliveryDate(0)
					
						FedExEstTimeDateStampF = tmpEstDeliveryMonth&", "&tmpEstDeliveryDay&" "&tmpEstDeliveryYear&" "&tmpEstTimeHour&":"&tmpEstTimeMinutes&" "&tmpEstAMPM
					else
						FedExEstTimeDateStampF = "N/A"
					end if
					
					select case pcv_strService
						case "PRIORITY_OVERNIGHT"
							pcv_strService="FedEx Priority Overnight<sup>&reg;</sup>"
						case "STANDARD_OVERNIGHT"
							pcv_strService="FedEx Standard Overnight<sup>&reg;</sup>"
						case "FIRST_OVERNIGHT"
							pcv_strService="FedEx First Overnight<sup>&reg;</sup>"
						case "FEDEX_2_DAY"
							pcv_strService="FedEx 2Day<sup>&reg;</sup>"
						case "FEDEX_EXPRESS_SAVER"
							pcv_strService="FedEx Express Saver<sup>&reg;</sup>"
						case "INTERNATIONAL_PRIORITY"
							pcv_strService="FedEx International Priority<sup>&reg;</sup>"
						case "INTERNATIONAL_ECONOMY"
							pcv_strService="FedEx International Economy<sup>&reg;</sup>"
						case "INTERNATIONAL_FIRST"
							pcv_strService="FedEx International First<sup>&reg;</sup>"
						case "FEDEX_1_DAY_FREIGHT"
							pcv_strService="FedEx 1Day<sup>&reg;</sup> Freight"
						case "FEDEX_2_DAY_FREIGHT"
							pcv_strService="FedEx 2Day<sup>&reg;</sup> Freight"
						case "FEDEX_3_DAY_FREIGHT"
							pcv_strService="FedEx 3Day<sup>&reg;</sup> Freight"
						case "FEDEX_GROUND"
							pcv_strService="FedEx Ground<sup>&reg;</sup>"
						case "GROUND_HOME_DELIVERY"
							pcv_strService="FedEx Home Delivery<sup>&reg;</sup>"
						case "INTERNATIONAL_PRIORITY_FREIGHT"
							pcv_strService="FedEx International Priority<sup>&reg;</sup> Freight"
						case "INTERNATIONAL_ECONOMY_FREIGHT"
							pcv_strService="FedEx International Economy<sup>&reg;</sup> Freight"
						case "SMART_POST"
							pcv_strService="FedEx SmartPost<sup>&reg;</sup>"
					end select
					%>
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Tracking Number <%=pcv_strTrackingNumber%> Summary</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
								  <tr>
									<td colspan="3">

										<table width="100%" border="0" cellpadding="0" cellspacing="0">
											<tr>
												<td width="19%" align="right">Tracking Number:</td>
												<td width="32%" align="left"><%=pcv_strTrackingNumber%></td>
												<td align="right">Service Type:</td>
											<td align="left"><%=pcv_strService%></td>
											</tr>
											<tr>
												<td align="right">Signed For By:</td>
												<td align="left">
												<% if pcv_strSignedForBy<>"" then %>
													<%=pcv_strSignedForBy%>
												<% else %>
													N/A
												<% end if %>
												</td>
												<td align="right">Destination:</td>
												<td align="left">
												<% if pcv_strEventAddressCity<>"" then %>
													<%=pcv_strEventAddressCity%>, <%=pcv_strEventAddressStateOrProvinceCode%>
												<% else %>
													N/A
												<% end if %>												</td>
											</tr>
											<tr>
												<td align="right">Ship Date:</td>
												<td align="left"><%=FedExShipTimeDateStampF%></td>
												<td align="right">Packaging:</td>
												<td align="left">
												<% if pcv_strPackagingDescription<>"" then %>
													<%=pcv_strPackagingDescription%>
												<% else %>
													N/A
												<% end if %>
												</td>
											</tr>
											<tr>
												<td align="right" nowrap="nowrap">Delivery Date/Time:</td>
												<td align="left">
												<% = FedExActualTimeDateStampF %>
												</td>
												<td align="right" nowrap="nowrap">Estimated Delivery Date:</td>
												<td align="left">
													<%=FedExEstTimeDateStampF%>
												</td>
											</tr>
										  <tr>
											<td align="right">Status:</td>
											<td align="left"><%=pcv_strStatusDescription%></td>

											<td width="18%" align="right">&nbsp;</td>
												<td width="31%" align="left">&nbsp;</td>
											</tr>
										</table>

									</td>
								  </tr>
								  <tr>
									<th width="44%"><strong>Date/ Time : Location</strong></th>
									<th width="16%"><strong>Scan Activity </strong></th>
									<th width="40%"><strong>Comments </strong></th>
								  </tr>
								<%

								'// Generate/ Trim Event Type
								arrayFedExEventType = objFedExClass.ReadResponseasArray("//v6:Events", "v6:EventType")
								'arrayFedExEventType = objFedExClass.pcf_FedExTrimArray(arrayFedExEventType)
								'// Generate/ Trim Event Description
								arrayFedExEventDescription = objFedExClass.ReadResponseasArray("//v6:Events", "v6:StatusExceptionDescription")
								'arrayFedExEventDescription = objFedExClass.pcf_FedExTrimArray(arrayFedExEventDescription)
								'// Generate/ Trim Event Status Exception Description
								arrayFedExEventStatusExcDes = objFedExClass.ReadResponseasArray("//v6:Events", "v6:StatusExceptionDescription")
								'arrayFedExEventStatusExcDes = objFedExClass.pcf_FedExTrimArray(arrayFedExEventStatusExcDes)
								'// Generate/ Trim Event Date
								arrayFedExTimestamp = objFedExClass.ReadResponseasArray("//v6:Events", "v6:Timestamp")
								arrayFedExEventDescription2 = objFedExClass.ReadResponseasArray("//v6:Events", "v6:EventDescription")
								arrayFedExCity = objFedExClass.ReadResponseasArray("//v6:Events", "v6:Address/v6:City")
								arrayFedExStateOrProvinceCode = objFedExClass.ReadResponseasArray("//v6:Events", "v6:Address/v6:StateOrProvinceCode")
								arrayFedExPostalCode = objFedExClass.ReadResponseasArray("//v6:Events", "v6:Address/v6:PostalCode")
								arrayFedExCountryCode = objFedExClass.ReadResponseasArray("//v6:Events", "v6:Address/v6:CountryCode")
								arrayFedExArrivalLocation = objFedExClass.ReadResponseasArray("//v6:Events", "v6:ArrivalLocation")

								arrayFedExEventType = split(arrayFedExEventType, ",")
								arrayFedExEventDescription = split(arrayFedExEventDescription, ",")
								arrayFedExEventStatusExcDes = split(arrayFedExEventStatusExcDes, ",")

								arrayFedExTimestamp = split(arrayFedExTimestamp, ",")
								if arrayFedExEventDescription&""="" Then
									arrayFedExEventDescription = split(arrayFedExEventDescription2, ",")
								end if
								arrayFedExCity = split(arrayFedExCity, ",")
								arrayFedExStateOrProvinceCode = split(arrayFedExStateOrProvinceCode, ",")
								arrayFedExPostalCode = split(arrayFedExPostalCode, ",")
								arrayFedExCountryCode = split(arrayFedExCountryCode, ",")
								arrayFedExArrivalLocation = split(arrayFedExArrivalLocation, ",")


								for bIdOptCounter = 0 to Ubound(arrayFedExEventType)-1

								tmpDeliveryTimestamp = arrayFedExTimestamp(bIdOptCounter)
								arrDeliveryTimestamp = split(tmpDeliveryTimestamp, "T")
								tmpDeliveryTime = arrDeliveryTimestamp(1)
								arrTimeFormat = split(tmpDeliveryTime,":")
								tmpTimeHour = Cint(arrTimeFormat(0))
								tmpTimeMinutes = arrTimeFormat(1)
								tmpTimeSeconds = arrTimeFormat(2)
								'//Format hour and check for AM/PM
								if tmpTimeHour < 12 then
									tmpAMPM = "AM"
									tmpHour = Cint(tmpTimeHour)
								else
									tmpAMPM = "PM"
									tmpHour = Cint(tmpTimeHour) - Cint(12)
								end if
								tmpDeliveryDate = arrDeliveryTimestamp(0)
								arrDeliveryDate = split(tmpDeliveryDate,"-")
								tmpDeliveryDay = arrDeliveryDate(2)
								tmpDeliveryMonth = arrDeliveryDate(1)
								select case tmpDeliveryMonth
									case "01"
										tmpDeliveryMonth = "January"
									case "02"
										tmpDeliveryMonth = "February"
									case "03"
										tmpDeliveryMonth = "March"
									case "04"
										tmpDeliveryMonth = "April"
									case "05"
										tmpDeliveryMonth = "May"
									case "06"
										tmpDeliveryMonth = "June"
									case "07"
										tmpDeliveryMonth = "July"
									case "08"
										tmpDeliveryMonth = "August"
									case "09"
										tmpDeliveryMonth = "September"
									case "10"
										tmpDeliveryMonth = "October"
									case "11"
										tmpDeliveryMonth = "November"
									case "12"
										tmpDeliveryMonth = "December"
								end select
								tmpDeliveryYear = arrDeliveryDate(0)

								FedExEventTimeDateStampF = tmpDeliveryMonth&", "&tmpDeliveryDay&" "&tmpDeliveryYear&" "&tmpTimeHour&":"&tmpTimeMinutes&" "&tmpAMPM
									'2012-07-11T00:00:00 %>
								  <tr>
									<td>
									<%
									select case arrayFedExArrivalLocation(bIdOptCounter)
										case "AIRPORT"
											FedExLocation = "Airport"
										case "CUSTOMER"
											FedExLocation = "Customer"
										case "CUSTOMS_BROKER"
											FedExLocation = "Customs Broker"
										case "DELIVERY_LOCATION"
											FedExLocation = "Delivery Location"
										case "DESTINATION_AIRPORT"
											FedExLocation = "Destination Airport"
										case "DESTINATION_FEDEX_FACILITY"
											FedExLocation = "Destination FedEx Facility"
										case "DROP_BOX"
											FedExLocation = "Drop Box"
										case "ENROUTE"
											FedExLocation = "Enroute"
										case "FEDEX_FACILITY"
											FedExLocation = "FedEx Facility"
										case "FEDEX_OFFICE_LOCATION"
											FedExLocation = "FedEx Office Location"
										case "INTERLINE_CARRIER"
											FedExLocation = "Interline Carrier"
										case "NON_FEDEX_FACILITY"
											FedExLocation = "Non-FedEx Facility"
										case "ORIGIN_AIRPORT"
											FedExLocation = "Origin Airport"
										case "ORIGIN_FEDEX_FACILITY"
											FedExLocation = "Origin FedEx Facility"
										case "PICKUP_LOCATION"
											FedExLocation = "Pickup Location"
										case "PLANE"
											FedExLocation = "Plane"
										case "PORT_OF_ENTRY"
											FedExLocation = "Port of Entry"
										case "SORT_FACILITY"
											FedExLocation = "Sort Facility"
										case "TURNPOINT"
											FedExLocation = "Turnpoint"
										case "VEHICLE"
											FedExLocation = "Vehicle"
										case else
											FedExLocation = "Unknown"
									end Select


									if FedExEventTimeDateStampF<>"" then
										%>
										<%=FedExEventTimeDateStampF%>
							    <% else %>
										N/A
									<% end if %>
									: <%=FedExLocation%></td>
									<td nowrap><%=arrayFedExEventDescription(bIdOptCounter)%>
									</td>
									<td><%=arrayFedExEventStatusExcDes(bIdOptCounter)%></td>
								  </tr>
								</table>
								<%
								next
								%>
							</td>
						</tr>
					</table>
				<%
				end if
				'set objFedExClass = nothing

			next
			'***************************************************************************
			' END LOOP THROUGH TRACKING
			'***************************************************************************
			%>
			</td>
		</tr>
	</form>
</table>
<%
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>
<!--#include file="AdminFooter.asp"-->