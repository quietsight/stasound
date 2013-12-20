<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Center for UPS" %>
<% Section="mngAcc" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/pcUPSClass.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/pcShipTestModes.asp" -->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<% 
Dim query, rs, conntemp
Dim iPageCurrent, varFlagIncomplete, uery, strORD, pcv_intOrderID
Dim pcv_strMethodName, pcv_strMethodReply, CustomerTransactionIdentifier, pcv_strAccountNumber, pcv_strMeterNumber, pcv_strCarrierCode
Dim pcv_strValue, pcv_strType, pcv_strTrackingNumberUniqueIdentifier, pcv_strShipDateRangeBegin, pcv_strShipDateRangeEnd, pcv_strShipmentAccountNumber
Dim pcv_strDestinationCountryCode, pcv_strDestinationPostalCode, pcv_strLanguageCode, pcv_strLocaleCode, pcv_strDetailScans, pcv_strPagingToken
Dim UPS_postdata, objUPSClass, objOutputXMLDoc, srvUPSXmlHttp, UPS_result, UPS_URL, pcv_strErrorMsg, pcv_strAction

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
pcPageName="UPS_ManageShipmentsTrack.asp"
ErrPageName="UPS_ManageShipmentsResults.asp"

'// ACTION
pcv_strAction = request("Action")

'// OPEN DATABASE
call openDb()

'// SET THE UPS OBJECT
set objUPSClass = New pcUPSClass

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
		
'// UPS CREDENTIALS
query="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then
	ups_active=rs("active")
	ups_userid=trim(rs("userID"))
	ups_password=trim(rs("password"))
	ups_license_key=trim(rs("AccessLicense"))
end if

set rs=nothing

'// CREATE ARRAY OF PACKAGES
Dim xIdOptCounter, pcArrayTrackingNumbers
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
		<th colspan="2">UPS OnLine&reg; Tools - Tracking Shipments for Order Number <%=(scpre+int(Session("pcAdminOrderID")))%></th>
	</tr>
    <% if UPS_TESTMODE="1" then %>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td colspan="2" valign="top">
            <div class="pcCPmessage">UPS Shipping Wizard is currently running in Test Mode<br>
Tracking is not available in Test Mode </div>
           </td>
        </tr>
	<% end if %>
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
				<% '***************************************************************************
				' START LOOP THROUGH TRACKING
				'***************************************************************************
				for xIdOptCounter = 0 to Ubound(pcArrayTrackingNumbers)
			
					pcv_strTmpNumber = pcArrayTrackingNumbers(xIdOptCounter)
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Set Required Data
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' SELECT DATA SET
					' >>> Tables: pcPackageInfo
					query = 		"SELECT pcPackageInfo.pcPackageInfo_ID, pcPackageInfo.pcPackageInfo_TrackingNumber "
					query = query & "FROM pcPackageInfo "
					query = query & "WHERE pcPackageInfo.pcPackageInfo_ID=" & pcv_strTmpNumber &" "	
				
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=conntemp.execute(query)
					
					if NOT rs.eof then		
						pcv_strValue=rs("pcPackageInfo_TrackingNumber")
					end if
					set rs=nothing
				
					UPS_postdata=""
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' START: Build Transaction
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					pcv_strTrackingNumber=pcv_strValue
	
					objUPSClass.NewXMLTransaction ups_license_key, ups_userid, ups_password
					objUPSClass.NewXMLShipmentTrackRequest "Tracking", pcv_strTrackingNumber
					
					'//Clear illegal ampersand characters from XML
					UPS_postdata=replace(UPS_postdata, "&", "and")
					UPS_postdata=replace(UPS_postdata, "andamp;", "and")

					'// Print out our newly formed request xml
					'response.write UPS_postdata
					'response.end
				
					'get URL to post to
					if UPS_TESTMODE="1" then
						UPS_URL="https://wwwcie.ups.com/ups.app/xml/Track"
					else
					UPS_URL="https://www.ups.com/ups.app/xml/Track"
					end if
					call objUPSClass.SendXMLRequest(UPS_postdata, UPS_URL)

					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Send Our Transaction.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
					call objUPSClass.SendXMLRequest(UPS_postdata, UPS_URL)
					'// Print out our response
					'response.write UPS_result&"<HR>"
					'response.end
		
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Load Our Response.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					call objUPSClass.LoadXMLResults(UPS_result)
		
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' Check for errors from UPS.
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~				
					call objUPSClass.XMLResponseVerify(ErrPageName)
				
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
						
						'// HEADER
						pcv_strResponseStatusCode = objUPSClass.ReadResponseNode("//Response", "ResponseStatusCode")	
						
						'// ERROR
						pcv_strErrorCode = objUPSClass.ReadResponseNode("//Error", "ErrorCode")
						pcv_strErrorMessage = objUPSClass.ReadResponseNode("//Error", "ErrorDescription")
			
						pcv_strServiceDescription = objUPSClass.ReadResponseNode("//Shipment", "Service/Description")
	
	
						pcv_strAddressLine1 = objUPSClass.ReadResponseNode("//ShipTo", "Address/AddressLine1")
						pcv_strCity = objUPSClass.ReadResponseNode("//ShipTo", "Address/City")
						pcv_strStateProvinceCode = objUPSClass.ReadResponseNode("//ShipTo", "Address/StateProvinceCode")
						pcv_strPostalCode = objUPSClass.ReadResponseNode("//ShipTo", "Address/PostalCode")
						pcv_strCountryCode = objUPSClass.ReadResponseNode("//ShipTo", "Address/CountryCode")
						
						pcv_strShipmentIdentificationNumber = objUPSClass.ReadResponseNode("//Shipment", "ShipmentIdentificationNumber")
						pcv_strPickupDate = objUPSClass.ReadResponseNode("//Shipment", "PickupDate")
						pcv_strScheduledDeliveryDate = objUPSClass.ReadResponseNode("//Shipment", "ScheduledDeliveryDate")
	
						'// PACKAGE												
						pcv_strTrackingNumber = objUPSClass.ReadResponseNode("//Package", "TrackingNumber")
						
						pcv_strActivityCity = objUPSClass.ReadResponseNode("//Package", "Activity/ActivityLocation/Address/City")
						pcv_strActivityStateProvinceCode = objUPSClass.ReadResponseNode("//Package", "Activity/ActivityLocation/Address/StateProvinceCode")
						pcv_strActivityPostalCode = objUPSClass.ReadResponseNode("//Package", "Activity/ActivityLocation/Address/PostalCode")
						pcv_strActivityCountryCode = objUPSClass.ReadResponseNode("//Package", "Activity/ActivityLocation/Address/CountryCode")
						pcv_strActivityLocationCode = objUPSClass.ReadResponseNode("//Package", "Activity/ActivityLocation/Code")
						pcv_strActivityLocationDescription = objUPSClass.ReadResponseNode("//Package", "Activity/ActivityLocation/Description")
						
						pcv_strActivityStatusCode = objUPSClass.ReadResponseNode("//Package", "Activity/Status/StatusType/Code")
						pcv_strActivityStatusDescription = objUPSClass.ReadResponseNode("//Package", "Activity/Status/StatusType/Description")
						
						pcv_strActivityDate = objUPSClass.ReadResponseNode("//Package", "Activity/Date")
						pcv_strActivityTime = objUPSClass.ReadResponseNode("//Package", "Activity/Time")
						
						pcv_strUnitOfMeasurement = objUPSClass.ReadResponseNode("//Package", "PackageWeight/UnitOfMeasurement/Code")
						pcv_strPackageWeight = objUPSClass.ReadResponseNode("//Package", "PackageWeight/Weight")
						
						strActivityYear = left(pcv_strActivityDate, 4)
						strActivityDay = right(pcv_strActivityDate, 2)
						strActivityMonth = MonthName(mid(pcv_strActivityDate, 5, 2))
						strActivityDisplayDate = MonthName(mid(pcv_strActivityDate, 5, 2))&" "&strActivityDay&", "&strActivityYear
					
						%>
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2">UPS Tracking Number <%=pcv_strTrackingNumber%> Summary</th>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>			
							<tr> 
								<td colspan="2">
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td colspan="4">
									
												<table width="100%" border="0" cellpadding="0" cellspacing="0">
													<tr>
														<td width="12%" align="left"><span style="font-weight: bold">Status: </span></td>
														<td width="88%" align="left"><%=pcv_strActivityStatusDescription%></td>
														</tr>
													<tr>
													  <td align="left"><span style="font-weight: bold">Details: </span></td>
													  <td align="left">Date: <%=strActivityDisplayDate%></td>
													  </tr>
													<tr>
													  <td align="right">&nbsp;	</td>
													  <td align="left">Shipped To: <%=pcv_strActivityCity&", "&pcv_strActivityStateProvinceCode&" "&pcv_strActivityPostalCode %> </td>
													  </tr>
													<tr>
													  <td align="right">&nbsp;</td>
													  <td align="left">Package Weight: <%=pcv_strPackageWeight&" "&pcv_strUnitOfMeasurement %></td>
													  </tr>
													<tr>
													  <td align="right">&nbsp;</td>
													  <td align="left">&nbsp;</td>
													  </tr>
												</table>											</td>
										</tr>
										<tr>
											<th width="17%" align="left"><strong>Date  </strong></th>
											<th width="18%" align="left"><strong>Time</strong></th>
											<th width="25%" align="left"><strong>Location </strong></th>
											<th width="40%" align="left"><strong>Activity </strong></th>
										</tr>
										<%
										arrayActivityCity = objUPSClass.ReadResponseasArray("//Activity", "ActivityLocation/Address/City")
										arrayActivityCity = objUPSClass.pcf_UPSTrimArray(arrayActivityCity)
										arrayActivityStateProvinceCode = objUPSClass.ReadResponseasArray("//Activity", "ActivityLocation/Address/StateProvinceCode")
										arrayActivityStateProvinceCode = objUPSClass.pcf_UPSTrimArray(arrayActivityStateProvinceCode)
										arrayActivityPostalCode = objUPSClass.ReadResponseasArray("//Activity", "ActivityLocation/Address/PostalCode")
										arrayActivityPostalCode = objUPSClass.pcf_UPSTrimArray(arrayActivityPostalCode)
										arrayActivityCountryCode = objUPSClass.ReadResponseasArray("//Activity", "ActivityLocation/Address/CountryCode")
										arrayActivityCountryCode = objUPSClass.pcf_UPSTrimArray(arrayActivityCountryCode)
										arrayrActivityLocationCode = objUPSClass.ReadResponseasArray("//Activity", "ActivityLocation/Code")
										arrayActivityLocationCode = objUPSClass.pcf_UPSTrimArray(arrayActivityLocationCode)
										arrayActivityLocationDescription = objUPSClass.ReadResponseasArray("//Activity", "ActivityLocation/Description")
										arrayActivityLocationDescription = objUPSClass.pcf_UPSTrimArray(arrayActivityLocationDescription)
									
										arrayActivityStatusCode = objUPSClass.ReadResponseasArray("//Activity", "Status/StatusType/Code")
										arrayActivityStatusCode = objUPSClass.pcf_UPSTrimArray(arrayActivityStatusCode)
										arrayActivityStatusDescription = objUPSClass.ReadResponseasArray("//Activity", "Status/StatusType/Description")
										arrayActivityStatusDescription = objUPSClass.pcf_UPSTrimArray(arrayActivityStatusDescription)
										
										arrayActivityDate = objUPSClass.ReadResponseasArray("//Activity", "Date")
										arrayActivityDate = objUPSClass.pcf_UPSTrimArray(arrayActivityDate)
										arrayActivityTime = objUPSClass.ReadResponseasArray("//Activity", "Time")
										arrayActivityTime = objUPSClass.pcf_UPSTrimArray(arrayActivityTime)
								
										'// Generate/ Trim Event Date
										pcArrActivityDate = split(arrayActivityDate, ",")
										pcArrActivityTime = split(arrayActivityTime, ",")
										pcArrActivityCity = split(arrayActivityCity, ",")
										pcArrActivityStateProvinceCode = split(arrayActivityStateProvinceCode, ",")
										pcArrActivityPostalCode = split(arrayActivityPostalCode, ",")
										pcArrActivityCountryCode = split(arrayActivityCountryCode, ",")
										pcArrActivityLocationCode = split(arrayActivityLocationCode, ",")
										pcArrActivityLocationDescription = split(arrayActivityLocationDescription, ",")
										pcArrActivityStatusCode =split(arrayActivityStatusCode, ",")
										pcArrActivityStatusDescription = split(arrayActivityStatusDescription, ",")
		
										pcArrUnitOfMeasurement = split(arrayUnitOfMeasurement, ",")
										pcArrPackageWeight = split(arrayPackageWeight, ",")
					

										
										for bIdOptCounter = 0 to Ubound(pcArrActivityDate)
											strActivityYear = left(pcArrActivityDate(bIdOptCounter), 4)
											strActivityDay = right(pcArrActivityDate(bIdOptCounter), 2)
											strActivityMonth = MonthName(mid(pcArrActivityDate(bIdOptCounter), 5, 2))
											strActivityDisplayDate = strActivityMonth&" "&strActivityDay&", "&strActivityYear
	
											%>
											<tr>
												<td><b><%=strActivityDisplayDate%></b></td>
												<td><%=pcArrActivityTime(bIdOptCounter)%></td>
												<td><%=pcArrActivityCity(bIdOptCounter)&", "&pcArrActivityStateProvinceCode(bIdOptCounter)&" "&pcArrActivityPostalCode(bIdOptCounter) %></td>
												<td><%=pcArrActivityStatusDescription(bIdOptCounter)%></td>
											</tr>								 
										<% next %>
									</table>
								</td>
							</tr>
						</table>
					<% end if
					'set objUPSClass = nothing	  
				next
				'***************************************************************************
				' END LOOP THROUGH TRACKING
				'***************************************************************************
				%>
			</td>
		</tr>
	</form>
</table>
<p>
<table align="center">
	<tr><td>&nbsp;</td></tr>
	<tr>
		<td valign="top"><div align="center">
		<%= pcf_UPSWriteLegalDisclaimers %>
		</div></td>
	</tr>
</table>
</p>
<%
'// DESTROY THE UPS OBJECT
set objUPSClass = nothing
%>
<!--#include file="AdminFooter.asp"-->