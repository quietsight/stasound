<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard - Close Manifest" %>
<% Section="mngAcc" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->

<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<% 
Dim query, rs, conntemp
Dim iPageCurrent, varFlagIncomplete, uery, strORD, pcv_intOrderID
Dim pcv_strMethodName, pcv_strMethodReply, CustomerTransactionIdentifier, pcv_strAccountNumber, pcv_strMeterNumber, pcv_strCarrierCode
Dim pcv_strTrackingNumber, pcv_strShipmentAccountNumber
Dim pcv_strDestinationCountryCode, pcv_strDestinationPostalCode, pcv_strLanguageCode, pcv_strLocaleCode, pcv_strDetailScans, pcv_strPagingToken
Dim fedex_postdata, objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEX_result, FEDEX_URL, pcv_strErrorMsg, pcv_strAction
Dim objFEDEXXmlDoc, objFedExStream, strFileName, GraphicXML


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
pcPageName="FedEx_ManageShipmentsClose.asp"
ErrPageName="FedEx_ManageShipmentsClose.asp"

'// ACTION
pcv_strAction = request("Action")

'// OPEN DATABASE
call openDb()

'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExClass

'// GET PACKAGE ID NUMBERS
PackageInfo_ID = Request("PackageInfo_ID")
SessionPackageInfo_ID = Session("pcAdminPackageInfo_ID")
if SessionPackageInfo_ID="" OR len(PackageInfo_ID)>0 then
	pcv_intPackageInfo = PackageInfo_ID
	Session("pcAdminPackageInfo_ID")=pcv_intPackageInfo
else
	pcv_intPackageInfo = SessionPackageInfo_ID
end if
	
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

'// SET REQUIRED VARIABLES
pcv_strMethodName = "FDXCloseRequest"
pcv_strMethodReply = "FDXCloseReply"
CustomerTransactionIdentifier = "ProductCart_CloseManifest"	


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'***************************************************************************
' START: POST BACK
'***************************************************************************
if request.form("submit")<>"" then

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Get all of the required information.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	
	'// Generic error for page
	pcv_strGenericPageError = "At least one required field was empty. "
	
	pcs_ValidateTextField	"Date", true, 10
	pcs_ValidateTextField	"Time", false, 8
	pcs_ValidateTextField	"ReportIndicator", false, 0
	pcs_ValidateTextField	"ReportOnly", false, 1
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for Validation Errors. Do not proceed if there are errors.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If pcv_intErr>0 Then
		response.redirect pcPageName & "?msg=" & pcv_strGenericPageError
	Else
			
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Build Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		objFedExClass.NewXMLTransaction pcv_strMethodName, pcv_strAccountNumber, pcv_strMeterNumber, "FDXG", CustomerTransactionIdentifier
			objFedExClass.WriteSingleParent "Date", Session("pcAdminDate")
			objFedExClass.WriteSingleParent "Time", Session("pcAdminTime")	
			if Session("pcAdminReportIndicator")<>"" AND  Session("pcAdminReportOnly")<>"" then				
				objFedExClass.WriteSingleParent "ReportIndicator", Session("pcAdminReportIndicator")		
				objFedExClass.WriteSingleParent "ReportOnly", Session("pcAdminReportOnly")
			end if
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
		
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Check for errors from FedEx.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~				
		call objFedExClass.XMLResponseVerify(ErrPageName)
		

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
			
			'// REPORTS
			pcv_strMultiweightReport = objFedExClass.ReadResponseNode("//Manifest", "MultiweightReport")	
			pcv_strFileName = objFedExClass.ReadResponseNode("//Manifest", "FileName")	
			pcv_strFile = objFedExClass.ReadResponseNode("//Manifest", "File")
			
			if pcv_strFile<>"" then				
				
				'=======================
				'// Start Label Decoding
				'=======================
				'// Create XML for Label 
				'TrackingNumber, EncodedLabelString, FileType, FilePreFix
				objFedExClass.NewXMLLabel pcv_strFileName, pcv_strFile, "TXT", "CLOSE"
		
				'// Load label from the request stream
				call objFedExClass.LoadXMLLabel(GraphicXML)
		
				'// Use ADO stream to save the binary data
				objFedExClass.SaveBinaryLabel
				'=======================
				'// End Label Decoding
				'=======================
				pcv_strURLFile = "FedExLabels/CLOSE"&pcv_strFileName&".TXT"
			end if
			%>
			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th colspan="2">FedEx<sup>&reg;</sup> Close Confirmation</th>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>	
				<% if pcv_strMultiweightReport<>"" then %>
				<!--				
				<tr> 
					<td>Multi-weight Report:  </td>
					<td><a href="<%=pcv_strURLpcv_strMultiweightReport%>" target="_blank">Click Here to Save</a>.</td>
				</tr> 
				-->
				<% end if %>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>	
				<% if pcv_strFile<>"" then %>
				<tr> 
					<td>Manifest Report:  </td>
					<td><a href="<%=pcv_strURLFile%>" target="_blank">Click Here to Save</a>.</td>
				</tr>
				<% end if %>
			</table>
			<%
		end if
	end if	  
	
end if
'***************************************************************************
' END: POST BACK
'***************************************************************************
%>
<% 
msg=request.querystring("msg")

if msg<>"" then 
	%>
	<div class="pcCPmessage">
		<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
	</div>
	<% 
end if 
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">FedEx<sup>&reg;</sup> End of Day Closeout & Print Manifest</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
			<p>
			This service allows a customer to close out all shipments made for the day. Closing is a function to be used only for FedEx Ground shipments. Customers cannot cancel
			any shipments once they are closed out. However, shipments can be added to a day's shipment after a Close has been performed and multiple Closes can be performed in
			a day.
			</p>
		</td>
	</tr>
</table>

<table class="pcCPcontent">
	<form name="form1" action="<%=pcPageName%>" method="post" class="pcForms">
		<input name="PackageInfo_ID" type="hidden" value="<%=pcv_intPackageInfo%>">
		<input name="id" type="hidden" value="<%=pcv_intOrderID%>">
		<%
		dtShippedDate=date()
		if SQL_Format="1" then
			dtShippedDate=(day(dtShippedDate)&"/"&month(dtShippedDate)&"/"&year(dtShippedDate))
		else
			dtShippedDate=(month(dtShippedDate)&"/"&day(dtShippedDate)&"/"&year(dtShippedDate))
		end if
		function pad(thevalue)
			x=len(thevalue)
			if x = 1 then
				pad="0"&thevalue
			else
				pad=thevalue
			end if
		end function
		dtShippedDate=(year(dtShippedDate)&"-"&pad(month(dtShippedDate))&"-"&pad(day(dtShippedDate)))
		%>
		<input name="Date" type="hidden" id="Date" value="<%=dtShippedDate%>">
		<input name="Time" type="hidden" id="Time" value="<%=FormatDateTime(now(),4) & ":00"%>">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>		
		<tr>
			<th colspan="2">Closeout Options</th>
		</tr>		
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>Report Only:</td>
			<td>
			<input type="checkbox" name="ReportOnly" value="1" class="clearBorder" <%=pcf_CheckOption("ReportOnly", "1")%>>
			</td>
		</tr>	
		<tr>
			<td>Report Type:</td>
			<td>
				<select name="ReportIndicator" id="ReportIndicator">
					<option value="" <%=pcf_SelectOption("ReportIndicator","")%>>NONE</option>
					<option value="MANIFEST" <%=pcf_SelectOption("ReportIndicator","MANIFEST")%>>Ground Manifest Report</option>					
					<!--<option value="MULTIWEIGHT" <%=pcf_SelectOption("ReportIndicator","MULTIWEIGHT")%>>Ground Multiweight Report</option> -->
				</select>
			</td>
		</tr>		
		<tr>
			<td></td>
			<td>
			<input type=submit name="submit" value="Close Ground Shipments" class="ibtnGrey">
			</td>
		</tr>
	</form>
</table>
<%
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>
<!--#include file="AdminFooter.asp"-->