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
dim pcv_strOrderID, pcv_strSessionOrderID, pcv_intOrderID, pcPageName, ErrPageName
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
pcPageName="UPS_ManageShipmentsCancel.asp"
ErrPageName="UPS_ManageShipmentsCancel.asp"

'// ACTION
pcv_strAction = request("Action")

'// OPEN DATABASE
call openDb()

'// SET THE UPS OBJECT
set objUPSClass = New pcUPSClass

dim PackageInfo_ID, SessionPackageInfo_ID, pcv_intPackageInfo

'// GET PACKAGE ID NUMBERS
PackageInfo_ID = Request("PackageInfo_ID")
SessionPackageInfo_ID = Session("pcAdminPackageInfo_ID")
if SessionPackageInfo_ID="" OR len(PackageInfo_ID)>0 then
	pcv_intPackageInfo = PackageInfo_ID
	Session("pcAdminPackageInfo_ID")=pcv_intPackageInfo
else
	pcv_intPackageInfo = SessionPackageInfo_ID
end if

dim query, rs, conntemp
	
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

'// SELECT DATA SET
' >>> Tables: pcPackageInfo
query = "SELECT pcPackageInfo_ID, pcPackageInfo_TrackingNumber, pcPackageInfo_ShippedDate "
query = query & "FROM pcPackageInfo "
query = query & "WHERE pcPackageInfo_ID=" & pcv_intPackageInfo &" "	

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then		
	dim pcv_strTrackingNumber, pcv_strShipDate
	'// LOOKUP THE PACKAGE INFO
	pcv_strTrackingNumber=rs("pcPackageInfo_TrackingNumber")
	pcv_strShipDate=rs("pcPackageInfo_ShippedDate")
end if
set rs=nothing

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Page Load
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
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for Validation Errors. Do not proceed if there are errors.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If pcv_intErr>0 Then
		response.redirect pcPageName & "?msg=" & pcv_strGenericPageError
	Else
			
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Build Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		objUPSClass.NewXMLTransaction ups_license_key, ups_userid, ups_password
		objUPSClass.NewXMLShipmentVoidRequest	 "Cancel", pcv_strTrackingNumber

		'//Clear illegal ampersand characters from XML
		UPS_postdata=replace(UPS_postdata, "&", "and")
		UPS_postdata=replace(UPS_postdata, "andamp;", "and")
		
		'// Print out our newly formed request xml
		'response.write UPS_postdata
		'response.end
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Send Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if UPS_TESTMODE="1" then
			UPS_URL="https://wwwcie.ups.com/ups.app/xml/Void"
		else
			UPS_URL="https://www.ups.com/ups.app/xml/Void"
		end if
		call objUPSClass.SendXMLRequest(UPS_postdata, UPS_URL)
		'// Print out our response
		'response.write UPS_result
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
		
			pcv_strHideForm="true"
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
			pcv_strCustomerTransactionIdentifier = objUPSClass.ReadResponseNode("//Response", "ResponseStatusCode")	
			
			'// ERROR
			pcv_strErrorCode = objUPSClass.ReadResponseNode("//Error", "ErrorCode")
			pcv_strErrorMessage = objUPSClass.ReadResponseNode("//Error", "ErrorDescription")
			
			'// Reset Products
			query="UPDATE ProductsOrdered SET PcPrdOrd_Shipped=0 WHERE idOrder="&pcv_intOrderID&" AND pcPackageInfo_ID="&pcv_intPackageInfo&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)

			'// Insert Code that will delete the package shipment info
			query="DELETE FROM pcPackageInfo "
			query = query & "WHERE pcPackageInfo_ID=" & pcv_intPackageInfo &";"	
			
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
			call closedb()
			if err.number<>0 then
				response.redirect pcPageName & "?msg=There was an error processing your request. Please try again."
			else
				pcs_ClearAllSessions()
				response.redirect "UPS_ManageShipmentsResults.asp?id=" & pcv_intOrderID & "&msg=Your Shipment has been deleted.&del=YES"
				response.end					
			end if			
			%>			
			
			
			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th colspan="2">UPS OnLine&reg; Tools Shipping - Shipment Canceled</th>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>	
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
<% if pcv_strHideForm <> "true" then %>
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">UPS OnLine&reg; Tools Shipping - Void Shipment Request</th>
	</tr>
    <% if UPS_TESTMODE="1" then %>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td colspan="2" valign="top">
            <div class="pcCPmessage">UPS Shipping Wizard is currently running in Test Mode<br>
This feature is disabled while in Test Mode </div>
           </td>
        </tr>
	<% end if %>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
			<p>
			UPS<sup>&reg;</sup> Void Shipment Request is used to cancel a shipping request after the acceptance phase. A shipping request can be voided until the end of the following day (23:59 Eastern Time) after a shipment has been accepted.<b> <br>
      </b><br />
			<b>NOTE:</b>  Void is only valid before a shipment is picked up by the UPS service provider. 
			</p>
		</td>
	</tr>
</table>

<table class="pcCPcontent">

	<form name="form1" action="<%=pcPageName%>" method="post" class="pcForms">
	<input name="PackageInfo_ID" type="hidden" value="<%=pcv_intPackageInfo%>">
	<input name="id" type="hidden" value="<%=pcv_intOrderID%>">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>		
		<tr>
			<th colspan="2">Are you sure you want to void this shipment?</th>
		</tr>		
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>	
		<tr>
			<td></td>
			<td>
			<input type=submit name="submit" value="Request Void Shipment" class="ibtnGrey">
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
<% end if
'// DESTROY THE UPS OBJECT
set objUPSClass = nothing
%>
<!--#include file="AdminFooter.asp"-->