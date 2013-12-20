<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEX<sup>&reg;</sup> Shipping Configuration: Enable FedEx&reg; API" %>
<% Section="shipOpt" %>
<% pcPageName = "ConfigureOption3.asp" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->

<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<% 
Dim query, rs, conntemp
Dim pcv_strMethodName, pcv_strMethodReply, fedex_postdata, objFedExClass, objOutputXMLDoc
Dim srvFEDEXXmlHttp, FEDEX_result, FEDEX_URL, pcv_strErrorMsg, objFEDEXXmlDoc, objFedExStream, strFileName, GraphicXML

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

pcv_strMethodName = "FDXSubscriptionRequest"
pcv_strMethodReply = "FDXSubscriptionReply"	
CustomerTransactionIdentifier = "Subscription_Request"
	 
if request.form("submit")<>"" then
	
	'// Set error count
	pcv_intErr=0	
	
	'// generic error for page
	pcv_strGenericPageError = "At least one required field was empty."
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ValidateTextField	"FedExMode", true, 4	
	pcs_ValidateTextField	"FedEx_AccountNumber", true, 20	
	pcs_ValidateTextField	"FedEx_PersonName", true, 20	
	pcs_ValidateTextField	"FedEx_CompanyName", false, 20
	pcs_ValidateTextField	"FedEx_Department", false, 20
	pcs_ValidatePhoneNumber	"FedEx_PhoneNumber", true, 16
	pcs_ValidatePhoneNumber	"FedEx_PagerNumber", false, 16
	pcs_ValidatePhoneNumber	"FedEx_FaxNumber", false, 16
	pcs_ValidateEmailField	"FedEx_eMailAddress", false, 250	
	pcs_ValidateTextField	"FedEx_Line1", true, 20	
	pcs_ValidateTextField	"FedEx_Line2", false, 20
	pcs_ValidateTextField	"FedEx_City", true, 20
	pcs_ValidateTextField	"FedEx_CountryCode", true, 2	
	pcs_ValidateTextField	"FedEx_StateOrProvinceCode", true, 20
		
	if len(Session("pcAdminFedEx_Line1"))<1 AND (Session("pcAdminFedEx_CountryCode")="US" OR Session("pcAdminFedEx_CountryCode")="CA") then
		pcv_intErr=pcv_intErr+1
	end if	
	
	pcs_ValidateTextField	"FedEx_PostalCode", true, 20	
	
	if len(Session("pcAdminFedEx_PostalCode"))<1 AND (Session("pcAdminFedEx_CountryCode")="US" OR Session("pcAdminFedEx_CountryCode")="CA") then
		pcv_intErr=pcv_intErr+1
	end if	
	
	Session("pcAdminFedEx_AccountNumber") = replace(Session("pcAdminFedEx_AccountNumber"),"-","")
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
		FedExAPI_ID=getUserInput(request("FedExAPI_ID"),4)
		
		'// Open the DB
		call opendb()
		
		'// Generate the Query (Save form data)
		if FedExAPI_ID=0 then		
			query="INSERT INTO FedExAPI (FedExAPI_PersonName, FedExAPI_CompanyName, FedExAPI_Department, FedExAPI_PhoneNumber, FedExAPI_PagerNumber, FedExAPI_FaxNumber, FedExAPI_EmailAddress, FedExAPI_Line1, FedExAPI_Line2, FedExAPI_city, FedExAPI_State, FedExAPI_PostalCode, FedExAPI_Country) VALUES ('"&Session("pcAdminFedEx_PersonName")&"', '"&Session("pcAdminFedEx_CompanyName")&"', '"&Session("pcAdminFedEx_Department")&"', '"&Session("pcAdminFedEx_PhoneNumber")&"', '"&Session("pcAdminFedEx_PagerNumber")&"', '"&Session("pcAdminFedEx_FaxNumber")&"', '"&Session("pcAdminFedEx_EmailAddress")&"', '"&Session("pcAdminFedEx_Line1")&"', '"&Session("pcAdminFedEx_Line2")&"', '"&Session("pcAdminFedEx_City")&"', '"&Session("pcAdminFedEx_StateOrProvinceCode")&"', '"&Session("pcAdminFedEx_PostalCode")&"', '"&Session("pcAdminFedEx_CountryCode")&"');"
		else		
			query="UPDATE FedExAPI SET FedExAPI_PersonName='"&Session("pcAdminFedEx_PersonName")&"', FedExAPI_CompanyName='"&Session("pcAdminFedEx_CompanyName")&"', FedExAPI_Department='"&Session("pcAdminFedEx_Department")&"', FedExAPI_PhoneNumber='"&Session("pcAdminFedEx_PhoneNumber")&"', FedExAPI_PagerNumber='"&Session("pcAdminFedEx_PagerNumber")&"', FedExAPI_FaxNumber='"&Session("pcAdminFedEx_FaxNumber")&"', FedExAPI_EmailAddress='"&Session("pcAdminFedEx_EmailAddress")&"', FedExAPI_Line1='"&Session("pcAdminFedEx_Line1")&"', FedExAPI_Line2='"&Session("pcAdminFedEx_Line2")&"', FedExAPI_city='"&Session("pcAdminFedEx_City")&"', FedExAPI_State='"&Session("pcAdminFedEx_StateOrProvinceCode")&"', FedExAPI_PostalCode='"&Session("pcAdminFedEx_PostalCode")&"', FedExAPI_Country='"&Session("pcAdminFedEx_CountryCode")&"' WHERE FedExAPI_ID=1;"
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
		set objFedExClass = New pcFedExClass
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Build Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		objFedExClass.NewXMLSubscription pcv_strMethodName, Session("pcAdminFedEx_AccountNumber"), CustomerTransactionIdentifier		
			objFedExClass.WriteSingleParent "CSPSolutionType", "120"
			objFedExClass.WriteSingleParent "CSPIndicator", "01"
			objFedExClass.WriteParent "Contact", ""
				objFedExClass.AddNewNode "PersonName", Session("pcAdminFedEx_PersonName")
				objFedExClass.AddNewNode "CompanyName", Session("pcAdminFedEx_CompanyName")
				objFedExClass.AddNewNode "Department", Session("pcAdminFedEx_Department")
				objFedExClass.AddNewNode "PhoneNumber", fnStripPhone(Session("pcAdminFedEx_PhoneNumber"))
				objFedExClass.AddNewNode "PagerNumber", fnStripPhone(Session("pcAdminFedEx_PagerNumber"))
				objFedExClass.AddNewNode "FaxNumber", fnStripPhone(Session("pcAdminFedEx_FaxNumber"))
				objFedExClass.AddNewNode "E-MailAddress", Session("pcAdminFedEx_eMailAddress")
			objFedExClass.WriteParent "Contact", "/"
			objFedExClass.WriteParent "Address", ""
				objFedExClass.AddNewNode "Line1", Session("pcAdminFedEx_Line1")
				objFedExClass.AddNewNode "Line2", Session("pcAdminFedEx_Line2")
				objFedExClass.AddNewNode "City", Session("pcAdminFedEx_City")
				objFedExClass.AddNewNode "StateOrProvinceCode", Session("pcAdminFedEx_StateOrProvinceCode")
				objFedExClass.AddNewNode "PostalCode", Session("pcAdminFedEx_PostalCode")
				objFedExClass.AddNewNode "CountryCode", Session("pcAdminFedEx_CountryCode")
			objFedExClass.WriteParent "Address", "/"		
		objFedExClass.EndXMLTransaction pcv_strMethodName	
		
		'// Print out our newly formed request xml
		'response.write fedex_postdata
		'response.end
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Send Our Transaction.
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		call objFedExClass.SendXMLRequest(fedex_postdata, Session("pcAdminFedExMode"))
		
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
		If NOT len(pcv_strErrorMsg)>0 Then
	
			'// HEADER
			pcv_strCustomerTransactionIdentifier = objFedExClass.ReadResponseNode("//ReplyHeader", "CustomerTransactionIdentifier")	
			
			'// ERROR
			pcv_strErrorCode = objFedExClass.ReadResponseNode("//Error", "Code")
			pcv_strErrorMessage = objFedExClass.ReadResponseNode("//Error", "Message")
			
			'// METER
			pcv_strMeterNumber = objFedExClass.ReadResponseParent(pcv_strMethodReply, "MeterNumber")
			
			'// Ensure that the MeterNumber exists
			if len(pcv_strMeterNumber)<1 OR len(Session("pcAdminFedEx_AccountNumber"))<1 then	
				response.redirect pcPageName & "?msg=There was an error activating your FedEx account. The FedEx servers may be down temporarily. Please try again later."
			else
				'// Save MeterNumber in database along with AccountNumber
				call opendb()
				query="UPDATE ShipmentTypes SET password='"&pcv_strMeterNumber&"', userID='"&Session("pcAdminFedEx_AccountNumber")&"', AccessLicense='"&Session("pcAdminFedExMode")&"' WHERE (((ShipmentTypes.idShipment)=1));"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing				
				if err.number<>0 then
					call closedb()
					response.redirect pcPageName & "?msg=There was an error activating your FedEx account. Please submit your registration request again."
				else					
					'// Generate the Query (Update shipment types)
					query="UPDATE ShipmentTypes SET AccessLicense='"&Session("pcAdminFedExMode")&"' WHERE (((ShipmentTypes.idShipment)=1));"
					set rs=server.CreateObject("ADODB.RecordSet")
					'// Execute the Query
					set rs=conntemp.execute(query)
					set rs=nothing
					call closedb()			
					'// No errors, redirect to next step					
					session("FedExSetUP")="YES"
					pcs_ClearAllSessions()
					response.redirect "FEDEX_EditShipOptions.asp"
					response.end	
									
				end if
			end if		
			
		End If '// If NOT len(pcv_strErrorMsg)>0 Then
	
	end if
end if
'**************************************************************************
' END: If registration request was submitted, process request
'**************************************************************************




'**************************************************************************
' START: Was FedEx was previously registered by querying the database ?
'**************************************************************************
if request("changeMode")="" then
	call opendb()
	query="SELECT userID FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=1));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	If rs.eof then
		'FedEx was previously activated - redirect
		session("FedExSetUP")="YES"
		response.redirect "FEDEX_EditShipOptions.asp"
		response.end
	end if 
	set rs=nothing
	call closedb()
end if
'**************************************************************************
' END: Was FedEx was previously registered by querying the database ?
'**************************************************************************
%>

<%  
'**************************************************************************
' START: Get Fed Ex credentials
'**************************************************************************
call opendb()

'// Get Access License
query="SELECT ShipmentTypes.AccessLicense, ShipmentTypes.userID FROM ShipmentTypes WHERE idShipment=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

Session("pcAdminFedEx_AccountNumber") = rs("userID")
strAccessLicense=rs("AccessLicense")

if len(strAccessLicense)<1 then
	strAccessLicense="TEST"
end if

'// Get Form Data
query="SELECT FedExAPI_ID, FedExAPI_PersonName, FedExAPI_CompanyName, FedExAPI_Department, FedExAPI_PhoneNumber, FedExAPI_PagerNumber, FedExAPI_FaxNumber, FedExAPI_EmailAddress, FedExAPI_Line1, FedExAPI_Line2, FedExAPI_city, FedExAPI_State, FedExAPI_PostalCode, FedExAPI_Country FROM FedExAPI;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if NOT rs.eof then
	FedExAPI_ID=rs("FedExAPI_ID")
	if request("changeMode")="Y" then		
		Session("pcAdminFedEx_PersonName")=rs("FedExAPI_PersonName")
		Session("pcAdminFedEx_CompanyName")=rs("FedExAPI_CompanyName")
		Session("pcAdminFedEx_Department")=rs("FedExAPI_Department")
		Session("pcAdminFedEx_PhoneNumber")=rs("FedExAPI_PhoneNumber")
		Session("pcAdminFedEx_PagerNumber")=rs("FedExAPI_PagerNumber")
		Session("pcAdminFedEx_FaxNumber")=rs("FedExAPI_FaxNumber")
		Session("pcAdminFedEx_EmailAddress")=rs("FedExAPI_EmailAddress")
		Session("pcAdminFedEx_Line1")=rs("FedExAPI_Line1")
		Session("pcAdminFedEx_Line2")=rs("FedExAPI_Line2")
		Session("pcAdminFedEx_city")=rs("FedExAPI_city")
		Session("pcAdminFedEx_StateOrProvinceCode")=rs("FedExAPI_State")
		Session("pcAdminFedEx_PostalCode")=rs("FedExAPI_PostalCode")
		Session("pcAdminFedEx_Country")=rs("FedExAPI_Country")
	end if
else
	FedExAPI_ID=0
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
	
<form name="form1" method="post" action="ConfigureOption3.asp" class="pcForms">
	<input type="hidden" name="FedExAPI_ID" value="<%=FedExAPI_ID%>">
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
			<input name="FedExMode" type="hidden" value="LIVE">
			</td>
		</tr>		
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="23%">FedEx Account Number:</td>
			<td width="77%"><input name="FedEx_AccountNumber" type="text" value="<%=pcf_FillFormField("FedEx_AccountNumber", true)%>" size="15" maxlength="25">
			<%pcs_RequiredImageTag "FedEx_AccountNumber", true%></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td width="23%">Contact Name: </td>
			<td width="77%">
			  <input name="FedEx_PersonName" type="text" value="<%=pcf_FillFormField("FedEx_PersonName", true)%>" size="30" maxlength="100">
			  <%pcs_RequiredImageTag "FedEx_PersonName", true%></td>
		</tr>
		<tr>
			<td width="23%">Company Name: </td>
			<td width="77%">
			  <input name="FedEx_CompanyName" type="text" value="<%=pcf_FillFormField("FedEx_CompanyName", false)%>" size="30" maxlength="100"></td>
		</tr>
		<tr>
			<td width="23%">Department: </td>
			<td width="77%">
			  <input name="FedEx_Department" type="text" value="<%=pcf_FillFormField("FedEx_Department", false)%>" size="30" maxlength="100"></td>
		</tr>
		<% if len(Session("ErrFedEx_PhoneNumber"))>0 then %>
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
			  <input name="FedEx_PhoneNumber" type="text" value="<%=pcf_FillFormField("FedEx_PhoneNumber", true)%>" size="16" maxlength="16">
			  <%pcs_RequiredImageTag "FedEx_PhoneNumber", true%></td>
		</tr>
		<tr>
			<td width="23%">Pager Number:  </td>
			<td width="77%">
			  <input name="FedEx_PagerNumber" type="text" value="<%=pcf_FillFormField("FedEx_PagerNumber", false)%>" size="16" maxlength="16">
			  </td>
		</tr>
		<tr>
			<td width="23%">Fax Number:  </td>
			<td width="77%">
			  <input name="FedEx_FaxNumber" type="text" value="<%=pcf_FillFormField("FedEx_FaxNumber", false)%>" size="16" maxlength="16">
			  </td>
		</tr>
		<% if len(Session("ErrFedEx_eMailAddress"))>0 then %>
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
			  <input name="FedEx_eMailAddress" type="text" value="<%=pcf_FillFormField("FedEx_eMailAddress", true)%>" size="40" maxlength="250">
			  <%pcs_RequiredImageTag "FedEx_eMailAddress", true%></td>
			</tr>
		<tr>
			<td width="23%">Address:  </td>
			<td width="77%">
			  <input name="FedEx_Line1" type="text" value="<%=pcf_FillFormField("FedEx_Line1", true)%>" size="40" maxlength="250">
			  <%pcs_RequiredImageTag "FedEx_Line1", true%></td>
			</tr>
		<tr>
			<td width="23%">&nbsp;</td>
			<td width="77%">
			<input name="FedEx_Line2" type="text" value="<%=pcf_FillFormField("FedEx_Line2", false)%>" size="40" maxlength="250"></td>
			</tr>
		<tr>
			<td width="23%">City:  </td>
			<td width="77%">
			  <input name="FedEx_City" type="text" value="<%=pcf_FillFormField("FedEx_City", true)%>" size="30" maxlength="250">
			  <%pcs_RequiredImageTag "FedEx_City", true%></td>
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
				<select name="FedEx_StateOrProvinceCode">
				<option value=""></option>
				<% do while not rs.eof
					strStateCode=rs("stateCode")
					strStateName=rs("stateName")
					%>
					<option value="<%=strStateCode%>" <%=pcf_SelectOption("FedEx_StateOrProvinceCode",strStateCode)%>><%=strStateName%></option>
					<%rs.movenext
				loop
				set rs=nothing
				call closedb() %>
				</select>
				<%pcs_RequiredImageTag "FedEx_StateOrProvinceCode", true%>
				</td>
			</tr>
		<tr>
			<td width="23%">Postal Code:  </td>
			<td width="77%">
			  <input name="FedEx_PostalCode" type="text" value="<%=pcf_FillFormField("FedEx_PostalCode", true)%>" size="15" maxlength="20">
			  <%pcs_RequiredImageTag "FedEx_PostalCode", true%></td>
			</tr>
		<tr>
			<td width="23%">Country Code:  </td>
			<td width="77%">
			  US<input type="hidden" name="FedEx_CountryCode" value="US"></td>
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