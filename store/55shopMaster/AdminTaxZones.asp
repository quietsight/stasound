<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Store Settings" %>
<% Section="shipOpt" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/languages.asp" --> 
<% pcPageName="AdminTaxZones.asp" 

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
Dim pcv_strAdminPrefix
pcv_strAdminPrefix="1"
			
dim conntemp %>
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<html>
<head>
<title><%=title%></title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:470px; background-image: none;">
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer" align="center">
			<% 
			msg=getUserInput(request.querystring("msg"),0)
			if msg<>"" then %>
				<div class="pcCPmessage">
					<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
				</div>
			<% end if %>			
		</td>
	</tr>
</table>
<% 
pcv_isZoneStateRequired=true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isZoneStateRequired=pcv_strStateCodeRequired
end if
pcv_isZoneProvinceRequired=false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isZoneProvinceRequired=pcv_strProvinceCodeRequired
end if
pcv_isZoneCountryRequired=true

if Session("pcAdminZoneProvince")<>"" then
	pcStrZoneState = Session("pcAdminZoneProvince")
else
	pcStrZoneState = Session("pcAdminZoneState")
end if
pcStrZoneCountry = Session("pcAdminZoneCountry")
pcStrNewZoneName = Session("pcAdminNewZoneName")

if request("submit")="Add New Zone" then
	call opendb()
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = "One of more fields were not filled in correctly."
	
	'// validate all fields
	pcs_ValidateTextField "NewZoneName", true, 250
	pcs_ValidateTextField	"ZoneState", pcv_isZoneStateRequired, 50
	pcs_ValidateTextField	"ZoneProvince", pcv_isZoneProvinceRequired, 50
	pcs_ValidateTextField	"ZoneCountry", pcv_isZoneCountryRequired, 50
	If pcv_intErr>0 Then	
		response.redirect pcStrPageName&"?reID="&reID&"&msg=" & pcv_strGenericPageError
	End if
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	
	'/////////////////////////////////////////////////////
	'// Set Local Variables for Setting
	'/////////////////////////////////////////////////////
	if Session("pcAdminZoneProvince")<>"" then
		pcStrZoneState = Session("pcAdminZoneProvince")
	else
		pcStrZoneState = Session("pcAdminZoneState")
	end if
	pcStrZoneCountry = Session("pcAdminZoneCountry")
	pcStrNewZoneName = Session("pcAdminNewZoneName")
	
	
	'// INSERT INTO DATABASE
	intStateCnt=0
	if instr(pcStrZoneState,",") then
		pcStrZoneStateArray=split(pcStrZoneState,", ")
		intStateCnt=ubound(pcStrZoneStateArray)
	end if

	on error resume next
	
	for i=0 to intStateCnt
		if intStateCnt=0 then
			pcStrZoneStateCode=pcStrZoneState
		else
			pcStrZoneStateCode=pcStrZoneStateArray(i)
		end if

		if pcStrZoneStateCode<>"" then
			query="INSERT INTO pcTaxZones (pcTaxZone_CountryCode,pcTaxZone_Province) VALUES ('"&pcStrZoneCountry&"', '"&pcStrZoneStateCode&"');"
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=conntemp.execute(query)

			query="SELECT pcTaxZone_ID FROM pcTaxZones WHERE pcTaxZone_CountryCode='"&pcStrZoneCountry&"' AND pcTaxZone_Province='"&pcStrZoneStateCode&"';"
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=conntemp.execute(query)
			
			intTaxZoneID=rstemp("pcTaxZone_ID")

			'// Check if this tax zone description already exists
			tempQuery="SELECT pcTaxZoneDesc_ID FROM pcTaxZoneDescriptions WHERE pcTaxZoneDesc='"&pcStrNewZoneName&"';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(tempQuery)
			
			'// If it doesn't already exists - create record
			if rs.eof then
				query="INSERT INTO pcTaxZoneDescriptions (pcTaxZoneDesc) VALUES ('"&pcStrNewZoneName&"');"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)

				set rstemp=conntemp.execute(tempQuery)
				intTaxZoneDescID=rstemp("pcTaxZoneDesc_ID")
			else
				intTaxZoneDescID=rs("pcTaxZoneDesc_ID")
			end if
			
			'// Check to ensure tax zone and tax desc id don't already exist in tax groups
			query="SELECT * FROM pcTaxGroups WHERE pcTaxZoneDesc_ID="&intTaxZoneDescID&" AND pcTaxZone_ID="&intTaxZoneID&" ;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			
			if rs.eof then	
				query="INSERT INTO pcTaxGroups (pcTaxZoneDesc_ID, pcTaxZone_ID) VALUES ("&intTaxZoneDescID&", "&intTaxZoneID&");"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
			end if
			
			set rs=nothing
			set rstemp=nothing
			
		end if
	next

	pcs_ClearAllSessions()
	
	response.redirect "pcTaxViewZones.asp?success=1"
	
	'/////////////////////////////////////////////////////
	'// Update database with new Settings
	'/////////////////////////////////////////////////////
end if 

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf
response.write "<!--"&vbcrlf	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf

StrGenericJSError="One or more fields were not filled in correctly"

pcs_JavaTextField	"ZoneState", pcv_isZoneStateRequired, StrGenericJSError
pcs_JavaDropDownList "ZoneCountry", pcv_isZoneCountryRequired, StrGenericJSError

response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "//-->"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' End Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

%>
<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">
	<table class="pcCPcontent">	
		<tr>
			<td valign="top">
				<table class="pcCPcontent">
					<tr>
					  <td colspan="2" class="pcCPspacer"></td>
					  </tr>
					<tr>
					  <th colspan="2">Tax Zones</th>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td colspan="2">
							<div>If a zone contains more than one state or province, select the first one here. You will be able to add new ones when you edit the tax zone.</div>
							<div style="padding-top:8px;">For example, British Columbia and Manitoba (Canada) both use a 7% Provincial Sales Tax (or Retail Sales Tax). You can add the first province here and the second one when you edit the zone.</div></td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
					  <td><p>Zone Name: </p></td>
					  <td><p><input type="text" name="NewZoneName" id="NewZoneName" size="20" value="<%=pcStrNewZoneName %>">
							<% pcs_RequiredImageTag "NewZoneName", true %></p></td>
					</tr>
					<%
					call openDB()
					'///////////////////////////////////////////////////////////
					'// START: COUNTRY AND STATE/ PROVINCE CONFIG
					'///////////////////////////////////////////////////////////
					' 
					' 1) Place this section ABOVE the Country field
					' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
					' 3) Additional Required Info
					
					'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
					pcv_isStateCodeRequired = pcv_isZoneStateRequired '// determines if validation is performed (true or false)
					pcv_isProvinceCodeRequired = pcv_isZoneProvinceRequired '// determines if validation is performed (true or false)
					pcv_isCountryCodeRequired = pcv_isZoneCountryRequired '// determines if validation is performed (true or false)
					
					'// #3 Additional Required Info
					pcv_strTargetForm = "form1" '// Name of Form
					pcv_strCountryBox = "ZoneCountry" '// Name of Country Dropdown
					pcv_strTargetBox = "ZoneState" '// Name of State Dropdown
					pcv_strProvinceBox =  "ZoneProvince" '// Name of Province Field
					
					'// Set local Country to Session
					if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
						Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrZoneCountry
					end if
					
					'// Set local State to Session
					if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
						Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrZoneState
					end if
					
					'// Set local Province to Session
					if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
						Session(pcv_strSessionPrefix&pcv_strProvinceBox) = pcStrZoneState
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

					'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
					pcs_StateProvince
					call closeDB()
					%>	
					<tr>
						<td class="pcCPspacer" colspan="2"></td>
					</tr>
					<tr>
					  <td></td>
					  <td>&nbsp;</td>
					  </tr>
					<tr>
						<td></td>
						<td>
						<p>
							<input type="submit" name="submit" value="Add New Zone" class="submit2">&nbsp;
						  <input name="button" type="button" onClick="location.href='pcTaxViewZones.asp'" value="Back">&nbsp;
						  <input type="button" name="Back" value="Close" onClick="self.close();">
						</p>
						</td>
					</tr>  
				</table>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
			</td>
		</tr>
	</table>
</form>
</div>
</body>
</html>