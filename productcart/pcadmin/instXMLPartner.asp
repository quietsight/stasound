<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="XML Tools - Add New Partner" %>
<% section="layout"%>
<%PmAdmin=19%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: PAGE CONFIG
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
dim conntemp, query, rstemp, rstemp2, rs

call openDb()

'// Set Page Name
pcStrPageName = "instXMLPartner.asp"

'// Set Required Fields

'// General Info
pcv_isUserStatusRequired= false
pcv_isPartnerNameRequired = false
pcv_isPartnerCompanyRequired = false
pcv_isPartnerPhoneRequired = false
pcv_isPartnerFaxRequired = false
pcv_isPartnerEmailRequired = false

'// Partner Address Info
pcv_isPartnerAddressRequired = false
pcv_isPartnerPostalCodeRequired = false
pcv_isPartnerCityRequired = false
pcv_isPartnerCountryCodeRequired = false
pcv_isPartnerAddress2Required = false
pcv_isPartnerStateCodeRequired = false
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isPartnerStateCodeRequired=pcv_strStateCodeRequired
end if
pcv_isPartnerProvinceRequired = false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isPartnerProvinceRequired=pcv_strProvinceCodeRequired
end if

'// XML Info
pcv_isUserRequired= true
pcv_isPasswordRequired= true
pcv_isUserKeyRequired= true

'//Export XML Info
pcv_isExportAdminRequired = false
pcv_isFTPHostRequired = false
pcv_isFTPDirectoryRequired = false
pcv_isFTPUsernameRequired = false
pcv_isFTPPasswordRequired = false
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: PAGE CONFIG
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IF request.Form("Modify")<>"" THEN

	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	
	'// General Info
	pcs_ValidateTextField "pcUserStatus", pcv_isUserStatusRequired, 0
	pcs_ValidateTextField "pcPartnerName", pcv_isPartnerNameRequired, 0
	pcs_ValidateTextField "pcPartnerCompany", pcv_isPartnerCompanyRequired, 0
	pcs_ValidatePhoneNumber "pcPartnerPhone", pcv_isPartnerPhoneRequired, 0
	pcs_ValidatePhoneNumber "pcPartnerFax", pcv_isPartnerFaxRequired, 0
	pcs_ValidateEmailField "pcPartnerEmail", pcv_isPartnerEmailRequired, 0
	
	'// Partner Address Info
	pcs_ValidateTextField "pcPartnerAddress", pcv_isPartnerAddressRequired, 0
	pcs_ValidateTextField "pcPartnerPostalCode", pcv_isPartnerPostalCodeRequired, 0
	pcs_ValidateTextField "pcPartnerStateCode", pcv_isPartnerStateCodeRequired, 0
	pcs_ValidateTextField "pcPartnerProvince", pcv_isPartnerProvinceRequired, 0
	pcs_ValidateTextField "pcPartnerCity", pcv_isPartnerCityRequired, 0
	pcs_ValidateTextField "pcPartnerCountryCode", pcv_isPartnerCountryCodeRequired, 0
	pcs_ValidateTextField "pcPartnerAddress2", pcv_isPartnerAddress2Required, 0

	'// Misc.
	pcs_ValidateTextField "pcUser", pcv_isUserRequired, 0
	pcs_ValidateTextField "pcPassword", pcv_ispasswordRequired, 0
	pcs_ValidateTextField "pcUserKey", pcv_isUserKeyRequired, 0
	
	'// Export XML
	pcs_ValidateTextField "ExportAdmin", pcv_isExportAdminRequired, 0
	pcs_ValidateTextField "FTPHost", pcv_isFTPHostRequired, 0
	pcs_ValidateTextField "FTPDirectory", pcv_isFTPDirectoryRequired, 0
	pcs_ValidateTextField "FTPUsername", pcv_isFTPUsernameRequired, 0
	pcs_ValidateTextField "FTPPassword", pcv_isFTPPasswordRequired, 0

		
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError&"&idPartner="&pidPartner
	Else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'PartnerID Already in Database
		query="SELECT pcXP_PartnerID FROM pcXMLPartners WHERE pcXP_PartnerID like '"&trim(Session("pcAdminpcUser"))&"';"	
		set rstemp=conntemp.execute(query)
		if err.number <> 0 then
			response.redirect "techErr.asp?error="& Server.Urlencode("Error occurred while checking for duplicate XML Partner IDs: "&Err.Description) 
	 	end if		
		if NOT rstemp.eof then
			response.redirect pcStrPageName&"?msg=The Partner ID you have chosen is already in use by another Partner.&idPartner=" & pidPartner
		end if
		
		'Partner Key Already in Database
		query="SELECT pcXP_Key FROM pcXMLPartners WHERE pcXP_Key like '"&trim(Session("pcAdminpcUserKey"))&"';"	
		set rstemp=conntemp.execute(query)
		if err.number <> 0 then
			response.redirect "techErr.asp?error="& Server.Urlencode("Error occurred while checking for duplicate XML Partner Keys: "&Err.Description) 
	 	end if		
		if NOT rstemp.eof then
			response.redirect pcStrPageName&"?msg=The Partner Key you have chosen is already in use by another Partner.&idPartner=" & pidPartner
		end if
		
		'// Password
		Session("pcAdminpcPassword")=enDeCrypt(Session("pcAdminpcPassword"), scCrypPass)
		if Session("pcAdminFTPPassword")<>"" then
			Session("pcAdminFTPPassword")=enDeCrypt(Session("pcAdminFTPPassword"), scCrypPass)
		end if
		
		if Session("pcAdminExportAdmin")="" then
			Session("pcAdminExportAdmin")="0"
		end if
		if Session("pcAdminExportAdmin")="1" then
			query="UPDATE pcXMLPartners Set pcXP_ExportAdmin=0;"
			set rstemp=conntemp.execute(query)	
			if err.number <> 0 then
				response.redirect "techErr.asp?error="& Server.Urlencode("Error occurred while updating table pcXMLPartners: "&Err.Description) 
			end if
			set rstemp=nothing
		end if
		
		'// Insert Partner
		query="INSERT INTO pcXMLPartners (pcXP_PartnerID, pcXP_Password, pcXP_Key, pcXP_Name, pcXP_Email, pcXP_Company, pcXP_Address, pcXP_Address2, pcXP_City, pcXP_StateCode, pcXP_Province, pcXP_Zip, pcXP_CountryCode, pcXP_Phone, pcXP_Fax, pcXP_Status,pcXP_ExportAdmin,pcXP_FTPHost,pcXP_FTPDirectory,pcXP_FTPUsername,pcXP_FTPPassword) VALUES ('" &Session("pcAdminpcUser")& "','" &Session("pcAdminpcPassword")& "','" &Session("pcAdminpcUserKey")& "','" &Session("pcAdminpcPartnerName")& "','" &Session("pcAdminpcPartnerEmail")& "','" &Session("pcAdminpcPartnerCompany")& "','" &Session("pcAdminpcPartnerAddress")& "','" &Session("pcAdminpcPartnerAddress2")& "','" &Session("pcAdminpcPartnerCity")& "','" &Session("pcAdminpcPartnerStateCode")& "','" &Session("pcAdminpcPartnerProvince")& "','" &Session("pcAdminpcPartnerPostalCode")& "','" &Session("pcAdminpcPartnerCountryCode")& "','" &Session("pcAdminpcPartnerPhone")& "','" &Session("pcAdminpcPartnerFax")& "'," &Session("pcAdminpcUserStatus")& "," & Session("pcAdminExportAdmin") & ",'" & Session("pcAdminFTPHost") & "','" & Session("pcAdminFTPDirectory") & "','" & Session("pcAdminFTPUsername") & "','" & Session("pcAdminFTPPassword") & "');"
		
		set rstemp=conntemp.execute(query)	
		if err.number <> 0 then
			response.redirect "techErr.asp?error="& Server.Urlencode("Error occurred while adding new Partner into database: "&Err.Description) 
		end if
		
		set rstemp=nothing	
		call closedb()	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Clear the sessions
		pcs_ClearAllSessions
		
		'// Redirect
		response.redirect "AdminManageXMLPartner.asp?msg=added"
		
	End If	
END IF	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

if Session("adminCountryCode")="" then
	Session("adminCountryCode")=scShipFromPostalCountry
end if

if Session("adminshippingCountryCode")="" then
	Session("adminshippingCountryCode")=scShipFromPostalCountry
end if

%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<script>

function GenPartnerKey()
{
var tmp1="PXP-";
var i=0;
	for (i=1;i<=26;i++)
	{
		tmpCType=Math.floor(Math.random()*2);
		if (eval(tmpCType)==0)
		{
			tmp1=tmp1 + "" + Math.floor(Math.random()*10);
		}
		else
		{
			if (eval(tmpCType)==1)
			{
				tmp1=tmp1 + String.fromCharCode(Math.floor(Math.random()*26)+65);
			}
		}
	}

	return(tmp1);
}

function GetPartnerKey(xfield)
{
	xfield.value=GenPartnerKey();
}

</script>

<form method="post" name="modPartner" action="<%=pcStrPageName%>" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">General Information</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td nowrap><p>Partner Status:</p></td>
			<td><p>
				<select name="pcUserStatus">
					<option value="1" <%if pcf_FillFormField("pcUserStatus", true)="1" then%>selected<%end if%>>Active</option>
					<option value="0" <%if pcf_FillFormField("pcUserStatus", true)="0" then%>selected<%end if%>>Inactive</option>
					<option value="2" <%if pcf_FillFormField("pcUserStatus", true)="2" then%>selected<%end if%>>Locked</option>
					<option value="3" <%if pcf_FillFormField("pcUserStatus", pcv_isUserStatusRequired)="3" then%>selected<%end if%>>Suspended</option>
				</select>
				</p>
			</td>
		</tr>
		<tr>
			<td><p>
				Partner ID:
			</p></td>
			<td><p>
				<input type="text" name="pcUser" value="<% =pcf_FillFormField ("pcUser", pcv_isPartnerNameRequired) %>" size="20" />
				<%pcs_RequiredImageTag "pcUser", pcv_isUserRequired %>
			</p></td>
		</tr>
		<tr> 
			<td><p>Password:</p></td>
			<td><p>
				<input type="password" name="pcPassword" value="<% =pcf_FillFormField ("pcPassword", pcv_ispasswordRequired) %>" size="25" maxlength="50">
				<%pcs_RequiredImageTag "pcPassword", pcv_isPasswordRequired %>
			</p>
			</td>
		</tr>
		<tr>
			<td valign="top"><p>
				Partner Key:
			</p></td>
			<td><p>
				<input type="text" name="pcUserKey" value="<% =pcf_FillFormField ("pcUserKey", pcv_isPartnerNameRequired) %>" size="40" />
				<%pcs_RequiredImageTag "pcUserKey", pcv_isUserKeyRequired %>
				<br>
				<input type="button" name="genKey" value="Generate Key" onclick="javascript:GetPartnerKey(document.modPartner.pcUserKey);" class="submit2">
			</p></td>
		</tr>
		<tr>
			<td><p>
				Partner Name:
			</p></td>
			<td><p>
				<input type="text" name="pcPartnerName" value="<% =pcf_FillFormField ("pcPartnerName", pcv_isPartnerNameRequired) %>" size="20" />
				<%pcs_RequiredImageTag "pcPartnerName", pcv_isPartnerNameRequired %>
			</p></td>
		</tr>
		<tr>
			<td><p>
				Company:
			</p></td>
			<td><p>
				<input type="text" name="pcPartnerCompany" value="<% =pcf_FillFormField ("pcPartnerCompany", pcv_isPartnerCompanyRequired) %>" size="30" />
				<%pcs_RequiredImageTag "pcPartnerCompany", pcv_isPartnerCompanyRequired %>
			</p></td>
		</tr>
		
		
		<%	'// Phone Custom Error
		if session("ErrpcPartnerPhone")<>"" then %>
		<tr>
			<td>&nbsp;</td>
			<td><img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%></td>
		</tr>
		<% 
		session("ErrpcPartnerPhone")=""
		end if 
		%>
					
		<tr>
			<td><p>
				Phone:
			</p></td>
			<td><p>
				<input type="text" name="pcPartnerPhone" value="<% =pcf_FillFormField ("pcPartnerPhone", pcv_isPartnerPhoneRequired) %>" size="15" />
				<%pcs_RequiredImageTag "pcPartnerPhone", pcv_isPartnerPhoneRequired %>
			</p></td>
		</tr>
		<tr>
			<td><p>
				Fax:
			</p></td>
			<td><p>
				<input type="text" name="pcPartnerFax" value="<% =pcf_FillFormField ("pcPartnerFax", pcv_isPartnerFaxRequired) %>" size="15" />
				<%pcs_RequiredImageTag "pcPartnerFax", pcv_isPartnerFaxRequired %>
			</p></td>
		</tr>
		
		<% if Session("pcAdminErremail")<>"" then %>
			<tr> 
				<td>&nbsp;</td>
				<td><img src="images/next.gif" width="10" height="10"> <%=Session("pcAdminErremail")%></td>
			</tr>
		<% end if %>
		<tr> 
			<td><p>E-mail:</p></td>
			<td>
			<p>
				<input type="text" name="pcPartnerEmail" value="<% =pcf_FillFormField ("pcPartnerEmail", pcv_isPartnerEmailRequired) %>" size="25" maxlength="150">
				<%pcs_RequiredImageTag "pcPartnerEmail", pcv_isPartnerEmailRequired %>
			</p>
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>               
		<tr> 
			<th colspan="2">Partner Address</th>
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
		pcv_isStateCodeRequired = pcv_isPartnerStateCodeRequired '// determines if validation is performed (true or false)
		pcv_isProvinceCodeRequired = pcv_isPartnerProvinceRequired '// determines if validation is performed (true or false)
		pcv_isCountryCodeRequired = pcv_isPartnerCountryCodeRequired '// determines if validation is performed (true or false)					
		
		'// #3 Additional Required Info
		pcv_strTargetForm = "modPartner" '// Name of Form
		pcv_strCountryBox = "pcPartnerCountryCode" '// Name of Country Dropdown
		pcv_strTargetBox = "pcPartnerStateCode" '// Name of State Dropdown
		pcv_strProvinceBox =  "pcPartnerProvince" '// Name of Province Field
		
		'// Set local Country to Session
		if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
			Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
		end if
		
		'// Set local State to Session
		if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
			Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session(pcv_strSessionPrefix&pcv_strTargetBox)
		end if
		
		'// Set local Province to Session
		if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
			Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session(pcv_strSessionPrefix&pcv_strProvinceBox)
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
		
		
		<tr>
			<td><p>
				Address:
			</p></td>
			<td><p>
				<input type="text" name="pcPartnerAddress" value="<% =pcf_FillFormField ("pcPartnerAddress", pcv_isPartnerAddressRequired) %>" size="30" />
				<%pcs_RequiredImageTag "pcPartnerAddress", pcv_isPartnerAddressRequired %>
			</p></td>
		</tr>
		<tr>
			<td><p>&nbsp;</p></td>
			<td><p>
				<input type="text" name="pcPartnerAddress2" value="<% =pcf_FillFormField ("pcPartnerAddress2", pcv_isPartnerAddress2Required) %>" size="30" />
				<%pcs_RequiredImageTag "pcPartnerAddress2", pcv_isPartnerAddress2Required %>
			</p></td>
		</tr>
		<tr>
			<td><p>
				City:
			</p></td>
			<td><p>
				<input type="text" name="pcPartnerCity" value="<% =pcf_FillFormField ("pcPartnerCity", pcv_isPartnerCityRequired) %>" size="30" />
				<%pcs_RequiredImageTag "pcPartnerCity", pcv_isPartnerCityRequired %>
			</p></td>
		</tr>
		
		<%
		'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
		pcs_StateProvince
		%>
		
		<tr>
			<td><p>
				Postal Code:
			</p></td>
			<td><p>
				<input type="text" name="pcPartnerPostalCode" value="<% =pcf_FillFormField ("pcPartnerPostalCode", pcv_isPartnerPostalCodeRequired) %>" size="10" />
				<%pcs_RequiredImageTag "pcPartnerPostalCode", pcv_isPartnerPostalCodeRequired %>
			</p></td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">XML Export Settings</th>
		</tr>  
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td align="right">
				<input type="checkbox" name="ExportAdmin" value="1" class="clearBorder" <%if pcf_FillFormField ("ExportAdmin", pcv_isExportAdminRequired)="1" then%>checked<%end if%> />
				<%pcs_RequiredImageTag "ExportAdmin", pcv_isExportAdminRequired %>
			</td>
			<td>Make this account my &quot;Export to XML&quot; administrator account</td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">FTP Server Information</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>Host:</td>
			<td>
				<input type="text" name="FTPHost" value="<% =pcf_FillFormField ("FTPHost", pcv_isFTPHostRequired) %>" size="30">
				<%pcs_RequiredImageTag "FTPHost", pcv_isFTPHostRequired %>
			</td>
		</tr>
		<tr>
			<td>Directory:</td>
			<td>
				<input type="text" name="FTPDirectory" value="<% =pcf_FillFormField ("FTPDirectory", pcv_isFTPDirectoryRequired) %>" size="30">
				<%pcs_RequiredImageTag "FTPDirectory", pcv_isFTPDirectoryRequired %>
			</td>
		</tr>
		<tr>
			<td>User Name:</td>
			<td>
				<input type="text" name="FTPUsername" value="<% =pcf_FillFormField ("FTPUsername", pcv_isFTPUsernameRequired) %>" size="30">
				<%pcs_RequiredImageTag "FTPUsername", pcv_isFTPUsernameRequired %>
			</td>
		</tr>
		<tr>
			<td>Password:</td>
			<td>
				<input type="password" name="FTPPassword" value="<% =pcf_FillFormField ("FTPPassword", pcv_isFTPPasswordRequired) %>" size="30">
				<%pcs_RequiredImageTag "FTPPassword", pcv_isFTPPasswordRequired %>
			</td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"> 
				<input type="submit" name="Modify" value="Add XML Partner" class="submit2">&nbsp;
				<input type="button" name="Main" value="Manage XML Partners" onClick="location.href='AdminManageXMLPartner.asp'">&nbsp;
				<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<%
call closedb()
'// Clear the sessions
pcs_ClearAllSessions
%><!--#include file="AdminFooter.asp"-->