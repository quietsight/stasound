<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="AffLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/rc4.asp" -->
<!--#include file="header.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<%


'// Check if store is turned off and return message to customer
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If 

' Load affiliate ID
affVar=session("pc_idaffiliate")
if not validNum(affVar) then
	response.redirect "AffiliateLogin.asp"
end if

dim mySQL, conntemp, rstemp

pcStrPageName = "pcmodAffa.asp"

'// Set Required Fields
pcv_isnameRequired=true
pcv_iscompanyRequired=false
pcv_isemailRequired=true
pcv_ispasswordRequired=true
pcv_iscountryRequired=true

'// Use the Request object to toggle State (based of Country selection)
pcv_isstateRequired=true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isstateRequired=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
pcv_isprovinceRequired=false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isprovinceRequired=pcv_strProvinceCodeRequired
end if

pcv_isaddressRequired=true
pcv_isaddress2Required=false  
pcv_iscityRequired=true
pcv_iszipRequired=true
pcv_isphoneRequired=true
pcv_isfaxRequired=false
pcv_iswebsiteRequired=false
%>
<% 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf
response.write "<!--"&vbcrlf	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf
pcs_JavaTextField	"name", pcv_isnameRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"company", pcv_iscompanyRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"email", pcv_isemailRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"password", pcv_ispasswordRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"country", pcv_iscountryRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
'// Do not show, invisible controls
'pcs_JavaTextField	"state", pcv_isstateRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
'pcs_JavaTextField	"province", pcv_isprovinceRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"address", pcv_isaddressRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"city", pcv_iscityRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"zip", pcv_iszipRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"phone", pcv_isphoneRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"fax", pcv_isfaxRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"website", pcv_iswebsiteRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "//-->"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

call openDB()


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IF request.Form("Submit.x")<>"" THEN
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	

	pcs_ValidateTextField	"name", pcv_isnameRequired, 150
	pcs_ValidateTextField	"company", pcv_iscompanyRequired, 150
	pcs_ValidateEmailField	"email", pcv_isemailRequired, 50
	pcs_ValidateTextField	"password", pcv_ispasswordRequired, 100
	pcs_ValidateTextField	"address", pcv_isaddressRequired, 70
	pcs_ValidateTextField	"address2", pcv_isaddress2Required, 150 
	pcs_ValidateTextField	"country", pcv_iscountryRequired, 150
	pcs_ValidateTextField	"state", pcv_isstateRequired, 150
	pcs_ValidateTextField	"province", pcv_isprovinceRequired, 150
	pcs_ValidateTextField	"city", pcv_iscityRequired, 150
	pcs_ValidateTextField	"zip", pcv_iszipRequired, 12
	pcs_ValidatePhoneNumber	"phone", pcv_isphoneRequired, 30
	pcs_ValidatePhoneNumber	"fax", pcv_isfaxRequired, 30
	pcs_ValidateTextField	"website", pcv_iswebsiteRequired, 150

	'// run additional checks and functions on the our sessions
	'if NOT validNum(Session("pcSFzip")) then
	'	Session("pcSFzip")=0
	'end if	
	
	
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError
	Else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		If Session("pcSFprovince")<>"" then
			pcv_strStateOrProvince = Session("pcSFprovince")
		Else
			pcv_strStateOrProvince = Session("pcSFstate")
		End If
		
		Session("pcSFPassword")=enDeCrypt(Session("pcSFPassword"), scCrypPass)
		
		mySQL="UPDATE affiliates SET affiliateName='" &Session("pcSFname")& "', affiliateEmail='" &Session("pcSFemail")& "'"
		
		if Session("pcSFcompany") <> "" then
			mySQL=mySQL & ", affiliatecompany='" &Session("pcSFcompany")& "'"
		end if
			mySQL=mySQL & ", affiliateaddress='" &Session("pcSFaddress")& "'"
		
		if Session("pcSFaddress2") <> "" then
			mySQL=mySQL & ", affiliateaddress2='" &Session("pcSFaddress2")& "'"
		end if
			mySQL=mySQL & ", affiliatecity='" &Session("pcSFcity")& "', affiliatestate='" &pcv_strStateOrProvince& "', affiliateCountryCode='" &Session("pcSFcountry")& "', affiliatezip='" &Session("pcSFzip")& "'"
		
		if Session("pcSFphone") <> "" then
				mySQL=mySQL & ", affiliatephone='" &Session("pcSFphone")& "'"
		end if
		if Session("pcSFfax") <> "" then
				mySQL=mySQL & ", affiliatefax='" &Session("pcSFfax")& "'"
		end if
		if Session("pcSFwebsite") <> "" then
				mySQL=mySQL & ", pcAff_website='" &Session("pcSFwebsite")& "'"
		end if
		
		mySQL=mySQL & ",pcAff_Password='" & Session("pcSFpassword") & "' WHERE idaffiliate=" & session("pc_IDAffiliate")
		
		set rstemp=conntemp.execute(mySQL)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	

		'// Clear the sessions
		pcs_ClearAllSessions	
		

		call closeDB()

		response.redirect "AffiliateMain.asp?msg=" & dictLanguage.Item(Session("language")&"_ModAffb_1")
		
		
	End If
END IF	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



mySQL="SELECT * FROM Affiliates WHERE Affiliates.idAffiliate=" & session("pc_IDAffiliate")
set rstemp=conntemp.execute(mySQL)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
		
pIdAffiliate=session("pc_IDAffiliate")
Session("pcSFname")= pcf_ResetFormField(Session("pcSFname"), rstemp("affiliateName"))
Session("pcSFemail")= pcf_ResetFormField(Session("pcSFemail"), rstemp("affiliateEmail"))
Session("pcSFcompany")= pcf_ResetFormField(Session("pcSFcompany"), rstemp("affiliatecompany"))
Session("pcSFaddress")= pcf_ResetFormField(Session("pcSFaddress"), rstemp("affiliateaddress"))
Session("pcSFaddress2")= pcf_ResetFormField(Session("pcSFaddress2"), rstemp("affiliateaddress2"))
Session("pcSFcity")= pcf_ResetFormField(Session("pcSFcity"), rstemp("affiliatecity"))
Session("pcSFstate")= pcf_ResetFormField(Session("pcSFstate"), rstemp("affiliatestate"))
Session("pcSFprovince")= pcf_ResetFormField(Session("pcSFprovince"), rstemp("affiliatestate"))
Session("pcSFcountry")= pcf_ResetFormField(Session("pcSFcountry"), rstemp("affiliateCountryCode"))
Session("pcSFphone")= pcf_ResetFormField(Session("pcSFphone"), rstemp("affiliatephone"))
Session("pcSFfax")= pcf_ResetFormField(Session("pcSFfax"), rstemp("affiliatefax"))
Session("pcSFzip")= pcf_ResetFormField(Session("pcSFzip"), rstemp("affiliatezip"))
pcommission = rstemp("commission")
if Session("pcSFpassword") = "" then
	Session("pcSFpassword")= rstemp("pcAff_Password")
	Session("pcSFpassword")= enDeCrypt(Session("pcSFpassword"), scCrypPass)
end if
Session("pcSFwebsite")= pcf_ResetFormField(Session("pcSFwebsite"), rstemp("pcAff_website"))
%>
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td> 
			<h1><%response.write dictLanguage.Item(Session("language")&"_NewAffa_1")%></h1>
			<p><%response.write dictLanguage.Item(Session("language")&"_NewAffa_1b")%></p>
		</td>
	</tr>
	<tr>
		<td>
		<form method="post" name="addaffiliate" action="<%=pcStrPageName%>" onSubmit="return Form1_Validator(this)" class="pcForms">
			
			<table class="pcShowContent">
				<tr class="normal"> 
					<td height="21"><p><%=dictLanguage.Item(Session("language")&"_NewAffa_2")%></p></td>
					<td height="21">  
						<p><input type="text" name="name" value="<% =pcf_FillFormField ("name", pcv_isnameRequired) %>" size="30" maxlength="50"> 
						<% pcs_RequiredImageTag "name", pcv_isnameRequired %></p>
					</td>
				</tr>
									  
				<tr class="normal"> 
					<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_3")%></p></td>
					<td><p> 
						<input type="text" name="company" value="<% =pcf_FillFormField ("company", pcv_iscompanyRequired) %>" size="30" maxlength="50"> 
						<% pcs_RequiredImageTag "company", pcv_iscompanyRequired %>
						</p>
					</td>
				</tr>
				
				<%	'// Email Custom Error
				if session("Erremail")<>"" then %>
					<tr> 
						<td>&nbsp;</td>
						<td>
						<img src="<%=pcf_GenerateIconURL(rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_16")%>
						</td>
					</tr>
					<% 
					session("Erremail") = ""
				end if 
				%>
									
				<tr class="normal"> 
					<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_4")%></p></td>
					<td><p>
						<input type="text" name="email" value="<% =pcf_FillFormField ("email", pcv_isemailRequired) %>" size="30" maxlength="150">
						<% pcs_RequiredImageTag "email", pcv_isemailRequired %>
						</p>
					</td>
				</tr>
									  
				<tr class="normal"> 
					<td>
						<p><%=dictLanguage.Item(Session("language")&"_NewAffa_5")%></p>
					</td>
					<td><p>
						<input type="password" name="password" value="<% =pcf_FillFormField ("password", pcv_ispasswordRequired) %>" size="30" maxlength="150">
						<% pcs_RequiredImageTag "password", pcv_ispasswordRequired %>
						</p>
					</td>
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
				pcv_isStateCodeRequired = pcv_isstateRequired '// determines if validation is performed (true or false)
				pcv_isProvinceCodeRequired = pcv_isprovinceRequired '// determines if validation is performed (true or false)
				pcv_isCountryCodeRequired = pcv_iscountryRequired '// determines if validation is performed (true or false)
				
				'// #3 Additional Required Info
				pcv_strTargetForm = "addaffiliate" '// Name of Form
				pcv_strCountryBox = "country" '// Name of Country Dropdown
				pcv_strTargetBox = "state" '// Name of State Dropdown
				pcv_strProvinceBox =  "province" '// Name of Province Field
				
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
				 
				<tr class="normal">                         
					<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_6")%></p></td>
					<td><p> 
						<input type="text" name="address" value="<% =pcf_FillFormField ("address", pcv_isaddressRequired) %>" size="30" maxlength="150"> 
						<% pcs_RequiredImageTag "address", pcv_isShipAddressRequired %>
						</p>
					</td>
				</tr>                      
				<tr class="normal">                         
					<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_7")%></p></td>
					<td><p>                       
						<input type="text" name="address2" value="<% =pcf_FillFormField ("address2", pcv_isaddress2Required) %>" size="30" maxlength="150">
						<% pcs_RequiredImageTag "address2", pcv_isaddress2Required %>
						</p>
					</td>
				</tr>                      
				<tr class="normal">                         
				<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_8")%></p></td>
					<td><p>  
						<input type="text" name="city" value="<% =pcf_FillFormField ("city", pcv_iscityRequired) %>" size="20" maxlength="50">
						<% pcs_RequiredImageTag "city", pcv_iscityRequired %>
						</p>
					</td>
				</tr>               
				<%
				'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
				pcs_StateProvince
				%>
				<tr class="normal"> 
					<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_11")%></p></td>
					<td><p>  
						<input type="text" name="zip" value="<% =pcf_FillFormField ("zip", pcv_iszipRequired) %>" size="10" maxlength="50">
						<% pcs_RequiredImageTag "zip", pcv_iszipRequired %>
						</p>
					</td>
				</tr>
				
				<%	'// Phone Custom Error
				if session("Errphone")<>"" then %>
					<tr> 
						<td>&nbsp;</td>
						<td>
						<img src="<%=pcf_GenerateIconURL(rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
						</td>
					</tr>
					<% 
					session("Errphone") = ""
				end if 
				%>
														 
				<tr class="normal"> 
					<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_12")%></p></td>
					<td><p>  
						<input type="text" name="phone" value="<% =pcf_FillFormField ("phone", pcv_isphoneRequired) %>" size="20" maxlength="20"> 
						<% pcs_RequiredImageTag "phone", pcv_isphoneRequired %>
						</p>
					</td>
				</tr>
				
				
				<%	'// Fax Custom Error
				if session("Errfax")<>"" then %>
					<tr> 
						<td>&nbsp;</td>
						<td>
						<img src="<%=pcf_GenerateIconURL(rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
						</td>
					</tr>
					<% 
					session("Errfax") = ""
				end if 
				%>
									 
				<tr class="normal"> 
					<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_13")%></p></td>
					<td><p>  
						<input type="text" name="fax" value="<% =pcf_FillFormField ("fax", pcv_isfaxRequired) %>" size="20" maxlength="20">
						<% pcs_RequiredImageTag "fax", pcv_isfaxRequired %>
						</p>
					</td>
				</tr>
				
				
				<tr class="normal"> 
					<td><p><%=dictLanguage.Item(Session("language")&"_security_1")%></p></td>
					<td><p> 
						<input type="text" name="website" value="<% =pcf_FillFormField ("website", pcv_iswebsiteRequired) %>" size="30" maxlength="50"> 
						<% pcs_RequiredImageTag "website", pcv_iswebsiteRequired %>
						</p>
					</td>
				</tr>
				<tr>
					<td>
					<p><%response.write dictLanguage.Item(Session("language")&"_NewAffa_14")%></p>
					</td>
				 <td><p><%=pcommission%>&nbsp;%</p></td>
				</tr> 
				              
				<tr>
					<td colspan="2" align="center">&nbsp;</td>
				</tr>
				<tr> 									
					<td colspan="2" align="center">  
						<input type="image" src="<%=rslayout("submit")%>" border="0" name="Submit" id="submit">&nbsp;
						<a href="javascript:history.go(-1)"><img src="<%=rslayout("back")%>" border=0></a>
					</td>
				</tr>
				<tr> 
					<td colspan="2">&nbsp;</td>
				</tr>
			</table>
		</form>
	</td>
	</tr>
</table>
</div>
<%
call closeDB()
%>
<!--#include file="Footer.asp"-->