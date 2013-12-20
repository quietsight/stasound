<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Edit Affiliate" %>
<% section="misc" %>
<%PmAdmin=8%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<% 
dim f, mySQL, conntemp, rstemp, pIdAffiliate, pDescription, pDetails, pPrice, pImageUrl, pListPrice, pstock, psku, plisthidden, pactive, pweight

' form parameter 
pIdAffiliate=request.Querystring("idAffiliate")

if trim(pidAffiliate)="" then
   response.redirect "msg.asp?message=7"
end if

pcStrPageName = "modAffa.asp"

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
pcv_iscommissionRequired=true
pcv_isactiveRequired=false
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
pcs_JavaTextField	"active", true, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"address", pcv_isaddressRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"city", pcv_iscityRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"zip", pcv_iszipRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"phone", pcv_isphoneRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"fax", pcv_isfaxRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"website", pcv_iswebsiteRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
response.write "if (theForm.commission.value == '0')"&vbcrlf
response.write "{"&vbcrlf
response.write "alert('Please enter a value greater than zero for this field.');"&vbcrlf
response.write "theForm.commission.focus();"&vbcrlf
response.write "return (false);"&vbcrlf
response.write "}"&vbcrlf
response.write "if (allDigit(theForm.commission.value) == false)"&vbcrlf
response.write "{"&vbcrlf
response.write "alert('Please enter a number for this field.');"&vbcrlf
response.write "theForm.commission.focus();"&vbcrlf
response.write "return (false);"&vbcrlf
response.write "}"&vbcrlf
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
IF request.Form("Submit")<>"" THEN
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
	pcs_ValidateTextField	"active", pcv_isactiveRequired, 150
	pcs_ValidateTextField	"commission", pcv_iswebsiteRequired, 150


	'// run additional checks and functions on the our sessions
	'if NOT validNum(Session("pcAdminzip")) then
	'	Session("pcAdminzip")=0
	'end if	
	
	
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?idAffiliate="& pIDAffiliate &"&msg="&pcv_strGenericPageError
	Else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		if (Session("pcAdminActive")<>"") and (Session("pcAdminActive")="1") then
		else
		Session("pcAdminActive")="0"
		end if
		
		If Session("pcAdminprovince")<>"" then
			pcv_strStateOrProvince = Session("pcAdminprovince")
		Else
			pcv_strStateOrProvince = Session("pcAdminstate")
		End If		
		
		Session("pcAdminPassword")=enDeCrypt(Session("pcAdminPassword"), scCrypPass)
		
		mySQL="UPDATE affiliates SET affiliateName='" &Session("pcAdminname")& "', affiliateEmail='" &Session("pcAdminemail")& "'"
		
		if Session("pcAdmincompany") <> "" then
			mySQL=mySQL & ", affiliatecompany='" &Session("pcAdmincompany")& "'"
		end if
			mySQL=mySQL & ", affiliateaddress='" &Session("pcAdminaddress")& "'"
		
		if Session("pcAdminaddress2") <> "" then
			mySQL=mySQL & ", affiliateaddress2='" &Session("pcAdminaddress2")& "'"
		end if
			mySQL=mySQL & ", affiliatecity='" &Session("pcAdmincity")& "', affiliatestate='" &pcv_strStateOrProvince& "', affiliateCountryCode='" &Session("pcAdmincountry")& "', affiliatezip='" &Session("pcAdminzip")& "'"
		
		if Session("pcAdminphone") <> "" then
				mySQL=mySQL & ", affiliatephone='" &Session("pcAdminphone")& "'"
		end if
		if Session("pcAdminfax") <> "" then
				mySQL=mySQL & ", affiliatefax='" &Session("pcAdminfax")& "'"
		end if
		if Session("pcAdminwebsite") <> "" then
				mySQL=mySQL & ", pcAff_website='" &Session("pcAdminwebsite")& "'"
		end if
		
		mySQL=mySQL & ", commission='" &Session("pcAdmincommission")& "',pcAff_Password='" & Session("pcAdminpassword") & "', pcAff_Active="& Session("pcAdminactive") &" WHERE idaffiliate=" & session("pcadmin_IDAffiliate")
		'response.write mySQL
		'response.end
		
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
	
		set rstemp=nothing
		call closeDB()

		'// Clear the sessions
		pcs_ClearAllSessions
		
		'// Redirect
		response.redirect "AdminAffiliates.asp?s=1&msg=" & server.URLEncode("The selected affiliate account was updated successfully.")
		
	End If
END IF	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



mySQL="SELECT * FROM Affiliates WHERE Affiliates.idAffiliate=" &pIdAffiliate
set rstemp=conntemp.execute(mySQL)
if err.number <> 0 then
    response.redirect "techErr.asp?error="& Server.Urlencode("Error in loaditemform: "&Err.Description) 
end if

session("pcadmin_IDAffiliate")=rstemp("idAffiliate")
Session("pcAdminname")= pcf_ResetFormField(Session("pcAdminname"), rstemp("affiliateName"))
Session("pcAdminemail")= pcf_ResetFormField(Session("pcAdminemail"), rstemp("affiliateEmail"))
Session("pcAdmincompany")= pcf_ResetFormField(Session("pcAdmincompany"), rstemp("affiliatecompany"))
Session("pcAdminaddress")= pcf_ResetFormField(Session("pcAdminaddress"), rstemp("affiliateaddress"))
Session("pcAdminaddress2")= pcf_ResetFormField(Session("pcAdminaddress2"), rstemp("affiliateaddress2"))
Session("pcAdmincity")= pcf_ResetFormField(Session("pcAdmincity"), rstemp("affiliatecity"))
Session("pcAdminstate")= pcf_ResetFormField(Session("pcAdminstate"), rstemp("affiliatestate"))
Session("pcAdminprovince")= pcf_ResetFormField(Session("pcAdminprovince"), rstemp("affiliatestate"))
Session("pcAdmincountry")= pcf_ResetFormField(Session("pcAdmincountry"), rstemp("affiliateCountryCode"))
Session("pcAdminphone")= pcf_ResetFormField(Session("pcAdminphone"), rstemp("affiliatephone"))
Session("pcAdminfax")= pcf_ResetFormField(Session("pcAdminfax"), rstemp("affiliatefax"))
Session("pcAdminzip")= pcf_ResetFormField(Session("pcAdminzip"), rstemp("affiliatezip"))
Session("pcAdmincommission")= pcf_ResetFormField(Session("pcAdmincommission"), rstemp("commission"))
if Session("pcAdminpassword") = "" then
	Session("pcAdminpassword")= rstemp("pcAff_Password")
	Session("pcAdminpassword")= enDeCrypt(Session("pcAdminpassword"), scCrypPass)
end if

Session("pcAdminwebsite")= pcf_ResetFormField(Session("pcAdminwebsite"), rstemp("pcAff_website"))
	if trim(Session("pcAdminwebsite"))<> "" then
		if instr(Session("pcAdminwebsite"),"http://")=0 and instr(Session("pcAdminwebsite"),"https://")=0 then
			tempURL="http://" & Session("pcAdminwebsite")
			tempURL=replace(tempURL,"//","/")
			tempURL=replace(tempURL,"https:/","https://")
			tempURL=replace(tempURL,"http:/","http://")
			Session("pcAdminwebsite") = tempURL
		end if
	end if
Session("pcAdminActive")= pcf_ResetFormField(Session("pcAdminActive"), rstemp("pcAff_Active"))
if (Session("pcAdminActive")<>"") and (Session("pcAdminActive")="1") then
else
Session("pcAdminActive")="0"
end if
%>
<script language="JavaScript">
<!--
function isDigit(s)
{
var test=""+s;
if(test=="."||test==","||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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
//-->
</script>
<form method="post" name="addaffiliate" action="<%=pcStrPageName%>?idAffiliate=<%=pIDAffiliate%>" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">

<% 
msg=request.QueryString("msg")
if msg<>"" then %>
<tr> 
	<td colspan="2">
		<div class="pcCPmessage"><%=msg%></div>  
	</td>
</tr>   
<% end if %>

<tr> 
	<input type="hidden" name="idAffiliate" value="<%=pIdAffiliate%>">
	<td colspan="2"><p><b><%=Session("pcAdminname")%></b> - Affiliate ID: <%=pidAffiliate%></p></td>
</tr>                      
<tr> 
	<td colspan="2"><p><input type="checkbox" name="active" value="1" <%if Session("pcAdminActive")="1" then%>checked<%end if%> class="clearBorder">&nbsp;Active</p></td>
</tr>
                      
<tr> 
	<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_2")%></p></td>
	<td>  
		<p><input type="text" name="name" value="<%=pcf_FillFormField("name", pcv_isnameRequired)%>" size="30" maxlength="50"> 
		<% pcs_RequiredImageTag "name", pcv_isnameRequired %></p>
	</td>
</tr>
					  
<tr> 
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
		<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> <%=dictLanguage.Item(Session("language")&"_Custmoda_16")%>
		</td>
	</tr>
	<% 
	session("Erremail") = ""
end if 
%>
					
<tr> 
	<td><p><a href="mailto:<%=Session("pcAdminemail")%>"><%=dictLanguage.Item(Session("language")&"_NewAffa_4")%></a></p></td>
	<td><p>
		<input type="text" name="email" value="<% =pcf_FillFormField ("email", pcv_isemailRequired) %>" size="30" maxlength="150">
		<% pcs_RequiredImageTag "email", pcv_isemailRequired %>
		</p>
	</td>
</tr>
					  
<tr> 
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
 
<tr>                         
	<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_6")%></p></td>
	<td><p> 
		<input type="text" name="address" value="<% =pcf_FillFormField ("address", pcv_isaddressRequired) %>" size="30" maxlength="150"> 
		<% pcs_RequiredImageTag "address", pcv_isShipAddressRequired %>
		</p>
	</td>
</tr>                      
<tr>                         
	<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_7")%></p></td>
	<td><p>                       
		<input type="text" name="address2" value="<% =pcf_FillFormField ("address2", pcv_isaddress2Required) %>" size="30" maxlength="150">
		<% pcs_RequiredImageTag "address2", pcv_isaddress2Required %>
		</p>
	</td>
</tr>                      
<tr>                         
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
<tr> 
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
		<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
		</td>
	</tr>
	<% 
	session("Errphone") = ""
end if 
%>
										 
<tr> 
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
		<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
		</td>
	</tr>
	<% 
	session("Errfax") = ""
end if 
%>
					 
<tr> 
	<td><p><%=dictLanguage.Item(Session("language")&"_NewAffa_13")%></p></td>
	<td><p>  
		<input type="text" name="fax" value="<% =pcf_FillFormField ("fax", pcv_isfaxRequired) %>" size="20" maxlength="20">
		<% pcs_RequiredImageTag "fax", pcv_isfaxRequired %>
		</p>
	</td>
</tr>
<tr> 
	<td><p><a href="<%=Session("pcAdminwebsite")%>" target="_blank"><%=dictLanguage.Item(Session("language")&"_NewAffa_15")%></a></p></td>
	<td><p> 
		<input type="text" name="website" value="<%=pcf_FillFormField ("website", pcv_iswebsiteRequired) %>" size="30" maxlength="50"> 
		<% pcs_RequiredImageTag "website", pcv_iswebsiteRequired %>
        &nbsp;
        <span class="pcSmallText">Full URL, starting with &quot;http://&quot; or &quot;https://&quot;.</span>
		</p>
	</td>
</tr>

<tr> 
	<td><p><%response.write dictLanguage.Item(Session("language")&"_NewAffa_14")%></p></td>
	<td>
		<p>
		<input type="text" name="commission" value="<%=pcf_FillFormField("commission", pcv_iscommissionRequired) %>" size="20">
		<% pcs_RequiredImageTag "commission", pcv_iscommissionRequired %>
		&nbsp;
		<span class="pcSmallText">For 20%, enter 20.</span>
		</p>
	</td>
</tr>
                      
<tr>
	<td colspan="2" align="center">&nbsp;</td>
</tr>

<tr>                         
	<td colspan="2" align="center">                            
	<input type="submit" name="Submit" value="Save" class="submit2">
	&nbsp;
	<input type="button" value="Back" onClick="javascript:history.back()" class="ibtnGrey">
     </td>
</tr>

<tr>
	<td colspan="2" align="center">&nbsp;</td>
</tr>                    
</table>
</form>
<!--#include file="AdminFooter.asp"-->