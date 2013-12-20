<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Affiliate Program - General Settings" %>
<% Section = "misc" %>
<%PmAdmin=8%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/pcAffConstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<% pcPageName="pcAffiliateSettings.asp"

dim query, mySQL, conntemp, rstemp

pcv_isAffProgramActiveRequired=true
pcv_isAffAutoApprove=true
pcv_isAffDefaultCom=false
pcv_isSaveAffiliateRequired=true
pcv_isSaveAffiliateDaysRequired=false
pcv_isAllowedAffOrdersRequired=false

if request("Submit")="Save" then
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = "One of more fields were not filled in correctly."
	
	'// validate all fields
	pcs_ValidateTextField	"AffProgramActive", pcv_isAffProgramActiveRequired, 2
	pcs_ValidateTextField	"AffAutoApprove", pcv_isAffAutoApproveRequired, 2
	pcs_ValidateTextField	"AffDefaultCom", pcv_isAffDefaultComRequired, 5
	pcs_ValidateTextField	"SaveAffiliate", pcv_isSaveAffiliateRequired, 2
	pcs_ValidateTextField	"SaveAffiliateDays", pcv_isSaveAffiliateDaysRequired, 4
	pcs_ValidateTextField	"AllowedAffOrders", pcv_isAllowedAffOrdersRequired, 4
	pcs_ValidateTextField	"ExcludeWholesaleAff", pcv_isExcludeWholesaleAff, 2

	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError& "&lmode=" & pcLoginMode
	End If
	
	'/////////////////////////////////////////////////////
	'// Set Local Variables for Setting
	'/////////////////////////////////////////////////////
	pcIntAffProgramActive = Session("pcAdminAffProgramActive")
	pcIntAffAutoApprove = Session("pcAdminAffAutoApprove")
	pcIntAffDefaultCom = Session("pcAdminAffDefaultCom")
	pcIntSaveAffiliate = session("pcAdminSaveAffiliate")
	pcIntSaveAffiliateDays = session("pcAdminSaveAffiliateDays")
	pcIntAllowedAffOrders = session("pcAdminAllowedAffOrders")
	pcIntExcludeWholesaleAff = session("pcAdminExcludeWholesaleAff")
	
	if NOT isNumeric(pcIntAffDefaultCom) then
		pcIntAffDefaultCom="0"
	end if

	if NOT validNum(pcIntAllowedAffOrders) then
		pcIntAllowedAffOrders="0"
	end if

	if NOT validNum(pcIntSaveAffiliateDays) then
		pcIntSaveAffiliateDays="0"
	end if
	
	if NOT validNum(pcIntExcludeWholesaleAff) then
		pcIntExcludeWholesaleAff="0"
	end if

	'/////////////////////////////////////////////////////
	'// Write all changes to Settings.asp file
	'/////////////////////////////////////////////////////
	Dim objFS
	Dim objFile
	
	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
	
	if PPD="1" then
		pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/pcAffConstants.asp")
	else
		pcStrFileName=Server.Mappath ("../includes/pcAffConstants.asp")
	end if
	
	Set objFile = objFS.OpenTextFile (pcStrFileName, 2, True, 0)
	objFile.WriteLine CHR(60)&CHR(37)&"'// Affiliate Program Settings //" & vbCrLf
	objFile.WriteLine "private const scAffProgramActive = """&pcIntAffProgramActive&"""" & vbCrLf
	objFile.WriteLine "private const scAffAutoApprove = """&pcIntAffAutoApprove&"""" & vbCrLf
	objFile.WriteLine "private const scAffDefaultCom = """&pcIntAffDefaultCom&"""" & vbCrLf
	objFile.WriteLine "private const scSaveAffiliate = """&pcIntSaveAffiliate&"""" & vbCrLf
	objFile.WriteLine "private const scSaveAffiliateDays = """&pcIntSaveAffiliateDays&"""" & vbCrLf
	objFile.WriteLine "private const scAllowedAffOrders = """&pcIntAllowedAffOrders&"""" & vbCrLf
	objFile.WriteLine "private const scExcludeWholesaleAff = """&pcIntExcludeWholesaleAff&"""" & vbCrLf
	objFile.WriteLine "'// Affiliate Program Settings // " &CHR(37)&CHR(62)& vbCrLf
	objFile.Close
	set objFS=nothing
	set objFile=nothing
	msgSuccess = "Affiliate settings updated successfully"
	response.redirect "pcAffiliateSettings.asp?s=1&msg=" & msgSuccess
end if 

%>
<script language="JavaScript">
<!--
function Form1_Validator(theForm)
{
if (allDigit(theForm.AffDefaultCom.value) == false)
{
alert('Please enter a valid number for the default commission field.');
theForm.AffDefaultCom.focus();
return (false);
}
return (true);
}

	
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

<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms" onSubmit="return Form1_Validator(this)">
    <table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
        <tr> 
          <td nowrap>Affiliate program is active:</td>
          <td width="80%"><input name="AffProgramActive" type="radio" value="1" <%if scAffProgramActive="1" Then Response.Write "checked"%> class="clearBorder"> Yes <input type="radio" name="AffProgramActive" value="0" <%if scAffProgramActive="0" Then Response.Write "checked"%> onClick="javascript:alert('Are you sure you want to inactivate the Affiliate Program?')" class="clearBorder"> No</td>
        </tr>
        <tr> 
          <td nowrap>Automatically approve affiliates:</td>
          <td><input type="radio" name="AffAutoApprove" value="1" <%if scAffAutoApprove="1" Then Response.Write "checked"%> onClick="javascript:alert('Are you sure you want to automatically approve new affiliate accounts?')" class="clearBorder"> Yes <input type="radio" name="AffAutoApprove" value="0" <%if scAffAutoApprove="0" Then Response.Write "checked"%> class="clearBorder"> No &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=210')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
        </tr>
        <tr>
          <td>Default commission:</td>
          <td><input name="AffDefaultCom" type="text" id="AffDefaultCom" value="<%=scAffDefaultCom%>" size="6" maxlength="6"> % (e.g. for a 20% commission, enter 20) &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=211')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
        </tr>
        <tr>
          <td nowrap>Exclude wholesale customers:</td>
          <td><input name="ExcludeWholesaleAff" type="radio" id="ExcludeWholesaleAff" value="1" <%if scExcludeWholesaleAff="1" then Response.Write "checked"%> class="clearBorder"> Yes <input name="ExcludeWholesaleAff" type="radio" id="ExcludeWholesaleAff" value="0" <%if scExcludeWholesaleAff="0" then Response.Write "checked"%> class="clearBorder"> No &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=212')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
        </tr>
        <tr>
          <td>Save Affiliate ID:</td>
           <td><input type="radio" name="SaveAffiliate" value="0" <%if scSaveAffiliate="0" then Response.Write "checked"%> class="clearBorder"> No &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=208')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
        </tr>
        <tr> 
          <td></td>
          <td><input type="radio" name="SaveAffiliate" value="1" <%if scSaveAffiliate="1" then Response.Write "checked"%> class="clearBorder"> Yes - Enter the number of days after which the cookie expires:
            <input name="SaveAffiliateDays" type="text" id="scSaveAffiliateDays" value="<%=scSaveAffiliateDays %>" size="6" maxlength="6">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=208')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
          </td>
        </tr>						
        <tr>
          <td nowrap>Max number of orders per customer:</td>
          <td><input name="AllowedAffOrders" type="text" id="AllowedAffOrders" value="<%=scAllowedAffOrders%>" size="4" maxlength="4"> order(s) on which commissions will be earned &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=209')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
        </tr>
        <tr> 
            <td colspan="2"><hr></td>
        </tr>
        <tr> 
            <td align="center" colspan="2"> 		
                <input type="submit" name="Submit" value="Save" class="submit2">&nbsp;
                <input type="button" name="back" value="Back" onClick="javascript:history.back()">
            </td>
        </tr>
    </table>
</form>
<!--#include file="Adminfooter.asp"-->