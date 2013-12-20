<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Reward Points - General Settings" %>
<% Section = "specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<% 
pcStrPageName="RPSettings.asp"

dim query, mySQL, conntemp, rstemp

'/////////////////////////////////////////////////////
'// Retrieve current database data
'/////////////////////////////////////////////////////
%>
<!--#include file="pcAdminRetrieveSettings.asp"-->
<%
pcv_isRewardsActiveRequired=false
pcv_isRewardsIncludeWholesaleRequired=false
pcv_isRewardsPercentRequired=false
pcv_isRewardsLabelRequired=false
pcv_isRewardsReferralRequired=false
pcv_isRewardsFlatValueRequired=false
pcv_isRewardsPercValueRequired=false

' End update referrer list
if request("Submit")="Save" then
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = "One of more fields were not filled in correctly."
	
	'// validate all fields
	pcs_ValidateTextField	"RewardsActive", pcv_isRewardsActiveRequired, 2
	pcs_ValidateTextField	"RewardsIncludeWholesale", pcv_isRewardsIncludeWholesaleRequired, 2
	pcs_ValidateTextField	"RewardsPercent", pcv_isRewardsPercentRequired, 20
	pcs_ValidateTextField	"RewardsLabel", pcv_isRewardsLabelRequired, 250
	pcs_ValidateTextField	"RewardsReferral", pcv_isRewardsReferralRequired, 2
	pcs_ValidateTextField	"RewardsFlatValue", pcv_isRewardsFlatValueRequired, 20
	pcs_ValidateTextField	"RewardsPercValue", pcv_isRewardsPercValueRequired, 20

	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError& "&lmode=" & pcLoginMode
	End If
	
	'/////////////////////////////////////////////////////
	'// Set Local Variables for Setting
	'/////////////////////////////////////////////////////
	pcIntRewardsFlat=0
	pcIntRewardsPerc=0
	pcIntRewardsPercValue =0
	pcIntRewardsFlatValue = 0

	pcIntRewardsActive = Session("pcAdminRewardsActive")
	pcIntRewardsIncludeWholesale = Session("pcAdminRewardsIncludeWholesale")
	pcIntRewardsPercent = replace(Session("pcAdminRewardsPercent"),"%","")
	pcStrRewardsLabel = Session("pcAdminRewardsLabel")
	pcIntRewardsReferral = Session("pcAdminRewardsReferral")
	pcIntRewardsFlatValue = Session("pcAdminRewardsFlatValue")
	pcIntRewardsPercValue = replace(Session("pcAdminRewardsPercValue"),"%","")

	if pcIntRewardsReferral=0 then
		pcIntRewardsPercValue =0
		pcIntRewardsFlatValue = 0
	end if
	
	if pcIntRewardsReferral = 1 then
		pcIntRewardsFlat = 1
		pcIntRewardsReferral= 1
		pcIntRewardsPerc = 0
		pcIntRewardsPercValue = 0
		if pcIntRewardsFlatValue = "" then
			pcIntRewardsFlatValue=0
		end if
	end if
	
	if pcIntRewardsReferral=2 then
		pcIntRewardsPerc=1
		pcIntRewardsReferral=1
		pcIntRewardsFlat=0
		pcIntRewardsFlatValue = 0
		if pcIntRewardsPercValue = "" then
			pcIntRewardsPercValue=0
		end if
	end if
	
	'/////////////////////////////////////////////////////
	'// Update database with new Settings
	'/////////////////////////////////////////////////////
	%>
	<!--#include file="pcAdminSaveSettings.asp"-->
	<!--#include file="pcAdminRetrieveSettings.asp"-->
    
    <%
    msg="Reward Points settings updated successfully."
    msgtype=1
	%>

<% end if %>

<form name="form1" method="post" action="<%=pcStrPageName%>" class="pcForms">
<table class="pcForms">
    <tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr> 
		<td nowrap><strong>Program is active</strong>:</td>
        <td>
		<input name="RewardsActive" type="radio" value="1" <%if pcIntRewardsActive=1 Then Response.Write "checked"%> class="clearBorder"> Yes 
		<input type="radio" name="RewardsActive" value="0" <%if pcIntRewardsActive=0 Then Response.Write "checked"%> onClick="javascript:alert('Are you sure you want to inactivate Reward Points?  If you do, customers will no longer be able to use their accrued reward points.')" class="clearBorder"> No</td>
	</tr>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td valign="top">Points definition:</td>
        <td><input name="RewardsLabel" type="text" id="RewardsLabel" value="<%=pcStrRewardsLabel%>">
        <div class="pcSmallText" style="padding-top: 6px;">This text is shown next to the number of reward points on pages such as product details pages, order details pages, etc.</div></td>
	</tr>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td valign="top">Conversion rate:</td>
        <td><input name="RewardsPercent" type="text" id="RewardsPercent" value="<%=pcIntRewardsPercent%>%"> (e.g. 10%)
        <div class="pcSmallText" style="padding-top: 6px;">Indicates how points translate into dollars. 100% means that each point equals one dollar. 150% means that each point translates into 1.5 dollars. 20% indicates that each point is equal to $0.20. And so on.</div></td>
	</tr>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>		
	<tr> 
		<td valign="top">Referral points:</td>
        <td>
        <input type="radio" name="RewardsReferral" value="0" <%if pcIntRewardsReferral=0 then Response.Write "checked"%> class="clearBorder"> No
        <div style="padding-top: 6px;">
        <input name="RewardsReferral" type="radio" value="1" <%if pcIntRewardsFlat=1 then Response.Write "checked"%> class="clearBorder"> Yes, flat amount. Amount:	<input name="RewardsFlatValue" type="text" id="RewardsFlatValue" size="15" value="<%=pcIntRewardsFlatValue%>">
        </div>
        <div style="padding-top: 6px;">
        <input type="radio" name="RewardsReferral" value="2" <%if pcIntRewardsPerc=1 then Response.Write "checked"%> class="clearBorder"> Yes, percentage of referred customer's first order. Value: <input name="RewardsPercValue" type="text" id="RewardsPercValue" size="15" value="<%=pcIntRewardsPercValue%>%">
        </div>
        <div class="pcSmallText" style="padding-top: 6px;">Indicates how many points a customer receives when he/she refers to the store a new customer. Points are awarded based on a flat value (e.g. 75) or percentage value (e.g. 25% * order amount). Points are awarded only if the referred customer makes a purchase during its first visit.</div>
        </td>
	</tr>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td nowrap>Include wholesale customers:</td>
        <td>
        <input name="RewardsIncludeWholesale" type="radio" value="1" <%if pcIntRewardsIncludeWholesale=1 then Response.Write "checked"%> class="clearBorder"> Yes <input name="RewardsIncludeWholesale" type="radio" value="0" <%if pcIntRewardsIncludeWholesale=0 then Response.Write "checked"%> class="clearBorder">No
        </td>
	</tr>	
	<tr> 
		<td colspan="2"><hr></td>
	</tr>
	<tr> 
		<td align="center" colspan="2"> 		
			<input type="submit" name="Submit" value="Save" class="submit2">&nbsp;
			<input type="button" name="back" value="Back" onClick="document.location.href='RpStart.asp'">
        </td>
	</tr>
</table>
</form>
<!--#include file="Adminfooter.asp"-->