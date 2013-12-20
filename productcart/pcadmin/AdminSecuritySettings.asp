<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<%

Set fso=server.CreateObject("Scripting.FileSystemObject")

if request("action")="update" then

	pcv_Security=request("pcv_Security")
	if pcv_Security="" then
		pcv_Security=0
	end if
	pcv_UserLogin=request("pcv_UserLogin")
	if pcv_UserLogin="" then
		pcv_UserLogin=0
	end if
	pcv_UserReg=request("pcv_UserReg")
	if pcv_UserReg="" then
		pcv_UserReg=0
	end if
	pcv_AffLogin=request("pcv_AffLogin")
	if pcv_AffLogin="" then
		pcv_AffLogin=0
	end if
	pcv_AffReg=request("pcv_AffReg")
	if pcv_AffReg="" then
		pcv_AffReg=0
	end if
	pcv_Review=request("pcv_Review")
	if pcv_Review="" then
		pcv_Review=0
	end if
	pcv_Contact=request("pcv_Contact")
	if pcv_Contact="" then
		pcv_Contact=0
	end if
	pcv_AdminLogin=request("pcv_AdminLogin")
	if pcv_AdminLogin="" then
		pcv_AdminLogin=0
	end if
	pcv_UseImgs=request("pcv_UseImgs")
	if pcv_UseImgs="" then
		pcv_UseImgs=0
	end if
	pcv_UseImgs2=request("pcv_UseImgs2")
	if pcv_UseImgs2="" then
		pcv_UseImgs2=0
	end if
	pcv_AlarmMsg=request("pcv_AlarmMsg")
	if pcv_AlarmMsg="" then
		pcv_AlarmMsg=0
	end if
	pcv_AttackCount=request("pcv_AttackCount")
	if pcv_AttackCount="" then
		pcv_AttackCount=0
	end if

	if PPD="1" then
		findit=Server.MapPath("/"&scPcFolder&"/includes/securitysettings.asp")
	else
		findit=Server.MapPath("../includes/securitysettings.asp")
	end if

	Set f = fso.CreateTextFile(FindIt,True)
	fBody=CHR(60)&CHR(37)
	fBody=fBody&"private const scSecurity=" & pcv_Security & VBCrlf
	fBody=fBody&"private const scUserLogin=" & pcv_UserLogin & VBCrlf
	fBody=fBody&"private const scUserReg=" & pcv_UserReg & VBCrlf
	fBody=fBody&"private const scAffLogin=" & pcv_AffLogin & VBCrlf
	fBody=fBody&"private const scAffReg=" & pcv_AffReg & VBCrlf
	fBody=fBody&"private const scReview=" & pcv_Review & VBCrlf
	fBody=fBody&"private const scContact=" & pcv_Contact & VBCrlf
	fBody=fBody&"private const scAdminLogin=" & pcv_AdminLogin & VBCrlf
	fBody=fBody&"private const scUseImgs=" & pcv_UseImgs & VBCrlf
	fBody=fBody&"private const scUseImgs2=" & pcv_UseImgs2 & VBCrlf
	fBody=fBody&"private const scAlarmMsg=" & pcv_AlarmMsg & VBCrlf
	fBody=fBody&"private const scAttackCount=" & pcv_AttackCount & VBCrlf
	fBody=fBody&CHR(37)&CHR(62)
	f.write fBody
	f.Close
	Set f=nothing

	response.redirect "AdminSecuritySettings.asp?s=1&msg=Security Settings were updated successfully!"

end if

set fso=nothing


pageTitle="Advanced Security Settings" 
pageIcon="pcv4_security.png"
section="layout" 
%>
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="AdminSecuritySettings.asp?action=update" class="pcForms">
	<table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr>
			<th colspan="3">Overview</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td colspan="3">
		<p>Use this feature to tun on and off additional security filters when customers and Control Panel users login. The CAPTCHA feature requires a working XML Parser: <a href="pcTSUtility.asp" target="_blank">review your XML Parser settings</a>. The settings listed below apply only when &quot;Advanced Security&quot; is turned &quot;On&quot;, <u>except for</u> the one for the &quot;Contact Us&quot; form, which works independently of the others.&nbsp;<a href="http://wiki.earlyimpact.com/productcart/settings-security-settings" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="Learn more about this topic"></a></p></td>
	</tr>
		<tr>
			<td colspan="3">
			<p> Turn Advanced Security Settings <input type="radio" name="pcv_Security" value="1" <%if scSecurity=1 then%>checked<%end if%> class="clearBorder">On <input type="radio" name="pcv_Security" value="0" <%if scSecurity<>1 then%>checked<%end if%> class="clearBorder">Off</p>
            </td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="3"><p></p></td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">Storefront</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_UserLogin" value="1" <%if scUserLogin=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to User <strong>Login</strong> pages</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_UserReg" value="1" <%if scUserReg=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to User <strong>Registration</strong> pages</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_UseImgs" value="1" <%if scUseImgs=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%" valign="top">Add <strong>CAPTCHA</strong> (random security code) to <strong>Login/Registration</strong> pages in the storefront.<br /><a href="pcTSUtility.asp" target="_blank" style="color:#666;">Review your XML parser settings</a> to ensure your store is setup to use a XML parser supported by your Web server.</td>
		</tr>
 		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_Review" value="1" <%if scReview=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add <strong>CAPTCHA</strong> (random security code) to the <strong>Product Review</strong> submission page. <br /> <a href="pcTSUtility.asp" target="_blank" style="color:#666;">Review your XML parser settings</a> to ensure your store is setup to use a XML parser supported by your Web server.</td>
		</tr>
 		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_Contact" value="1" <%if scContact=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add <strong>CAPTCHA</strong> (random security code) to the <strong><a href="../pc/contact.asp" target="_blank">Contact Us</a></strong> form. <br /><a href="pcTSUtility.asp" target="_blank" style="color:#666;">Review your XML parser settings</a> to ensure your store is setup to use a XML parser supported by your Web server.</td>
		</tr>
    	<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_AffLogin" value="1" <%if scAffLogin=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to Affiliate Login pages</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_AffReg" value="1" <%if scAffReg=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to Affiliate Registration pages</td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">Control Panel</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td><input type="checkbox" name="pcv_AdminLogin" value="1" <%if scAdminLogin=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Add advanced security to Control Panel Login page</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_UseImgs2" value="1" <%if scUseImgs2=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%" valign="top">Add <strong>CAPTCHA</strong> (random security code) to the Control Panel <strong>Login</strong> page.<br /><a href="pcTSUtility.asp" target="_blank" style="color:#666;">Review your XML parser settings</a> to ensure your store is setup to use a XML parser supported by your Web server.</td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">Alerts</th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td valign="top"><input type="checkbox" name="pcv_AlarmMsg" value="1" <%if scAlarmMsg=1 then%>checked<%end if%> class="clearBorder"></td>
		<td width="95%">Send a notification e-mail to the store administrator when someone attempts to log into the store more than the number of attempts listed below. This feature can alert you of a script-based attacked performed against the store. This applies to any login form in the storefront and in the Control Panel.</td>
		</tr>
		<tr>
		<td width="5%">&nbsp;</td>
		<td>&nbsp;</td>
		<td width="95%">Number of Consecutive Attempts: <input type="text" name="pcv_AttackCount" size="4" value="<%=scAttackCount%>"></td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="3"><hr></td>
		</tr>
		<tr>
		<td>&nbsp;</td>
		<td></td>
		<td>
			<input type="submit" name="submit" value="Update Settings" class="submit2">
            &nbsp;<input type="button" name="back" value="Back" onClick="JavaScript:history.go(-1);">
        </td>
        </tr>
		<tr>
			<td colspan="3">&nbsp;</td>
		</tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->