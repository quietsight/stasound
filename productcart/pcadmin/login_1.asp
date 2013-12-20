<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="pcCPLog.asp" -->
<% 
if scRegistered="124587" then 
	Response.redirect "../setup/default.asp"
end if

'Checks for cookie
Dim CookieVar, ShowAgreement
ShowAgreement=0
CookieVar=Request.Cookies("AgreeLicense45")

if request("RedirectURL")<>"" then
	Session("RedirectURL")=getUserInput(request("RedirectURL"),0)
end if

If CookieVar="Agreed" then
Else
	ShowAgreement=1
End If
If request.form("Submit2")<>"" then
	AgreeVar=request.form("agree")
	If AgreeVar=1 then
		'place cookie
		Response.Cookies("AgreeLicense45")="Agreed"
		Response.Cookies("AgreeLicense45").Expires=Date() + 365
		MyCookiePath=Request.ServerVariables("PATH_INFO")
		do while not (right(MyCookiePath,1)="/")
		MyCookiePath=mid(MyCookiePath,1,len(MyCookiePath)-1)
		loop
		Response.Cookies("AgreeLicense45").Path=MyCookiePath
		response.redirect "login_1.asp"
	else
		'send message to agree
		AgreeMsg="Agree to the terms and conditions of the ProductCart End User License Agreement to continue."
		response.redirect "login_1.asp?AM="&server.URLEncode(AgreeMsg)
	end if
End If
' verifies if admin is logged, so as not send to login page
if session("admin")<>0 then
if Session("RedirectURL")<>"" then
	RedirectURL=Session("RedirectURL")
	Session("RedirectURL")=""
	response.redirect RedirectURL
else
 response.redirect "menu.asp"
end if 
end if
pageTitle="Login"
pageIcon="pcv4_icon_login.png"
%>
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/languagesCP.asp"-->
<%
if (request.form("submitf")="1") and (Session("cp_Adminlogin")="1") then
	if (scSecurity=1) and (scAdminLogin=1) and (scUseImgs2=1) then %>
    <!-- Include file for CAPTCHA configuration -->
    <!-- #include file="../CAPTCHA/CAPTCHA_configuration.asp" --> 
     
    <!-- Include file for CAPTCHA form processing -->
    <!-- #include file="../CAPTCHA/CAPTCHA_process_form.asp" -->   
	<%	
		If not blnCAPTCHAcodeCorrect then
			If scAlarmMsg=1 then
				if session("AttackCount")="" then
					session("AttackCount")=0
				end if
				session("AttackCount")=session("AttackCount")+1
				if session("AttackCount")>=scAttackCount then
					session("AttackCount")=0%>
					<!--#include file="../includes/sendAlarmEmail.asp" -->
				<%end if	
			End if
			Session("cp_postnum")=""
			response.redirect "login_1.asp?msg="& Server.Urlencode(dictLanguage.Item(Session("language")&"_security_3"))
		End if
	End if
End if
%>
<% validateForm "login.asp" %>
<!--#include file="AdminHeader.asp"-->
<% if ShowAgreement=0 then %>

	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
    <% If (Request.ServerVariables("HTTPS") = "on") Then
		pcv_TwitterJsURL = "https://widgets.twimg.com/j/2/widget.js"
	Else
		pcv_TwitterJsURL = "http://widgets.twimg.com/j/2/widget.js"
	End If %>
	<div style="float: right; width: 300px">
    	<div style="float:right; margin: 0 5px;"><img src="images/twitter_newbird_boxed_blueonwhite.png" alt="Twitter"></div>
    	<div style="margin-bottom: 10px;">The latest news from the ProductCart world, via Twitter</div>
		<script src="<%=pcv_TwitterJsURL%>"></script>
        <script>
        new TWTR.Widget({
          version: 2,
          type: 'profile',
          rpp: 3,
          interval: 6000,
          width: 300,
          height: 300,
          theme: {
            shell: {
              background: '#f1f1f1',
              color: '#555555'
            },
            tweets: {
              background: '#ffffff',
              color: '#777777',
              links: '#0790eb'
            }
          },
          features: {
            scrollbar: false,
            loop: false,
            live: false,
            hashtags: true,
            timestamp: true,
            avatars: false,
            behavior: 'all'
          }
        }).render().setUser('productcart').start();
        </script>
    </div>

    <form method="post" name="login" class="pcForms">
        <table class="pcCPcontent" style="width: 400px;">
            <tr> 
                <td colspan="4">&nbsp;<% validateError %></td>
            </tr>
            <tr> 
                <td align="right">User:</td>
                <td>
				<% textbox "idadmin", "", 12, "textbox"
				validate "idadmin", "positiveNumber" %>
				</td>
                <td align="right">Password:</td>
                <td> 
				<% textbox "password","", 12, "password"
                validate "password", "required" %>
                </td>
            </tr>
            <tr>
            	<td></td>
            	<td><a href="forgot_username.asp" style="text-decoration: none; color:#777;">Forgot User Name?</a></td>
            	<td></td>
                <td><a href="forgot_password.asp" style="text-decoration: none; color:#777;">Forgot Password?</a></td>
            </tr>
			<%
            Session("cp_Adminlogin")="1"
            Session("cp_postnum")=""
            session("cp_num")="      "%>
            <%if (scSecurity=1) and (scAdminLogin=1) and (scUseImgs2=1) then%>
            	<tr>
                	<td colspan="4" class="pcCPspacer"></td>
                </tr>
                <tr>
                    <td></td>
                    <td colspan="3"><!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" --></td>
                </tr>
            <%end if%>
            <tr> 
                <td colspan="4">&nbsp;</td>
            </tr>
            <tr> 
                <td>&nbsp;</td>
                <td colspan="3">
                <input type="hidden" name="submitf" value="1">
                <input type="submit" name="Submit" value="Submit" class="submit2">
                </td>
            </tr>
        </table>
    </form>
    <script type="text/javascript">
     document.login.idadmin.focus();
    </script>
    
<% else %>
    <form action="login_1.asp" method="post" name="IAgree" id="IAgree" class="pcForms">
        <table class="pcCPcontent">
            <tr> 
                <td colspan="2">
                <% if request.querystring("AM")<>"" then %>
                <div class="pcCPmessage">
                <%=request.querystring("AM")%>
                </div>
                <% end if %>
                </td>
            </tr>
            <tr>
                <td colspan="2">
				<!--#include file="inc_EULA.asp"-->
                </td>
            </tr>
            <tr> 
                <td colspan="2">
                <input type="checkbox" name="agree" value="1" class="clearBorder"> I agree to the terms and conditions of the <strong>ProductCart End User License Agreement</strong>, which are listed above.
                </td>
            </tr>
            <tr> 
                <td colspan="2" style="padding-top: 10px;">
                <input type="submit" name="Submit2" value="Continue" class="submit2">
                </td>
            </tr>
        </table>
    </form> 
<% end if %>
<!--#include file="AdminFooter.asp"-->