<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
Session("RedirectURL")= "" 
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
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/languagesCP.asp"-->
<% 
Dim SPath
SPath=Request.ServerVariables("PATH_INFO")
SPath=mid(SPath,1,InStrRev(SPath,"/")-1)
If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
	strSiteURL="http://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
Else
	strSiteURL="https://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
End if

pageTitle="Forgot Your User Name?"

if session("ForgotAttackCount")="" then
  session("ForgotAttackCount")=0
end if
	
validateForm "getUserName.asp" 

if (request.form("submitf")="1") then	   
	session("ForgotAttackCount")=session("ForgotAttackCount")+1
	if InStr(ucase(Request.servervariables("HTTP_REFERER")),ucase(strSiteURL & "forgot_password.asp")) <>1 then
		Session("cp_Forgotpassword")=""
		session("ForgotAttackCount")=session("ForgotAttackCount")+1		
	end if
End if
%>

<!--#include file="AdminHeader.asp"-->
<% 
If session("ForgotAttackCount") < 5 Then
   
	If request.Form("submitf") <> "1"  Then %>
		<% ' START show message, if any
		'====================================================================================== 
		'//Prevent an XSS attack on pcv4_showMessage.asp (only needed in this non-secure page)
		'====================================================================================== 
		If msg&""="" Then
			pcStrMsg=trim(GetUserInput(request.querystring("msg"),0))
			if pcStrMsg&""="" then
				pcStrMsg=trim(GetUserInput(request.querystring("message"),0))
			end if
			msg = pcStrMsg
		End If
		'====================================================================================== 
		%>
        <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
	<% End If %>
		  
	<form method="post" name="forgot_username" class="pcForms">
		<table class="pcCPcontent">
			<tr> 
				<td colspan="2">&nbsp;<% validateError %></td>
			</tr>
			<tr> 
				<td width="15%" align="right">Admin Email:</td>
				<td width="85%">
				<% textbox "email", "",25, "textbox"
				   validate "email", "email"
				%>
				<i></i></td>
			</tr>
			
			<%
			Session("cp_ForgotUserName")="1"
			Session("cp_postnum")=""
			session("cp_num")=""
			%>
		   
			<tr> 
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr> 
				<td>&nbsp;</td>
				<td>
				<input type=hidden name=submitf value="1">
				<input type="submit" name="Submit" value="Submit" class="submit2">
				</td>
			</tr>
			
		</table>
	</form> 
	<script type="text/javascript">
     document.forgot_username.email.focus();
    </script>
<%Else%>
    <div class="pcCPmessage">
        <%=dictLanguageCP.Item(Session("language")&"_forgotpassword_securityMessage") %>
    </div>
<%End if %>
<!--#include file="AdminFooter.asp"-->