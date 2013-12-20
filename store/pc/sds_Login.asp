<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="pcStartSession.asp"-->
<%
If scStoreOff="1" then
	response.redirect "msg.asp?message=31"
End If
%>
<% if request.form("SubmitCO.y")<>"" then
	ErrCnt=0
	EP=0
	'form is submitted
	sds_username=replace(request.form("sds_username"),"'","''")
	session("sds_username")=sds_username
	if sds_username="" then
		ErrCnt=ErrCnt+1
	End if
	sds_password=request.form("sds_password")
	if sds_password="" then
		ErrCnt=ErrCnt+1
		EP=1
	End if
	
	If ErrCnt>0 then
		If (scSecurity=1) and (scAlarmMsg=1) then
				if session("AttackCount")="" then
					session("AttackCount")=0
				end if
				session("AttackCount")=session("AttackCount")+1
				if session("AttackCount")>=scAttackCount then
				session("AttackCount")=0%>
				<!--#include file="../includes/sendAlarmEmail.asp" -->
				<%end if	
		End if
		response.redirect "sds_Login.asp?EP="&EP&"&msg="& Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	Else
		erypassword=encrypt(sds_password, 9286803311968)
		session("sds_erypassword")=erypassword
		response.redirect "sds_LoginB.asp" 
	End if

end if
%>
<% ' if Drop-Shipper already login
if (Session("pc_idsds")<>"") and (Session("pc_idsds")<>"0") then
 response.redirect "sds_MainMenu.asp"
end if
%>
<!--#include file="header.asp"-->
<%
pcRequestRedirect=trim(getUserInput(request("redirectUrl"),250))
if len(pcRequestRedirect)>0 then
	session("redirectUrlLI")=pcRequestRedirect
end if
%>
<%
'// START - Check for SSL and redirect to SSL login if not already on HTTPS
	If scSSL="1" And scIntSSLPage="1" Then
		If (Request.ServerVariables("HTTPS") = "off") Then
		Dim xredir__, xqstr__
		xredir__ = "https://" & Request.ServerVariables("SERVER_NAME") & _
		Request.ServerVariables("SCRIPT_NAME")
		xqstr__ = Request.ServerVariables("QUERY_STRING")
		if xqstr__ <> "" Then xredir__ = xredir__ & "?" & xqstr__
		Response.redirect xredir__
		End if
	End If
'// END - check for SSL
%>
<div id="pcMain">
	<form method="post" name="auth" action="sds_Login.asp" class="pcForms">

		<% msg=server.HTMLEncode(request.querystring("msg"))
			If msg<>"" then	%>
				<div class="pcErrorMessage">
				<%=msg%>
				</div>
		<% end if %>
		
		<table class="pcMainTable">
			<tr> 
				<td valign="top"> 

				<!-- start of login form -->
					<table class="pcShowContent"> 
						<tr>
							<td colspan="2">
							<h2><%response.write dictLanguage.Item(Session("language")&"_sdsLogin_1")%></h2>
							<p><%response.write dictLanguage.Item(Session("language")&"_sdsLogin_2")%></p>
							</td>
						</tr>
						<tr>
							<td width="10%" nowrap="nowrap">
								<p><%response.write dictLanguage.Item(Session("language")&"_sdsLogin_3")%></p>
							</td>
							<td width="90%">
								<input type="text" name="sds_username" size="15" maxlength="150" value="<%=session("sds_username")%>">
								<% if msg="" then %>
									<img src="<%=rsIconObj("requiredicon")%>">
                <% else
										if session("email")="" then %>
                    <img src="<%=rsIconObj("errorfieldicon")%>">
                    <% end if %>
                <% end if %>
              </td>
						</tr>
						<tr>
							<td nowrap="nowrap">
								<p><%response.write dictLanguage.Item(Session("language")&"_sdsLogin_4")%></p>
							</td>
							<td>
								<input type="password" name="sds_password" size="15" maxlength="150">
								<% if msg="" then %>
								<img src="<%=rsIconObj("requiredicon")%>">
								<% end if %>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<input type="image" src="<%=rslayout("login")%>" border="0" name="SubmitCO" value="Submit" id="submit"></td>
								</tr>
							</table>
						<!-- end of login table -->
					</td>
				</tr>
			</table>
	</form>
	<hr>
	<!-- start of password request -->
		<table class="pcMainTable">
			<% if request.querystring("s")="1" then %>
			<tr>
				<td>
					<div class="pcErrorMessage">
						<%response.write dictLanguage.Item(Session("language")&"_sdsLogin_5")%>
					</div>
				</td>
			</tr>
			<% else %>
			<tr> 
				<td>
					<p><%response.write dictLanguage.Item(Session("language")&"_sdsLogin_6")%>
					<a href="sds_fpass.asp?redirectUrl=<%if request("redirectUrl")<>"" then%><%=Server.URLEnCode(session("redirectUrlLI"))%><%else%><%=Server.URLEncode(session("redirectUrlLI"))%><%end if%>&frURL=sds_Login.asp"><%response.write dictLanguage.Item(Session("language")&"_sdsLogin_7")%></a></p>
				</td>
			</tr>
			<% end if %>
		</table>
<!-- end of password request -->
</div>
<!--#include file="footer.asp"-->
<%call clearLanguage()%>