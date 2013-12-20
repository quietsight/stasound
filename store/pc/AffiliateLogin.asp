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
<!--#include file="../includes/productcartFolder.asp"-->
<%
If scStoreOff="1" then
	response.redirect "msg.asp?message=31"
End If
%>
<% if request.form("SubmitCO.y")<>"" then
	ErrCnt=0
	EP=0
	if (scSecurity=1) and (scAffLogin=1) and (scUseImgs=1) then
	Session("store_affpostnum")=replace(request("postnum"),"'","''")
	else
	Session("store_affpostnum")=""
	end if
	'form is submitted
	Email=replace(request.form("Email"),"'","''")
	session("Email")=Email
	if Email="" then
		ErrCnt=ErrCnt+1
	End if
	password=request.form("password")
	if password="" then
		ErrCnt=ErrCnt+1
		EP=1
	End if
	
	if (scSecurity=1) and (scAffLogin=1) and (scUseImgs=1) then
	%>
	
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
			Session("store_affpostnum")=""
			response.redirect "AffiliateLogin.asp?EP="&EP&"&msg="& Server.Urlencode(dictLanguage.Item(Session("language")&"_security_3"))
		End if
	End if

	If ErrCnt>0 then
	If (scSecurity=1) and (scAffLogin=1) and (scAlarmMsg=1) then
				if session("AttackCount")="" then
					session("AttackCount")=0
				end if
				session("AttackCount")=session("AttackCount")+1
				if session("AttackCount")>=scAttackCount then
				session("AttackCount")=0%>
				<!--#include file="../includes/sendAlarmEmail.asp" -->
				<%end if	
	End if
		response.redirect "AffiliateLogin.asp?EP="&EP&"&msg="& Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	Else
		erypassword=encrypt(password, 9286803311968)
		session("erypassword")=erypassword
		response.redirect "AffiliateLoginB.asp" 
	End if

end if
%>
<% ' if customer already login
if (Session("pc_idAffiliate")<>0) then
 response.redirect "AffiliateMain.asp"
end if
%>
<!--#include file="header.asp"-->
<%
pcRequestRedirect=getUserInput(request("redirectUrl"),250)
if len(pcRequestRedirect)>0 then
	session("redirectUrlLI")=pcRequestRedirect
end if
%>
<div id="pcMain">
	<form method="post" name="auth" action="AffiliateLogin.asp" class="pcForms">

		<% msg=server.HTMLEncode(request.querystring("msg"))
			If msg<>"" then	%>
				<div class="pcErrorMessage">
				<%=msg%>
				</div>
		<% end if %>
		
		<table class="pcMainTable">
			<tr> 
				<td width="50%" valign="top"> 

				<!-- start of login form -->
					<table class="pcShowContent"> 
						<tr>
							<td colspan="2">
							<h2><%response.write dictLanguage.Item(Session("language")&"_AffLogin_1")%></h2>
							<p><%response.write dictLanguage.Item(Session("language")&"_AffLogin_2")%></p>
							</td>
						</tr>
						<tr>
							<td>
								<p><%response.write dictLanguage.Item(Session("language")&"_AffLogin_3")%></p>
							</td>
							<td>
								<input type="text" name="Email" size="15" maxlength="150" value="<%=session("Email")%>">
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
							<td>
								<p><%response.write dictLanguage.Item(Session("language")&"_AffLogin_4")%></p>
							</td>
							<td>
								<input type="password" name="password" size="15" maxlength="150">
								<% if msg="" then %>
                <img src="<%=rsIconObj("requiredicon")%>">
                <% else
										if request("EP")="1" then %>
                    <img src="<%=rsIconObj("errorfieldicon")%>">
                    <% end if %>
                <% end if %>
							</td>
						</tr>
						<%
						Session("store_afflogin")="1"
						Session("store_affpostnum")=""
						session("store_affnum")="      "
						%>
						<%if (scSecurity=1) and (scAffLogin=1) and (scUseImgs=1) then%>
                            <!--#include file="../CAPTCHA/CAPTCHA_form_inc.asp" -->
						<%end if%>
						<tr>
							<td colspan="2">
								<input type="image" src="<%=rslayout("login")%>" border="0" name="SubmitCO" value="Submit" id="submit"></td>
								</tr>
							</table>
						<!-- end of login table -->
					
					</td>

					<td width="50%" valign="top"> 
						<!-- start of register table -->
						<table class="pcShowContent">
							<tr> 
								<td colspan="2" valign="top">
								<h2><%response.write dictLanguage.Item(Session("language")&"_AffLogin_5")%></h2>
								<p><%response.write dictLanguage.Item(Session("language")&"_AffLogin_6")%></p>
								</td>
							</tr>
							<tr> 
								<td colspan="2">
									<p><a href="NewAffa.asp"><img src="<%=rslayout("register")%>"></a></p>
									<p>&nbsp;</p>
								</td>
							</tr>
						</table>
						<!-- end of register table -->
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
						<%response.write dictLanguage.Item(Session("language")&"_AffLogin_7")%>
					</div>
				</td>
			</tr>
			<% else 
			pcRequestRedirect=getUserInput(request("redirectUrl"),250)
			%>
			<tr> 
				<td>
					<p><%response.write dictLanguage.Item(Session("language")&"_AffLogin_8")%>
					<a href="Affiliatefpassword.asp?redirectUrl=<%if trim(pcRequestRedirect)<>"" then%><%=Server.URLEnCode(pcRequestRedirect)%><%else%><%=Server.URLEncode(session("redirectUrlLI"))%><%end if%>&frURL=AffiliateLogin.asp"><%response.write dictLanguage.Item(Session("language")&"_AffLogin_9")%></a></p>
				</td>
			</tr>
			<% end if %>
		</table>
<!-- end of password request -->
</div>
<%call clearLanguage()%>
<!--#include file="footer.asp"-->