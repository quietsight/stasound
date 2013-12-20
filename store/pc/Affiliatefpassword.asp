<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<%
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

validateForm "Affretreivepassword.asp"%>
<!--#include file="header.asp"-->
<%
pcRequestRedirect=getUserInput(request("redirectUrl"),250)
if len(pcRequestRedirect)>0 then
	Session("pcSF_redirectUrl")=pcRequestRedirect
end if
pcfrUrl=getUserInput(request("frUrl"),250)
if len(pcfrUrl)>0 then
	Session("pcSF_pcfrUrl")=pcfrUrl
end if
%>
<div id="pcMain">
	<form method="post" name="auth" class="pcForms">
		<table class="pcMainTable">
			<tr> 
				<td>
					<h2><%response.write dictLanguage.Item(Session("language")&"_AffLogin_8")%></h2>
				</td>
			</tr>
			<tr> 
				<td>
					<p>
					<%response.write dictLanguage.Item(Session("language")&"_AffLogin_3")%>
					<%textbox "Email", "", 30, "textbox"%>
					<%validate "email", "email"%>
					</p>
                    <%validateError%>
				</td>
			</tr>
			<tr> 
				<td> 
                	<hr>
					<input type="image" src="<%=rslayout("submit")%>" name="Submit" value="Submit" id="submit">
				</td>
			</tr>
		</table>
	</form>
</div>
<%call clearLanguage()%>
<!--#include file="footer.asp"-->