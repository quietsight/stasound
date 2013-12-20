<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="AffLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<%
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

' Load affiliate ID
affVar=session("pc_idaffiliate")
if not validNum(affVar) then
	response.redirect "AffiliateLogin.asp"
end if
%>
<!--#include file="header.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td> 
			<h1><%=dictLanguage.Item(Session("language")&"_AffMain_1")%></h1>
		</td>
	</tr>
	<tr>
		<td>
			<%If request.querystring("msg")<>"" then %>
			<div class="pcErrorMessage">
				<%response.write server.HTMLEncode(request.querystring("msg")) %>
			</div>
			<%end if%>
		
		<ul>
			<li><a href="pcmodAffA.asp"><%=dictLanguage.Item(Session("language")&"_AffMain_2")%></a></li>
			<li><a href="Affgenlinks.asp"><%=dictLanguage.Item(Session("language")&"_AffMain_3")%></a></li>
			<li><a href="AffCommissions.asp"><%=dictLanguage.Item(Session("language")&"_AffMain_4")%></a></li>
			<li><a href="AffLO.asp"><%=dictLanguage.Item(Session("language")&"_AffMain_5")%></a></li>
		</ul>                    
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->