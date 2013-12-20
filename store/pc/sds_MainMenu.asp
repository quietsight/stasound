<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If
%>
<!--#include file="header.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td> 
			<h1><%=dictLanguage.Item(Session("language")&"_sdsMain_1")%></h1>
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
			<li><a href="pcmodsdsA.asp"><%=dictLanguage.Item(Session("language")&"_sdsMain_2")%></a></li>
			<li><a href="sds_viewPast.asp"><%=dictLanguage.Item(Session("language")&"_sdsMain_3")%></a></li>
			<li><a href="contact.asp"><%=dictLanguage.Item(Session("language")&"_sdsMain_4")%></a></li>
			<li><a href="sds_LO.asp"><%=dictLanguage.Item(Session("language")&"_sdsMain_5")%></a></li>
		</ul>                    
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->