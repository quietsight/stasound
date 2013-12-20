<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="header.asp"-->
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<p><%response.write dictLanguage.Item(Session("language")&"_sorry_1")%></p>
				<hr>
				<p><% response.write server.HTMLEncode(request.querystring("ErrMsg")) %></p>
				<br>
				<br>
				<p><a href="javascript: history.back(-1)"><img src="<%=rslayout("back")%>" border=0></a></p>
			</td>
		</tr>
	</table>
</div>
<!--#include file="footer.asp"-->