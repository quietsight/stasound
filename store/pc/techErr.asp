<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<% dim pcStrCustRefID
pcStrCustRefID=getUserInput(request("err"),0)
%>
<!--#include file="header.asp"-->
<div id="pcMain">
	<table class="pcMainTable">
			<tr valign="top"> 
				<td>
					<div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_techErr_1")%></div>
				</td>
			</tr>
			<tr>
				<td class="pcSpacer"></td>
			</tr>
			<tr>
				<td>
					<p><strong><%=dictLanguage.Item(Session("language")&"_techErr_6")%></strong></p>
					<p><%=dictLanguage.Item(Session("language")&"_techErr_7")%><b><%=pcStrCustRefID%></b><%=dictLanguage.Item(Session("language")&"_techErr_8")%></p>
					<p>&nbsp;</p>
					<%=dictLanguage.Item(Session("language")&"_techErr_9")%>
				</td>
			</tr>
			<tr>
				<td class="pcSpacer"></td>
			</tr>
	</table>
</div>
<%call clearLanguage()%>
<!--#include file="footer.asp"-->