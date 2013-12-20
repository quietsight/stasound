<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="header.asp"-->
<div id="pcMain">
	<table class="pcMainTable">
		<tr> 
			<td>
				<div class="pcErrorMessage">
					<%
					'Fixes for current strings
					msg=getUserInput(request.querystring("message"),0)
					msg=replace(msg, "&lt;BR&gt;", "<BR>")
					msg=replace(msg, "&lt;br&gt;", "<br>")
					msg=replace(msg, "&lt;b&gt;", "<b>")
					msg=replace(msg, "&lt;/b&gt;", "</b>")
					msg=replace(msg, "&lt;/font&gt;", "</font>")
					msg=replace(msg, "&lt;a href", "<a href")
					msg=replace(msg, "&gt;Back&lt;/a&gt;", ">Back</a>")
					msg=replace(msg, "&lt;font", "<font")
					msg=replace(msg, "&gt;<b>Error&nbsp;</b>:", "><b>Error&nbsp;</b>:")
					msg=replace(msg, "&gt;&lt;img src=", "><img src=")
					msg=replace(msg, "&gt;&lt;/a&gt;", "></a>")
					msg=replace(msg, "&gt;<b>", "><b>")
					msg=replace(msg, "&lt;/a&gt;", "</a>")
					msg=replace(msg, "&gt;View Cart", ">View Cart")
					msg=replace(msg, "&gt;Continue", ">Continue")
					msg=replace(msg, "&lt;u>", "<u>")
					msg=replace(msg, "&lt;/u>", "</u>")
					msg=replace(msg, "&lt;ul&gt;", "<ul>")
					msg=replace(msg, "&lt;/ul&gt;", "</ul>")
					msg=replace(msg, "&lt;li&gt;", "<li>")
					msg=replace(msg, "&lt;/li&gt;", "</li>")
					msg=replace(msg, "&gt;", ">") 
					%>
					<%=msg%>
				</div>
			</td>
		</tr>
	</table>
</div>
<!--#include file="footer.asp"-->