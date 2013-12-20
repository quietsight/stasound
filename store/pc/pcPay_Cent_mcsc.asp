<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<% Response.Buffer=True%> 
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<HTML>
<HEAD><TITLE>Introducing MasterCard® SecureCode</TITLE>
	<SCRIPT language=JavaScript>
		btn_close = new Image(); btn_close.src = "/assets/secure_code/btn_close.gif";
		btn_close_over = new Image(); btn_close_over.src = "/assets/secure_code/btn_close_over.gif";
	
		function changeImage() {
			if (document.images) {
				for (var i=0; i < changeImage.arguments.length; i+=2) {
					document[changeImage.arguments[i]].src = eval(changeImage.arguments[i+1] + ".src");
					}
				}
			}
	</SCRIPT>
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</HEAD>
<body>
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td colspan="3">
				<img height="44" alt="MasterCard® SecureCode" src="images/pc_secureCode_logo.gif" width="116">
			</td>
		</tr>
		<tr>
			<td colspan="3">
				<p>MasterCard<SUP>®</SUP> SecureCode™ is a new service from MasterCard and your card issuer that provides added protection when you buy online. There is no need to get a new MasterCard or Maestro<SUP>®</SUP> card. You choose your own personal MasterCard SecureCode and it is never shared with any merchant. A private code means added protection against unauthorized use of your credit or debit card when you shop online.</p>
				<p>Every time you pay online with your MasterCard or Maestro card, a box pops up from your card issuer asking you for your personal SecureCode, just like the bank does at the ATM. In seconds, your card issuer confirms it's you and allows your purchase to be completed.</p>
        <p>&nbsp;</p>
			</td>
		</tr>
		<tr valign="top">
			<td><img height=43 alt=MasterCard src="images/pc_mc_logo.gif" width=72 align=left></td>
			<td align="middle">
				<p>To find out more about MasterCard SecureCode go to <a href="http://www.mastercardsecurecode.com" target="_blank">www.mastercardsecurecode.com</a> 
         </p>
         <p>&nbsp;</p></td>
			<td><img height=43 alt="Maestro International" src="images/pc_maestro_logo.gif" width=72 align=left> </td>
		</tr>
		<tr>
			<td></td>
			<td align="center"><A href="javascript:window.close();"><img src="images/close.gif" alt="Close" name="button" width="32" height="25" hspace=0 border=0></A>
			</td>
			<td></td>
		</tr>
	</table>
</div>
</body>
</html>