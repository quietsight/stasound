<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=0%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<html>
<head>
<title>SSL Check</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:275px; background-image: none;">
<form name="form1" method="post" class="pcForms">
<table class="pcCPcontent">
  	<tr>
    	<td>
			<h1>Secure Socket Layer (SSL) Check</h1><br>
			<%
			if 	session("SSLCheck")="SSL Check Successful." then 
			%>
			<table>
				<tr> 
					<td width="4%"><img src="images/pc_checkmark_sm.gif" width="16" height="16"></td>
					<td width="96%">Your Store's SSL is configured properly.</td>
				</tr>
			</table>
			<%
			else
			%>
			<table>
				<tr> 
					<td width="4%"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
					<td width="96%">Your Store's SSL is <b>NOT</b> configured properly. Make sure that the kind of SSL certificate used on your store is in compliance with the <a href="http://www.earlyimpact.com/productcart/system_req.asp#ssl" target="_blank">ProductCart system requirements</a>.</td>
				</tr>
			</table>
			<%
			end if
			%>
		</td>
	</tr>
  	<tr>
    	<td colspan="2" align="center">
			<input type="button" name="Back" value="Close Window" onClick="self.close();">
		</td>
	</tr>
</table>
</form>
</div>
</body>
</html>