<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials"
pageTitle="Sales Manager"
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<table class="pcCPcontent">
<tr>
	<td>
		<div class="pcCPmessageInfo">
			The Sales Manager features only work on a SQL Server database.<br>
			Please upgrade your database if you want to use them.
		</div>
	</td>
</tr>
<tr>
	<td class="pcCPspacer">&nbsp;</td>
</tr>
<tr>
	<td class="pcCPspacer">&nbsp;</td>
</tr>
<tr>
	<td>
		<input type="button" name="Go" value=" Back to Main page " onclick="location='menu.asp';" class="submit2">
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->