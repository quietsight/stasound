<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Bing Cashback" %>
<% section="genRpts" %>
<%PmAdmin=3%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/CashbackConstants.asp"-->

<%
dim conntemp, rs, query
Dim pcv_strPageName
pcv_strPageName="pcCashback_settings.asp"

If request("action")="add" Then

	Session("LSCB_KEY")=Request("key")
	Session("LSCB_STATUS")=Request("status")

	response.Redirect("../includes/PageCreateCashbackConstants.asp")

End If
%>
<!--#include file="AdminHeader.asp"-->
<style>
.pcCPOverview {
	background-color: #F5F5F5;
	border: 1px solid #FF9900;
	margin: 5px;
	padding: 5px;
	color: #666666;
	font-size:11px;
	text-align: left;
}
.pcCodeStyle {
	font-family: "Courier New", Courier, monospace;
	color: #FF0000;
	font-size: 9;
}
</style>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form method="post" name="form1" action="<%=pcv_strPageName%>?action=add" class="pcForms">
	<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Bing Cashback Settings</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<p>To activate Bing Cashback (formerly <em>Live Search Cashback</em>) enter your &quot;Merchant ID&quot; in the text field below. Next, set the Cashback status to the "On" position. When you are ready click &quot;Save Bing Cashback Settings&quot; to save your settings and return home. You will then be able to export a product feed that you can upload to your Cashback account.</p>
            </td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td><p>Bing Merchant ID: 
		    <input name="key" type="text" value="<%=LSCB_KEY%>" size="45" maxlength="45">
            <input name="status" type="hidden" value="1">
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"><hr></td>
	</tr>
	<!--	
    <tr>
		<td><p>Cashback Status: 
		    <input type="radio" name="status" value="1" class="clearBorder" <% if LSCB_STATUS="1" then response.Write("checked") %>> 
		    On&nbsp;&nbsp;
		    <input type="radio" name="status" value="0" class="clearBorder" <% if LSCB_STATUS="0" then response.Write("checked") %>> 
		    Off</p></td>
	</tr> 
	<tr>
		<td class="pcCPspacer"><hr></td>
	</tr>
    -->
	<tr> 
		<td style="text-align: center;">
			<input name="submit" type="submit" class="submit2" value="Save Bing Cashback Settings">
            &nbsp;
            <input name="back" type="button" value="Back" onClick="document.location.href='pcCashback_main.asp';">
       	</td>
	</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->