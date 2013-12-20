<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="pcCalculateBTODefaultPrices.asp" -->
<!--#include file="inc_UpdateDates.asp" -->
<%
dim pageTitle, section, f, query, conntemp, rs, rstemp
pageTitle="Gateway"
pageIcon="pcv4_icon_pg.png"
section="paymntOpt" 

pcMessage = getUserInput("msg", 0)
id = getUserInput(request("id"), 0)
gwchoice = getUserInput(request("gwchoice"), 0)
%>
<!--#include file="AdminHeader.asp"-->

<table class="pcCPcontent">
	<tr>
		<td colspan="2">
			<div class="pcCPmessageSuccess">&quot;<%=request.QueryString("msg")%>
		</td>
	</tr>
	<tr>
		<td valign="top">
			<ul class="pcListIcon">
				<li style="padding-bottom: 10px;"><a href="pcConfigurePayment.asp?mode=Edit&id=<%=id%>&gwchoice=<%=gwchoice%>">Edit this gateway</a></li>
				<li style="padding-bottom: 10px;"><a href="PaymentOptions.asp">View/Modify Payment Options</a></li>
				<li style="padding-bottom: 10px;"><a href="OrderPaymentOptions.asp">Set Display Order</a></li>
          </ul>
          </td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->