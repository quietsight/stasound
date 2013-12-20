<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Place an order for a new customer" %>
<% section="mngAcc" %>
<%PmAdmin=7%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="AdminHeader.asp"-->

	<table class="pcCPcontent">
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td><p>There are two ways to place an order for a <strong>new</strong> customer:</p>
            <ul class="pcListIcon">
            <li><span style="font-weight: bold"><a href="instCusta.asp">Add the customer</a> now</span>, and then use the &quot;Place Order&quot; feature to be logged into the storefront as him/her. For example, if you need to place an order for a new customer that will belong to a specific <a href="AdminCustomerCategory.asp">pricing category</a>, you should use this method. You will add the customer, assign it to the correct pricing category, and then use the &quot;Place Order&quot; link on the customer details page.</li>
            <li style="padding-top: 10px;"><span style="font-weight: bold"><a href="../pc/default.asp" target="_blank">Use the storefront</a></span> and act as if you were the new customer.</li>
            </ul></td>
	</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	</table>

<!--#include file="AdminFooter.asp"-->