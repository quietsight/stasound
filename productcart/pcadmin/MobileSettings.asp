<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Mobile Commerce Settings" %>
<% Section="layout" %>
<%PmAdmin=1%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim rs, conntemp
pcStrPageName = "MobileSettings.asp"

%>
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 
	<tr>
		<th>Mobile Commerce Add-on</th>
	</tr> 
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 
	<tr> 		
		<td>
    	<p><a href="http://www.productcart.com/mobile-commerce.asp" target="_blank"><img src="images/Mobile-Commerce.png" alt="ProductCart Mobile Commerce Add-on" width="220" height="160" align="right" style="margin-left: 15px;"></a>The Mobile Commerce Add-on allows you to add to your ProductCart-powered ecommerce Web site a new set of storefront pages that have been optimized for mobile devices such as the iPhone, Adroid phones, Blackberry, Windows 7 smartphones, etc.</p>
		  <p style="padding-top: 6px;">Installation is fast and painless: just upload a new set of files to your Web store. From then on, when a mobile device visits your regular storefront, they will be automatically redirected to the mobile-optimized pages.</p>
		  <p style="padding-top: 6px;">Quick, easy shopping for mobile devices</p>
		  <ul>
			<li>Simple user interface optimized for mobile devices: <a href="http://www.productcart.com/mobile-commerce-features.asp" target="_blank">learn more about its features</a></li>
			<li><a href="http://www.productcart.com/mobile-commerce-demo.asp" target="_blank">See it at work</a> on our software store</li>
			<li><a href="https://www.earlyimpact.com/eistore/productcart/pc/Mobile-Commerce-Add-on-125p475.htm" target="_blank">Buy a license</a> and add mobile commerce to your ecommerce Web site in no time.</li>
      	    <li><a href="http://www.productcart.com/mobile-commerce.asp" target="_blank">Learn more...</a></li>
      </ul>
    </td>
	</tr>	
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 	
</table>
<!--#include file="AdminFooter.asp"-->
