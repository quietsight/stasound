<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="List your products on eBay" %>
<% Section="genRpts" %>
<%PmAdmin=3%><!--#include file="adminv.asp"-->
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
pcStrPageName = "eb_home.asp"

'// START - Check for eBay and redirect to Add-on Home page
Set fs=Server.CreateObject("Scripting.FileSystemObject")
If (fs.FileExists(Server.MapPath("ebay_Listings.asp")))=0 Then
	isEbayApplied="0"
	else
	isEbayApplied="1"
End If
set fs=nothing

if isEbayApplied="1" then
	response.Redirect("ebay_home.asp")
end if
'// END
%>
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 
	<tr>
		<th>Easily list your products on eBay</th>
	</tr> 
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 
	<tr> 		
		<td>
    	<p><a href="http://www.earlyimpact.com/productcart/eBay/" target="_blank"><img src="images/pc2008-eBay-Addon.jpg" alt="ProductCart Recurring Billing Add-on" width="258" height="180" align="right" style="margin-left: 15px;"></a>The eBay Add-on for ProductCart allows you to easily list your products for sale on eBay, the world's largest marketplace.</p>
		  <p style="padding-top: 6px;">It contains features that will make it easy for you to leverage the work you've already done. For example, product names, description, images, etc. are pre-filled when you create a new listing to be posted on eBay.</p>
		  <ul>
      	<li>Pick products and start eBay auctions in minutes</li>
      	<li>Live listing data synchronization</li>
        <li>Layout templates and bulk listing features</li>
        <li>Support for international eBay sites</li>
        <li><a href="http://www.earlyimpact.com/productcart/eBay/" target="_blank">Learn more...</a></li>
      </ul>
    </td>
	</tr>	
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 	
</table>
<!--#include file="AdminFooter.asp"-->
