<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export your orders to QuickBooks" %>
<% Section="genRpts" %>
<%PmAdmin=10%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="QuickBooks/spry_includes.asp"-->
<%
Dim rs, conntemp
pcStrPageName = "qb_home.asp"

'// START - Check for QuickBooks and redirect to Add-on Home page
Set fs=Server.CreateObject("Scripting.FileSystemObject")
If (fs.FileExists(Server.MapPath("QB_Default.asp")))=0 Then
   isQBKApplied="0"
   else
   isQBKApplied="1"   
End If
set fs=nothing

if isQBKApplied="1" then
	response.Redirect("QB_Default.asp")
end if
'// END
%>
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 
	<tr>
		<th>Export your orders to QuickBooks</th>
	</tr> 
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 
	<tr> 		
		<td>
    	<p><a href="http://www.earlyimpact.com/productcart/quickbooks" target="_blank"><img src="images/export_to_quickbooks.gif" alt="ProductCart Recurring Billing Add-on" width="220" height="160" align="right" style="margin-left: 15px;"></a>The QuickBooks® Add-On for our ProductCart shopping cart software gives you the ability to quickly and easily export order and customer information in an organized, time-saving manner.</p>
		  <p style="padding-top: 6px;">With a few clicks, you will be able to select which orders to export and create a QuickBooks data file to be imported into the popular accounting package published by Intuit.</p>
		  <ul>
      	<li>Easily export orders from your ProductCart-powered store</li>
      	<li>Quickly import them into QuickBooks as invoices or sales receipts</li>
        <li>Save, print, email the invoices or sales receipts: no double entries!</li>
        <li>Use new QuickBooks tools to process FedEx and UPS shipments</li>
        <li><a href="http://www.earlyimpact.com/productcart/quickbooks" target="_blank">Learn more...</a></li>
      </ul>
    </td>
	</tr>	
	<tr>
		<td class="pcCPspacer"></td>
	</tr> 	
</table>
<!--#include file="AdminFooter.asp"-->
