<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="About Purging Credit Card Numbers" %>
<% Section="" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/rc4.asp"--> 
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr><td colspan="3" class="pcCPspacer"></td></tr>
    <tr><th colspan="3">Purge Credit Card Numbers</th></tr>
    <tr><td colspan="3" class="pcCPspacer"></td></tr>
    <tr>
		<td colspan="3">
        	<p>ProductCart saves credit card information to the store database, in an encrypted format, <u>only when using:</u></p>
            <ul>
                <li><strong>Offline credit card</strong> processing</li>
                <li><strong>Authorize.Net</strong> in  &quot;Authorize Only&quot; mode</li>
                <li><strong>PayPal Payflow Pro</strong> in  &quot;Authorize Only&quot; mode</li>
                <li><strong>Netbilling</strong> in  &quot;Authorize Only&quot; mode</li>
                <li><strong>USAePay</strong> in  &quot;Authorize Only&quot; mode</li>
                <li><strong>LinkPoint API</strong> in  &quot;Authorize Only&quot; mode</li>
                <li><strong>NetSource Commerce Gateway</strong> in  &quot;Authorize Only&quot; mode /w "Secure Vault" disabled</li>
            </ul>
            <p>For more information about the Purge Credit Card Number feature, please <a href="http://wiki.earlyimpact.com/productcart/orders_purging_cc" target="_blank">see our documentation</a>.</p>        </td>
    </tr>
    <tr><td colspan="3" class="pcCPspacer"></td></tr>
    <tr>
		<td colspan="3">
			<input type="button" class="submit2" value="Select orders for which to purge c/c numbers" onClick="location='creditCardPurge.asp';">
        </td>
    </tr>
    <tr><td colspan="3" class="pcCPspacer"></td></tr>			
</table>
<!--#include file="AdminFooter.asp"-->