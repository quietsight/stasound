<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="About ProductCart&reg;" %>
<% Section="about" %>
<%PmAdmin=0%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
pcPageName="about.asp"
%>
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer">ProductCart is a registered trademark property of NetSource Commerce. Since 2001, NetSource Commerce has been developing and updating the ProductCart family of shopping cart software for Internet merchants.</td>
	</tr>
	<tr> 		
		<td>
    	<ul class="pcListIcon">
        	<li><a href="about_credits.asp">Copyright &amp; Credits</a></li>
        	<li><a href="about_terms.asp">Terms &amp; Conditions</a></li>
            <li><a href="http://www.productcart.com" target="_blank">About NetSource Commerce</a></li>
		</ul>
    </td>
	</tr>	
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
    <tr>
    	<td>There is a very active community around ProductCart.</td>
    </tr>
    <tr>
    	<td>
            <ul class="pcListIcon">
                <li>Visit and participate in the <a href="http://www.earlyimpact.com/forum/" target="_blank">ProductCart Forums</a></li>
                <li>Read and contribute to the <a href="http://wiki.earlyimpact.com/" target="_blank">ProductCart WIKI</a> (or get an <a href="http://wiki.earlyimpact.com/feed.php" target="_blank">RSS feed of recent changes</a>)</li>
                <li><a href="http://www.earlyimpact.com/productcart/addons.asp" target="_blank">Extend ProductCart</a> with the many available add-on's.</li>
                <li><a href="http://www.earlyimpact.com/cpd/" target="_blank">Hire a developer</a> to customize it for you.</li>
            </ul>
        </td>
    </tr>
</table>
<!--#include file="AdminFooter.asp"-->
