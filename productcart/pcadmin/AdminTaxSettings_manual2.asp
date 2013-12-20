<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Edit Tax Settings - Manual Entry Method - Step 2" %>
<% Section="taxmenu" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr> 
		<td>
        	For more information about the following options, see the <a href="http://wiki.earlyimpact.com/productcart/tax_manual" target="_blank">User Guide</a>:
<ul class="pcListIcon">
			<li><a href="AddTaxPerPlace.asp" style="font-weight: bold;">Add tax by location</a><br>
		    You can specify a tax rate for a specific location, such as a country, a state or province, or a particular postal code. </li>
			<li style="padding-top: 10px;"><a href="AddTaxPerZone.asp" style="font-weight: bold;">Add tax by zone</a><br>
			  You can specify a tax rate for a group of states or provinces.</li>
			<li style="padding-top: 10px;"><a href="AddTaxPerPrd.asp" style="font-weight: bold;">Add tax by product</a><br>
			  You can specify a tax rate that applies to a specific product.</li>
			<li style="padding-top: 10px;"><a href="viewTax.asp">List current tax rules</a></li>
          </ul>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->