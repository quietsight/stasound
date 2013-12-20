<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Custom Shipping Option" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr> 
		<td>
		Select one of the following methods to create a new custom shipping option. You can create multiple shipping options based on the same, or different calculation criteria.
		<ul class="pcListIcon">
		<li><a href="AddFlatShippingRates.asp?type=P">Flat rate based on order amount</a><br>
		Select this option if you wish to charge a flat rate based on the total order amount (e.g. for orders between $50	and $100, charge $7.50).<br><br></li>
		<li><a href="AddFlatShippingRates.asp?type=O">Percentage of order amount</a><br>
		Select this option if you wish to create a shipping option that charges a percentage of the total order amount (e.g. for orders between $50 and $100, charge 11% of the order	amount).<br><br>
		</li>
		<li><a href="AddFlatShippingRates.asp?type=Q">Flat rate based on order quantity</a><br>
		Select this option if you wish to create a shipping option that charges a flat rate  based on the total number of the items in the cart (e.g. for order between 10 and 20 units, charge $5).<br><br>
		</li>
		<li><a href="AddFlatShippingRates.asp?type=W">Flat rate based on order weight</a><br>
		Select this option if you wish to create a shipping option that charges a flat rate  based on the total	weight of the items in the cart (e.g. for order between 5 and 10 pounds, charge $12).<br><br>
		</li>
		<li><a href="AddFlatShippingRates.asp?type=I">Incremental calculation based on order quantity</a><br>
		Example: charge $5.00 for the first item, then an additional $1 on the next 9 items, then $0.50 on all items over 9. If the order contained 22 units, shipping would be  calculated as follows: ($5	+ ($1*9) + (.50*12))=$20.00.</li>
		</ul>
		</td>
	</tr>
</table> 
<!--#include file="AdminFooter.asp"-->