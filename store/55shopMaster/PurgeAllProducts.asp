<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle="Purge All Products and Orders from Database" %>
<!--#include file="AdminHeader.asp"-->
<% 
on error resume next 



If request("purgeproducts")<>"" then
	dim mySQL, conntemp, rstemp
	call openDb()

	' delete from taxPrd
	mySQL="DELETE FROM taxPrd"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If

	' delete product from configSpec_products
	mySQL="DELETE FROM configSpec_products"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	
	' delete product from configSpec_categories
	mySQL="DELETE FROM configSpec_categories"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	
	' delete product from cs_relationships
	mySQL="DELETE FROM cs_relationships"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	
	' delete product from categories_products
	mySQL="DELETE FROM categories_products"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If

	' delete product from wishlist
	mySQL="DELETE FROM wishList"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	
	' delete product from options_optionsGroups
	mySQL="DELETE FROM options_optionsGroups"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	
	' delete product from configSessions	
	mySQL="DELETE FROM configSessions"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	
	' delete product from discountsPerQuantity	
	mySQL="DELETE FROM discountsPerQuantity"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	
	mySQL="DELETE FROM ProductsOrdered"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	
	mySQL="DELETE FROM creditCards"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If

	mySQL="DELETE FROM offlinepayments"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If

	mySQL="DELETE FROM Orders"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	
	' delete products from products table
	mySQL="DELETE FROM products"
	set rstemp=conntemp.execute(mySQL)
	if err.number <> 0 then
		pcErrorNumber = err.number
		pcErrorDescription = err.description
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in Purging All Products: " & pcErrorNumber & " - " & pcErrorDescription) 
	end If
	%>

	<table class="pcCPcontent">
			<tr>
				<td>
					<div class="pcCPmessageSuccess">All products and orders successfully deleted. <a href="menu.asp">Return to the Start Page</a></div>
				</td>
			</tr>
	</table>

<% else
	if request("confirm") <> "" then %>
	
	<form action="PurgeAllProducts.asp" method="post" name="form1" id="form1" class="pcForms">
  	<table class="pcCPcontent">
		<tr> 
		<td align="center">
			<div class="pcCPmessage" style="color: #F00; font-weight: bold; font-size: 18px;">CONFIRM PERMANENT DELETION?</div>
		</td>
		</tr>
		<tr>
			<td align="center">
			Are you absolutely sure that you want to <u>permanently remove all products</u> and <u>permanently remove all orders</u>?<br>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
		<td align="center">This action is permanent and cannot be undone.</td>
		</tr>
		<tr>
			<td class="pcCPspacer"><hr></td>
		</tr>
		<tr> 
		<td align="center">
		<input name="purgeproducts" type="submit" value="Yes" class="submit2">
        &nbsp;
		<input type="button" value="Cancel" onClick="javascript:history.back()">
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
</table>
</form>
<% else %>

<form action="PurgeAllProducts.asp" method="post" name="form2" id="form2" class="pcForms">
<table class="pcCPcontent">
	<tr> 
		<td colspan="2" align="left">
       	  <div class="pcCPmessage" style="color: #F00; font-weight: bold; font-size: 18px;">WARNING: This is a dangerous feature.</div>
          <p>This feature allows you to completely remove all products from your database. When you take this action you will <u>completely <strong>remove all products</strong></u> and <u><strong>all orders</strong></u>. This action is permanent and cannot be undone.</p>
          <p style="margin-top: 8px;">The main purpose of this feature is to purge products and orders that were entered into your database for <strong>testing purposes</strong>. </p>
          <p style="margin-top: 8px;">Once your store is &quot;live&quot; and there are real orders in your database, you should not perform this action. Instead, you can <strong>delete an item</strong> to hide the product form both the storefront and the Control Panel, without affecting the integrity of previous orders that might contain that product. Here are links to this and other options:  </p>
          <ul>
            <li><a href="LocateProducts.asp?cptype=0">Locate a product and delete it</a> (<em>the product remains in the database, but it's hidden</em>)</li>
            <li><a href="PurgeProducts.asp">Permanently remove a previously deleted product</a></li>
            <li><a href="PurgeRemovedProducts.asp">Permanently remove all previously deleted products</a></li>
            <li><a href="PurgeInactiveProducts.asp">Permanently remove all inactive products</a></li>
          </ul></td>
	</tr>
    <tr>
        <td class="pcCPspacer"><hr></td>
    </tr>
	<tr> 
	<td align="center" colspan="2">
		<input name="confirm" type="submit" value="Permanently Remove All Products &amp; Orders" class="submit2">
		&nbsp;<input type="button" value="Cancel" onClick="javascript:history.back()">
	</td>
	</tr>
    <tr>
        <td class="pcCPspacer"></td>
    </tr>
	</table>
</form>
	<% end if %>
<% end if %>
<!--#include file="AdminFooter.asp"-->