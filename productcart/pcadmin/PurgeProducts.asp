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
<% pageTitle="Purge a Product and Related Orders" %>
<!--#include file="AdminHeader.asp"-->

<% 
on error resume next 

dim query, conntemp, rs

call openDb()

If request("purgeproduct")<>"" then
	' form parameters
	pIdProduct=request("ID")
	if pIdProduct="" then
		pIdProduct=request("idProduct")
	end if
	
	if Cint(pIdProduct)=0 then
		
		call closeDb()
		response.redirect "PurgeProducts.asp?purgeproduct=&message="&Server.Urlencode("You must specify a product to delete from the dropdown list.")
	end if
	
	' delete from taxPrd
	query="DELETE FROM taxPrd WHERE idProduct=" &pidproduct
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 1 purging product on PurgeProducts.asp") 
	end If

	' delete product from configSpec_products
	query="DELETE FROM configSpec_products WHERE configProduct=" &pIdProduct
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 2 purging product on PurgeProducts.asp") 
	end If
	
	' delete product from configSpec_categories
	query="DELETE FROM configSpec_categories WHERE idProduct=" &pIdProduct
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 3 purging product on PurgeProducts.asp") 
	end If
	
	' delete product from cs_relationships
	query="DELETE FROM cs_relationships WHERE idProduct=" &pIdProduct
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 4 purging product on PurgeProducts.asp") 
	end If
	
	' delete product from categories_products
	query="DELETE FROM categories_products WHERE idProduct=" &pIdProduct
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 5 purging product on PurgeProducts.asp") 
	end If
	
	' delete product from wishlist
	query="DELETE FROM wishList WHERE idProduct=" &pIdProduct
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 6 purging product on PurgeProducts.asp") 
	end If
	
	' delete product from options_optionsGroups
	query="DELETE FROM options_optionsGroups WHERE idProduct=" &pIdProduct
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 7 purging product on PurgeProducts.asp") 
	end If
	
	' delete product from configSessions	
	query="DELETE FROM configSessions WHERE idProduct=" &pIdProduct
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 8 purging product on PurgeProducts.asp") 
	end If
	
	' delete product from discountsPerQuantity	
	query="DELETE FROM discountsPerQuantity WHERE idProduct=" &pIdProduct
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 9 purging product on PurgeProducts.asp") 
	end If
	
	' delete product from ProductsOrdered
	query="SELECT idOrder FROM ProductsOrdered WHERE idProduct=" &pIdProduct
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 10 purging product on PurgeProducts.asp") 
	end If
	do until rs.eof
		tempIdOrder=rs("idOrder")
		query="DELETE FROM ProductsOrdered WHERE idOrder=" &tempIdOrder
		set rs2=Server.CreateObject("ADODB.Recordset")
		set rs2=conntemp.execute(query)
		query="DELETE FROM creditCards WHERE idOrder=" &tempIdOrder
		set rs2=conntemp.execute(query)
		query="DELETE FROM offlinepayments WHERE idOrder=" &tempIdOrder
		set rs2=conntemp.execute(query)
		query="DELETE FROM Orders WHERE idOrder=" &tempIdOrder
		set rs2=conntemp.execute(query)
		set rs2=nothing
	rs.movenext
	loop
	
	' delete product from products table
	query="DELETE FROM products WHERE idProduct=" &pIdProduct
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error 11 purging product on PurgeProducts.asp") 
	end If
	
	set rs=nothing
	call closeDb()
	
	%>
	<table class="pcCPcontent">
			<tr>
				<td align="center"><div class="pcCPmessageSuccess"><% response.write "The following product ID was permanently removed: "&pidproduct %>. <a href="PurgeProducts.asp">Remove another product</a>.</div></td>
			</tr>
	</table>
<% else %>

	<form action="PurgeProducts.asp" method="post" name="form" id="form" class="pcForms">
		<table class="pcCPcontent">
			 
		<% if request.querystring("message")<>"" then %>
				
		<tr> 
			<td><div class="pcCPmessage">You must select at least one product to delete.</div></td>
		</tr>
		
		<% end if %>
		
		<%
			' get products
			call openDb()
			query="SELECT products.idProduct, products.description, products.sku FROM products WHERE products.removed <> 0;"
			set rs = Server.CreateObject("ADODB.Recordset")
			set rs = conntemp.execute(query)
			if rs.eof then %>
			
		<tr> 
			<td align="center"><div class="pcCPmessage">There are currently no products in your database that have been deleted from the Control Panel. You can only purge products that have been previously deleted from the store. <br /><br /><a href="LocateProducts.asp?cptype=0">Locate a product</a> to delete it from the Control Panel. <br /><br />If you wish to clean up the store from test products &amp; orders, you can use the <a href="PurgeAllProducts.asp">Purge All Products</a> feature (administrator access only).</div></td>
		</tr>
			
		<% else%>
    
		<tr> 
		<td>
		<div class="pcCPmessage" style="color: #F00; font-weight: bold; font-size: 18px;">WARNING: This is a dangerous feature</div>
          <p>This feature allows you to completely remove a product from your database. When you take this action you will <u>completely <strong>remove the product</strong></u> and <u><strong>all orders</strong> that may contain the product</u>. This action is permanent and cannot be undone.</p>
          <p style="margin-top: 8px;">The main purpose of this feature is to purge products and orders that were entered into your database for <strong>testing purposes</strong>.</p>
          <p style="margin-top: 8px;">Once your store is &quot;live&quot; and there are real orders in your database, you should not perform this action. Instead, you can <strong>delete an item</strong> to hide the product form both the storefront and the Control Panel, without affecting the integrity of previous orders that might contain that product. Here are links to this and other options:  </p>
        <ul>
          <li><a href="LocateProducts.asp?cptype=0">Locate a product and delete it</a> (<em>the product remains in the database, but it's hidden</em>)</li>
          <li><a href="PurgeRemovedProducts.asp">Permanently remove all previously deleted products</a></li>
          <li><a href="PurgeInactiveProducts.asp">Permanently remove all inactive products</a></li>
          <li><a href="PurgeAllProducts.asp">Permanently remove all products</a></li>
        </ul>
        <hr>
<p style="margin-top: 8px;">To completely remove products from the database, enter the product ID or select the product by name using the drop down list below. Only products that have been <strong>deleted</strong> from the Control Panel are shown. If you wish to permanently remove all products, you can use the <a href="purgeAllProducts.asp"><strong>Purge All Products</strong></a> feature.</p>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td><p>Enter a Product ID:	<input name="ID" type="text" id="ID" size="5" maxlength="20"> ... or select a product: 
				<select name="idProduct" id="idProduct">
				<option value="0">Select a Product</option>
				<% do until rs.eof %>
				<option value="<%=rs("idProduct")%>"><%=rs("description")%> (<%=rs("sku")%>)</option>
				<%
						rs.movenext
						loop
						set rs=nothing
						call closeDb()
				%>
				</select>
				</p>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"><hr></td>
	</tr>			
	<tr> 
		<td align="center">
		<input name="purgeproduct" type="submit" id="purgeproduct" value="Permanently Remove Selected Product" class="submit2">
		&nbsp;
		<input type="button" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>
			
<% end if %>

	<tr>
		<td class="pcCPspacer"></td>
	</tr>	
</table>
</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->