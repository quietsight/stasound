<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle="Remove Product from Control Panel" %>
<% 
dim query, conntemp, rs

' form parameters
pIdProduct=request.Querystring("idProduct")

if not ValidNum(pIdProduct) then
  response.redirect "msg.asp?message=2"
end if

call openDb()

' delete from taxPrd
query="DELETE FROM taxPrd WHERE idProduct=" &pidproduct
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	call closeDb()
  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

' delete product from configSpec_products
query="DELETE FROM configSpec_products WHERE configProduct=" &pIdProduct
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	call closeDb()
  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

' delete product from cs_relationships
query="DELETE FROM cs_relationships WHERE idProduct=" &pIdProduct
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	call closeDb()
  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

' delete product from categories_products
query="DELETE FROM categories_products WHERE idProduct=" &pIdProduct
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	call closeDb()
  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

' delete product from products table
query="UPDATE products SET active=0, removed=-1 WHERE idproduct=" &pidproduct
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	call closeDb()
  response.redirect "techErr.asp?error="& Server.Urlencode("Error deleting a product - delPrdb.asp") 
end If

set rs=nothing
call closeDb()

response.redirect "srcPrds.asp"

%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr>
		<td><div class="pcCPmessageSuccess">Product ID <%response.write pidproduct%> successfully removed from the Control Panel.</div></td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>Note: to preserve the integrity of previous orders, when you delete a product <strong>the item is not completely deleted from your store</strong>, but rather only removed from the Control Panel. The product still exists in the &quot;products&quot; table in your store database. To completely remove a product from your database, you can use the &quot;<a href="PurgeProducts.asp">Purge Products</a>&quot; feature (only the Master Administrator has access to this feature).</td>
	</tr>
	<tr>
		<td>
		<ul>
		<li><a href="LocateProducts.asp?cptype=0">Locate another product</a></li>
		<li><a href="menu.asp">Start page</a></li>
		</ul>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->