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
<!--#include file="../includes/dateinc.asp"-->
<% pageTitle="Purge Orders from the Store Database" %>
<!--#include file="AdminHeader.asp"-->
<% 
on error resume next 

dim query, conntemp, rstemp
call openDb()

If request("purgeorder")<>"" then
	' form parameters
	pIdOrder=request("idOrder")
	if pIdOrder="0" then
		pIdOrder=request("ID")
		pIdOrder=(int(pIdOrder)-scpre)
	end if
	pIdOrderShow=(int(pIdOrder)+scpre)
	
	If Not validNum(pIdOrder) then
		call closeDb()
		response.redirect "PurgeOrders.asp?purgeorder=&message="&Server.Urlencode("You must specify a valid order ID to delete an order from the database.")
	end if
	
	' delete product from ProductsOrdered
		query="DELETE FROM ProductsOrdered WHERE idOrder=" &pIdOrder
		set rstemp=server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		query="DELETE FROM creditCards WHERE idOrder=" &pIdOrder
		set rstemp=conntemp.execute(query)
		query="DELETE FROM offlinepayments WHERE idOrder=" &pIdOrder
		set rstemp=conntemp.execute(query)
		query="DELETE FROM Orders WHERE idOrder=" &pIdOrder
		set rstemp=conntemp.execute(query)
	
		if err.number <> 0 then
			pcvErrDescription = err.description
			set rstemp = nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in PurgeOrders: "&pcvErrDescription) 
		end If
		
		set rstemp=nothing
	
	%>
	<table class="pcCPcontent">
		<tr> 
			<td>
				<div class="pcCPmessageSuccess"><% response.write "Order successfully deleted: " & pIdOrderShow %>. <a href="PurgeOrders.asp">Delete another order</a>.</div>
			</td>
		</tr>
	</table>
<% else %>
	<form action="PurgeOrders.asp" method="post" name="form" id="form" class="pcForms">
		<table class="pcCPcontent">
			<tr> 
				<td>
					<% if request.querystring("message")<>"" then %>
						<div class="pcCPmessage">Either the order number you entered is not valid or it could not be located in the database</div>
					<% else %>
					<h2 style="color:#F00;">WARNING: this is a dangerous feature</h2>
					You may use this form to completely purge an order from your database. <u>This action is permanent and cannot be reversed</u>. The main purpose of this feature is to purge orders that were entered into the store database for testing purposes. Once your store is &quot;live&quot; and there are real orders in your database, you should &quot;cancel&quot; orders instead, which you can do from the Order Details page.
         	<p>&nbsp;</p>
        </td>
       </tr>
         <% end if 
			' get orders
			query="SELECT orders.idOrder, orders.orderDate FROM orders ORDER BY idOrder DESC;"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			if rs.eof then %>
				<tr> 
					<td>
						<div class="pcCPmessage">No Orders Found.</div>
					</td>
				</tr>
      <% else %>
				<tr> 
					<td>Enter the <b>Order ID:</b> <input name="ID" type="text" id="ID" size="5" maxlength="20"> ... or <b>Select an Order:</b> 
						<select name="idOrder">
							<option value="0">Select an Order</option>
							<%
								do until rs.eof
								pIdOrder=rs("idOrder")
								pIdOrderShow=(pIdOrder+scpre)
								pOrderDate=rs("orderDate") %>
								<option value="<%=pIdOrder%>"><%response.write "Order #: "&pIdOrderShow & " - Date: " & ShowDateFrmt(pOrderDate) %></option>
							<% 
								rs.movenext
								loop
								set rs=nothing
							%>
						</select>
					</td>
				</tr>
				<tr> 
					<td class="pcCPspacer"></td>
				</tr>
				<tr> 
					<td><input name="purgeorder" type="submit" id="purgeorder" value="Permanetly remove this order" class="submit2"></td>
				</tr>
				<% end if %>
				<tr> 
					<td class="pcCPspacer"></td>
				</tr>
     </table>
	</form>
<% end if
call closeDb()
%>
<!--#include file="AdminFooter.asp"-->