<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle="Delete Customer" %>
<!--#include file="AdminHeader.asp"-->
<% 

dim query, conntemp, rstemp, rs

pIdCustomer=request.Querystring("idcustomer")

'find all orders associated with customer
if request.QueryString("delconfirm")="YES" then
	call openDb()
	on error resume next
	If request.QueryString("delALL")="YES" then
		query="SELECT idOrder FROM orders WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		do until rs.eof
			query="DELETE FROM prdOrders WHERE idOrder="&rs("idOrder")
			set rstemp=conntemp.execute(query)
			query="DELETE FROM ProductsOrdered WHERE idOrder="&rs("idOrder")
			set rstemp=conntemp.execute(query)
			query="DELETE FROM creditcards WHERE idOrder="&rs("idOrder")
			set rstemp=conntemp.execute(query)
		rs.moveNext
		loop

		query="DELETE FROM orders WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM recipients WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM DPRequests WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
	
		query="DELETE FROM used_discounts WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM wishlist WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcCustomerSessions WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
	
		query="DELETE FROM pcEvProducts WHERE pcEP_IDEvent IN (SELECT pcEv_IDEvent FROM pcEvents WHERE pcEv_IDCustomer="&pIdCustomer&");"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcEvents WHERE pcEv_IDCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcCustomerFieldsValues WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM authorders WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcPay_EIG_Authorize WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM netbillorders WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pfporders WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcPay_eMerch_Orders WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)

		query="DELETE FROM pcPay_LinkPointAPI WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcPay_PayPal_Authorize WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcPay_USAePay_Orders WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcTaxEptCust WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcSavedCarts WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcCustomerTermsAgreed WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)
		
		query="DELETE FROM pcQB_CMaps WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query)

		query="DELETE FROM customers WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query) %>
			<div class="pcCPmessageSuccess">Customer successfully deleted. <a href="viewCusta.asp">Locate another customer</a>.</div>
	<% end if
	if request.QueryString("delALL")="NO" then
		'update customer table
		query="UPDATE customers SET email='REMOVED' WHERE idCustomer="&pIdCustomer&";"
		set rs=conntemp.execute(query) %>
		<table class="pcCPcontent">
			<tr>
				<td>
					<p>&nbsp;</p>
					<p><br>Customer successfully removed.</p>
					<p><a href="viewCusta.asp">Locate another customer</a></p>
					<p>&nbsp;</p>
				</td>
			</tr>
		</table>
	<% end if
	call closeDb()
else 
	call openDb()
	query="SELECT idOrder FROM orders WHERE orderStatus <> 1 AND idCustomer="&pIdCustomer&";"
	set rs=conntemp.execute(query)
	if NOT rs.eof then %>
		<table class="pcCPcontent">
			<tr>
				<td>
					<p>&nbsp;</p>
					<p><strong>Please note:</strong> there are orders associated with this customer's account.</p>
				  <p><a href="viewCustOrders.asp?idcustomer=<%=pIdCustomer%>">View Orders</a></p>
				  <hr>
				  <p>If you opt to '<strong>Remove Customer</strong>', the customer's account will be removed from the Control Panel, but customer information will be kept in the store database in order not to alter order details.</p>
				  <p>The customer will no longer be able to log into the store.</p>
					<p>The customer's e-mail and password will be replaced with a generic text string for both privacy protection and to allow the customer to register again in the future using the same e-mail address.<br>
					</p>
				  <p><a href="delCustomer.asp?delconfirm=YES&delAll=NO&idcustomer=<%=pIdCustomer%>">Remove Customer</a></p>
				  <hr size="1" color="#e1e1e1">
				  <p>If you opt to '<strong>Delete Customer</strong>', the customer's account will be permanently deleted from the database and all associated orders will also be permanently removed. This action can not be undone. This is a good option when you are removing test customers and the associated test orders.</p>
				  <p><a href="delCustomer.asp?delconfirm=YES&delAll=YES&idcustomer=<%=pIdCustomer%>">Delete Customer</a></p>
					<hr>
				  <p>
					<a href="viewCusta.asp">Locate another customer</a> - <a href="javascript:history.back()">Back</a></p></td>
			</tr>
		</table>
	<% else 
		response.redirect "delCustomer.asp?delconfirm=YES&delAll=YES&idcustomer="&pIdCustomer
	end if
end if %><!--#include file="AdminFooter.asp"-->