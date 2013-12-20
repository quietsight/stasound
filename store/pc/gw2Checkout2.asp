<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<% '//Check if this is a post-back
CartOrderID = request("cart_order_id")

if CartOrderID<>"" then
	gwTransID=request("order_number")

	session("GWAuthCode")=""
	session("GWTransId")=gwTransID
	if session("GWOrderId")="" then
		session("GWOrderId")=CartOrderID
	end if
	session("GWSessionID")=Session.SessionID 
	Response.redirect "gwReturn.asp?s=true&gw=twoCheckout"
end if

'//Set redirect page to the current file name
session("redirectPage")="gw2Checkout2.asp"
session("redirectPage2")="https://www.2checkout.com/2co/buyer/purchase"

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
		
'//Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if

'//Retrieve customer data from the database using the current session id		
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

dim connTemp, rs
call openDb()

'//Retrieve any gateway specific data from database or hard-code the variables
query="SELECT store_id, v2co_TestMode FROM twoCheckout Where id_twoCheckout=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
store_id=rs("store_id")
v2co_TestMode=rs("v2co_TestMode")
set rs=nothing
call closedb()

%>
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<img src="images/checkout_bar_step5.gif" alt="">
			</td>
		</tr>
		<tr>
			<td>
				<form method="POST" action="<%=session("redirectPage2")%>" name="PaymentForm" class="pcForms">
					<input type="hidden" name="sid" value="<%=store_id%>">
					<input type="hidden" name="cart_order_id" value="<%=session("GWOrderId")%>">
					<input type="hidden" name="email" value="<%=pcCustomerEmail%>">
					<input type="hidden" name="total" value="<%=pcBillingTotal%>">
					<% if v2co_TestMode=1 then %>
						<input type="hidden" name="demo" value="Y">
					<% end if %>
					<input type="hidden" name="twoCheckout" value="twoCheckout">
					<% 'select all products from the ProductsOrdered table to insert them into the 2Checkout db.
					call opendb()
					query="SELECT products.idproduct, products.description, quantity, unitPrice FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& session("GWOrderId")
					set rs=server.CreateObject("ADODB.Recordset")
					set rs=connTemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rstemp=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
					IntProdCnt=0
					do until rs.eof
						tempIntIdProduct=rs("idproduct")
						tempStrDescription=rs("description")
						tempIntQuantity=rs("quantity")
						tempDblUnitPrice=rs("unitPrice")
						IntProdCnt=IntProdCnt+1
						%>
						<input type="hidden" name="c_prod_<%=IntProdCnt%>" value="Product_<%=tempIntIdProduct%>,<%=tempIntQuantity%>">
						<input type="hidden" name="id_type" value="1">
						<input type="hidden" name="c_name_<%=IntProdCnt%>" value="<%=tempStrDescription%>">
						<input type="hidden" name="c_description_<%=IntProdCnt%>" value="<%=tempStrDescription%>">
						<input type="hidden" name="c_price_<%=IntProdCnt%>" value="<%=tempDblUnitPrice%>">
						<% rs.moveNext
					loop 
					set rs=nothing
					call closedb() 
					%>
					<input type="hidden" name="card_holder_name" size=35 value="<%=pcBillingFirstName&" "&pcBillingLastName%>">
					<input type="hidden" name="street_address" value="<%=pcBillingAddress%>"> 
					<input type="hidden" name="city" value="<%= pcBillingCity%>"> 
					<input type="hidden" name="state" value="<%=pcBillingState%>"> 
					<input type="hidden" name="zip" value="<%= pcBillingPostalCode %>">
					<input type="hidden" name="country" value="<%= pcBillingCountryCode %>"> 
					<input type="hidden" name="phone" value="<%= pcBillingPhone %>"> 


					<table class="pcShowContent">
					<% if Msg<>"" then %>
						<tr valign="top"> 
							<td colspan="2">
								<div class="pcErrorMessage"><%=Msg%></div>
							</td>
						</tr>
					<% end if %>
					<% if len(pcCustIpAddress)>0 AND CustomerIPAlert="1" then %>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
							<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></p></td>
						</tr>
					<% end if %>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr class="pcSectionTitle">
						<td colspan="2"><p><%=dictLanguage.Item(Session("language")&"_GateWay_1")%></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
						<td colspan="2"><p><%=pcBillingFirstName&" "&pcBillingLastName%></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><%=pcBillingAddress%></p></td>
					</tr>
					<% if pcBillingAddress2<>"" then %>
					<tr>
						<td colspan="2"><p><%=pcBillingAddress2%></p></td>
					</tr>
					<% end if %>
					<tr>
						<td colspan="2"><p><%=pcBillingCity&", "&pcBillingState%><% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><a href="pcModifyBillingInfo.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p></td>
					</tr>
					<% if v2co_TestMode=1 then %>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
							<td colspan="2"><div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_3")%></div></td>
						</tr>
					<% end if %>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr> 
						<td><p><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></p></td>
						<td><%=money(pcBillingTotal)%></td>
					</tr>
					<tr> 
						<td colspan="2" align="center">
							<!--#include file="inc_gatewayButtons.asp"-->
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->