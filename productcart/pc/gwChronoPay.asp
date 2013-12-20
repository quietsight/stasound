<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'Gateway File: gwChronopay.asp
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
<% 
'// See if this is a response back from ChronoPay
Transaction_Id = request("transaction_id")
if Transaction_Id <> "" then
	session("GWAuthCode")=""
	session("GWTransId")=request("transaction_id")
	if session("GWOrderId")="" then
		session("GWOrderId")=request("idOrder")
	end if
	session("GWSessionID")=Session.SessionID 
	'// Payment is received - redirect to gwReturn.asp
	response.Redirect("gwReturn.asp?s=true&gw=ChronoPay")
end if

'======================================================================================
'// Set redirect page
'======================================================================================
session("redirectPage")="gwChronoPay.asp"
session("redirectPage2")="https://secure.chronopay.com/index_shop.cgi"

'======================================================================================
'// DO NOT ALTER BELOW THIS LINE	
'======================================================================================
': Declare and Retrieve Customer's IP Address
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
': End Declare and Retrieve Customer's IP Address	

': Declare URL path to gwSubmit.asp	
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
': End Declare URL path to gwSubmit.asp

': Get Order ID and Set to session
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if
': End Get Order ID

': Get customer and order data from the database for this order	
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<%
': End Get customer and order data


': Reset customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer
': End Reset customer session

': Open Connection to the DB
dim connTemp, rs
call openDb()
'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE	
'======================================================================================

'======================================================================================
'// Retrieve ALL required gateway specific data from database

query="SELECT CP_ProdID, CP_Currency, CP_testmode FROM pcPay_Chronopay WHERE CP_id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
CP_ProdID=rs("CP_ProdID")
CP_Currency=rs("CP_Currency")
CP_testmode=rs("CP_testmode")

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
				<form method="POST" action="<%=session("redirectPage2")%>" name="payment_form" class="pcForms">
					<input type="hidden" name="product_id" value="<%=CP_ProdID%>">
					<% 'select all products from the ProductsOrdered table to post them into the Chronopay.
					call opendb()
					query="SELECT products.idproduct, products.description FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idOrder="& session("GWOrderId")
					set rs=server.CreateObject("ADODB.Recordset")
					set rs=connTemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rstemp=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
					tempStrDescription=""
					do until rs.eof
						if(tempStrDescription="") then
							tempStrDescription=tempStrDescription & rs("description")
						else
							tempStrDescription=tempStrDescription & " , " & rs("description")
						end if
						rs.moveNext
					loop 
					set rs=nothing
					call closedb() 
					%>

					<input type="hidden" name="product_name" value="<%=tempStrDescription%>">
					<input type="hidden" name="product_price" value="<%=pcBillingTotal%>">
					<input type="hidden" name="language" value="En">
					<input type="hidden" name="f_name" value="<%=pcBillingFirstName%>">
					<input type="hidden" name="s_name" value="<%=pcBillingLastName%>">
					<input type="hidden" name="street" value="<%=pcBillingAddress%>">
					<input type="hidden" name="city" value="<%=pcBillingCity%>">
					<input type="hidden" name="state" value="<%=pcBillingState%>">
					<input type="hidden" name="zip" value="<%=pcBillingPostalCode%>">
					<input type="hidden" name="country" value="<%=pcBillingCountryCode%>">
					<input type="hidden" name="phone" value="<%=pcBillingPhone%>">
					<input type="hidden" name="email" value="<%=pcCustomerEmail%>">
					<input type="hidden" name="cb_url" value="<%=replace((scStoreURL&"/"&scPcFolder&"/pc/gwreturn.asp"),"//","/")%>">
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
						<td colspan="2"><p><%=pcBillingCity&", "&pcBillingState%>
						<% if pcBillingPostalCode<>"" then%>&nbsp;<%=pcBillingPostalCode%><% end if %></p></td>
					</tr>
					<tr>
						<td colspan="2"><p><a href="pcModifyBillingInfo.asp"><%=dictLanguage.Item(Session("language")&"_GateWay_2")%></a></p></td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% 
					'======================================================================================
					'// if the cart is in testmode, alert the customer that this is not a live transaction.
					'======================================================================================
					if CP_testmode="YES" then %>
						<tr>
							<td colspan="2">
								<div class="pcErrorMessage">
									<%=dictLanguage.Item(Session("language")&"_GateWay_3")%>
								</div>
							</td>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
					<% end if
					'======================================================================================
					'// End Testing environment variable
					'======================================================================================
					%>

					
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