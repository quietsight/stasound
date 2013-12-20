<%@ Language=VBScript %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp"-->
<%
if Session.SessionID<>session("GWSessionID") then
	response.redirect "techErr.asp?error=You do not have proper rights to access this page."
end if

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
call openDB()
query="SELECT ssl_merchant_id, ssl_pin, CVV,ssl_avs,testmode, ssl_user_id FROM klix Where idKlix=1"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
	
pcv_SSLMerchantID=rs("ssl_merchant_id")
pcv_SSLPin=rs("ssl_pin")
pcv_SSLPin=enDeCrypt(pcv_SSLPin, scCrypPass)
pcv_CVV=rs("CVV")
pcv_SSLAVS=rs("ssl_avs")
pcv_TestMode=rs("testmode")
pcv_UserID=rs("ssl_user_id")

if err.number <> 0 then
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
end If 

set rs=nothing
call closedb() %>
<HTML>
<HEAD>
</HEAD>
<body onLoad="document.PaymentInfo.submit();">
<form action="https://www.viaKLIX.com/process.asp" method="post" name="PaymentInfo">
<input type="hidden" name="orderTotal" value="<%=session("GWOrderId")%>"><br>
<input type="hidden" name="ssl_transaction_type" value="SALE">
<input type="hidden" name="ssl_salestax" value="0.00">
<input type="hidden" name="ssl_merchant_id" value="<%=pcv_SSLMerchantID%>">
<input type="hidden" name="ssl_pin" value="<%=pcv_SSLPin%>">
<input type="hidden" name="ssl_amount" value="<%=pcBillingTotal%>">
<input type="hidden" name="ssl_show_form" value="false">
<% if pcv_TestMode="1" then %>
	<input type="hidden" name="ssl_test_mode" value="TRUE">
<%else%>
	<input type="hidden" name="ssl_test_mode" value="FALSE">
<%end if%>
<% 
if scSSL="1" then
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwKlixReceipt.asp"),"//","/")
else
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwKlixReceipt.asp"),"//","/")
end if
tempURL=replace(tempURL,"http:/","http://")
tempURL=replace(tempURL,"https:/","https://")
%>
<input type="hidden" name="ssl_result_format" value="HTML">
<input type="hidden" name="ssl_receipt_decl_method" value="REDG">
<input type="hidden" name="ssl_receipt_decl_get_url" value="<%=tempURL%>">
<input type="hidden" name="ssl_receipt_apprvl_method" value="REDG">
<input type="hidden" name="ssl_receipt_apprvl_get_url" value="<%=tempURL%>">
<input type="hidden" name="ssl_invoice_number" value="<%=session("GWOrderId")%>">
<input type="hidden" name="ssl_customer_code" value="<%=session("GWOrderId")%>">
<input type="hidden" name="ssl_user_id" value="<%=pcv_UserID%>">
<input type="hidden" name="ssl_email" value="<%=pcCustomerEmail%>">
<input type="hidden" name="ssl_city" value="<%=pcBillingCity%>">
<input type="hidden" name="ssl_state" value="<%=pcBillingState%>">
<input type="hidden" name="ssl_billto_postal_name_first" value="<%=pcBillingFirstName%>">
<input type="hidden" name="ssl_billto_postal_name_last" value="<%=pcBillingLastName%>">
<input type="hidden" name="ssl_billto_business_name" value="<%=pcBillingCompany%>">
<input type="hidden" name="ssl_avs_address" value="<%=pcBillingAddress%>">
<input type="hidden" name="ssl_billto_postal_city" value="<%=pcBillingCity%>"> 
<input type="hidden" name="ssl_billto_postal_stateprov" value="<%=pbillingState%>"> 
<input type="hidden" name="ssl_avs_zip" value="<%=pcBillingPostalCode%>">
<input type="hidden" name="ssl_billto_postal_countrycode" value="<%=pcBillingCountryCode%>"> 
<input type="hidden" name="ssl_phone_number" value="<%=pcBillingPhone%>"> 
<input type="hidden" name="ssl_shipto_same_as_billto" value="">
<input type="hidden" name="ssl_shipto_postal_name_first" value="<%=pcShippingFirstName%>">
<input type="hidden" name="ssl_shipto_postal_name_last" value="<%=pcShippingLastName%>">
<input type="hidden" name="ssl_shipto_business_name" value="<%=pcShippingCompany%>">
<input type="hidden" name="ssl_shipto_postal_street_line1" value="<%=pcShippingAddress%>">
<input type="hidden" name="ssl_shipto_postal_street_line2" value="<%=pcShippingAddress2%>">
<input type="hidden" name="ssl_shipto_postal_city" value="<%=pcShippingCity%>">
<input type="hidden" name="ssl_shipto_postal_stateprov" value="<%=shippingState%>">
<input type="hidden" name="ssl_shipto_postal_postalcode" value="<%=pcShippingPostalCode%>">
<input type="hidden" name="ssl_shipto_postal_countrycode" value="<%=pcShippingCountryCode%>">
<input type="hidden" name="ssl_phone_number" value="<%=pcShippingPhone%>">
<input type="hidden" name="ssl_receipt_link_text" value="Continue">
<input type="hidden" name="ssl_card_name" value="<%=pcBillingFirstName&" "&pcBillingLastName%>">
<input type="hidden" name="ssl_card_number" value="<%=request.Form("CardNumber")%>">
<input type="hidden" name="ssl_exp_date" value="<%=request.Form("expMonth")&request.Form("expYear")%>">
<% if pcv_CVV="1" then %>
	<input type="hidden" name="ssl_cvv2" value="<%=request.Form("ssl_cvv2")%>">
	<input type="hidden" name="ssl_cvv2cvc2" value="<%=request.form("ssl_cvv2cvc2")%>">
<% end if %>
</form>
</body>
</html>