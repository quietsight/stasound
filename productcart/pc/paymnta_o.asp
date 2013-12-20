<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

response.Buffer=true

Response.Expires = 60
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<% dim conntemp, query, rs
'//Set redirect page to the current file name
session("redirectPage")="paymnta_o.asp"

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")

' Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=getUserInput(request("idOrder"),0)
end if
if not validNum(session("GWOrderId")) then
   response.redirect("msg.asp?message=64")
end if

dim pcTempIdPayment
pcTempIdPayment=getUserInput(request("idPayment"),0)

if session("GWPaymentId")="" then
	session("GWPaymentId")=pcTempIdPayment
else
	if pcTempIdPayment<>session("GWPaymentId") AND pcTempIdPayment<>"" then
		session("GWPaymentId")=pcTempIdPayment
	end if
end if

' extract real idorder (without prefix)
pTrueOrderId=(int(session("GWOrderId"))-scpre)
If Not validNum(pTrueOrderId) then
	response.redirect "msg.asp?message=10"
End If

'redirect to gwReturn.asp
Response.redirect "gwReturn.asp?s=true&gw=OFFLINECC"
%>