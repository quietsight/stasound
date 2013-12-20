<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<% On Error Resume Next
Dim connTemp, query, rs
call openDb()

Session("pcCustomerTermsAgreed")="1"

if session("idCustomer")<>"" AND  session("idCustomer")<>"0" then
	query="UPDATE Customers SET pcCust_AgreeTerms=1 WHERE idCustomer="&session("idCustomer")&";"
	set rs=connTemp.execute(query)
	set rs=nothing
end if

OKmsg="OK"
response.write OKmsg
%>


