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
dim query, conntemp, rs
call openDb()

pcEmail=getUserInput(request("custemail"),0)
pcOrderKey=getUserInput(request("ordercode"),0)

pcErrMsg=""

if pcEmail="" OR pcOrderKey="" then
	pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checkorv_1")
else
	query="SELECT customers.idcustomer,customers.pcCust_Guest,orders.idorder,customers.suspend,customers.pcCust_Locked FROM customers INNER JOIN Orders ON customers.idcustomer=orders.idcustomer WHERE customers.email like '" & pcEmail & "' AND orders.pcOrd_OrderKey like '" & pcOrderKey & "';"
	set rs=connTemp.execute(query)
	if rs.eof then
		pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checkorv_2")
	else
		if rs("suspend")="1" then
			pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checkorv_3")
		end if
		if rs("pcCust_Locked")="1" then
			pcErrMsg=dictLanguage.Item(Session("language")&"_opc_checkorv_4")
		end if
		if rs("pcCust_Guest")="2" AND pcErrMsg="" then
			Session("JustPurchased")="1"
		end if
		pidCustomer=rs("idCustomer")
		pCustomerGuest=rs("pcCust_Guest")
		pidOrder=rs("idOrder")
	end if
end if
set rs=nothing
if pcErrMsg="" then
	if pCustomerGuest<>"0" then
		session("idCustomer")=pidCustomer
		session("CustomerGuest")=pCustomerGuest
	else
		session("REGidCustomer")=pidCustomer
	end if
	pidOrder=clng(pidOrder)+scpre
	tmpURL="CustViewPastD.asp?idorder=" & pidOrder
	response.write "OK|*|" & tmpURL
else
	response.write pcErrMsg
end if
call closeDb()
%>
