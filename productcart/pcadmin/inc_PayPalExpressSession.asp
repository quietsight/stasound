<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/pcPayPalClass.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<%
dim rs,connTemp,query

if session("admin") = 0 then
	response.clear
	response.write "SECURITY"
	response.End
end if

call openDb()

if request("checked") = 1 then
	session("pcSetupPayPalExpress") = 1
	response.Write("OK")
	response.End()
else
	session("pcSetupPayPalExpress") = ""
	response.Write("NOTOK")
	response.End()
end if

call closedb()
%>