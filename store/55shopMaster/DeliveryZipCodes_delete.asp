<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp"-->
<%PmAdmin="1*4*"%><!--#include file="adminv.asp"--> 
<%
	Dim rs, connTemp, query
	zipcode=request("zipcode")
	call openDB()
	query="delete from ZipCodeValidation where zipcode='" & zipcode & "'"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
	call closeDb()
	response.redirect "DeliveryZipCodes_main.asp?s=1&msg=This ZIP Code was deleted successfully!"
%>