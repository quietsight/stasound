<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/rc4.asp"-->
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<%
Dim rs, connTemp, query, IDBrand

call openDB()

	IDBrand=request.querystring("idbrand")
	
	if not validNum(IDBrand) then response.redirect "BrandsManage.asp?msg="&Server.URLEncode("Not a valid brand ID.")
	
	query="DELETE FROM Brands WHERE IDBrand=" & IDBrand
	set rs=server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error removing brand from Brands table") 
	end if
	
	query="UPDATE Products SET IDBrand=0 WHERE IDBrand=" & IDBrand
	set rs=connTemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error removing brand from Products table") 
	end if
	
	set rs=nothing
	
call closeDB()
response.redirect "BrandsManage.asp?s=1&msg="&Server.URLEncode("Brand deleted successfully.")
response.End()
%>