<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Cross Selling - General Settings" %>
<% Section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<% 
dim mySQL, conntemp, rs
call openDb() 
set rs=Server.CreateObject("ADODB.Recordset")
ptype=request.QueryString("type")
idmain=request.QueryString("idmain")
if ptype="1" then
	'delete all relationships in cs_relationships table
	query="DELETE FROM cs_relationships WHERE idproduct="&idmain
	set rs=conntemp.execute(query)
else
	idcrosssell=request.QueryString("idcrosssell")
	query="DELETE FROM cs_relationships WHERE idcrosssell="&idcrosssell&";"
	set rs=conntemp.execute(query)
end if
set rs=nothing
call closeDb()
If ptype="1" then
	response.redirect "crossSellView.asp?idmain="&idmain
else
	response.redirect "crossSellEdit.asp?idmain="&idmain&"&s=1&msg=" & server.URLEncode("Item successfully removed from the cross-selling relationship.")
end if
%>
<!--#include file="AdminFooter.asp"-->
