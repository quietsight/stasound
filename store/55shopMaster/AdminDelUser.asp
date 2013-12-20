<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<%
Dim rs, connTemp, query

IDAdmin=request("ID")
if not validNum(IDAdmin) then
	response.redirect "AdminUserManager.asp?msg=" & Server.Urlencode("The user does not exist.")
end if

if session("PmAdmin")<>"19" then
	response.redirect "AdminUserManager.asp?r=1&msg=" & Server.Urlencode("You don't have permissions to delete this user.")
else
	call openDb()
	query="delete from Admins where ID=" & IDAdmin
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
	call closeDb()
	response.redirect "AdminUserManager.asp?s=1&msg=" & Server.Urlencode("This user was deleted successfully!")
end if
%>