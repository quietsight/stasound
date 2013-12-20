<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% Dim pageTitle, Section
pageTitle="Delete Feedback Type"
Section="layout" %>
<%
lngIDPro=getUserInput(request("IDPro"),0)
if not validNum(lngIDPro) then
	response.redirect "adminFBTypeManager.asp?msg=Not a valid Feedback Type ID"
end if
Dim rstemp, connTemp, query
call openDB()
query="delete from pcFTypes where pcFType_IDType=" & lngIDPro
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
query="update pcComments set pcComm_FType=0 where pcComm_FType=" & lngIDPro
set rstemp=connTemp.execute(query)
call closedb()
response.redirect "adminFBTypeManager.asp?s=1&msg=The feedback type has been removed successfully!"
%>