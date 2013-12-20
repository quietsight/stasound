<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% Dim pageTitle, Section
pageTitle="Delete Feedback Status"
Section="layout" %>
<%
Dim rstemp, connTemp, mySQL
call openDB()
lngIDPro=getUserInput(request("IDPro"),0)

If (lngIDPro="1") or (lngIDPro="2") then
	response.redirect "adminFBStatusManager.asp?msg=Default Message Status: cannot be edited or removed"
End If

mySQL="delete from pcFStatus where pcFStat_IDStatus=" & lngIDPro
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(mySQL)

mySQL="update pcComments set pcComm_FStatus=0 where pcComm_FStatus=" & lngIDPro
set rstemp=connTemp.execute(mySQL)
set rstemp=nothing
call closedb()

response.redirect "adminFBStatusManager.asp?s=1&msg=Message Status removed successfully!"
%>