<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% Dim pageTitle, Section
pageTitle="Delete Message Priority Level"
Section="layout" %>
<%
Dim rstemp, connTemp, query
call openDB()

lngIDPro=getUserInput(request("IDPro"),0)

query="delete from pcPriority where pcPri_IDPri=" & lngIDPro
set retemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)

query="update pcComments set pcComm_Priority=0 where pcComm_Priority=" & lngIDPro
set rstemp=connTemp.execute(query)
set rstemp=nothing

call closedb()

response.redirect "adminFBPriorityManager.asp?s=1&msg=Message Priority Level removed successfully!"
%>