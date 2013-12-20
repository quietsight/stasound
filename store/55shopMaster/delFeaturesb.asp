<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle="Delete Item" %>
<% 
dim mySQL, conntemp, rstemp
pIdproduct=request.Querystring("idproduct")
if not validNum(pIdproduct) then
   response.redirect "msg.asp?message=2"
end if

call openDb()
mySQL="UPDATE products SET showInHome=0 WHERE idproduct=" &pidproduct
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(mySQL)
set rstemp=nothing
call closeDb()

if err.number <> 0 then
  response.redirect "techErr.asp?error="& Server.Urlencode("Error in delSpc: "&Err.Description) 
end If

response.redirect "AdminFeatures.asp?s=1&msg=" & Server.URLEncode("The selected product was removed from the list of Featured Products.")
%>