<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Special Customer Fields" %>
<% Section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->

<% 
Dim connTemp,rs,query

pcv_ID=request("idSearchField")
if pcv_ID="" or pcv_ID="0" then
	response.redirect "ManageSearchFields.asp"
end if

call opendb()
query="DELETE FROM pcSearchFields_Products WHERE idSearchData IN (SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & pcv_ID & ");"
Set rstemp=conntemp.execute(query)
Set rstemp=nothing

query="DELETE FROM pcSearchData WHERE idSearchField=" & pcv_ID & ";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
set rs=nothing

query="DELETE FROM pcSearchFields WHERE idSearchField=" & pcv_ID & ";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
set rs=nothing

call closedb()
response.redirect "ManageSearchFields.asp?s=1&msg=" & server.URLEncode("Custom Search Field deleted successfully.")
%>