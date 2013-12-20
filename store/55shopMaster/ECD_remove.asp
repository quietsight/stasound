<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Remove Endicia Account" %>
<% response.Buffer=true %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<%
Dim connTemp,rs,query

call opendb()

query="Delete FROM pcEDCSettings WHERE pcES_Reg=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
set rs=nothing

call closedb()

response.redirect "viewShippingOptions.asp?msg=Endicia Postage Label Services has been successfully removed.&s=1"
%>
