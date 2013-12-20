<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Special Customer Fields" %>
<% Section="layout" %>
<%PmAdmin=7%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 
Dim connTemp,rs,query
pcv_ID=request("idcustfield")
if not validNum(pcv_ID) then
	response.redirect "menu.asp"
end if

call opendb()

query="DELETE FROM pcCustomerFieldsValues WHERE pcCField_ID=" & pcv_ID & ";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
set rs=nothing

query="DELETE FROM pcCustomerFields WHERE pcCField_ID=" & pcv_ID & ";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
set rs=nothing

query="DELETE FROM pcCustFieldsPricingCats WHERE pcCField_ID=" & pcv_ID & ";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
set rs=nothing

call closedb()
response.redirect "manageCustFields.asp?message=2"
%>