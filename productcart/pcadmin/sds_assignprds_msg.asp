<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
if (request("pagetype")="1") or (request("src_pagetype")="1") then
	pcv_PageType=1
	pcv_Title="Drop-Shipper"
	pcv_Table="pcDropShipper"
else
	pcv_PageType=0
	pcv_Title="Supplier"
	pcv_Table="pcSupplier"
end if

pageTitle="Assign Products to the " & pcv_Title%>
<% Section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
IF request("action")="add" THEN
	%>
	<div class="pcCPmessageSuccess">
		Selected products were assigned to the <%if pcv_PageType="1" then%>Drop-Shipper<%else%>Supplier<%end if%> successfully!
		<br /><br />
		<a href="sds_manage.asp?pagetype=1">Manage Drop-Shippers</a>
		&nbsp;|&nbsp;
		<a href="sds_manage.asp?pagetype=0">Manage Suppliers</a>
	</div>
<%END IF%>
<!--#include file="AdminFooter.asp"-->