<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
if request("pagetype")="1" then
	pcv_PageType=1
	pcv_Title="Drop-Shipper"
	pcv_Table="pcDropShipper"
else
	pcv_PageType=0
	pcv_Title="Supplier"
	pcv_Table="pcSupplier"
end if

pageTitle="Delete a " & pcv_Title %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="AdminHeader.asp"-->

<% Dim connTemp,rs,query

IF (request("action")="del") THEN
	pcv_idsds=request("idsds")
	if not validNum(pcv_idsds) then
		response.redirect "menu.asp"
	end if

	call opendb()

	query="DELETE FROM " & pcv_Table & "s WHERE " & pcv_Table & "_ID=" & pcv_idsds
	set rs=connTemp.execute(query)
	set rs=nothing
	
	call closedb()		 
	
	pcMessage=pcv_Title & " was deleted successfully!"
%>

		<div class="pcCPmessageSuccess">
			<%=pcMessage%>
			<br /><br />
			<a href="sds_manage.asp?pagetype=1">Manage Drop-Shippers</a>
			&nbsp;|&nbsp;
			<a href="sds_manage.asp?pagetype=0">Manage Suppliers</a>
		</div>

<%
ELSE
	response.redirect "menu.asp"
END IF
%>
<!--#include file="AdminFooter.asp"-->