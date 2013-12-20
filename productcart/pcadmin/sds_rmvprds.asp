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

pageTitle="Delete Products from the Selected " & pcv_Title%>
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
Dim connTemp,rs,query
IF request("action")="del" THEN

	pcv_PageType=request("pagetype")
	if pcv_PageType="" then
		pcv_PageType=0
	end if
	pcv_IDSDS=request("idsds")
	pcv_IsDropShipper=request("isdropshipper")
	pcv_PrdList=request("prdlist")
	if (trim(pcv_PrdList)="") or (pcv_IDSDS="") or (pcv_IDSDS="0") then
		response.redirect "menu.asp"
	end if
	pcArr=split(pcv_PrdList,",")
	call opendb()
	pcIntPrdCount = ubound(pcArr)
	For i=lbound(pcArr) to ubound(pcArr)
	if trim(pcArr(i)<>"") then
		if pcv_PageType="0" then
			if pcv_IsDropShipper="1" then
				query="DELETE FROM pcDropShippersSuppliers WHERE idproduct=" & pcArr(i)
				set rs=connTemp.execute(query)
				set rs=nothing
				query="UPDATE Products SET pcProd_IsDropShipped=0,pcDropShipper_ID=0 WHERE idproduct=" & pcArr(i)
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			query="UPDATE Products SET pcSupplier_ID=0 WHERE idproduct=" & pcArr(i)
		else
			query="DELETE FROM pcDropShippersSuppliers WHERE idproduct=" & pcArr(i)
			set rs=connTemp.execute(query)
			set rs=nothing
			query="UPDATE Products SET pcProd_IsDropShipped=0,pcDropShipper_ID=0 WHERE idproduct=" & pcArr(i)
		end if
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	Next
	call closedb()
	%>
	<div class="pcCPmessageSuccess">
		<%=pcIntPrdCount%> product(s) were successfully deleted from the selected <%if pcv_PageType="1" then%>Drop-Shipper<%else%>Supplier<%end if%>!
		<br /><br />
		<a href="sds_manage.asp?pagetype=1">Manage Drop-Shippers</a>
		&nbsp;|&nbsp;
		<a href="sds_manage.asp?pagetype=0">Manage Suppliers</a>
	</div>
<%END IF%>
<!--#include file="AdminFooter.asp"-->