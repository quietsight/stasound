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
<%Dim connTemp,rs,query
IF request("action")="add" THEN

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
	For i=lbound(pcArr) to ubound(pcArr)
	if trim(pcArr(i)<>"") then
		if pcv_PageType="0" then
			if pcv_IsDropShipper="1" then
				query="UPDATE Products SET pcProd_IsDropShipped=1,pcDropShipper_ID=" & pcv_IDSDS & " WHERE idproduct=" & pcArr(i)
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			query="UPDATE Products SET pcSupplier_ID=" & pcv_IDSDS & " WHERE idproduct=" & pcArr(i)
		else
			query="UPDATE Products SET pcProd_IsDropShipped=1,pcDropShipper_ID=" & pcv_IDSDS & " WHERE idproduct=" & pcArr(i)
		end if
		set rs=connTemp.execute(query)
		set rs=nothing
		if (pcv_PageType="1") or (pcv_IsDropShipper="1") then
			query="DELETE FROM pcDropShippersSuppliers WHERE idproduct=" & pcArr(i)
			set rs=connTemp.execute(query)
			set rs=nothing
			query="INSERT INTO pcDropShippersSuppliers (idproduct,pcDS_IsDropShipper) VALUES (" & pcArr(i) & "," & pcv_IsDropShipper & ");"
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	end if
	Next
	call closedb()
	response.redirect "sds_assignprds_msg.asp?action=add&pagetype=" & pcv_PageType
	%>
<%END IF%>
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
			
				'--- Get Search Form parameters ---

				src_IncNormal=getUserInput(request("src_IncNormal"),0)
				src_IncBTO=getUserInput(request("src_IncBTO"),0)
				src_IncItem=getUserInput(request("src_IncItem"),0)
				src_Special=getUserInput(request("src_Special"),0)
				src_Featured=getUserInput(request("src_Featured"),0)
				src_DisplayType=getUserInput(request("src_DisplayType"),0)
				src_ShowLinks=getUserInput(request("src_ShowLinks"),0)
				src_FromPage=getUserInput(request("src_FromPage"),0)
				src_ToPage=getUserInput(request("src_ToPage"),0)
				src_Button2=getUserInput(request("src_Button2"),0)
				src_Button3=getUserInput(request("src_Button3"),0)

				'Start SDBA
				src_PageType=getUserInput(request("src_PageType"),0)
				src_IDSDS=getUserInput(request("src_IDSDS"),0)
				src_IsDropShipper=getUserInput(request("src_IsDropShipper"),0)
				src_sdsAssign=getUserInput(request("src_sdsAssign"),0)
				src_sdsStockAlarm=getUserInput(request("src_sdsStockAlarm"),0)
				'End SDBA
				
				src_PageSize=getUserInput(request("resultCnt"),0)
				
				src_FormTitle1="Find products"
				src_FormTitle2="Assign Products to the "
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select one or more products that you want to assign to the "
				src_Button1=" Search "
				UseSpecial=0
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->