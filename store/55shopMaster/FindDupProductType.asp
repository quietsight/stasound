<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=0%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<%
Dim rs, connTemp, query, idProduct, pcIntConfigOnly, pcIntserviceSpec

	' This page is used to link to the close product page
	' The page is different depending on the product type
	
	' Get product id
	idProduct = request.QueryString("id")
		if idProduct = "" then
			idProduct = request.QueryString("idproduct")
		end if
		
	' Get "tab" querystring, if it exists:
	tab = request.QueryString("tab")
		if validNum(tab) then
			tabQS = "&tab=" & tab & "#TabbedPanels1"
			else
			tabQS = ""
		end if
		
	' Load data from database
	call openDb()
	query="SELECT configOnly, serviceSpec FROM products WHERE idproduct="&idProduct
	Set rs=Server.CreateObject("ADODB.Recordset")   
	Set rs=connTemp.execute(query)
	pcIntConfigOnly = rs("configOnly")
	if IsNull(pcIntConfigOnly) or pcIntConfigOnly="" then
		pcIntConfigOnly=0
	end if
	pcIntserviceSpec = rs("serviceSpec")
	if IsNull(pcIntserviceSpec) or pcIntserviceSpec="" then
		pcIntserviceSpec=0
	end if
	Set rs=nothing
	call closeDb()
	
	' Find out the type of product
	if Cint(pcIntConfigOnly) <> 0 Then ' This is a Build To Order Only item
			response.redirect "AddDupProduct.asp?idProduct="&idProduct&"&prdType=item" & tabQS
		elseif Cint(pcIntserviceSpec) <> 0 Then ' This is a Built To Order product
			response.redirect "AddDupProduct.asp?idProduct="&idProduct&"&prdType=bto" & tabQS
		else
			response.redirect "AddDupProduct.asp?idProduct="&idProduct&"&prdType=std" & tabQS
	end if
%>