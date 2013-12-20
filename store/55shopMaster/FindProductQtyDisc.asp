<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<%
Dim rs, connTemp, query, idProduct

	' This page is used to link to the quantity discount page
	' The page is different depending on the product type
	
	' Get product id
	idProduct = request.QueryString("id")
		if idProduct = "" then
			idProduct = request.QueryString("idproduct")
		end if

	' Sanitize product id
	if not validNum(idProduct) then
		response.redirect "menu.asp"
		response.End()
	end if
		
	' Load data from database
	call openDb()
	pcv_HaveQtyDisc=0
	query= "SELECT idproduct FROM discountsPerQuantity WHERE discountDesc='PD' AND idproduct="&idProduct
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)
	if NOT rs.eof then
		pcv_HaveQtyDisc=1
	end if
	set rs=nothing
	call closeDb()
	
	' Redirect to the right page
	if pcv_HaveQtyDisc = 1 Then
			' Edit existing discounts
			response.redirect "ModDctQtyPrd.asp?idproduct="&idProduct
		else
			' Add new discounts
			response.redirect "AdminDctQtyPrd.asp?idproduct="&idProduct
	end if
%>