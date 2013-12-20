<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/CPconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/SocialNetworkWidgetConstants.asp"-->
<%  Dim query, conntemp, rstemp, rs

call opendb()

'// Category ID
pIdCategory=SNW_CATEGORY
if pIdCategory="" OR isNULL(pIdCategory) then
	pIdCategory=0
end if

'// Affiliate ID
idaffiliate=Request("idaffiliate")

'// Sort
if ProdSort="" then
	ProdSort="19"
end if

select case ProdSort
	Case "19": query1 = " ORDER BY categories_products.POrder Asc"
	Case "0": query1 = " ORDER BY products.SKU Asc"
	Case "1": query1 = " ORDER BY products.description Asc" 	
	Case "2": 
		If Session("customerType")=1 then
			if Ucase(scDB)="SQL" then
				query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) DESC"
			else
				query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) DESC"
			end if
		else
			if Ucase(scDB)="SQL" then
				query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) DESC"
			else
				query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) DESC"
			end if
		End if
	Case "3":
		If Session("customerType")=1 then
			if Ucase(scDB)="SQL" then
				query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) ASC"
			else
				query1 = " ORDER BY (iif(iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),iif(IsNull(Products.pcProd_BTODefaultPrice),0,Products.pcProd_BTODefaultPrice),Products.pcProd_BTODefaultWPrice)=0,iif(Products.btoBPrice=0,Products.Price,Products.btoBPrice),iif((Products.pcProd_BTODefaultWPrice=0) OR (IsNull(Products.pcProd_BTODefaultWPrice)),Products.pcProd_BTODefaultPrice,Products.pcProd_BTODefaultWPrice))) ASC"
			end if
		else
			if Ucase(scDB)="SQL" then
				query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) ASC"
			else
				query1 = " ORDER BY (iif((Products.pcProd_BTODefaultPrice=0) OR (IsNull(Products.pcProd_BTODefaultPrice)),Products.Price,Products.pcProd_BTODefaultPrice)) ASC"
			end if
		End if
end select

'// Query Products
query="SELECT products.idProduct, products.sku, products.description, products.price, products.listhidden, products.listprice, products.serviceSpec, products.bToBPrice, products.smallImageUrl,products.noprices,products.stock, products.noStock,products.pcprod_HideBTOPrice, POrder,products.FormQuantity,products.pcProd_BackOrder FROM products, categories_products WHERE products.idProduct=categories_products.idProduct AND categories_products.idCategory="& pIdCategory &" AND active=-1 AND configOnly=0 and removed=0 " & query1
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)		
if NOT rs.EOF then
	pcArray_Products = rs.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)
end If	
set rs=nothing

'// Send the headers
Response.ContentType = "text/xml"
Response.AddHeader "Pragma", "public"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
%><?xml version='1.0' encoding='iso-8859-1'?>
<Products>
<%
pCnt=0
do while (pCnt <= pcv_intProductCount) AND (pCnt<cint(SNW_MAX))

	pidProduct=""
	pSku=""
	pDescription=""   
	pPrice=""
	pListHidden=""
	pListPrice=""						   
	pserviceSpec=""
	pBtoBPrice=""   
	pSmallImageUrl="" 
	pnoprices=0
	pStock=""
	pNoStock=""
	pcv_intHideBTOPrice=0
	pFormQuantity=""
	pcv_intBackOrder=""
	pidrelation=""						
	psDesc=""
	pSmallImageUrl=""
	pcv_URL=""

	pidProduct=pcArray_Products(0,pCnt) '// rs("idProduct")
	pSku=pcArray_Products(1,pCnt) '// rs("sku")
	pDescription=pcArray_Products(2,pCnt) '// rs("description")   
	pPrice=pcArray_Products(3,pCnt) '// rs("price")
	pListHidden=pcArray_Products(4,pCnt) '// rs("listhidden")
	pListPrice=pcArray_Products(5,pCnt) '// rs("listprice")						   
	pserviceSpec=pcArray_Products(6,pCnt) '// rs("serviceSpec")
	pBtoBPrice=pcArray_Products(7,pCnt) '// rs("bToBPrice")   
	pSmallImageUrl=pcArray_Products(8,pCnt) '// rs("smallImageUrl")   
	pnoprices=pcArray_Products(9,pCnt) '// rs("noprices")
	pStock=pcArray_Products(10,pCnt) '// rs("stock")
	pNoStock=pcArray_Products(11,pCnt) '// rs("noStock")
	pcv_intHideBTOPrice=pcArray_Products(12,pCnt) '// rs("pcprod_HideBTOPrice")
	pFormQuantity=pcArray_Products(14,pCnt) '// rs("FormQuantity")
	pcv_intBackOrder=pcArray_Products(15,pCnt) '// rs("pcProd_BackOrder")
	pidrelation=pcArray_Products(0,pCnt) '// rs("idProduct")
	
	if isNULL(pnoprices) OR pnoprices="" then
		pnoprices=0
	end if
	if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
		pcv_intHideBTOPrice="0"
	end if
	if pnoprices=2 then
		pcv_intHideBTOPrice=1
	end if						
							
	'// Get sDesc
	query="SELECT sDesc FROM products WHERE idProduct="&pidrelation&";"  
	set rsDescObj=server.CreateObject("ADODB.RecordSet")
	set rsDescObj=conntemp.execute(query)
	psDesc=rsDescObj("sDesc")
	set rsDescObj=nothing
	
	if pSmallImageUrl="" OR isNULL(pSmallImageUrl) then
		pSmallImageUrl="no_image.gif"
	end if	

	pcv_URL=replace((scStoreURL&"/"&scPcFolder&"/pc/"), "//", "/")  
	pcv_URL=replace(pcv_URL,"http:/","http://")

	pDescription=ClearHTMLTags2(pDescription,2)
	pDescription=replace(pDescription,"&quot;","""")
	If 44<len(pDescription) then
		pDescription=trim(left(pDescription,44)) & "..."
	End If
	
	pCnt=pCnt+1
	%>
	<Product>
    	<Description><![CDATA[ <%=pDescription%> ]]></Description>
        <Price><![CDATA[ <%=money(pPrice)%> ]]></Price>
        <SmallImage><![CDATA[ <%=pcv_URL%>catalog/<%=pSmallImageUrl%> ]]></SmallImage>
        <URL><![CDATA[<%=pcv_URL%>viewPrd.asp?idproduct=<%=pidProduct%>]]></URL>
	</Product>
	<%
loop
call closeDb()
%>
</Products>
