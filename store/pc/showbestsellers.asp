<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include FILE="../includes/pcProductOptionsCode.asp"-->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include FILE="prv_incFunctions.asp"-->
<% 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim query, conntemp, rsProducts, rsDisc, pDiscountPerQuantity, pcStrPageName
pcStrPageName = "showbestsellers.asp"
call openDb()
%>
<!--#include FILE="prv_getSettings.asp"-->
<%

'*******************************
' GET PAGE SETTINGS FROM DB
'*******************************
Dim pcIntBestSellCount, pagesize, pcStrBestSellDesc, pcIntBestSellNFS, queryNFS, pcIntBestSellInStock, queryInStock, pcIntBestSellSales, pShowSKU, pShowSmallImg, pcPageStyle

pcIntBestSellSales=0
pcIntBestSellCount=0

query="SELECT pcBSS_BestSellCount,pcBSS_Style,pcBSS_PageDesc,pcBSS_NSold,pcBSS_NotForSale,pcBSS_OutOfStock,pcBSS_SKU,pcBSS_ShowImg FROM pcBestSellerSettings;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)
if not rs.eof then
	pcIntBestSellCount=rs("pcBSS_BestSellCount")
	pcPageStyle=rs("pcBSS_Style")
	pcStrBestSellDesc=rs("pcBSS_PageDesc")
	pcIntBestSellSales=rs("pcBSS_NSold")
	pcIntBestSellNFS=rs("pcBSS_NotForSale")
	pcIntBestSellInStock=rs("pcBSS_OutOfStock")
	pShowSKU=rs("pcBSS_SKU")
	pShowSmallImg=rs("pcBSS_ShowImg")
end if
set rs=nothing

if isNULL(pcIntBestSellCount) or (pcIntBestSellCount="0") then
	pcIntBestSellCount= 14
end if
pagesize = pcIntBestSellCount

if isNULL(pcIntBestSellSales) or (pcIntBestSellSales="0") then
	pcIntBestSellSales=2
end if

if pcIntBestSellNFS<> 0 and NotForSaleOverride(session("customerCategory"))=0 then
	queryNFS = " AND ((products.formQuantity)=0)"
else
	queryNFS = " "
end if

if isNULL(pShowSKU) OR (pShowSKU="") then
	pShowSKU=0
end if

if isNULL(pShowSmallImg) OR (pShowSmallImg="") then
	pShowSmallImg=0
end if

if pcPageStyle = "" then
	pcPageStyle = LCase(Request.QueryString("pageStyle"))
	if pcPageStyle = "" then
		pcPageStyle = LCase(Request.Form("pageStyle"))
	end if
end if

if pcPageStyle = "" then
	pcPageStyle = LCase(bType)
end if
		
if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
	pcPageStyle = LCase(bType)
end if

'*******************************
' GET Best Sellers from DB
'*******************************
if session("CustomerType")<>"1" then
	query1= " AND ((categories.pccats_RetailHide)=0)"
else
	query1=""
end if

if pcIntBestSellInStock<> 0 and scOutOfStockPurchase<>0 then
	query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.NoPrices, products.stock, products.noStock, products.pcprod_HideBTOPrice, products.sales, products.formQuantity, products.pcProd_BackOrder FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.stock)>0) AND ((products.sales)>="&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") OR (((products.noStock)=-1) AND ((products.sales)>"&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.sales DESC;"
else
	query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.NoPrices, products.stock, products.noStock, products.pcprod_HideBTOPrice, products.sales, products.formQuantity, products.pcProd_BackOrder FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.sales)>="&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.sales DESC;"
end if

set rsProducts=server.CreateObject("ADODB.Recordset")
set rsProducts=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsProducts=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if NOT rsProducts.eof then
	pcArray_Products = rsProducts.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1
else
	set rsProducts = nothing
	call closeDb()
	response.redirect "msg.asp?message=94"
end if
set rsProducts = nothing

'*******************************
' Start: Set variables for "M" display
'*******************************
if pcPageStyle = "m" then
	'Check if customers are allowed to order products
	dim iShow
	iShow=0
	If scOrderlevel=0 then
		iShow=1
	end if
	If scOrderlevel=1 AND session("customerType")="1" then
		iShow=1
	End if
	
	Dim pCnt, pAddtoCart, pAllCnt
	'reset count variables
	pCnt=Cint(0)
	pAllCnt=Cint(0)

	'// Run through the products to count all products, products with options, and BTO products
	do while (pCnt < pcv_intProductCount) and (pCnt < pagesize)		
		
		pidrelation=pcArray_Products(0,pCnt) '// rsCount("idProduct")
		pserviceSpec=pcArray_Products(6,pCnt) '// rsCount("serviceSpec")	
		pStock=pcArray_Products(10,pCnt) '// rsCount("stock")
		pNoStock=pcArray_Products(11,pCnt) '// rsCount("noStock")
		pcv_intBackOrder=pcArray_Products(15,pCnt) '// rs("pcProd_BackOrder")
		
		pCnt=pCnt+1
		
		' Check which items will have multi qty enabled,
		pcv_SkipCheckMinQty=-1 
		If pcf_AddToCart(pidrelation)=False Then
			pAllCnt=pAllCnt+1
		End If	
		
	loop
	
	pcv_SkipCheckMinQty=0
		
	' If all items on the page are either BTO or have options,
	' do not show the quantity column or the Add to Cart button.						
	if cint(pAllCnt) <> cint(pCnt) then 
		pAddtoCart = 1
	end if
end if	
'*******************************
' End: Set variables for "M" display
'*******************************


'*******************************
' Build the page
'*******************************
%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->
<!--#include file="inc_PrdCatTip.asp"-->
<!--#include file="inc_AddThis.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td> 
			<%
            '// PC v4.5 AddThis integration
            if scAddThisDisplay=1 then pcs_AddThis
            %>
			<h1><%response.write dictLanguage.Item(Session("language")&"_viewBestSellers_2")%></h1>
			<% ' Show New Best Sellers description, if any
					if pcStrBestSellDesc <> "" then %>
					<div class="pcPageDesc"><%=pcStrBestSellDesc%></div>
			<% 	end if %>
		</td>
	</tr>
	<tr>
		<td>
			<% if pcPageStyle = "m" then %>
				<form action="instPrd.asp" method="post" name="m" id="m" class="pcForms">
			<% end if %>
			<table class="pcShowProducts">
			<%
					'*******************************
					' Add table headers for display
					' styles L and M
					'*******************************
			%>
					<% if pcPageStyle = "l" then	%>
							<tr class="pcShowProductsLheader">
							<% if pShowSmallImg <> 0 then %>
								<td>&nbsp;</td>
							<% end if %>
								<td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_9") %></td>
							<% if pShowSku <> 0 then %>
								<td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_8") %></td>
							<% end if %>
								<td><% response.write dictLanguage.Item(Session("language")&"_viewCat_P_10") %></td>
							</tr>
					<% elseif pcPageStyle = "m" then %>
							<tr class="pcShowProductsMheader">
								<td colspan="<%if iShow=1 then%>5<%else%>4<%end if%>">
									<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_12") %>
								</td>
							</tr>
							<tr>
							<tr class="pcShowProductsMheader">
									<% if iShow=1 then %> 
										<% if pAddtoCart = 1 then %>
											<td width="8%">
												<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_7") %>
											</td>
										<% end if %>
									<% end if %>
									<% if pShowSmallImg <> 0 then %>
									<td width="8%">&nbsp;</td>
									<% end if %>
									<% if pShowSku <> 0 then %>
									<td width="11%">
										<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_8") %>
									</td>
									<% end if %>
									<td width="47%">
										<% response.write dictLanguage.Item(Session("language")&"_viewCat_P_9") %>
									</td>
									<td width="16%" align="center">
										<% If session("customerType")="1" then
												response.write dictLanguage.Item(Session("language")&"_viewCat_P_11")
											 else
												response.write dictLanguage.Item(Session("language")&"_viewCat_P_10")
										end if %>
									</td>
								</tr>
			 			<% else %>
						<tr>
					<% end if %>
				<%
				'*******************************
				' End table headers
				'*******************************
			
				'*******************************
				' Load product information
				' Loop through the products
				'*******************************
				
				'Set the product count to zero
				count=0
					
				if pcPageStyle = "m" then
					pCnt=Cint(0)
					pSQty=0
					pAllCnt=Cint(0)
				end if

				tCnt=Cint(0)
					
				do while (tCnt < pcv_intProductCount) and (count < pagesize)

					pidProduct=pcArray_Products(0,tCnt) '// rs("idProduct")
					pSku=pcArray_Products(1,tCnt) '// rs("sku")
					pDescription=pcArray_Products(2,tCnt) '// rs("description")   
					pPrice=pcArray_Products(3,tCnt) '// rs("price")
					pListHidden=pcArray_Products(4,tCnt) '// rs("listhidden")
					pListPrice=pcArray_Products(5,tCnt) '// rs("listprice")						   
					pserviceSpec=pcArray_Products(6,tCnt) '// rs("serviceSpec")
					pBtoBPrice=pcArray_Products(7,tCnt) '// rs("bToBPrice")   
					pSmallImageUrl=pcArray_Products(8,tCnt) '// rs("smallImageUrl")   
					pnoprices=pcArray_Products(9,tCnt) '// rs("noprices")
					if isNULL(pnoprices) OR pnoprices="" then
						pnoprices=0
					end if
					pStock=pcArray_Products(10,tCnt) '// rs("stock")
					pNoStock=pcArray_Products(11,tCnt) '// rs("noStock")
					pcv_intHideBTOPrice=pcArray_Products(12,tCnt) '// rs("pcprod_HideBTOPrice")
					if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
						pcv_intHideBTOPrice="0"
					end if
					if pnoprices=2 then
						pcv_intHideBTOPrice=1
					end if
					pFormQuantity=pcArray_Products(14,tCnt) '// rs("FormQuantity")
					pcv_intBackOrder=pcArray_Products(15,tCnt) '// rs("pcProd_BackOrder")
					pidrelation=pcArray_Products(0,tCnt) '// rs("idProduct")						
											
					'// Get sDesc
					query="SELECT sDesc FROM products WHERE idProduct="&pidrelation&";"
					set rsDescObj=server.CreateObject("ADODB.RecordSet")
					set rsDescObj=conntemp.execute(query)
					psDesc=rsDescObj("sDesc")
					set rsDescObj=nothing
					
					if pcPageStyle = "m" then
						pCnt=pCnt+1
					end if
					tCnt=tCnt+1
					%>
					<!--#include file="pcGetPrdPrices.asp"-->
					<%
   				
						'*******************************
						' Show product information
						' depending on the page style
						'*******************************
					
						' FIRST STYLE - Show products horizontally, with images
						if pcPageStyle = "h" then	%>
							<td> 
								<!--#include file="pcShowProductH.asp" -->
							</td>
							<% i=i + 1
							If i > (scPrdRow-1) then 
								response.write "</TR><TR>"
								i=0
							End If
						end if
					
					' SECOND STYLE - Show products vertically, with images 
					if pcPageStyle = "p" then	%>
						<td> 
							<!--#include file="pcShowProductP.asp" -->
						</td>
					</tr>
					<% end if
					
					' THIRD STYLE - Show a list of products, with a small image 
					if pcPageStyle = "l" then	%>
							<!--#include file="pcShowProductL.asp" -->
					<% end if
					
					' FOURTH STYLE - Show a list of products, with multiple add to cart 
					if pcPageStyle = "m" then	%>
							<!--#include file="pcShowProductM.asp" -->
					<% end if %>
					<%	
					count=count + 1

				loop
			%>
	</table>      
	
	<%' If page style is M, show the Add to Cart button when
	' products can be added to the cart from this page.	
		if pcPageStyle = "m" then %>
			<input type="hidden" name="pCnt" value="<%=pCnt%>">
			<% if iShow=1 and clng(pSQty)<>0 then %>
				<div style="padding: 10px 0 10px 0;">
					<input name="submit" type="image" src="<%=rslayout("addtocart")%>" id="submit">
				</div>
			<% end if %>
			</form>
	<% 
	end if
	call closeDb() 
	%>
		</td>
	</tr>
</table>
<!--#include file="atc_viewprd.asp"-->
</div>
<!--#include file="footer.asp"-->