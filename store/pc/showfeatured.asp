<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%pcStrPageName="showfeatured.asp"%>
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
<!--#include FILE="prv_incFunctions.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
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
pcStrPageName = "showfeatured.asp"
call openDb()
%>
<!--#include FILE="prv_getSettings.asp"-->
<%

'*******************************
' LOAD SETTINGS (same as home page)
'*******************************

query="SELECT pcHPS_Style,pcHPS_ShowSKU,pcHPS_ShowImg FROM pcHomePageSettings;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query)

if not rs.eof then
	pcStrHPStyle=rs("pcHPS_Style")	
	pcIntHPShowSKU=rs("pcHPS_ShowSKU")
	pcIntHPShowImg=rs("pcHPS_ShowImg")
end if

set rs=nothing

pShowSKU = pcIntHPShowSKU
if pShowSKU = "" or isNull(pShowSKU) then
	pShowSKU = -1 ' If 0, then the SKU is hidden
end if

pShowSmallImg = pcIntHPShowImg
if pShowSmallImg = "" or isNull(pShowSmallImg) then
	pShowSmallImg = -1 ' If 0, then the Image is hidden
end if

' START - Not For Sale visibility
' This variable controls whether NOT FOR SALE items should be shown
' PC v4.1: copy from Best Sellers
	Dim pcIntFeaturedNFS, queryNFS
	query="SELECT pcBSS_NotForSale FROM pcBestSellerSettings;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcIntFeaturedNFS=rs("pcBSS_NotForSale")
	end if
	set rs=nothing

	' Or you can override the value manually by uncommenting one of the lines below
	'	pcIntFeaturedNFS = 0 ' Not for sale items are shown
	'	pcIntFeaturedNFS = -1 ' Not for sale items are not shown
		
	if pcIntFeaturedNFS <> 0 and NotForSaleOverride(session("customerCategory"))=0 then
		queryNFS = " AND formQuantity = 0 "
		else
		queryNFS = " "
	end if
'// END - Not For Sale visibility

'*******************************
' END LOAD PAGE SETTINGS
'*******************************

'*******************************
' GET page style
'*******************************
	' Load the page style: check to see if a querystring
	' or a form is sending the page style.
	Dim pcPageStyle
	pcPageStyle = LCase(Request.QueryString("pageStyle"))
		if pcPageStyle = "" then
			pcPageStyle = LCase(Request.Form("pageStyle"))
		end if
		
		if pcPageStyle = "" then
			pcPageStyle = pcStrHPStyle
		end if

		if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
			pcPageStyle = LCase(bType)
		end if

'*******************************
' GET page size
'*******************************
	Dim pcv_ViewAllVar, newCount
	pcv_ViewAllVar=getUserInput(request("VA"),1)
	if NOT isNumeric(pcv_ViewAllVar) OR pcv_ViewAllVar="" then
		pcv_ViewAllVar=0
	end if
	newCount=0
	
	
	Dim iPageSize
	iPageSize=(scPrdRow*scPrdRowsPerPage)
	if request.queryString("iPageCurrent")="" then
		if request.queryString("page")="" then
			iPageCurrent=1
		else
			iPageCurrent=server.HTMLEncode(request.queryString("page"))
		end if
	else
		iPageCurrent=server.HTMLEncode(request.queryString("iPageCurrent"))
	end if

'*******************************
' GET sorting criteria
'*******************************

 	Dim ProdSort, querySort
	ProdSort="" & request("prodsort")
 	if not validNum(ProdSort) then
		ProdSort="" & PCOrd
 	end if

 	if ProdSort="" then
		ProdSort="0"
 	end if
 	
 	select case ProdSort
		Case "0": querySort = " ORDER BY pcprod_OrdInHome asc"
		Case "1": querySort = " ORDER BY products.description Asc" 	
		Case "2": 
		If Session("customerType")=1 then
		querySort = " ORDER BY products.btoBprice desc, products.price Desc"
		else
		querySort = " ORDER BY products.price Desc"
		End if 	
		Case "3":
		If Session("customerType")=1 then
		querySort = " ORDER BY products.bToBprice Asc, products.price Asc" 	
		else
		querySort = " ORDER BY products.price Asc" 	
		end if 	
 	end select

'*******************************
' GET Featured Items from DB
'*******************************


if session("CustomerType")<>"1" then
	query1= " AND categories.pccats_RetailHide=0"
else
	query1=""
end if

query="SELECT distinct products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.noprices, products.stock, products.noStock, products.pcprod_HideBTOPrice, products.pcprod_OrdInHome, products.formQuantity, products.pcProd_BackOrder FROM products, categories_products, categories WHERE products.active=-1 AND products.showInHome=-1 AND products.configOnly=0 AND products.removed=0 " & queryNFS & " AND categories_products.idProduct=products.idProduct AND categories.idCategory=categories_products.idCategory AND categories.iBTOhide=0 " & query1 & querySort
set rsProducts=Server.CreateObject("ADODB.Recordset")     
rsProducts.CursorLocation=adUseClient
rsProducts.CacheSize=iPageSize
rsProducts.Open query, conntemp
	
if err.number<>0 then
	call LogErrorToDatabase()
	set rsProducts=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
dim iPageCount, count
if NOT rsProducts.eof then	
	
	rsProducts.MoveFirst
	rsProducts.PageSize=iPageSize
	pcv_strPageSize=iPageSize
	iPageCount=rsProducts.PageCount

	rsProducts.AbsolutePage=Cint(iPageCurrent)
	pcArray_Products = rsProducts.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1

else
	set rsProducts = nothing
	call closeDb()
  	response.redirect "msg.asp?message=89"         
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

	'// Loop until the total number of products to show
	if pcv_ViewAllVar=0 then
		newCount = pcv_strPageSize
	else
		newCount = 999999
	end if	
	
	'// Run through the products to count all products, products with options, and BTO products
	do while (pCnt < pcv_intProductCount) and (pCnt < newCount)		
		
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
			<h1><%response.write dictLanguage.Item(Session("language")&"_mainIndex_7")%></h1>
			
			<%if HideSortPro<>"1" then%>
				<div class="pcSortProducts">
				<form action="<%=pcStrPageName%>" method="post" class="pcForms">
				<%=dictLanguage.Item(Session("language")&"_viewCatOrder_5")%> <select name="prodSort" onChange="javascript:if (this.value != '') {this.form.submit();}">
						<option value="0" <%if ProdSort="0" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_1")%></option>
						<option value="1" <%if ProdSort="1" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_2")%></option>
						<option value="2" <%if ProdSort="2" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_3")%></option>
						<option value="3" <%if ProdSort="3" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_viewCatOrder_4")%></option>
								</select>
						<input type="hidden" value="<%=pcPageStyle%>" name="PageStyle">
                        <input type="hidden" value="<%=pcv_ViewAllVar%>" name="VA">
				</form>
				</div>
		 <%end if%>		 

        <% if pcv_ViewAllVar=0 then %>
		<!--#Include File="pcPageNavigation.asp"-->
        <% end if %>
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

				'// Loop until the total number of products to show
				if pcv_ViewAllVar=0 then
					newCount = pcv_strPageSize
				else
					newCount = 999999
				end if
				
				do while (tCnt < pcv_intProductCount) and (count < newCount)
					
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
					<% end if
					
					iRecordsShown=iRecordsShown + 1 %>
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
<%  end if %>


	<!-- Insert page navigation -->
	<% if pcv_ViewAllVar=0 then %>
		<!--#Include File="pcPageNavigation.asp"-->
	<% end if %>	

		<%	  
			call closeDb()
		%>
		</td>
	</tr>
</table>
<!--#include file="atc_viewprd.asp"-->
</div>
<!--#include file="footer.asp"-->