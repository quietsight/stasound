<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/pcProductOptionsCode.asp"--> 
<!--#include file="../includes/CashbackConstants.asp"--> 
<!--#INCLUDE file="HomeCode.asp"-->
<% 'PRV41 start %>
<!--#include FILE="prv_incFunctions.asp"-->
<% 'PRV41 end %>
<% 
'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "home.asp"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'-------------------------------
' declare local variables
'-------------------------------
dim pcStrHPStyle, pcStrHPDesc, pcIntHPFirst, pcIntHPShowSKU, pcIntHPShowImg, pcIntHPFeaturedCount, pcIntHPFeaturedOrder
dim pcIntHPSpcCount, pcIntHPSpcOrder, pcIntHPNewCount, pcIntHPSNewOrder, pcIntHPBestCount, pcIntHPBestOrder
Dim query, conntemp, rsProducts, rsDisc, pDiscountPerQuantity, pTotalCount
Dim pcv_intBackOrder, pStock, pNoStock, pFormQuantity, pserviceSpec
call opendb()
%>
<!--#include FILE="prv_getSettings.asp"-->
<%

'*******************************
' LOAD HOMEPAGE SETTINGS
'*******************************
' Refer to "pcadmin/manageHomePage.asp" to see features added to this page

query=  "SELECT pcHPS_FeaturedCount,pcHPS_Style,pcHPS_PageDesc,pcHPS_First,pcHPS_ShowSKU,pcHPS_ShowImg," &_
        "pcHPS_SpcCount,pcHPS_SpcOrder,pcHPS_NewCount,pcHPS_NewOrder,pcHPS_BestCount,pcHPS_BestOrder FROM pcHomePageSettings;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcIntHPFeaturedCount=rs("pcHPS_FeaturedCount")
	pcStrHPStyle=rs("pcHPS_Style")	
	pcStrHPDesc=replace(rs("pcHPS_PageDesc"),"''","'")
	pcIntHPFirst=rs("pcHPS_First")
	pcIntHPShowSKU=rs("pcHPS_ShowSKU")
	pcIntHPShowImg=rs("pcHPS_ShowImg")
	pcIntHPSpcCount=rs("pcHPS_SpcCount")
	pcIntHPSpcOrder=rs("pcHPS_SpcOrder")
	pcIntHPNewCount=rs("pcHPS_NewCount")
	pcIntHPNewOrder=rs("pcHPS_NewOrder")
	pcIntHPBestCount=rs("pcHPS_BestCount")
	pcIntHPBestOrder=rs("pcHPS_BestOrder")
end if

if pcIntHPFeaturedCount = "" or not validNum(pcIntHPFeaturedCount) or pcIntHPFeaturedCount < 0 then
	pcIntHPFeaturedCount = 3
end if

' // Note: 0 is an acceptable value and it indicates that Specials should not be shown
if pcIntHPSpcCount = "" or not validNum(pcIntHPSpcCount) or pcIntHPSpcCount < 0 then
	pcIntHPSpcCount = 4
end if

' // Note: 0 is an acceptable value and it indicates that New Arrivals should not be shown
if pcIntHPNewCount = "" or not validNum(pcIntHPNewCount) or pcIntHPNewCount < 0 then
	pcIntHPNewCount = 4
end if

' // Note: 0 is an acceptable value and it indicates that Best Sellers should not be shown
if pcIntHPBestCount = "" or not validNum(pcIntHPBestCount) or pcIntHPBestCount < 0 then
	pcIntHPBestCount = 4
end if

set rs=nothing
call closedb()

pShowSKU = pcIntHPShowSKU
if pShowSKU = "" or isNull(pShowSKU) then
	pShowSKU = -1 ' If 0, then the SKU is hidden
end if

pShowSmallImg = pcIntHPShowImg
if pShowSmallImg = "" or isNull(pShowSmallImg) then
	pShowSmallImg = -1 ' If 0, then the Image is hidden
end if

'*******************************
' END LOAD HOMEPAGE SETTINGS
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

if pcPageStyle = "" then
	pcPageStyle = LCase(bType)
end if

if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
	pcPageStyle = LCase(bType)
end if

'*******************************
' GET Featured Products from DB
'*******************************

call openDb()
if session("CustomerType")<>"1" then
	query1= " AND ((categories.pccats_RetailHide)=0)"
else
	query1=""
end if

'// START v4.1 - Not For Sale override
	if NotForSaleOverride(session("customerCategory"))=1 then
		queryNFSO=""
	else
		queryNFSO="AND formQuantity = 0 "
	end if
'// END v4.1

query="SELECT distinct products.idProduct,products.sku,products.description,products.price,products.listHidden,products.listPrice,products.serviceSpec,products.bToBPrice,products.smallImageUrl,products.noprices,products.stock,products.noStock, products.pcprod_HideBTOPrice,products.formQuantity,pcprod_OrdInHome,products.pcProd_BackOrder FROM products,categories_products,categories WHERE products.active=-1 AND products.showInHome=-1 AND products.configOnly=0 AND products.removed=0 " & queryNFSO & "AND categories_products.idProduct=products.idProduct AND categories.idCategory=categories_products.idCategory AND categories.iBTOhide=0 " & query1 & " order by pcprod_OrdInHome asc"

set rsProducts=server.CreateObject("ADODB.Recordset")
set rsProducts=conntemp.execute(query)
if Err.number <> 0 then
	set rsProducts=nothing
	call closeDb()  
	response.redirect "techErr.asp?error="&Server.UrlEncode("Error db in " & pcStrPageName & " - Error: "&Err.Description)
end If
if NOT rsProducts.eof then
	pcArray_Products = rsProducts.getRows()
	pcv_intProductCount = UBound(pcArray_Products,2)+1
end if
set rsProducts = nothing


'*******************************
' Set Total Count
'*******************************
pTotalCount=pcv_intProductCount

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
	
	'Loop until the total number of products to show
	if pcIntHPFirst<>0 then
		pCnt=pCnt+1
		pAllCnt=pAllCnt+1
		pcTempHPFeaturedCount=pcIntHPFeaturedCount+1
	else
		pcTempHPFeaturedCount=pcIntHPFeaturedCount
	end if

	'// Run through the products to count all products, products with options, and BTO products
	do while (pCnt < pcv_intProductCount) and (pCnt < pcTempHPFeaturedCount)		
		
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

<!--#include file="../includes/javascripts/pcValidateFormViewPrd.asp"-->
<!--#include file="inc_PrdCatTip.asp"-->
<div id="pcMain">
<table class="pcMainTable">
	<%
	'If there are no featured products and no page description, hide the table row
	if pcIntHPFeaturedCount > 0 or pcStrHPDesc <> "" then
	%>
	<tr>
		<td> 
			<%
			' If there are featured products, show that message, otherwise hide it
			if pcIntHPFeaturedCount > 0 then
			%>
				<h1><%response.write dictLanguage.Item(Session("language")&"_mainIndex_11")%></h1>
			<%
			end if
			
			' Show Home Page description, if any
			if pcStrHPDesc <> "" then %>
				<div class="pcPageDesc"><%=pcStrHPDesc%></div>
			<% 	
			end if
			%>
		</td>
	</tr>
	<%
	end if
	%>

<%
'*****************************************************************************************************
' 1) PRODUCT OF THE MONTH
'*****************************************************************************************************
    pcs_ProductOfTheMonth
'*****************************************************************************************************
' END PRODUCT OF THE MONTH
'*****************************************************************************************************
%>

<%
'*****************************************************************************************************
' 2) FEATURED PRODUCTS
'*****************************************************************************************************
    pcs_FeaturedProducts
'*****************************************************************************************************
' END FEATURED PRODUCTS
'*****************************************************************************************************
%>

	
<%
'*****************************************************************************************************
' 3) Best sellers, new arrivals, specials
'*****************************************************************************************************
    pcs_ShowProducts
'*****************************************************************************************************
' END Best sellers, new arrivals, specials 
'*****************************************************************************************************
%>
</table>
<!--#include file="atc_viewprd.asp"-->
</div>
<%	  
set rsProducts=Nothing
set iPageCurrent=Nothing
call closeDb()
%>
<!--#include file="orderCompleteTracking.asp"-->
<!--#include file="inc-Cashback.asp"-->
<!--#include file="footer.asp"-->