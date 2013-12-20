<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
dim starttime
starttime = Timer
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

Dim query, conntemp, rsProducts, rsDisc, pDiscountPerQuantity, pcStrPageName, pagesize, pcPageStyle
pcStrPageName = "pctest.asp"
pagesize = 50
call openDb()
%>
<!--#include FILE="prv_getSettings.asp"-->
<%

'*******************************
' GET Products from DB
'*******************************
if session("CustomerType")<>"1" then
	query1= " AND ((categories.pccats_RetailHide)=0)"
else
	query1=""
end if

query="SELECT TOP 50 products.idProduct, products.sku, products.description, products.price, products.listHidden, products.listPrice, products.serviceSpec, products.bToBPrice, products.smallImageUrl, products.NoPrices, products.stock, products.noStock, products.pcprod_HideBTOPrice, products.sales, products.formQuantity, products.pcProd_BackOrder FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.description DESC;"

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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="include-metatags.asp"-->
<html>
<head>
<%Session.LCID = 1033
if pcv_PageName<>"" then%>
<title><%=pcv_PageName%></title>
<%end if%>
<%GenerateMetaTags()%>
<%Response.Buffer=True%> 
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcHeaderFooter11.css" />
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<link type="text/css" rel="stylesheet" href="pcBTO.css" />
<!--#include file="inc_header.asp" -->
</head>
<body>
<div id="pcHeaderContainer">

    <div id="pcHeader">
    
        <div id="pcHeaderLeft">
        	<%=scCompanyName%>
        </div>
        
        <div id="pcHeaderCenter">

			<%            
              '// Locate preferred results count and load as default
                Dim pcIntPreferredCountSearch
                pcIntPreferredCountSearch =(scPrdRow*scPrdRowsPerPage)
            %>
            <form action="showsearchresults.asp" name="search" method="get" id="pcHSearchForm">
                <input type="hidden" name="pageStyle" value="<%=bType%>">
                <input type="hidden" name="resultCnt" value="<%=pcIntPreferredCountSearch%>">
                <input type="text" name="keyword" value="" id="pcHSearch">
                <input type="submit" name="submit" value="Go" id="pcHSearchSubmit">
                <div id="pcHSearchMore">
                    <a href="search.asp">More search options</a>
                </div>
            </form>
            <script language="JavaScript">
            <!--
            function pcf_CheckSearchBox() {
                pcv_strTextBox = document.getElementById("pcHSearch").value;
                if (pcv_strTextBox != "") {
                    document.getElementById('small_search').onclick();
                }
            }
            //-->
            </script>
            <%
            Response.Write(pcf_InitializePrototype())
            response.Write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_advSrca_23"), "small_search", 200))
            %>

        </div>
        
        <div id="pcHeaderRight">
			<span><img src="images/pc11-sampleAd2.png" width="134" height="50" alt="Hassle free returns"></span>
            <span><img src="images/pc11-sampleAd1.png" width="134" height="50" alt="Free shipping"></span>
        </div>
            </div>
        </div>
<div id="pcNavContainer45">

	<div id="pcNav45">
    	<a href="viewcategories.asp">Browse The Store</a><a href="showbestsellers.asp">Best Sellers</a><a href="showspecials.asp">Specials</a><a href="shownewarrivals.asp">New Arrivals</a><a href="showrecentlyreviewed.asp">Recently Reviewed</a><a href="search.asp">Adv. Search</a><a href="checkout.asp" style="border-right: none;">Checkout</a>
    </div>
    
</div>

<div id="pcIconBarContainer">

    <div id="pcIconBar">
    
    	<div id="pcIconBarLeft">
            <!--#include file="SmallShoppingCartSmall.asp"-->
        </div>
        
    	<div id="pcIconBarRight">
        	<span class="pcIconBarSeparator"><img src="images/pc11-icon-callus.png" alt="Call Us">Questions? Call us at 800.800.8000</span>
        	<%
			Dim pcMyAccountLink
			if session("idCustomer")="0" or session("idCustomer")="" then
				pcMyAccountLink="Register or Login"
				else
				pcMyAccountLink="Manage My Account"
			end if
			%>
            <a href="custPref.asp"><img src="images/pc11-icon-account.png" alt="<%=pcMyAccountLink%>"></a><a href="custPref.asp" class="pcIconBarSeparator"><%=pcMyAccountLink%></a>
            <a href="checkout.asp"><img src="images/pc11-icon-checkout.png" alt="Checkout"></a><a href="checkout.asp">Checkout</a>
            <a href="contact.asp"><img src="images/pc11-icon-contact.png" alt="Contact Us"></a><a href="contact.asp">Contact Us</a>
            <a href="default.asp"><img src="images/pc11-icon-home.png" alt="Home"></a><a href="default.asp">Home</a>
        </div>
	</div>
  
  </div>
  
<div id="pcMainArea">
  
	<div id="pcMainArea-LEFT">
		  
        <div id="pcMainArea-BROWSE">
        <h3>Browse by Category</h3>
		<!--#include file="inc_catsmenu.asp"-->
      </div>
      
        <div id="pcMainArea-LINKS">
        <h3>Useful Links</h3>
        <ul>
					<li><a href="default.asp">Store Home</a></li>
					<li><a href="viewcategories.asp">Browse Catalog</a></li>
					<% ' Show Browse by Brand link if there are brands in the store
						sdquery="SELECT BrandName FROM Brands"
						set rsSideCatObj=Server.CreateObject("ADODB.Recordset")
						set rsSideCatObj=conlayout.execute(sdquery)
						if not rsSideCatObj.eof then
					%>
							<li><a href="viewbrands.asp">Shop by Brands</a></li><%
						' End show Browse by Brand
						end if
						set rsSideCatObj=nothing
					%>
					<li><a href="search.asp">Advanced Search</a></li>
					<li><a href="viewcart.asp">View Cart</a></li>
					<% 
						'// START CONTENT PAGES
						'// Select pages compatible with customer type
						if session("customerCategory")<>0 then ' The customer belongs to a customer category
							' Load pages accessible by ALL, plus those accessible by the customer pricing category that the customer belongs to
							queryCustType = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
						else
							if session("customerType")=0 then ' Retail customer or customer not logged in
								queryCustType = " AND pcCont_CustomerType = 'ALL'" ' Load pages accessible by ALL
							else
								queryCustType = " AND pcCont_CustomerType = 'W'" ' Load pages accessible by Wholesale customers only
							end if
						end if
						
						'// Load pages from the database: active, not excluded from navigation, and compatible with customer type						
						sdquery="SELECT pcCont_IDPage, pcCont_PageName FROM pcContents WHERE pcCont_InActive=0 AND pcCont_MenuExclude<>1 " & queryCustType & " ORDER BY pcCont_Order ASC, pcCont_PageName ASC;"
						set rsSideCatObj=Server.CreateObject("ADODB.RecordSet")					
						set rsSideCatObj=conlayout.execute(sdquery)
						do while not rsSideCatObj.eof
							pcIntContentPageID=rsSideCatObj("pcCont_IDPage")
							pcvContentPageName=rsSideCatObj("pcCont_PageName")
							'// Call SEO Routine
							pcGenerateSeoLinks
							'//
						%>
							<li><a href="<%=pcStrCntPageLink%>"><%=pcvContentPageName%></a></li>
						<%
							rsSideCatObj.MoveNext
						loop
						set rsSideCatObj=nothing
						'// END CONTENT PAGES
					%>
					<li><a href="contact.asp">Contact Us</a></li>
         </ul>
      </div>
            
    </div>
    
  	<div id="pcMainArea-RIGHT">
    
        <!--#include file="CategorySearchFields.asp"-->
        
          <div id="pcMainArea-PRICE">
            <h3>Browse by Price</h3>
                <ul>
                    <li><a href="showsearchresults.asp?priceFrom=0&priceUntil=20">Below $20</a></li>
                    <li><a href="showsearchresults.asp?priceFrom=20&priceUntil=50">From $20 to $50</a></li>
                    <li><a href="showsearchresults.asp?priceFrom=50&priceUntil=100">From $50 to $100</a></li>
                    <li><a href="showsearchresults.asp?priceFrom=100&priceUntil=9999999">Over $100</a></li>
                </ul>
          </div>

		  <!--#include file="SmallRecentProducts.asp"-->
        
    </div>
    
    <div id="pcMainArea-CENTER">
    	
    	
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
			<h1>ProductCart Storefront Test Page</h1>
            <h2>Loading 50 products...</h2>
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
    </div>
    
    <div id="pcMainArea-SPACER">
    </div>
  
</div>
  
<div id="pcFooterContainer">
  
  <div id="pcFooter">
  	<div id="pcFooterLeft">
	<%
	if trim(scCompanyName)<>"" then
		response.write scCompanyName & " - " & scCompanyAddress & ", " & scCompanyCity & ", " & scCompanyState & " " & scCompanyZIP & ", " & scCompanyCountry
	end if
	%>	
    </div>
  	<div id="pcFooterRight">
    	PCI Compliance: this store uses PA-DSS validated <a href="http://www.earlyimpact.com/productcart/" target="_blank">shopping cart software</a>
	</div>    
  </div>
</div>

<!--#include file="inc_footer.asp" -->

<%
dim endtime,timetaken
endtime = Timer
timetaken = FormatNumber(endtime - starttime,4)
%>
<div style="width:100%; margin: 10px; padding: 10px; text-align: center; border-top: 1px dashed #CCC; background-color:#FFF;">This page loaded in <%=timetaken%> seconds</div>.
</body>
</html>