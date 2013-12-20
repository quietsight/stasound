<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
response.Buffer=true 
pageTitle="Add Product to an Existing Order"
%>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/shipFromSettings.asp"--> 
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include FILE="../includes/pcProductOptionsCode.asp"--> 
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../pc/inc_AddThis.asp"-->
<% 
Dim conntemp, tIndex, tUpdPrd, pIdCategory, strBreadCrumb, pIdProduct, query, rs, dblpcCC_Price
Dim pcv_strViewPrdStyle, pcv_strFormAction, pcv_intValidationFile, pcv_blnBTOisConfig, iRewardPoints, pDescription, pMainProductName
Dim pSku, pconfigOnly, pserviceSpec, pPrice, pBtoBPrice, pDetails, pListPrice, plistHidden, pimageUrl, pLgimageURL
Dim pArequired, pBrequired, pStock, pWeight, pEmailText, pFormQuantity, pnoshipping, pcustom1, pcontent1
Dim pcustom2, pcontent2, pcustom3, pcontent3, pxfield1, px1req, pxfield2, px2req, pxfield3, px3req, pnoprices
Dim pIDBrand, psDesc, pNoStock, pnoshippingtext, intIdProduct, intWeight, optionA, optionB
Dim pcv_intHideBTOPrice, pcv_intQtyValidate, pcv_lngMinimumQty, intpHideDefConfig, pnotax, BrandName
Dim FirstCnt, strDescription, intReward, pcv_BTORP, strConfigProductCategory, dblPrice, dblWPrice, intIdCategory
Dim VardiscGo, dblQuantityFrom, dblQuantityUntil, dblPercentage, dblDiscountPerWUnit, dblDiscountPerUnit
Dim intIdOptOptGrp, intIdOption, strOptionDescrip, OptInActive, optPrice, tempIdOptA, tempIdOptB
Dim xrequired, xfieldCnt, reqstring, TextArea, widthoffield, rowlength
Dim scCS, cs_showprod, cs_showcart, cs_showimage, crossSellText
Dim pcv_strOptionGroupDesc, pcv_intOptionGroupCount, pcv_strOptionGroupCount, pcv_strOptionGroupID, pcv_strOptionRequired
Dim xOptionsCnt, pcv_strNumberValidations, pcv_strFuntionCall, pcv_strReqOptString, xOtionrequired, pcv_strCSDiscounts , pcv_strPrdDiscounts 
Dim pcv_strProdImage_Url, pcv_strProdImage_LargeUrl, pcv_intProdImage_Columns, pcv_strShowImage_LargeUrl, pcv_strShowImage_Url, pcv_strCurrentUrl
Dim pcv_strAdditionalImages, cCounter, pcv_strWishListLink, BTOCharges, pcv_strCSString, pcv_strReqCSString, cs_RequiredIds, xCSCnt

pIdProduct=request.QueryString("idProduct")
pIdOrder=request.QueryString("ido")

if trim(pIdProduct)="" or IsNumeric(pIdProduct)=false then
   response.redirect "msg.asp?message=85"
end if


'Change paths of images
pcv_tmpNewPath="../pc/"

'--> open database connection
call opendb()

	query="SELECT customers.customerType,customers.idCustomerCategory FROM customers INNER JOIN orders ON customers.idcustomer=orders.idcustomer WHERE idorder=" & pIdOrder & ";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if not rs.eof then
		session("customerType")=rs("customerType")
		idcustomerCategory=rs("idcustomerCategory")
		if IsNull(idcustomerCategory) or idcustomerCategory="" then
			idcustomerCategory=0
		end if
	end if
	set rs=nothing

	query="SELECT idcustomerCategory, pcCC_Name, pcCC_Description, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory="&idcustomerCategory&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if NOT rs.eof then
		session("customerCategory")=rs("idcustomerCategory")
		strpcCC_Name=rs("pcCC_Name")
		session("customerCategoryDesc")=strpcCC_Name
		strpcCC_Description=rs("pcCC_Description")
		session("customerCategoryType")=rs("pcCC_CategoryType")
		if session("customerCategoryType")="ATB" then
			session("ATBCustomer")=1
			session("ATBPercentage")=rs("pcCC_ATB_Percentage")
			intpcCC_ATB_Off=rs("pcCC_ATB_Off")
			if intpcCC_ATB_Off="Retail" then
				session("ATBPercentOff")=0
			else
				session("ATBPercentOff")=1
			end if
		else
			session("ATBCustomer")=0
			session("ATBPercentage")=0
			session("ATBPercentOff")=0
		end if
	end if
	set rs=nothing
%>
<!--#INCLUDE FILE="../pc/viewPrdCode.asp"-->
<%
query="SELECT addtocart,customize From layout Where layout.ID=2;"
Set RSlayout = conntemp.Execute(query)
pAddToCartBtn=rslayout("addtocart")
pCustomizeBtn=rslayout("customize")
Set RSlayout = nothing

query="SELECT zoom, discount From icons WHERE id=1;"
Set rsIconObj = conntemp.Execute(query)
pZoomBtn=rsIconObj("zoom")
pDiscountBtn=rsIconObj("discount")
Set rsIconObj = nothing

' --> check for discount per quantity
query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" &pidProduct

set rs=conntemp.execute(query)

if err.number <> 0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

dim pDiscountPerQuantity
if not rs.eof then
 pDiscountPerQuantity=-1
else
 pDiscountPerQuantity=0
end if

' --> gets product details from db
query="SELECT iRewardPoints, description, sku, configOnly, serviceSpec, price, btobprice, details, listprice, listHidden, imageurl, largeImageURL, stock, weight, emailText, formQuantity, noshipping, custom1, content1, custom2, content2, custom3, content3, xfield1, x1req, xfield2, x2req, xfield3, x3req,noprices,IDBrand,noshippingtext,nostock,pcprod_HideDefConfig FROM products WHERE idProduct=" &pidProduct& " AND active=-1"

set rs=conntemp.execute(query)
if err.number <> 0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rs.eof then
	set rs = nothing
	call closeDb() 
  	response.redirect "msg.asp?message=34"
end if

' --> set product variables <---
iRewardPoints=rs("iRewardPoints")
pDescription=rs("description")
pSku= rs("sku")
pconfigOnly=rs("configOnly")
pserviceSpec=rs("serviceSpec")
pPrice=rs("price")
pBtoBPrice=rs("bToBPrice")
pDetails=rs("details")
pListPrice=rs("listPrice")
plistHidden=rs("listHidden")
pimageUrl=rs("imageUrl")
pLgimageURL=rs("largeImageURL")
pStock=rs("stock")
pWeight=rs("weight")
pEmailText=rs("emailText")
pFormQuantity=rs("formQuantity")
pnoshipping=rs("noshipping")
pcustom1=rs("custom1")
pcontent1=rs("content1")
pcustom2=rs("custom2")
pcontent2=rs("content2")
pcustom3=rs("custom3")
pcontent3=rs("content3")
pxfield1=rs("xfield1")
px1req=rs("x1req")
pxfield2=rs("xfield2")
px2req=rs("x2req")
pxfield3=rs("xfield3")
px3req=rs("x3req")
pnoprices=rs("noprices")
pIDBrand=rs("IDBrand")
pnoshippingtext=rs("noshippingtext")
pnostock=rs("nostock")
intHideDefConfig=rs("pcprod_HideDefConfig")

if intHideDefConfig<>"" then
else
	intHideDefConfig="0"
end if	

if pnoprices=1 then
	pPrice=0
	pBtoBPrice=0
	pListPrice=0
end if
if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
	query="select BrandName from Brands where IDBrand=" & pIDBrand
	set rstemp4=connTemp.execute(query)
	if not rstemp4.eof then
		BrandName=rstemp4("BrandName")
	end if
end if
%>

<!--#include file="adminheader.asp"-->
<!-- 58eed21a4b2adf0316e95c5c4ee68f13 -->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<!--#include file="../includes/javascripts/pcValidateViewPrd.asp"-->

<!-- Start Form -->
<% 
'/////////////////////////////////////////////////////////////////////////////////////////////////////
' GENERATE FORM																						//
' > BTO / BTO Configured / Standard Product															//
' > Each uses a different form action and JavaScript validation function                            //
'/////////////////////////////////////////////////////////////////////////////////////////////////////

'********************************************************************
' VALIDATION FILE
' pcv_intValidationFile = 1 // BTO
' pcv_intValidationFile = 2 // Standard
'
' FORM ACTION
' pcv_strFormAction = "instConfiguredPrd.asp" // BTO configured
' pcv_strFormAction = "instPrd.asp" // BTO NON configured and Standard
'********************************************************************
pcv_blnBTOisConfig = pcf_BTOisConfig '// returns true or false for Configured BTO

If pserviceSpec = "False" Then
	pserviceSpec = 0
End If

If pserviceSpec<>0 Then '// If its BTO Then
	if pcv_blnBTOisConfig then '// if its configured then
		pcv_strFormAction = "bto_instConfiguredPrd.asp"
		pcv_intValidationFile = 1
	else '// Its not configured
		pcv_strFormAction = "instPrdToOrd.asp"
		pcv_intValidationFile = 1
	end if
else '// Its standard
	pcv_strFormAction = "instPrdToOrd.asp"
	pcv_intValidationFile = 2
end if
%>
<div id="pcMain" style="margin: 10px;">
<form method="post" action="<%=pcv_strFormAction%>" name="additem" onsubmit="return checkproqty(document.additem.quantity);" class="pcForms">
<input name="idorder" type="hidden" value="<%=pIdOrder%>">
<!--#include file="../includes/javascripts/pcWindowsViewPrd.asp"-->
<table class="pcMainTable" width="100%">
	<tr valign="top"> 
		<td colspan="3">
		<%
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START:  Show product name 
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		pcs_ProductName
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END:  Show product name 
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		%>
 	</td>
</tr>

<tr> 
	<td>
	<%
'*****************************************************************************************************
' 2) GENERAL INFORMATION
'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show SKU
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ShowSKU	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show SKU
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Weight (If admin turned on)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_DisplayWeight
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Weight
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Brand (If assigned)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ShowBrand
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Brand
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Units in Stock (if on, show the stock level here)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_UnitsStock
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Units in Stock
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END GENERAL INFORMATION
'*****************************************************************************************************
	%>
	
<br />
	
	<%	

'*****************************************************************************************************
' 5) DESCRIPTION
'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Product Description
	'   >  If there is a short description, show it and link to the long description below.
	'   >  Otherwise show the long description
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_ProductDescription
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Product Description
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END DESCRIPTION
'*****************************************************************************************************
	%>
	
<br />
	
	<%
'*****************************************************************************************************
' 6) DEFAULT CONFIGURATION (BTO)
'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show BTO Defualt Configuration
	'   >  If this is a BTO product, then gather information about default configuration
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_BTOConfiguration
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show BTO Defualt Configuration
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END DEFAULT CONFIGURATION (BTO)
'*****************************************************************************************************
	%>
	
	<%
'*****************************************************************************************************
' 8) CUSTOM SEARCH FIELDS
'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Custom Search Fields
	'   >  Check to see if the product has been assigned Custom Search Fields. If so, display the values
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_CustomSearchFields
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Custom Search Fields
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END CUSTOM SEARCH FIELDS
'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Reward Points
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_RewardPoints
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Reward Points
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show product prices
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~				
	pcs_ProductPrices
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show product prices
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Free Shipping Text
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	pcs_FreeShippingText
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Free Shipping Text
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  BTO ADDON S 0r E
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_BTOADDON
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  BTO ADDON S 0r E
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Out of Stock Message
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_OutStockMessage
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Out of Stock Message
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	%>
	
	<%'Start SDBA%>
	<!-- Start Back-Order Message -->
	<%pcs_DisplayBOMsg%>
	<!-- End Back-Order Message -->
	<%'End SDBA%>
	</td>
	<td>&nbsp;</td>
	<td rowspan="2" valign="top">
	<% 

'*****************************************************************************************************
' 4) PRODUCT IMAGES
'*****************************************************************************************************

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show Product Image (If there is one)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	pcs_ProductImage
	if pcv_strUseEnhancedViews = True then 
	%>
		<script type="text/javascript">	
			hs.align = '<%=pcv_strHighSlide_Align%>';
			hs.transitions = [<%=pcv_strHighSlide_Effects%>];
			hs.outlineType = '<%=pcv_strHighSlide_Template%>';
			hs.fadeInOut = <%=pcv_strHighSlide_Fade%>;
			hs.dimmingOpacity = <%=pcv_strHighSlide_Dim%>;
			//hs.numberPosition = 'caption';
			<% if bCounter>0 then %>
				if (hs.addSlideshow) hs.addSlideshow({
					slideshowGroup: 'slides',
					interval: <%=pcv_strHighSlide_Interval%>,
					repeat: true,
					useControls: true,
					fixedControls: false,
					overlayOptions: {
						opacity: .75,
						position: 'top center',
						hideOnMouseOut: <%=pcv_strHighSlide_Hide%>
					}
				});	
			<% end if %>
			function pcf_initEnhancement(ele,img) {
				if (document.getElementById('1')==null) {
					hs.expand(ele, { src: img, minWidth: <%=pcv_strHighSlide_MinWidth%>, minHeight: <%=pcv_strHighSlide_MinHeight%> }); 
				} else {
					document.getElementById('1').onclick();			
				}
			}
		</script>
		
	<% end if   
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show Product Image (If there is one)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END PRODUCT IMAGES
'*****************************************************************************************************
	
response.write "<br />"
response.write "<br />"

'*****************************************************************************************************
' 15) QUANTITY DISCOUNTS ZONE
'*****************************************************************************************************

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Show quantity discounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	err.clear
	pcs_QtyDiscounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Show quantity discounts
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END QUANTITY DISCOUNTS ZONE
'*****************************************************************************************************
	%>
    </td>
	</tr>
	<tr> 
		<td valign="top">
		<%
'*****************************************************************************************************
' 7) PRODUCT OPTIONS
'*****************************************************************************************************
	response.write "<br />"
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Options A,B
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if pcf_VerifyShowOptions then '// IF [price =0 and BTO] DO NOT show Options	
		
		'/////////////////////////////////////////////////////////////
		'//      ORDERING OPTIONS									//
		'/////////////////////////////////////////////////////////////
		
		'*************************************************************
		' START: Options
		'*************************************************************
		pcs_OptionsN
		'*************************************************************
		' END: Options
		'*************************************************************
				
	end if  
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Options A,B
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END PRODUCT OPTIONS
'*****************************************************************************************************


'*****************************************************************************************************
' 9) CUSTOM INPUT FIELDS
'*****************************************************************************************************
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Options X
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if pcf_VerifyShowOptions then '// IF [price =0 and BTO] DO NOT show Custom Fields	
		
		'/////////////////////////////////////////////////////////////
		'//      CUSTOM INPUT FIELDS								//
		'/////////////////////////////////////////////////////////////
		
		'*************************************************************
		' START: Options X
		'*************************************************************
		pcs_OptionsX							
		'*************************************************************
		' END: Options X
		'*************************************************************
		
	end if  
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Options x
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END CUSTOM INPUT FIELDS
'*****************************************************************************************************
%>
	</td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td colspan="3">
	<!-- Start Quantity and Add to Cart -->
		<% 	
	'*****************************************************************************************************
	' 3) QUANTITY AND ADD TO CART
	'*****************************************************************************************************
		if pFormQuantity="-1" then
		'/////////////////////////////////////
		'// Product NOT For Sale 			//
		'/////////////////////////////////////
			if pEmailText<>"" then 
				response.write "<div class=pcShowProductNFS>" 
				response.write pEmailText '// reason why it's not for sale
				response.write "</div>" 
			end if
					
		else 
		'/////////////////////////////////////
		'// Product For Sale				//
		'/////////////////////////////////////
		
			' 2a) Check for order level permission		
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' NOTES:
			' Check for order level permission "scorderlevel".
			' scorderlevel = 0 // everybody
			' scorderlevel = 1 // wholesale only
			' scorderlevel = 2 // catalog only
			
			' Also check what level the current customer is classified.
			' session("customerType") = "" // not logged in
			' session("customerType") = 1  // wholesale
			' session("customerType") = 0  // retail
			
			' Verify level is 0 OR is 1 with a custmer type of 1	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
			' 2b) If out of stock AND out of stock purchase is allowed show button.
				if pcf_OutStockPurchaseAllow then
				' // out of stock purchase is allowed
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: Show CUSTOMIZE BUTTON or ADD TO CART
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						If ((pserviceSpec<>0) AND ((pnoprices>0) OR (pPrice=0) OR (scConfigPurchaseOnly=1))) or ((iBTOQuoteSubmitOnly=1) and (pserviceSpec<>0)) then 
						' // customize button only						
						
								'/////////////////////////////////////////////////////////////
								'//      CUSTOMIZE BUTTON									//
								'/////////////////////////////////////////////////////////////							 
								'*************************************************************
								' START: Customize Button Only
								'*************************************************************
								pcs_CustomizeButton
								'*************************************************************
								' END: Customize Button Only
								'*************************************************************
								
						else 
						' // show add to cart
						
								'/////////////////////////////////////////////////////////////
								'//      ADD TO CART										//
								'/////////////////////////////////////////////////////////////							 
								'*************************************************************
								' START: Add to Cart
								'*************************************************************
								pcs_AddtoCart
								'*************************************************************
								' END: Add to Cart
								'*************************************************************
						end if 
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: Show CUSTOMIZE BUTTON or ADD TO CART
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				end if ' end 2b			
		end if ' end if pFormQuantity="-1" then 	
	'*****************************************************************************************************
	' END QUANTITY AND ADD TO CART
	'*****************************************************************************************************
		%>
	<!-- End Quantity and Add to Cart -->

<%
'*****************************************************************************************************
' 12) LONG DESCRIPTION
'*****************************************************************************************************

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Display long product description if there is a short description
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcs_LongProductDescription
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Display long product description if there is a short description
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************************************************
' END LONG DESCRIPTION
'*****************************************************************************************************
	%>
	
	</td>
</tr>
</table>
<!--#include file="../includes/javascripts/pcValidateFormViewPrd.asp"-->
</form>
<!-- End Form -->
</div>
<!--#include file="adminfooter.asp"-->
<%
call closeDb()   
call clearLanguage()

set conntemp=Nothing
set rs=Nothing
set pIdproduct=Nothing
set pDescription=Nothing
set pDetails=Nothing
set pListPrice=Nothing
set pImageUrl=Nothing
set pWeight=Nothing
%>