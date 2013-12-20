
<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "viewPrdPopWindow.asp"
' This page is handles and displays all product-level info
' All product info is retreived from the database and
' displayed in its corresponding display zone.
'
'/////////////////////////////////////////////////////////////////
' NOTES:														//
'																//
' The "viewPrdCode.asp" include will hold the routines that 
' display the product information. Each segment of product 
' information has been divided into zone.
'
' View the commented sections of this page to
' find a particular zone.
'
'
'/////////////////////////////////////////////////////////////////
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#INCLUDE FILE="../includes/settings.asp"-->
<!--#INCLUDE FILE="../includes/storeconstants.asp"-->
<!--#INCLUDE FILE="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/languages.asp"-->
<!--#INCLUDE FILE="../includes/currencyformatinc.asp"-->
<!--#INCLUDE FILE="../includes/shipFromSettings.asp"-->
<!--#INCLUDE FILE="../includes/taxsettings.asp"-->
<!--#INCLUDE FILE="../includes/languages_ship.asp"-->
<!--#INCLUDE FILE="../includes/adovbs.inc"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include FILE="../includes/pcProductOptionsCode.asp"--> 
<!--#include file="../includes/ErrorHandler.asp"-->
<%
'--> open database connection
call opendb()
%>
<!--#INCLUDE FILE="viewPrdCode.asp"-->
<%
Response.Buffer = True
'-------------------------------
' declare local variables
'-------------------------------

Dim conntemp, tIndex, tUpdPrd, pIdCategory, strBreadCrumb, pIdProduct, query, rs, dblpcCC_Price
Dim pcv_strViewPrdStyle, pcv_strFormAction, pcv_intValidationFile, pcv_blnBTOisConfig, iRewardPoints, pDescription
Dim pSku, pconfigOnly, pserviceSpec, pPrice, pBtoBPrice, pDetails, pListPrice, plistHidden, pimageUrl, pLgimageURL
Dim pArequired, pBrequired, pStock, pWeight, pEmailText, pFormQuantity, pnoshipping, pcustom1, pcontent1
Dim pcustom2, pcontent2, pcustom3, pcontent3, pxfield1, px1req, pxfield2, px2req, pxfield3, px3req, pnoprices
Dim pcv_intHideBTOPrice, pcv_intQtyValidate, pcv_lngMinimumQty, intpHideDefConfig, pnotax, BrandName
Dim FirstCnt, strDescription, intReward, pcv_BTORP, strConfigProductCategory, dblPrice, dblWPrice, intIdCategory
Dim VardiscGo, dblQuantityFrom, dblQuantityUntil, dblPercentage, dblDiscountPerWUnit, dblDiscountPerUnit
Dim intIdOptOptGrp, intIdOption, strOptionDescrip, OptInActive, optPrice, tempIdOptA, tempIdOptB
Dim xrequired, xfieldCnt, reqstring, TextArea, widthoffield, rowlength
Dim scCS, cs_showprod, cs_showcart, cs_showimage, crossSellText
Dim pcv_strOptionGroupDesc, pcv_intOptionGroupCount, pcv_strOptionGroupCount, pcv_strOptionGroupID, pcv_strOptionRequired
Dim xOptionsCnt, pcv_strNumberValidations, pcv_strFuntionCall, pcv_strReqOptString, xOtionrequired
Dim pcv_strProdImage_Url, pcv_strProdImage_LargeUrl, pcv_intProdImage_Columns, pcv_strShowImage_LargeUrl, pcv_strShowImage_Url, pcv_strCurrentUrl
Dim pcv_strAdditionalImages, cCounter 

'// When the product has additional images, this variable defines how many thumbnails are shown per row, below the main product image
pcv_intProdImage_Columns = 10


'*****************************************************************************************************
' START PAGE ON-LOAD
'*****************************************************************************************************



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Product
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

pcv_strCurrentUrl=Request.Form("pcv_strCurrentUrl")
pIdProduct=Request.Form("idProduct")

	if not validNum(pIdProduct) then
		call closedb()
		response.redirect "msg.asp?message=207"
	end if

'--> Check if this customer is logged in with a customer category
dblpcCC_Price=0
if session("customerCategory")<>0 then
	query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&session("customerCategory")&" AND idProduct="&pIdProduct&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if NOT rs.eof then
		strcustomerCategory="YES"
		dblpcCC_Price=rs("pcCC_Price")
		dblpcCC_Price=pcf_Round(dblpcCC_Price, 2)
	else
		strcustomerCategory="NO"
	end if
	set rs=nothing
end if

'--> check for discount per quantity
query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" &pidProduct
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	call closeDb()
  'response.redirect "techErr.asp?error="& Server.Urlencode("Error in " & pcStrPageName & " - Line 149") 
end if

if not rs.eof then
	pDiscountPerQuantity=-1
else
	pDiscountPerQuantity=0
end if
set rs=nothing

'--> gets product details from db
query="SELECT iRewardPoints, description, sku, configOnly, serviceSpec, price, btobprice, details, listprice, listHidden, imageurl, largeImageURL, Arequired, Brequired, stock, weight, emailText, formQuantity, noshipping, custom1, content1, custom2, content2, custom3, content3,  xfield1, x1req, xfield2, x2req, xfield3, x3req, noprices, IDBrand, sDesc, noStock, noshippingtext,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcprod_HideDefConfig, notax FROM products WHERE idProduct=" &pidProduct& " AND active=-1"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	call closeDb()
  'response.redirect "techErr.asp?error="& Server.Urlencode("Error in " & pcStrPageName & " - Line 168") 
end if

if rs.eof then
	set rs=nothing
	call closeDb()
  response.redirect "msg.asp?message=95"
end if

'--> set our product information
iRewardPoints=rs("iRewardPoints")
pDescription=replace(rs("description"),"&quot;",chr(34))
pSku= rs("sku")
pconfigOnly=rs("configOnly")
pserviceSpec=rs("serviceSpec")
pPrice=rs("price")
pBtoBPrice=rs("bToBPrice")
pDetails=replace(rs("details"),"&quot;",chr(34))
pListPrice=rs("listPrice")
plistHidden=rs("listHidden")
pimageUrl=rs("imageUrl")
pLgimageURL=rs("largeImageURL")
pArequired=rs("Arequired")
pBrequired=rs("Brequired")
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
if isNull(pnoprices) OR pnoprices="" then
	pnoprices=0
end if
pIDBrand=rs("IDBrand")
psDesc=rs("sDesc")
pNoStock=rs("noStock")
pnoshippingtext=rs("noshippingtext")
pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
if isNull(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
	pcv_intHideBTOPrice="0"
end if
pcv_intQtyValidate=rs("pcprod_QtyValidate")
if isNull( pcv_intQtyValidate) OR  pcv_intQtyValidate="" then
	pcv_intQtyValidate="0"
end if				
pcv_lngMinimumQty=rs("pcprod_MinimumQty")
if isNull(pcv_lngMinimumQty) OR pcv_lngMinimumQty="" then
	pcv_lngMinimumQty="0"
end if
intpHideDefConfig=rs("pcprod_HideDefConfig")
if isNull(intpHideDefConfig) OR intpHideDefConfig="" then
	intpHideDefConfig="0"
end if
pnotax=rs("notax")

set rs=nothing

'--> Check to see if the product has been assigned to a brand. If so, get the brand name
if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
 	query="select BrandName from Brands where IDBrand=" & pIDBrand
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if not rs.eof then
		BrandName=rs("BrandName")
	end if
	set rs=nothing
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Product
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'*****************************************************************************************************
' END PAGE ON-LOAD
'*****************************************************************************************************

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=pDescription%></title>
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />

<style>
/*
-----------------------------------------------------------------------------------------
 ProductCart Product ViewPrd.asp Images
----------------------------------------------------------------------------------------
*/	
	/* To limit the size of the popup images with CSS uncomment the style below */
	
	/*
	#pcMain .pcShowMainImage img {
		width: 300px;
	}
	*/

</style>

<script>
function closeWindow(){
window.opener = top;
window.self.close();
}
</script>
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
</head>
<body onLoad='javascript:window.document.mainimg.src="<%=pcv_strCurrentUrl%>";'>
	
<div id="pcMain">

<table class="pcMainTable">
<form method="GET" action="<%=pcStrPageName%>" name="pcShowAdditionalImages">
<input type="hidden" name="Description" value="<%=pDescription%>">	
	<tr> 
		<td> 
		<div align="right">
		<a href="javascript:closeWindow();"><img src="images/close.gif" border="0"></a>
		</div>			
		
		<h1><%=pDescription%></h1>
		</td>
	</tr>
	<tr> 
		<td align="center">
		<% 
		pcv_strPopWindowOpen = 1
		pcs_AdditionalImages
		%>
		</td>
	</tr>
	<tr> 
		<td>
		<div align="center" style="height: 450px; overflow:auto">
		<%
		pLgimageURL = "" '// clear large image value so we dont get a zoom button
		pcs_ProductImage
		%>
		</div>
		</td>
	</tr>

	<tr> 
		<td>
		<div><!--links here --></div>
		</td>
	</tr>
</form>
</table>

</div>	
</body>
</html>