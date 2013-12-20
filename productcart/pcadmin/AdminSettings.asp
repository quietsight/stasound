<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="Store Settings"
pageIcon="pcv4_icon_settings.png"
%>
<% Section="layout" %>
<%PmAdmin=1%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<% pcPageName="AdminSettings.asp"

dim query, mySQL, conntemp, rstemp

'/////////////////////////////////////////////////////
'// Retrieve current database data
'/////////////////////////////////////////////////////
%>
<!--#include file="pcAdminRetrieveSettings.asp"-->

<link href="../includes/spry/SpryTabbedPanels-Settings.css" rel="stylesheet" type="text/css" />
<script src="../includes/spry/SpryTabbedPanels.js" type="text/javascript"></script>
<script src="../includes/spry/SpryURLUtils.js" type="text/javascript"></script>
<script type="text/javascript">
	var params = Spry.Utils.getLocationParamsAsObject();
</script>

<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer" align="center">
			<%
			msg=getUserInput(request.querystring("msg"),0)
			if msg<>"" then %>
				<div class="pcCPmessage"><%=msg%></div>
			<% end if %>
		</td>
	</tr>
</table>
<%
pcv_isCompanyNameRequired=false
pcv_isCompanyAddressRequired=false
pcv_isCompanyPhoneNumberRequired=false
pcv_isCompanyFaxNumberRequired=false
pcv_isCompanyZipRequired=false
pcv_isCompanyCityRequired=false
pcv_isCompanyStateRequired=false
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isCompanyStateRequired=pcv_strStateCodeRequired
end if
pcv_isCompanyProvinceRequired=false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isCompanyProvinceRequired=pcv_strProvinceCodeRequired
end if
pcv_isCompanyCountryRequired=false
pcv_isCompanyLogoRequired=false
pcv_isMetaTitleRequired=true
pcv_isMetaDescriptionRequired=false
pcv_isMetaKeywordsRequired=false
pcv_isQtyLimitRequired=true
pcv_isAddLimitRequired=true
pcv_isPreRequired=true
pcv_isCustPreRequired=true
pcv_isCatImagesRequired=true
pcv_isShowStockLmtRequired=true
pcv_isOutOfStockPurchaseRequired=true
pcv_isCurSignRequired=false
pcv_isDecSignRequired=false
pcv_isDateFrmtRequired=false
pcv_isMinPurchaseRequired=true
pcv_isWholesaleMinPurchaseRequired=true
pcv_isURLredirectRequired=false
pcv_isSSLRequired=false
pcv_isSSLUrlRequired=false
pcv_isIntSSLPageRequired=false
pcv_isPrdRowRequired=true
pcv_isPrdRowsPerPageRequired=true
pcv_isCatRowRequired=true
pcv_isCatRowsPerPageRequired=true
pcv_isBTypeRequired=true
pcv_isStoreOffRequired=false
pcv_isStoreMsgRequired=false
pcv_isWLRequired=false
pcv_isTFRequired=false
pcv_isorderLevelRequired=false
pcv_isDisplayStockRequired=false
pcv_isHideCategoryRequired=false
pcv_isPCOrdRequired=true
pcv_isHideSortProRequired=false
pcv_isViewPrdStyleRequired=false
pcv_isOrderNameRequired=false
pcv_isAllowCheckoutWRRequired=false
pcv_isHideDiscFieldRequired=false
pcv_isAllowSeparateRequired=false
pcv_isDisableDiscountCodesRequired=false
pcv_isShowSKURequired=false
pcv_isShowSmallImgRequired=false
pcv_isHideRMARequired=false
pcv_isShowHDRequired=false
pcv_isStoreUseToolTipRequired=false
pcv_isErrorHandlerRequired=false
pcv_isDisableGiftRegistryRequired=false
pcv_isBrandProRequired=false
pcv_isBrandLogoRequired=false
pcv_isSeoURLsRequired=false
pcv_isSeoURLs404Required=false
pcv_isQuickBuyRequired=false
pcv_isATCEnabledRequired=false
pcv_isRestoreCartRequired=false
pcv_isAddThisDisplayRequired=false
pcv_isAddThisCodeRequired=false
pcv_isPinterestDisplayRequired=false
pcv_isPinterestCounterRequired=false
pcv_isGoogleAnalyticsRequired=false

if request("updateSettings")<>"" then
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions
	'/////////////////////////////////////////////////////

	'// set errors to none
	pcv_intErr=0

	'// generic error for page
	pcv_strGenericPageError = dictLanguageCP.Item(Session("language")&"_cpCommon_403")

	'// validate all fields
	pcs_ValidateTextField	"CompanyName", pcv_isCompanyNameRequired, 150
	pcs_ValidateTextField	"CompanyAddress", pcv_isCompanyAddressRequired, 250
	pcs_ValidateTextField	"CompanyZip", pcv_isCompanyZipRequired, 20
	pcs_ValidateTextField	"CompanyPhoneNumber", pcv_isCompanyPhoneNumberRequired, 20
	pcs_ValidateTextField	"CompanyFaxNumber", pcv_isCompanyFaxNumberRequired, 20
	pcs_ValidateTextField	"CompanyCity", pcv_isCompanyCityRequired, 50
	pcs_ValidateTextField	"CompanyState", pcv_isCompanyProvinceRequired, 50
	pcs_ValidateTextField	"CompanyProvince", pcv_isCompanyStateRequired, 50
	pcs_ValidateTextField	"CompanyCountry", pcv_isCompanyCountryRequired, 50
	pcs_ValidateTextField	"CompanyLogo", pcv_isCompanyLogoRequired, 250
	pcs_ValidateTextField	"MetaTitle", pcv_isMetaTitleRequired, 250
	pcs_ValidateTextField	"MetaDescription", pcv_isMetaDescriptionRequired, 250
	pcs_ValidateTextField	"MetaKeywords", pcv_isMetaKeywordsRequired, 250
	pcs_ValidateTextField	"QtyLimit", pcv_isQtyLimitRequired, 6
	pcs_ValidateTextField	"AddLimit", pcv_isAddLimitRequired, 6
	pcs_ValidateTextField	"Pre", pcv_isPreRequired, 15
	pcs_ValidateTextField	"CustPre", pcv_isCustPreRequired, 15
	pcs_ValidateTextField	"CatImages", pcv_isCatImagesRequired, 2
	pcs_ValidateTextField	"ShowStockLmt", pcv_isShowStockLmtRequired, 2
	pcs_ValidateTextField	"OutOfStockPurchase", pcv_isOutOfStockPurchaseRequired, 2
	pcs_ValidateTextField	"CurSign", pcv_isCurSignRequired, 10
	pcs_ValidateTextField	"DecSign", pcv_isDecSignRequired, 4
	pcs_ValidateTextField	"DateFrmt", pcv_isDateFrmtRequired, 10
	pcs_ValidateTextField	"MinPurchase", pcv_isMinPurchaseRequired, 20
	pcs_ValidateTextField	"WholesaleMinPurchase", pcv_isWholesaleMinPurchaseRequired, 20
	pcs_ValidateTextField	"URLredirect", pcv_isURLredirectRequired, 250
	pcs_ValidateTextField	"SSL", pcv_isSSLRequired, 4
	pcs_ValidateTextField	"SSLUrl", pcv_isSSLUrlRequired, 250
	pcs_ValidateTextField	"IntSSLPage", pcv_isIntSSLPageRequired, 4
	pcs_ValidateTextField	"PrdRow", pcv_isPrdRowRequired, 20
	pcs_ValidateTextField	"PrdRowsPerPage", pcv_isPrdRowsPerPageRequired, 20
	pcs_ValidateTextField	"CatRow", pcv_isCatRowRequired, 20
	pcs_ValidateTextField	"CatRowsPerPage", pcv_isCatRowsPerPageRequired, 20
	pcs_ValidateTextField	"BType", pcv_isBTypeRequired, 4
	pcs_ValidateTextField	"StoreOff", pcv_isStoreOffRequired, 4
	pcs_ValidateHTMLField	"StoreMsg", pcv_isStoreMsgRequired, 0
	pcs_ValidateTextField	"WL", pcv_isWLRequired, 2
	pcs_ValidateTextField	"TF", pcv_isTFRequired, 2
	pcs_ValidateTextField	"orderLevel", pcv_isorderLevelRequired, 2
	pcs_ValidateTextField	"DisplayStock", pcv_isDisplayStockRequired, 2
	pcs_ValidateTextField	"HideCategory", pcv_isHideCategoryRequired, 2
	pcs_ValidateTextField	"PCOrd", pcv_isPCOrdRequired, 10
	pcs_ValidateTextField	"HideSortPro", pcv_isHideSortProRequired, 10
	pcs_ValidateTextField	"ViewPrdStyle",  pcv_isViewPrdStyleRequired, 10
	pcs_ValidateTextField	"OrderName", pcv_isOrderNameRequired, 4
	pcs_ValidateTextField	"AllowCheckoutWR", pcv_isAllowCheckoutWRRequired, 4
	pcs_ValidateTextField	"HideDiscField", pcv_isHideDiscFieldRequired, 4
	pcs_ValidateTextField	"AllowSeparate", pcv_isAllowSeparateRequired, 4
	pcs_ValidateTextField	"DisableDiscountCodes", pcv_isDisableDiscountCodesRequired, 4
	pcs_ValidateTextField	"ShowSKU", pcv_isShowSKURequired, 4
	pcs_ValidateTextField	"ShowSmallImg", pcv_isShowSmallImgRequired, 4
	pcs_ValidateTextField	"HideRMA", pcv_isHideRMARequired, 4
	pcs_ValidateTextField	"ShowHD", pcv_isShowHDRequired, 4
	pcs_ValidateTextField	"StoreUseToolTip", pcv_isStoreUseToolTipRequired, 4
	pcs_ValidateTextField	"ErrorHandler", pcv_isErrorHandlerRequired, 4
	pcs_ValidateTextField	"DisableGiftRegistry", pcv_isDisableGiftRegistryRequired, 4
	pcs_ValidateTextField	"BrandPro", pcv_isBrandProRequired, 2
	pcs_ValidateTextField	"BrandLogo", pcv_isBrandLogoRequired, 2
	pcs_ValidateTextField	"SeoURLs", pcv_isSeoURLsRequired, 2
	pcs_ValidateTextField	"SeoURLs404", pcv_isSeoURLs404Required, 50
	pcs_ValidateTextField	"QuickBuy", pcv_isQuickBuyRequired, 4
	pcs_ValidateTextField	"ATCEnabled", pcv_isATCEnabledRequired, 4
	pcs_ValidateTextField	"RestoreCart", pcv_isRestoreCartRequired, 4
	pcs_ValidateTextField	"AddThisDisplay", pcv_isAddThisDisplayRequired, 2
	pcs_ValidateTextField	"PinterestDisplay", pcv_isPinterestDisplayRequired, 2
	pcs_ValidateTextField	"PinterestCounter", pcv_isPinterestCounterRequired, 15
	pcs_ValidateHtmlField	"AddThisCode", pcv_isAddThisCodeRequired, 0
	pcs_ValidateTextField	"GoogleAnalytics", pcv_isGoogleAnalyticsRequired, 50
		if NOT validNum(Session("pcAdminCatRow")) OR Session("pcAdminCatRow")= "" OR Session("pcAdminCatRow")= "0" then
			Session("pcAdminCatRow")= 3
		end if

		if NOT validNum(Session("pcAdminCatRowsPerPage")) OR Session("pcAdminCatRowsPerPage")= "" OR Session("pcAdminCatRowsPerPage")= "0" then
			Session("pcAdminCatRowsPerPage")= 3
		end if
		if NOT validNum(Session("pcAdminPrdRow")) OR Session("pcAdminPrdRow")= "" OR Session("pcAdminPrdRow")= "0" then
			Session("pcAdminPrdRow")= 3
		end if

		if NOT validNum(Session("pcAdminPrdRowsPerPage")) OR Session("pcAdminPrdRowsPerPage")= "" OR Session("pcAdminPrdRowsPerPage")= "0" then
			Session("pcAdminPrdRowsPerPage")= 3
		end if

		if NOT validNum(Session("pcAdminBrandPro")) OR Session("pcAdminBrandPro")<>"1" then
			Session("pcAdminBrandPro")="0"
		end if
		if NOT validNum(Session("pcAdminBrandLogo")) OR Session("pcAdminBrandLogo")<>"1" then
			Session("pcAdminBrandLogo")="0"
		end if
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Set specific errors and default values instead of generic error message.
		'response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError& "&lmode=" & pcLoginMode
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		if Session("pcAdminQtyLimit")= "" or not validNum(Session("pcAdminQtyLimit")) then
			Session("pcAdminQtyLimit")= 100
		end if

		if Session("pcAdminAddLimit")= "" or not validNum(Session("pcAdminAddLimit")) then
			Session("pcAdminAddLimit")= 10
		end if

		if Session("pcAdminPre")= "" then
			Session("pcAdminPre")= 0
		end if

		if Session("pcAdminCustPre")= "" then
			Session("pcAdminCustPre")= 0
		end if

		if Session("pcAdminCatImages")= "" then
			Session("pcAdminCatImages")= 0
		end if

		if Session("pcAdminShowStockLmt")= "" then
			Session("pcAdminShowStockLmt")= 0
		end if

		if Session("pcAdminOutOfStockPurchase")= "" then
			Session("pcAdminOutOfStockPurchase")= 0
		end if

		if Session("pcAdminMinPurchase")= "" then
			Session("pcAdminMinPurchase")= 0
		end if

		if Session("pcAdminWholesaleMinPurchase")= "" then
			Session("pcAdminWholesaleMinPurchase")= 0
		end if

		if Session("pcAdminPrdRow")= "" or not validNum(Session("pcAdminPrdRow")) then
			Session("pcAdminPrdRow")= 3
		end if

		if Session("pcAdminPrdRowsPerPage")= "" or not validNum(Session("pcAdminPrdRowsPerPage")) then
			Session("pcAdminPrdRowsPerPage")= 3
		end if

		if NOT validNum(Session("pcAdminCatRow")) OR Session("pcAdminCatRow")= "" then
			Session("pcAdminCatRow")= 3
		end if

		if NOT validNum(Session("pcAdminCatRowsPerPage")) OR Session("pcAdminCatRowsPerPage")= "" then
			Session("pcAdminCatRowsPerPage")= 3
		end if

		if Session("pcAdminBType")= "" then
			Session("pcAdminBType")= "H"
		end if

		if Session("pcAdminPCOrd")= "" then
			Session("pcAdminPCOrd")= 0
		end if

	End If

	'/////////////////////////////////////////////////////
	'// Set Local Variables for Setting
	'/////////////////////////////////////////////////////
	pcStrCompanyName = Session("pcAdminCompanyName")
	pcStrCompanyName=Replace(pcStrCompanyName,"&quot;","""""")
	pcStrCompanyName=Replace(pcStrCompanyName,"""","""""")
	pcStrCompanyAddress = Session("pcAdminCompanyAddress")
	pcStrCompanyZip = Session("pcAdminCompanyZip")
	pcStrCompanyPhoneNumber = Session("pcAdminCompanyPhoneNumber")
	pcStrCompanyFaxNumber = Session("pcAdminCompanyFaxNumber")
	pcStrCompanyCity = Session("pcAdminCompanyCity")
	if Session("pcAdminCompanyProvince")<>"" then
		pcStrCompanyState = Session("pcAdminCompanyProvince")
	else
		pcStrCompanyState = Session("pcAdminCompanyState")
	end if
	pcStrCompanyCountry = Session("pcAdminCompanyCountry")
	pcStrCompanyLogo = Session("pcAdminCompanyLogo")
	pcStrMetaTitle = Session("pcAdminMetaTitle")
	pcStrMetaDescription = Session("pcAdminMetaDescription")
	pcStrMetaKeywords = Session("pcAdminMetaKeywords")
	pcIntQtyLimit = Session("pcAdminQtyLimit")
		if pcIntQtyLimit=0 or not validNum(pcIntQtyLimit) then
			pcIntQtyLimit=50
		end if
	pcIntAddLimit = Session("pcAdminAddLimit")
		if pcIntAddLimit=0 or not validNum(pcIntAddLimit) then
			pcIntAddLimit=1000
		end if
	pcIntPre = Session("pcAdminPre")
		if not validNum(pcIntPre) then
			pcIntPre=0
		end if
	pcIntCustPre = Session("pcAdminCustPre")
		if not validNum(pcIntCustPre) then
			pcIntCustPre=0
		end if
	pcIntCatImages = Session("pcAdminCatImages")
	pcIntShowStockLmt = Session("pcAdminShowStockLmt")
	pcIntOutOfStockPurchase = Session("pcAdminOutOfStockPurchase")
	pcStrCurSign = Session("pcAdminCurSign")
	pcStrDecSign = Session("pcAdminDecSign")
	pcStrDateFrmt = Session("pcAdminDateFrmt")

	'// Alert that integers are required
	if not validNum(Session("pcAdminMinPurchase")) or not validNum(Session("pcAdminWholesaleMinPurchase")) then
		response.Redirect "AdminSettings.asp?tab=2&msg=" & Server.URLEncode("The &quot;Minimum Order Amount&quot; and &quot;Wholesale Minimum Order Amount&quot; fields must be integers.")
	end if
	pcIntMinPurchase = pcf_ReplaceChars(Session("pcAdminMinPurchase"))
	pcIntWholesaleMinPurchase = pcf_ReplaceChars(Session("pcAdminWholesaleMinPurchase"))

	pcStrURLredirect = Session("pcAdminURLredirect")
	pcStrSSL = Session("pcAdminSSL")
	pcStrSSLUrl = Session("pcAdminSSLUrl")
	pcStrIntSSLPage = Session("pcAdminIntSSLPage")
	pcIntPrdRow = Session("pcAdminPrdRow")
	pcIntPrdRowsPerPage = Session("pcAdminPrdRowsPerPage")
	pcIntCatRow = Session("pcAdminCatRow")
	pcIntCatRowsPerPage = Session("pcAdminCatRowsPerPage")
	pcStrBType = Session("pcAdminBType")
	if len(pcStrBType)<1 then
		pcStrBType="H"
	end if
	pcStrStoreOff = Session("pcAdminStoreOff")
	pcStrStoreMsg = Session("pcAdminStoreMsg")
	if pcStrStoreMsg="" then
		pcStrStoreMsg=scStoreMsg
	end if
	pcStrStoreMsg=Replace(pcStrStoreMsg, vbCrLf, "<BR>")
	pcStrStoreMsg=Replace(pcStrStoreMsg,"""","""""")
	pcIntWL = Session("pcAdminWL")
	pcIntTF = Session("pcAdminTF")
	pcStrorderLevel = Session("pcAdminorderLevel")
	pcIntDisplayStock = Session("pcAdminDisplayStock")
	pcIntHideCategory = Session("pcAdminHideCategory")
	pcIntPCOrd = Session("pcAdminPCOrd")
	pcIntHideSortPro = Session("pcAdminHideSortPro")
	if pcIntHideSortPro="" then
		pcIntHidesortPro=0
	end if
	pcStrViewPrdStyle = Session("pcAdminViewPrdStyle")
	if len(pcStrViewPrdStyle)<1 then
		pcStrViewPrdStyle="L"
	end if
	pcStrOrderName = Session("pcAdminOrderName")
	pcIntAllowCheckoutWR = Session("pcAdminAllowCheckoutWR")
	if pcIntAllowCheckoutWR="" then
		pcIntAllowCheckoutWR=0
	end if
	pcStrHideDiscField = Session("pcAdminHideDiscField")
	pcStrAllowSeparate = Session("pcAdminAllowSeparate")
	if pcStrAllowSeparate="" then
		pcStrAllowSeparate=0
	end if
	pcIntDisableDiscountCodes = Session("pcAdminDisableDiscountCodes")
	if pcIntDisableDiscountCodes="" then
		pcIntDisableDiscountCodes=0
	end if
	pcIntShowSKU = Session("pcAdminShowSKU")
	pcIntShowSmallImg = Session("pcAdminShowSmallImg")
	pcIntHideRMA = Session("pcAdminHideRMA")
	pcIntShowHD = Session("pcAdminShowHD")
	pcIntStoreUseToolTip = Session("pcAdminStoreUseToolTip")
	pcIntErrorHandler = Session("pcAdminErrorHandler")
	pcIntDisableGiftRegistry = Session("pcAdminDisableGiftRegistry")
	pcIntBrandPro = Session("pcAdminBrandPro")
	pcIntBrandLogo = Session("pcAdminBrandLogo")
	pcIntSeoURLs = Session("pcAdminSeoURLs")
	if pcIntSeoURLs="" then
		pcIntSeoURLs=0
	end if
	pcStrSeoURLs404 = Session("pcAdminSeoURLs404")
	pcIntQuickBuy = Session("pcAdminQuickBuy")
	pcIntATCEnabled = Session("pcAdminATCEnabled")
	pcIntRestoreCart = Session("pcAdminRestoreCart")
	pcIntAddThisDisplay = Session("pcAdminAddThisDisplay")
	if pcIntAddThisDisplay="" then
		pcIntAddThisDisplay=0
	end if
	pcStrAddThisCode = Session("pcAdminAddThisCode")
	pcIntPinterestDisplay = Session("pcAdminPinterestDisplay")
	if pcIntPinterestDisplay&""="" then
		pcIntPinterestDisplay="0"
	end if
	pcStrPinterestCounter = Session("pcAdminPinterestCounter")

	pcStrGoogleAnalytics = Session("pcAdminGoogleAnalytics")

	pcs_ClearAllSessions()
	if pcStrDecSign="," then
		pcStrDivSign="."
	else
		pcStrDivSign=","
	end if

	'/////////////////////////////////////////////////////
	'// Update database with new Settings
	'/////////////////////////////////////////////////////
	%>
	<!--#include file="pcAdminSaveSettings.asp"-->
	<!--#include file="pcAdminRetrieveSettings.asp"-->
<% end if %>

<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf
response.write "<!--"&vbcrlf
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf

StrGenericJSError=dictLanguageCP.Item(Session("language")&"_cpCommon_403")

pcs_JavaTextField	"CompanyName", pcv_isCompanyNameRequired, StrGenericJSError
pcs_JavaTextField	"CompanyAddress", pcv_isCompanyAddressRequired, StrGenericJSError
pcs_JavaTextField	"CompanyZip", pcv_isCompanyZipRequired, StrGenericJSError
pcs_JavaTextField	"CompanyPhoneNumber", pcv_isCompanyPhoneNumberRequired, StrGenericJSError
pcs_JavaTextField	"CompanyFaxNumber", pcv_isCompanyFaxNumberRequired, StrGenericJSError
pcs_JavaTextField	"CompanyCity", pcv_isCompanyCityRequired, StrGenericJSError
pcs_JavaTextField	"CompanyState", pcv_isCompanyStateRequired, StrGenericJSError
pcs_JavaDropDownList "CompanyCountry", pcv_isCompanyCountryRequired, StrGenericJSError
pcs_JavaTextField	"CompanyLogo", pcv_isCompanyLogoRequired, StrGenericJSError
pcs_JavaTextField	"MetaTitle", pcv_isMetaTitleRequired, StrGenericJSError
pcs_JavaTextField	"MetaDescription", pcv_isMetaDescriptionRequired, StrGenericJSError
pcs_JavaTextField	"MetaKeywords", pcv_isMetaKeywordsRequired, StrGenericJSError
pcs_JavaTextField	"QtyLimit", pcv_isQtyLimitRequired, StrGenericJSError
pcs_JavaTextField	"AddLimit", pcv_isAddLimitRequired, StrGenericJSError
pcs_JavaTextField	"Pre", pcv_isPreRequired, StrGenericJSError
pcs_JavaTextField	"CustPre", pcv_isCustPreRequired, StrGenericJSError
pcs_JavaTextField	"CurSign", pcv_isCurSignRequired, StrGenericJSError
pcs_JavaDropDownList "DecSign", pcv_isDecSignRequired, StrGenericJSError
pcs_JavaDropDownList "PinterestCounter", pcv_isPinterestCounterRequired, StrGenericJSError
pcs_JavaDropDownList "DateFrmt", pcv_isDateFrmtRequired, StrGenericJSError
pcs_JavaTextField	"MinPurchase", pcv_isMinPurchaseRequired, StrGenericJSError
pcs_JavaTextField	"WholesaleMinPurchase", pcv_isWholesaleMinPurchaseRequired, StrGenericJSError
pcs_JavaTextField	"URLredirect", pcv_isURLredirectRequired, StrGenericJSError
pcs_JavaCheckedBox "SSL", pcv_isSSLRequired, StrGenericJSError
pcs_JavaTextField	"SSLUrl", pcv_isSSLUrlRequired, StrGenericJSError
pcs_JavaTextField	"PrdRow", pcv_isPrdRowRequired, StrGenericJSError
pcs_JavaTextField	"PrdRowsPerPage", pcv_isPrdRowsPerPageRequired, StrGenericJSError
pcs_JavaTextField	"CatRow", pcv_isCatRowRequired, StrGenericJSError
pcs_JavaTextField	"CatRowsPerPage", pcv_isCatRowsPerPageRequired, StrGenericJSError
pcs_JavaTextField	"StoreMsg", pcv_isStoreMsgRequired, StrGenericJSError
pcs_JavaCheckedBox	"AllowCheckoutWR", pcv_isAllowCheckoutWR, StrGenericJSError
pcs_JavaCheckedBox "HideSortPro", pcv_isHideSortProRequired, StrGenericJSError
pcs_JavaCheckedBox	"AllowSeparate", pcv_isAllowSeparateRequired, StrGenericJSError
pcs_JavaCheckedBox	"DisableDiscountCodes", pcv_isDisableDiscountCodesRequired, StrGenericJSError
pcs_JavaTextField	"ShowSKU", pcv_isShowSKURequired, StrGenericJSError
pcs_JavaTextField	"ShowSmallImg", pcv_isShowSmallImgRequired, StrGenericJSError
pcs_JavaTextField	"QuickBuy", pcv_isQuickBuyRequired, StrGenericJSError
pcs_JavaTextField	"ATCEnabled", pcv_isATCEnabledRequired, StrGenericJSError
pcs_JavaTextField	"RestoreCart", pcv_isRestoreCartRequired, StrGenericJSError

response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "//-->"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' End Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

%>

<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td valign="top">
				<div id="TabbedPanels1" class="TabbedPanels">
				  <ul class="TabbedPanelsTabGroup">
					<li class="TabbedPanelsTab" tabindex="100"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_1")%></li>
					<li class="TabbedPanelsTab" tabindex="200"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_2")%></li>
					<li class="TabbedPanelsTab" tabindex="300"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_3")%></li>
					<li class="TabbedPanelsTab" tabindex="400"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_4")%></li>
					<li class="TabbedPanelsTab" tabindex="500"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_5")%></li>
				  </ul>

					<div class="TabbedPanelsContentGroup">
						<div class="TabbedPanelsContent">
						<table class="pcCPcontent">
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_6")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="StoreOff" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_7")%>
									<input type="radio" name="StoreOff" value="1" <% if pcStrStoreOff="1" then%>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_8")%>
									 &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=412')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
								<td align="right" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_9")%></td>
								<%
								pcStrStoreMsg=Replace(pcStrStoreMsg, "<BR>", vbCrLf)
								pcStrStoreMsg=Replace(pcStrStoreMsg, "<br>", vbCrLf)
								pcStrStoreMsg=Replace(pcStrStoreMsg, """""", """")
								%>
								<td><textarea name="StoreMsg" cols="60" rows="6"><%=pcStrStoreMsg%></textarea></td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_10")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_11")%>:</td>
								<td align="left">
								<input type="text" name="CurSign" value="<%=pcStrCurSign%>" size="20">
								<% pcs_RequiredImageTag "CurSign", pcv_isCurSignRequired %>
								</td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_12")%>: </td>
								<td align="left">
									<select name="DecSign">
									<option value=""></option>
									<option value="." <% if (pcStrDecSign=".") and (pcStrDivSign=",") then %>selected<%end if%>>1,234,567.89</option>
									<option value="," <% if (pcStrDecSign=",") and (pcStrDivSign=".") then %>selected<% end if %>>1.234.567,89</option>
									</select>
									<% pcs_RequiredImageTag "DecSign", pcv_isDecSignRequired %>
							  </td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_13")%>:</td>
								<td align="left">
									<select name="DateFrmt">
									<option value="MM/DD/YY" selected><%=dictLanguageCP.Item(Session("language")&"_cpCommon_235")%></option>
									<option value="DD/MM/YY" <% if pcStrDateFrmt="DD/MM/YY" then %>selected<% end if %>><%=dictLanguageCP.Item(Session("language")&"_cpCommon_234")%></option>
									</select>
									<% pcs_RequiredImageTag "DateFrmt", pcv_isDateFrmtRequired %>
							  </td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_14")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_15")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% if pcStrSSL="1" then %>
									<input type="checkbox" name="ssl" value="1" checked class="clearBorder">
								<% else %>
									<input type="checkbox" name="ssl" value="1" class="clearBorder">
								<% end if %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_16")%></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_17")%>:
								<input type="text" name="sslURL" size="50" value="<%=pcStrSSLURL%>">
								<% pcs_RequiredImageTag "sslURL", pcv_issslURLRequired %>
								</td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_18")%></td>
							</tr>
							<tr>
							<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_19")%>:</td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;&nbsp;
								<input name="intSSLPage" type="radio" value="1" <% if pcStrIntSSLPage="1" then %>checked<%end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_20")%>&nbsp;<a href="javascript:win('AdminSettingsSSL.asp')"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_21")%> &gt;&gt;</a></td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;&nbsp;
								<input name="intSSLPage" type="radio" value="0" <% if pcStrIntSSLPage="0" OR pcStrIntSSLPage="" then %>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_22")%></td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_23")%></th>
							</tr>
							<tr>
							<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
							<td class="pcCPspacer" colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_24")%></td>
							</tr>
							<tr>
								<td colspan="3"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_25")%>:
								<input type="text" name="URLredirect" size="50" maxlength="250" value="<%=pcStrURLredirect%>">
								<% pcs_RequiredImageTag "URLredirect", pcv_isURLredirectRequired %>
								</td>
							</tr>
							<tr>
								<td></td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_26")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_27")%>:
								<% If pcIntWL=0 then %>
									<input type="radio" name="WL" value="-1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="WL" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% else %>
									<input type="radio" name="WL" value="-1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="WL" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% end if %>
								 &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=413')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_28")%>:
								<% If pcIntTF=0 then %>
									<input type="radio" name="TF" value="-1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="TF" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% else %>
									<input type="radio" name="TF" value="-1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="TF" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% end if %>
								 &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=414')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						</table>
					</div>

					<div class="TabbedPanelsContent">
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_3")%></th>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td nowrap width="20%"><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_2")%>:</p></td>
								<td width="80%">
								<%
								pcStrCompanyName=Replace(pcStrCompanyName, """""", "&quot;")
								%>
								<p>
								<input type="text" name="CompanyName" value="<%=pcStrCompanyName%>" size="40">
								<% pcs_RequiredImageTag "CompanyName", pcv_isCompanyNameRequired %>
								</p>
								</td>
							</tr>
							<%
							call openDB()
							'///////////////////////////////////////////////////////////
							'// START: COUNTRY AND STATE/ PROVINCE CONFIG
							'///////////////////////////////////////////////////////////
							'
							' 1) Place this section ABOVE the Country field
							' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
							' 3) Additional Required Info

							'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
							pcv_isStateCodeRequired = pcv_isstateRequired '// determines if validation is performed (true or false)
							pcv_isProvinceCodeRequired = pcv_isprovinceRequired '// determines if validation is performed (true or false)
							pcv_isCountryCodeRequired = pcv_iscountryRequired '// determines if validation is performed (true or false)

							'// #3 Additional Required Info
							pcv_strTargetForm = "form1" '// Name of Form
							pcv_strCountryBox = "CompanyCountry" '// Name of Country Dropdown
							pcv_strTargetBox = "CompanyState" '// Name of State Dropdown
							pcv_strProvinceBox =  "CompanyProvince" '// Name of Province Field

							'// Set local Country to Session
							if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrCompanyCountry
							end if

							'// Set local State to Session
							if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrCompanyState
							end if

							'// Set local Province to Session
							if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
								Session(pcv_strSessionPrefix&pcv_strProvinceBox) = pcStrCompanyState
							end if
							%>
							<!--#include file="../includes/javascripts/pcStateAndProvince.asp"-->
							<%
							'///////////////////////////////////////////////////////////
							'// END: COUNTRY AND STATE/ PROVINCE CONFIG
							'///////////////////////////////////////////////////////////
							%>
							<%
							'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
							pcs_CountryDropdown
							%>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_3")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyAddress" value="<%=pcStrCompanyAddress%>" size="40">
								<% pcs_RequiredImageTag "CompanyAddress", pcv_isCompanyAddressRequired %>
								</p>
								</td>
							</tr>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_22")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyCity" value="<%=pcStrCompanyCity%>" size="40">
								<% pcs_RequiredImageTag "CompanyCity", pcv_isCompanyCityRequired %>
								</p>
								</td>
							</tr>
							<%
							'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
							pcs_StateProvince
							call closeDB()
							%>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_25")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyZip" value="<%=pcStrCompanyZip%>" size="40">
								<% pcs_RequiredImageTag "CompanyZip", pcv_isCompanyZipRequired %>
								</p>
								</td>
							</tr>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_13")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyPhoneNumber" value="<%=pcStrCompanyPhoneNumber%>" size="40">
								<% pcs_RequiredImageTag "CompanyPhoneNumber", pcv_isCompanyPhoneNumberRequired %>
								</p>
								</td>
							</tr>
							<tr>
								<td><p><%=dictLanguageCP.Item(Session("language")&"_cpCommon_14")%>:</p></td>
								<td>
								<p>
								<input type="text" name="CompanyFaxNumber" value="<%=pcStrCompanyFaxNumber%>" size="40">
								<% pcs_RequiredImageTag "CompanyFaxNumber", pcv_isCompanyFaxNumberRequired %>
								</p>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_29")%>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=472')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></th>
							</tr>
								<tr>
									<td class="pcCPspacer" colspan="2"></td>
								</tr>
						  <script language="javascript">
							function chgWin(file,window) {
								msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
								if (msgWindow.opener == null) msgWindow.opener = self;
							}
							</script>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_312")%>:</td>
								<td>
								<input type="text" name="CompanyLogo" value="<%=pcStrCompanyLogo%>" size="40">
								<% pcs_RequiredImageTag "CompanyLogo", pcv_isCompanyLogoRequired %>
								<a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=CompanyLogo&fid=form1','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>
								<a href="javascript:;" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')"><img src="images/sortasc_blue.gif" alt="Upload Image"></a>
							&nbsp;(e.g.: <i>mylogo.gif</i>)
								</td>
							</tr>
							<tr>
								<td></td>
								<td>
								<% if trim(pcStrCompanyLogo)<>"" then %>
								<hr>
								Currently using:
								<div style="padding: 15px 0;"><img src="../pc/catalog/<%=pcStrCompanyLogo%>"></div>
								<% end if %>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2">Default Meta Tags&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=473')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td valign="top">Title:</td>
								<td>
								<textarea id="MetaTitle" name="MetaTitle" rows="2" cols="60" onkeyup="javascript:testchars(this,'1',250); javascript:document.getElementById('MetaTitleCounter').style.display='';"><%=pcStrMetaTitle%></textarea>
								<% pcs_RequiredImageTag "MetaTitle", pcv_isMetaTitleRequired %>
								<div id="MetaTitleCounter" style="margin-top: 5px; display: none; color:#666;">There are <span id="countchar1" name="countchar1" style="font-weight: bold"><%=maxlength%></span> characters left. Recommended length: around 60 characters.</div>
								</td>
							</tr>
							<tr>
								<td valign="top">Description:</td>
								<td>
								<textarea id="MetaDescription" name="MetaDescription" rows="2" cols="60" onkeyup="javascript:testchars(this,'2',250); javascript:document.getElementById('MetaDescriptionCounter').style.display='';"><%=pcStrMetaDescription%></textarea>
								<% pcs_RequiredImageTag "MetaTitle", pcv_isMetaDescriptionRequired %>
								<div id="MetaDescriptionCounter" style="margin-top: 5px; display: none; color:#666;">There are <span id="countchar2" name="countchar2" style="font-weight: bold"><%=maxlength%></span> characters left. Recommended length: around 150 characters.</div>
								</td>
							</tr>
							<tr>
								<td valign="top">Keywords:</td>
								<td>
								<textarea id="MetaKeywords" name="MetaKeywords" rows="2" cols="60" onkeyup="javascript:testchars(this,'3',250); javascript:document.getElementById('MetaKeywordsCounter').style.display='';"><%=pcStrMetaKeywords%></textarea>
								<% pcs_RequiredImageTag "MetaKeywords", pcv_isMetaKeywordsRequired %>
								<div id="MetaKeywordsCounter" style="margin-top: 5px; display: none; color:#666;">There are <span id="countchar3" name="countchar3" style="font-weight: bold"><%=maxlength%></span> characters left.</div>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
						</table>
					</div>

					<div class="TabbedPanelsContent">
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_31")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
							<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_32")%>:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=415')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td align="right">
								<input name="orderlevel" type="radio" value="0" checked class="clearBorder">
								</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_33")%></td>
							</tr>
							<tr>
								<td align="right">
								<input type="radio" name="orderlevel" value="1" <% if pcStrOrderlevel="1" then%>checked<%end if%> class="clearBorder">
								</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_34")%></td>
							</tr>
							<tr>
								<td align="right">
								<input type="radio" name="orderlevel" value="2" <% if pcStrOrderlevel="2" then%>checked<%end if%> class="clearBorder">
								</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_35")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_36")%>:</td>
								<td align="left">
								<input type="text" name="QtyLimit" value="<%=pcIntQtyLimit%>" size="20">
								<% pcs_RequiredImageTag "QtyLimit", pcv_isQtyLimitRequired %>
								&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=416')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2" align="right"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_37")%>:</td>
								<td align="left">
								<input type="text" name="AddLimit" value="<%=pcIntAddLimit%>" size="20">
								<% pcs_RequiredImageTag "AddLimit", pcv_isAddLimitRequired %>
								&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=417')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_38")%>:</td>
								<td align="left">
								<input type="text" name="MinPurchase" value="<%=pcIntMinPurchase%>" size="20">
								<% pcs_RequiredImageTag "MinPurchase", pcv_isMinPurchaseRequired %>
								</td>
							</tr>
							<tr>
								<td align="right">&nbsp;</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_39")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_40")%>:</td>
								<td align="left">
								<input type="text" name="WholesaleMinPurchase" value="<%=pcIntWholesaleMinPurchase%>" size="20">
								<% pcs_RequiredImageTag "WholesaleMinPurchase", pcv_isWholesaleMinPurchaseRequired %>
								</td>
							</tr>
							<tr>
								<td align="right">&nbsp;</td>
								<td align="left"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_41")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_42")%>:</td>
								<td>
								<input name="Pre" type="text" id="Pre" value="<%=pcIntPre%>" size="10">
								<% pcs_RequiredImageTag "Pre", pcv_isPreRequired %>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_43")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_44")%>:</td>
								<td>
								<input name="CustPre" type="text" id="CustPre" value="<%=pcIntCustPre%>" size="10">
								<% pcs_RequiredImageTag "CustPre", pcv_isCustPreRequired %>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_45")%></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_47")%>:
									<% If pcStrOrderName="1" then %>
									<input type="radio" name="OrderName" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="OrderName" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									<% else %>
									<input type="radio" name="OrderName" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="OrderName" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
									<% end if %>
									&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=434')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
							 </td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_48")%>: <input type="checkbox" name="AllowSeparate" value="1" <%if pcStrAllowSeparate="1" then%>checked<%end if%> class="clearBorder">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=108')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_76")%>: <% If pcIntDisableDiscountCodes="1" then %>
									<input type="radio" name="DisableDiscountCodes" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DisableDiscountCodes" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% else %>
									<input type="radio" name="DisableDiscountCodes" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DisableDiscountCodes" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% end if %>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=108')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_49")%>:
								<% If pcIntOutofstockpurchase="0" then %>
								<input type="radio" name="outofstockpurchase" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="outofstockpurchase" value="-1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% else %>
								<input type="radio" name="outofstockpurchase" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="outofstockpurchase" value="-1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% end if %>
								&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=411')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						</table>
					</div>

					<div class="TabbedPanelsContent">
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_50")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_51")%>:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=427')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_505")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="0" <% If trim(pcIntCatImages)="0" then  %>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_506")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="4" <% If trim(pcIntCatImages)="4" then  %>checked<% end if %> class="clearBorder">Thumbnails only
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="2" <% If trim(pcIntCatImages)="2" then  %>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_507")%>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_52")%>:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=428')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td align="right" nowrap="nowrap"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_508")%>:</td>
								<td align="left"><input type="text" name="CatRow" value="<%=pcIntCatRow%>">
								<% pcs_RequiredImageTag "CatRow", pcv_isCatRowRequired %>
								</td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
								<td width="556" align="left">
								<input type="text" name="CatRowsperPage" value="<%=pcIntCatRowsPerPage%>">
								<% pcs_RequiredImageTag "CatRowsperPage", pcv_isCatRowsperPageRequired %>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_53")%>:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=429')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td colspan="2">
									<% If ucase(trim(pcStrBType))="H" then  %>
									 <input type="radio" name="BType" value="H" checked class="clearBorder">
									<% Else %>
									 <input type="radio" name="BType" value="H" class="clearBorder">
									<% End If %>
								 <%=dictLanguageCP.Item(Session("language")&"_cpCommon_510")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrBType))="P" then  %>
								 <input type="radio" name="BType" value="P" checked class="clearBorder">
								<% Else %>
								 <input type="radio" name="BType" value="P" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_511")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrBType))="L" then  %>
									<input type="radio" name="BType" value="L" checked class="clearBorder">
								<% Else %>
									<input type="radio" name="BType" value="L" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_512")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrBType))="M" then  %>
									<input type="radio" name="BType" value="M" checked class="clearBorder">
								<% Else %>
									<input type="radio" name="BType" value="M" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_513")%></td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_54")%>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=430')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td align="right" width="20%" nowrap="nowrap"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_514")%>:</td>
								<td align="left" width="80%" nowrap="nowrap"><input type="text" name="PrdRow" value="<%=pcIntPrdRow%>">
								<% pcs_RequiredImageTag "PrdRow", pcv_isPrdRowRequired %></td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
								<td align="left">
								<input type="text" name="PrdRowsPerPage" value="<%=pcIntPrdRowsPerPage%>">
								<% pcs_RequiredImageTag "PrdRowsPerPage", pcv_isPrdRowsPerPageRequired %>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_55")%>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=431')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td colspan="2">
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_515")%>: <input type="radio" name="ShowSKU" value="-1" <%If pcIntShowSKU="-1" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="radio" name="ShowSKU" value="0" <%If pcIntShowSKU="0" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_516")%>: <input type="radio" name="ShowSmallImg" value="-1" <%If pcIntShowSmallImg="-1" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="radio" name="ShowSmallImg" value="0" <%If pcIntShowSmallImg="0" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><a name="brandSettings"></a>Brands Display Settings</th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2">These settings apply to top level brands. <a href="BrandsManage.asp" target="_blank">Additional display settings</a> are available for second level brands (subBrands) and products displayed within a brand. Items per row and rows per page use the Category Display settings set above.</td>
							</tr>
							<tr>
								<td colspan="2">
								<input name="BrandPro" type="checkbox" id="BrandPro" value="1" <%if sBrandPro=1 then%>checked<%end if%>>
								<% pcs_RequiredImageTag "BrandPro", pcv_isBrandProRequired %>
								Show brand on product details page
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<input type="checkbox" name="BrandLogo" value="1" <%if sBrandLogo=1 then%>checked<%end if%>>
								<% pcs_RequiredImageTag "BrandLogo", pcv_isBrandLogoRequired %>
								Show brand logo on &quot;<a href="../pc/viewbrands.asp" target="_blank">Browse By Brand</a>&quot; page</td>
							</tr>

							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_56")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_57")%>:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=432')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If (pcIntPCOrd="") or (pcIntPCOrd="0") then  %>
									<input type="radio" name="PCOrd" value="0" checked class="clearBorder">
								<%Else%>
									<input type="radio" name="PCOrd" value="0" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_58")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If (pcIntPCOrd="1") then  %>
									<input type="radio" name="PCOrd" value="1" checked class="clearBorder">
								<%Else%>
									<input type="radio" name="PCOrd" value="1" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_59")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If (pcIntPCOrd="2") then  %>
									<input type="radio" name="PCOrd" value="2" checked class="clearBorder">
								<%Else%>
									<input type="radio" name="PCOrd" value="2" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_60")%>&nbsp;</td>
							</tr>
							<tr>
								<td colspan="2">
								<% If (pcIntPCOrd="3") then  %>
									<input type="radio" name="PCOrd" value="3" checked class="clearBorder">
								<%Else%>
									<input type="radio" name="PCOrd" value="3" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_61")%></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_65")%>:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=433')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td colspan="2">
								<input type="checkbox" name="HideSortPro" value="1" <% If (pcIntHideSortPro="1") then  %>checked<%end if%> class="clearBorder">
								<%=dictLanguageCP.Item(Session("language")&"_cpSettings_64")%></td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_62")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_63")%>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=424')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
							</tr>
							<tr>
								<td colspan="2">
									<% If ucase(trim(pcStrViewPrdStyle))="C" then  %>
									 <input type="radio" name="ViewPrdStyle" value="C" checked class="clearBorder">
									<% Else %>
									 <input type="radio" name="ViewPrdStyle" value="C" class="clearBorder">
									<% End If %>
								 <%=dictLanguageCP.Item(Session("language")&"_cpCommon_502")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrViewPrdStyle))="L" then  %>
								 <input type="radio" name="ViewPrdStyle" value="L" checked class="clearBorder">
								<% Else %>
								 <input type="radio" name="ViewPrdStyle" value="L" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_503")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrViewPrdStyle))="O" then  %>
									<input type="radio" name="ViewPrdStyle" value="O" checked class="clearBorder">
								<% Else %>
									<input type="radio" name="ViewPrdStyle" value="O" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_504")%></td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
						</table>
					</div>

					<div class="TabbedPanelsContent">
						<table class="pcCPcontent">
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2">Show/Hide Storefront Elements</th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2">
								On a store with hundreds of categories and many category levels, you can use the following setting to improve the loading time for the storefront's <a href="../pc/search.asp" target="_blank">advanced search page</a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=410')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>:
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<input name="hideCategory" type="radio" value="0" class="clearBorder" <% If pcIntHideCategory="0" or pcIntHideCategory="" then %>checked<%end if%>> Show all categories in the drop-down <br />
								<input name="hideCategory" type="radio" value="1" class="clearBorder" <% If pcIntHideCategory="1" then %>checked<%end if%>> Only show top level categories (<u>recommended</u> for stores with large category trees) <br />
								<input name="hideCategory" type="radio" value="-1" class="clearBorder" <% If pcIntHideCategory="-1" then %>checked<%end if%>> Hide the categories drop-down completely
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_68")%>:</td>
								<td>
								<% If pcIntDisplayStock="-1" then %>
								<input type="radio" name="displayStock" value="-1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="displayStock" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
								<input type="radio" name="displayStock" value="-1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="displayStock" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
							 </td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_69")%>:</td>
								<td>
								<% If pcIntShowStockLmt="-1" then %>
								<input type="radio" name="ShowStockLmt" value="-1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="ShowStockLmt" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
								<input type="radio" name="ShowStockLmt" value="-1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
								<input type="radio" name="ShowStockLmt" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_70")%>: </td>
								<td>
								<% If pcStrHideDiscField="1" then %>
									<input type="radio" name="HideDiscField" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="HideDiscField" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="HideDiscField" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="HideDiscField" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
							 &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=418')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
								<td>Enable &quot;Quick Buy&quot; Feature:</td>
								<td>
								<% If pcIntQuickBuy="1" then %>
									<input type="radio" name="QuickBuy" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="QuickBuy" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="QuickBuy" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="QuickBuy" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=464')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
								<td>Enable &quot;Stay on Page when Adding To Cart&quot; Feature:</td>
								<td>
								<% If pcIntATCEnabled="1" then %>
									<input type="radio" name="ATCEnabled" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="ATCEnabled" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="ATCEnabled" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="ATCEnabled" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=465')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
								<td>Restore saved shopping cart on next visit:</td>
								<td>
								<% If pcIntRestoreCart="1" then %>
									<input type="radio" name="RestoreCart" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="RestoreCart" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="RestoreCart" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="RestoreCart" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
								&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=468')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>

							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
								<td valign="top">
									<a href="http://www.addthis.com/web-button-select" target="_blank"><img src="images/AddThisLogo.png" alt="AddThis" border="0"></a>
								</td>
								<td>
									<div style="margin-bottom: 10px;">Show <a href="http://www.addthis.com/web-button-select" target="_blank">AddThis</a> Buttons:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=471')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></div>
									<input type="radio" name="AddThisDisplay" value="0"<% If pcIntAddThisDisplay="0" then %> checked<% end if %> class="clearBorder"> Never &nbsp;
									<input type="radio" name="AddThisDisplay" value="1"<% If pcIntAddThisDisplay="1" then %> checked<% end if %> class="clearBorder"> Right of Page Title &nbsp;
									<input type="radio" name="AddThisDisplay" value="2"<% If pcIntAddThisDisplay="2" then %> checked<% end if %> class="clearBorder"> Below 'Add to Cart' section
								</td>
							</tr>
							<tr>
								<td></td>
								<td>
									<% if trim(pcStrAddThisCode)<>"" then %>
									<div style="margin-top: 10px; padding-top: 10px; border-top: 1px dashed #CCC;">
									You are currently using:
									<div style="margin: 10px 0;"><%=pcStrAddThisCode%></div>
									</div>
									<% end if %>
									<div style="margin-top: 10px; padding-top: 10px; border-top: 1px dashed #CCC;"><a href="http://www.addthis.com/web-button-select" target="_blank">Get the AddThis code</a> that best fits your needs, then <a href="JavaScript:;" onClick="document.getElementById('AddThisCodeDiv').style.display='';">paste it here</a>.</div>
									<div style="margin-top: 10px; display: none;" id="AddThisCodeDiv">
										<textarea name="AddThisCode" id="AddThisCode" cols="50" rows="6" onClick="selectFieldContent('AddThisCode')"><%=pcStrAddThisCode%></textarea>
									</div>
								</td>
							</tr>
							<tr>
							  <td colspan="2"><hr></td>
							</tr>
							<tr>
							  <td><img src="images/pinterest.JPG" width="112" height="38" alt="Pinterest ~ Pin It Button"></td>
                              <td><div style="margin-bottom: 10px;">Show Pinterest's &quot;Pin It&quot; Button:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=471')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></div>
							<input type="radio" name="PinterestDisplay" value="1"<% If pcIntPinterestDisplay="1" then %> checked<% end if %> class="clearBorder"> 
	Enable &nbsp;
							<input type="radio" name="PinterestDisplay" value="0"<% If pcIntPinterestDisplay="0" then %> checked<% end if %> class="clearBorder"> 
	Disable</td>
						  </tr>
							<tr>
							  <td></td>
							  <td>Pin Counter: 	
                              	<select name="PinterestCounter">
									<option value="none" selected>Don't Show</option>
									<option value="horizontal" <% if (pcStrPinterestCounter="horizontal") then %>selected<%end if%>>Beside</option>
									<option value="vertical" <% if (pcStrPinterestCounter="vertical") then %>selected<% end if %>>Above</option>
									</select>
									<% pcs_RequiredImageTag "PinterestCounter", pcv_isPinterestCounterRequired %>
</td>
						  </tr>
							<tr>
								<td class="pcCPspacer" colspan="2" style="padding-top: 20px;"></td>
							</tr>
							<tr>
								<th colspan="2">Turn Features On/Off</th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<%
							pcv_SEOURLCheck = scStoreURL&"/"&scPcFolder&"/pc/SEOCheck-c1.htm"
							pcv_SEOURLCheck = replace(pcv_SEOURLCheck,"//","/")
							pcv_SEOURLCheck = replace(pcv_SEOURLCheck,"http:/","http://")
							pcv_SEOURLCheck = replace(pcv_SEOURLCheck,"https:/","https://")

							'Send the transaction info as part of the querystring
							set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
							xml.open "GET", pcv_SEOURLCheck, false

							xml.send ""
							strStatus = xml.Status

							Set xml = Nothing

							If strStatus="200" then %>
							<tr>
								<td colspan="2">Use Keyword-rich URLs:
								<% If pcIntSeoURLs="1" then %>
									<input type="radio" name="SeoURLs" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="SeoURLs" value="0" class="clearBorder" onClick="JavaScript:alert('Turning off this feature can cause links to your store pages to return 404 Page Not Found errors. Make sure to update all navigation links so that they no longer use the keyword rich URLs as those URLs will no longer be working. Among other things, remember to re-generate the storefront navigation if you are using that feature, review header.asp and footer.asp, and any other page on your Web site that was linking to storefront pages.')"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% else %>
									<input type="radio" name="SeoURLs" value="1" class="clearBorder" onClick="JavaScript:alert('Make sure that you have changed the 404 error handler in your Web hosting account or directly on your dedicated Web server before enabling this feature. See the documentation for more information.')"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="SeoURLs" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% end if %>
								&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=460')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
								<td align="left">File name of &quot;Page Not Found&quot; page:</td>
								<td align="left">
								<input type="text" name="SeoURLs404" value="<%=pcStrSeoURLs404%>">
								<% pcs_RequiredImageTag "SeoURLs404", pcv_isSeoURLs404Required %>
								&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=461')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><div class="pcCPmessageInfo">This feature requires that you <a href="http://wiki.earlyimpact.com/productcart/seo-urls" target="_blank">carefully review the related documentation</a>.</div></td>
							</tr>
							<% else %>
								<tr>
									<td valign="top">Use Keyword-rich URLs:</td>
									<td><div class="pcCPmessageInfo" style="width: 300px; margin: 0;">This feature cannot be activated until you have configured your Web server to use &quot;404.asp&quot; as the customer error handler for &quot;404 Page Not Found&quot; errors. <a href="http://wiki.earlyimpact.com/productcart/seo-urls" target="_blank">Review the documentation &gt;&gt;</a></div></td>
								</tr>
							<% end if %>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td valign="top">
									<a href="https://www.google.com/analytics/" target="_blank"><img src="images/ga_logo.png" alt="Google Analytics" border="0"></a>
								</td>
								<td>
									Enter your <a href="https://www.google.com/analytics/" target="_blank">Google Analytics Profile ID</a> to activate the integration:
									<div style="margin-top: 10px;">Web site profile ID: <input type="text" name="GoogleAnalytics" id="GoogleAnalytics" value="<%=pcStrGoogleAnalytics%>" onClick="selectFieldContent('GoogleAnalytics')">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=470')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></div>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_75")%>:</td>
								<td>
								<% If pcIntDisableGiftRegistry="1" then %>
									<input type="radio" name="DisableGiftRegistry" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DisableGiftRegistry" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% else %>
									<input type="radio" name="DisableGiftRegistry" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="DisableGiftRegistry" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>&nbsp;
								<% end if %>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_71")%>:</td>
								<td>
								<% If pcIntHideRMA="1" then %>
									<input type="radio" name="HideRMA" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="HideRMA" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% else %>
									<input type="radio" name="HideRMA" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
									<input type="radio" name="HideRMA" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								<% end if %>
							 &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=421')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_72")%>:</td>
								<td>
								<% If pcIntShowHD="1" then %>
									<input type="radio" name="ShowHD" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_313")%>&nbsp;
									<input type="radio" name="ShowHD" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_314")%>&nbsp;
								<% else %>
									<input type="radio" name="ShowHD" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_313")%>&nbsp;
									<input type="radio" name="ShowHD" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_314")%>&nbsp;
								<% end if %>
							 &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=422')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td valign="top" nowrap>Category, product, and search previews (AJAX):&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=419')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></td>
								<td>
									<input type="radio" name="StoreUseToolTip" value="1" class="clearBorder"<% If pcIntStoreUseToolTip="1" then %> checked<%end if%>> Product, category, and search
									<div style="margin-top: 6px;"><input type="radio" name="StoreUseToolTip" value="2" class="clearBorder"<% If pcIntStoreUseToolTip="2" then %> checked<%end if%>> <strong>Product</strong> details preview only</div>
									<div style="margin-top: 6px;"><input type="radio" name="StoreUseToolTip" value="3" class="clearBorder"<% If pcIntStoreUseToolTip="3" then %> checked<%end if%>> <strong>Category</strong> content preview only</div>
									<div style="margin-top: 6px;"><input type="radio" name="StoreUseToolTip" value="4" class="clearBorder"<% If pcIntStoreUseToolTip="4" then %> checked<%end if%>> <strong>Search</strong> results preview only</div>
									<div style="margin-top: 6px;"><input type="radio" name="StoreUseToolTip" value="0" class="clearBorder"<% If pcIntStoreUseToolTip="0" or pcIntStoreUseToolTip="" then %> checked<%end if%>> Turn feature OFF</div>

								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpSettings_74")%>:</td>
								<td>
								<% If pcIntErrorHandler="1" then %>
									<input type="radio" name="ErrorHandler" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_313")%>&nbsp;
									<input type="radio" name="ErrorHandler" value="0" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_314")%>&nbsp;
								<% else %>
									<input type="radio" name="ErrorHandler" value="1" class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_313")%>&nbsp;
									<input type="radio" name="ErrorHandler" value="0" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_314")%>&nbsp;
								<% end if %>
							 &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=420')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a>
								</td>
							</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						</table>
						</div>
					</div>
				</div>
				<script type="text/javascript">
					var TabbedPanels1 = new Spry.Widget.TabbedPanels("TabbedPanels1", {defaultTab: params.tab ? params.tab : 0});
				</script>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
				<p>
				  <input type="submit" name="updateSettings" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_107")%>" class="submit2">
				</p>
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->