<% 
Call Opendb()

query="SELECT pcStoreVersion_Num, pcStoreVersion_Sub, pcStoreVersion_SP FROM pcStoreVersions WHERE pcStoreVersion_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if NOT rs.eof then

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	pcStrScVersion=rs("pcStoreVersion_Num")
	pcStrScSubVersion=rs("pcStoreVersion_Sub")
	pcStrScSP=rs("pcStoreVersion_SP")
end if 

if isNull(pcStrScVersion) OR pcStrScVersion&""="" then
	pcStrScVersion = scVersion
	pcStrScSubVersion = scSubVersion
	pcStrScSP = 0
end if

query="SELECT pcStoreSettings_CompanyName, pcStoreSettings_CompanyAddress, pcStoreSettings_CompanyZip, pcStoreSettings_CompanyCity, pcStoreSettings_CompanyState, pcStoreSettings_CompanyCountry, pcStoreSettings_CompanyLogo, pcStoreSettings_QtyLimit, pcStoreSettings_AddLimit, pcStoreSettings_Pre, pcStoreSettings_CustPre, pcStoreSettings_CatImages, pcStoreSettings_ShowStockLmt, pcStoreSettings_OutOfStockPurchase, pcStoreSettings_Cursign, pcStoreSettings_DecSign, pcStoreSettings_DivSign, pcStoreSettings_DateFrmt, pcStoreSettings_MinPurchase, pcStoreSettings_WholesaleMinPurchase, pcStoreSettings_URLredirect, pcStoreSettings_SSL, pcStoreSettings_SSLUrl, pcStoreSettings_IntSSLPage, pcStoreSettings_PrdRow, pcStoreSettings_PrdRowsPerPage, pcStoreSettings_CatRow, pcStoreSettings_CatRowsPerPage, pcStoreSettings_BType, pcStoreSettings_StoreOff, pcStoreSettings_StoreMsg, pcStoreSettings_WL, pcStoreSettings_TF, pcStoreSettings_orderLevel, pcStoreSettings_DisplayStock, pcStoreSettings_HideCategory, pcStoreSettings_AllowNews, pcStoreSettings_NewsCheckOut, pcStoreSettings_NewsReg, pcStoreSettings_NewsLabel, pcStoreSettings_PCOrd, pcStoreSettings_HideSortPro, pcStoreSettings_DFLabel, pcStoreSettings_DFShow, pcStoreSettings_DFReq, pcStoreSettings_TFLabel, pcStoreSettings_TFShow, pcStoreSettings_TFReq, pcStoreSettings_DTCheck, pcStoreSettings_DeliveryZip, pcStoreSettings_OrderName, pcStoreSettings_HideDiscField, pcStoreSettings_AllowSeparate, pcStoreSettings_DisableDiscountCodes, pcStoreSettings_ReferLabel, pcStoreSettings_ViewRefer, pcStoreSettings_RefNewCheckout, pcStoreSettings_RefNewReg, pcStoreSettings_BrandLogo, pcStoreSettings_BrandPro, pcStoreSettings_RewardsActive, pcStoreSettings_RewardsIncludeWholesale, pcStoreSettings_RewardsPercent, pcStoreSettings_RewardsLabel, pcStoreSettings_RewardsReferral, pcStoreSettings_RewardsFlat, pcStoreSettings_RewardsFlatValue, pcStoreSettings_RewardsPerc, pcStoreSettings_RewardsPercValue, pcStoreSettings_XML, pcStoreSettings_QDiscounttype, pcStoreSettings_BTODisplayType, pcStoreSettings_BTOOutofStockPurchase, pcStoreSettings_BTOShowImage, pcStoreSettings_BTOQuote, pcStoreSettings_BTOQuoteSubmit, pcStoreSettings_BTOQuoteSubmitOnly, pcStoreSettings_BTODetLinkType, pcStoreSettings_BTODetTxt, pcStoreSettings_BTOPopWidth, pcStoreSettings_BTOPopHeight, pcStoreSettings_BTOPopImage, pcStoreSettings_ConfigPurchaseOnly, pcStoreSettings_Terms, pcStoreSettings_TermsLabel, pcStoreSettings_TermsCopy, pcStoreSettings_TermsShown, pcStoreSettings_ShowSKU, pcStoreSettings_ShowSmallImg, pcStoreSettings_HideRMA, pcStoreSettings_ShowHD, pcStoreSettings_StoreUseToolTip, pcStoreSettings_ErrorHandler, pcStoreSettings_AllowCheckoutWR, pcStoreSettings_ViewPrdStyle, pcStoreSettings_CustomerIPAlert, pcStoreSettings_CompanyPhoneNumber, pcStoreSettings_CompanyFaxNumber, pcStoreSettings_DisableGiftRegistry,  pcStoreSettings_SeoURLs, pcStoreSettings_SeoURLs404, pcStoreSettings_QuickBuy, pcStoreSettings_ATCEnabled, pcStoreSettings_RestoreCart, pcStoreSettings_GuestCheckoutOpt, pcStoreSettings_AddThisDisplay, pcStoreSettings_AddThisCode, pcStoreSettings_GoogleAnalytics, pcStoreSettings_MetaTitle, pcStoreSettings_MetaDescription, pcStoreSettings_MetaKeywords, pcStoreSettings_PinterestDisplay, pcStoreSettings_PinterestCounter FROM pcStoreSettings WHERE (((pcStoreSettings_ID)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

pcStrCompanyName=rs("pcStoreSettings_CompanyName")
pcStrCompanyAddress=rs("pcStoreSettings_CompanyAddress")
pcStrCompanyZip=rs("pcStoreSettings_CompanyZip")
pcStrCompanyCity=rs("pcStoreSettings_CompanyCity")
pcStrCompanyState=rs("pcStoreSettings_CompanyState")
pcStrCompanyCountry=rs("pcStoreSettings_CompanyCountry")
pcStrCompanyLogo=rs("pcStoreSettings_CompanyLogo")
pcIntQtyLimit=rs("pcStoreSettings_QtyLimit")
pcIntAddLimit=rs("pcStoreSettings_AddLimit")
pcIntPre=rs("pcStoreSettings_Pre")
pcIntCustPre=rs("pcStoreSettings_CustPre")
pcIntCatImages=rs("pcStoreSettings_CatImages")
pcIntShowStockLmt=rs("pcStoreSettings_ShowStockLmt")
pcIntOutOfStockPurchase=rs("pcStoreSettings_OutOfStockPurchase")
pcStrCurSign=rs("pcStoreSettings_CurSign")
pcStrDecSign=rs("pcStoreSettings_DecSign")
pcStrDivSign=rs("pcStoreSettings_DivSign")
pcStrDateFrmt=rs("pcStoreSettings_DateFrmt")
pcIntMinPurchase=rs("pcStoreSettings_MinPurchase")
pcIntWholesaleMinPurchase=rs("pcStoreSettings_WholesaleMinPurchase")
pcStrURLredirect=rs("pcStoreSettings_URLredirect")
pcStrSSL=rs("pcStoreSettings_SSL")
pcStrSSLUrl=rs("pcStoreSettings_SSLUrl")
pcStrIntSSLPage=rs("pcStoreSettings_IntSSLPage")
pcIntPrdRow=rs("pcStoreSettings_PrdRow")
pcIntPrdRowsPerPage=rs("pcStoreSettings_PrdRowsPerPage")
pcIntCatRow=rs("pcStoreSettings_CatRow")
pcIntCatRowsPerPage=rs("pcStoreSettings_CatRowsPerPage")
pcStrBType=rs("pcStoreSettings_BType")
pcStrStoreOff=rs("pcStoreSettings_StoreOff")
pcStrStoreMsg=rs("pcStoreSettings_StoreMsg")
pcIntWL=rs("pcStoreSettings_WL")
pcIntTF=rs("pcStoreSettings_TF")
pcStrorderLevel=rs("pcStoreSettings_orderLevel")
pcIntDisplayStock=rs("pcStoreSettings_DisplayStock")
pcIntHideCategory=rs("pcStoreSettings_HideCategory")
pcIntAllowNews=rs("pcStoreSettings_AllowNews")
pcIntNewsCheckOut=rs("pcStoreSettings_NewsCheckOut")
pcIntNewsReg=rs("pcStoreSettings_NewsReg")
pcStrNewsLabel=rs("pcStoreSettings_NewsLabel")
pcIntPCOrd=rs("pcStoreSettings_PCOrd")
pcIntHideSortPro=rs("pcStoreSettings_HideSortPro")
pcStrDFLabel=rs("pcStoreSettings_DFLabel")
pcStrDFShow=rs("pcStoreSettings_DFShow")
pcStrDFReq=rs("pcStoreSettings_DFReq")
pcStrTFLabel=rs("pcStoreSettings_TFLabel")
pcStrTFShow=rs("pcStoreSettings_TFShow")
pcStrTFReq=rs("pcStoreSettings_TFReq")
pcStrDTCheck=rs("pcStoreSettings_DTCheck")
pcStrDeliveryZip=rs("pcStoreSettings_DeliveryZip")
pcStrOrderName=rs("pcStoreSettings_OrderName")
pcStrHideDiscField=rs("pcStoreSettings_HideDiscField")
pcStrAllowSeparate=rs("pcStoreSettings_AllowSeparate")
pcIntDisableDiscountCodes=rs("pcStoreSettings_DisableDiscountCodes")
pcStrReferLabel=rs("pcStoreSettings_ReferLabel")
pcIntViewRefer=rs("pcStoreSettings_ViewRefer")
pcIntRefNewCheckout=rs("pcStoreSettings_RefNewCheckout")
pcIntRefNewReg=rs("pcStoreSettings_RefNewReg")
pcIntBrandLogo=rs("pcStoreSettings_BrandLogo")
pcIntBrandPro=rs("pcStoreSettings_BrandPro")
pcIntRewardsActive=rs("pcStoreSettings_RewardsActive")
pcIntRewardsIncludeWholesale=rs("pcStoreSettings_RewardsIncludeWholesale")
pcIntRewardsPercent=rs("pcStoreSettings_RewardsPercent")
pcStrRewardsLabel=rs("pcStoreSettings_RewardsLabel")
pcIntRewardsReferral=rs("pcStoreSettings_RewardsReferral")
pcIntRewardsFlat=rs("pcStoreSettings_RewardsFlat")
pcIntRewardsFlatValue=rs("pcStoreSettings_RewardsFlatValue")
pcIntRewardsPerc=rs("pcStoreSettings_RewardsPerc")
pcIntRewardsPercValue=rs("pcStoreSettings_RewardsPercValue")
pcStrXML=rs("pcStoreSettings_XML")
if pcStrXML="" then
	pcStrXML=scXML
end if
pcIntQDiscounttype=rs("pcStoreSettings_QDiscountType")
pcIntBTODisplayType=rs("pcStoreSettings_BTODisplayType")
pcIntBTOOutofStockPurchase=rs("pcStoreSettings_BTOOutofStockPurchase")
pcIntBTOShowImage=rs("pcStoreSettings_BTOShowImage")
pcIntBTOQuote=rs("pcStoreSettings_BTOQuote")
pcIntBTOQuoteSubmit=rs("pcStoreSettings_BTOQuoteSubmit")
pcIntBTOQuoteSubmitOnly=rs("pcStoreSettings_BTOQuoteSubmitOnly")
pcIntBTODetLinkType=rs("pcStoreSettings_BTODetLinkType")
pcStrBTODetTxt=rs("pcStoreSettings_BTODetTxt")
pcIntBTOPopWidth=rs("pcStoreSettings_BTOPopWidth")
pcIntBTOPopHeight=rs("pcStoreSettings_BTOPopHeight")
pcIntBTOPopImage=rs("pcStoreSettings_BTOPopImage")
pcIntConfigPurchaseOnly=rs("pcStoreSettings_ConfigPurchaseOnly")
pcIntTerms=rs("pcStoreSettings_Terms")
pcStrTermsLabel=rs("pcStoreSettings_TermsLabel")
pcStrTermsCopy=rs("pcStoreSettings_TermsCopy")
pcIntTermsShown=rs("pcStoreSettings_TermsShown")
pcIntShowSKU=rs("pcStoreSettings_ShowSKU")
pcIntShowSmallImg=rs("pcStoreSettings_ShowSmallImg")
pcIntHideRMA=rs("pcStoreSettings_HideRMA")
pcIntShowHD=rs("pcStoreSettings_ShowHD")
pcIntStoreUseToolTip=rs("pcStoreSettings_StoreUseToolTip")
pcIntErrorHandler=rs("pcStoreSettings_ErrorHandler")
pcIntAllowCheckoutWR=rs("pcStoreSettings_AllowCheckoutWR")
pcStrViewPrdStyle=rs("pcStoreSettings_ViewPrdStyle")
pcStrCustomerIPAlert=rs("pcStoreSettings_CustomerIPAlert")
pcStrCompanyPhoneNumber=rs("pcStoreSettings_CompanyPhoneNumber")
pcStrCompanyFaxNumber=rs("pcStoreSettings_CompanyFaxNumber")
pcIntDisableGiftRegistry=rs("pcStoreSettings_DisableGiftRegistry")
pcIntSeoURLs=rs("pcStoreSettings_SeoURLs")
pcStrSeoURLs404=rs("pcStoreSettings_SeoURLs404")
pcIntQuickBuy=rs("pcStoreSettings_QuickBuy")
pcIntATCEnabled=rs("pcStoreSettings_ATCEnabled")
pcIntRestoreCart=rs("pcStoreSettings_RestoreCart")
pcIntGuestCheckoutOpt=rs("pcStoreSettings_GuestCheckoutOpt")
pcIntAddThisDisplay=rs("pcStoreSettings_AddThisDisplay")
pcStrAddThisCode=rs("pcStoreSettings_AddThisCode")
pcStrGoogleAnalytics=rs("pcStoreSettings_GoogleAnalytics")
pcStrMetaTitle=rs("pcStoreSettings_MetaTitle")
pcStrMetaDescription=rs("pcStoreSettings_MetaDescription")
pcStrMetaKeywords=rs("pcStoreSettings_MetaKeywords")
pcIntPinterestDisplay=rs("pcStoreSettings_PinterestDisplay")
if pcIntPinterestDisplay&""="" then
	pcIntPinterestDisplay="0"
end if
pcStrPinterestCounter=rs("pcStoreSettings_PinterestCounter")
pcStrScRegistered = scRegistered
set rs=nothing

Call Closedb()
%>
