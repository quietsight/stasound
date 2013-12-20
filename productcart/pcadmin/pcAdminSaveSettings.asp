<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/status.inc"-->
<!--#include file="../includes/statusAPP.inc"-->
<!--#include file="../includes/statusCM.inc"-->
<!--#include file="../includes/statusM.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% 
if pcStrPageName<>"RPSettings.asp" then
	if RewardsPercent<>0 then
		pcIntRewardsPercent=RewardsPercent
	end if
	if RewardsPerc = 1 then
		if RewardsPercValue<>0 then
			pcIntRewardsPercValue=RewardsPercValue
		end if
	end if
end if

Call Opendb()

'// START ProductCart Sub-Version

	'//BTO vs STD version
	pcIntBTO=statusBTO
	if len(pcIntBTO)<1 then
		pcIntBTO=0
	end if
	'// Add if missing
	if pcIntBTO=1 and InStr(pcStrScVersion,"b")=0 then
		pcStrScVersion=pcStrScVersion & "b"
	end if
	'// Remove if not needed
	if pcIntBTO=0 and InStr(pcStrScVersion,"b")=1 then
		pcStrScVersion=replace(pcStrScVersion, "b", "")
	end if
	
	'//Apparel Add-on status
	pcIntAPP=statusAPP
	if len(pcIntAPP)<1 then
		pcIntAPP=0
	end if
	'// Add if missing
	if pcIntAPP=1 and InStr(pcStrScSubVersion,"a")=0 then
		pcStrScSubVersion=pcStrScSubVersion & "a"
	end if
	'// Remove if not needed
	if pcIntAPP=0 and InStr(pcStrScSubVersion,"a")=1 then
		pcStrScSubVersion=replace(pcStrScSubVersion, "a", "")
	end if
	
	'//Conflict Management Add-on status
	pcIntCM=statusCM
	if len(pcIntCM)<1 then
		pcIntCM=0
	end if
	'// Add if missing
	if pcIntCM=1 and InStr(pcStrScSubVersion,"cm")=0 then
		pcStrScSubVersion=pcStrScSubVersion & "cm"
	end if
	'// Remove if not needed
	if pcIntCM=0 and InStr(pcStrScSubVersion,"cm")=1 then
		pcStrScSubVersion=replace(pcStrScSubVersion, "cm", "")
	end if
	
	'//Mobile Commerce Add-on status
	'// Ms = Mobile storefront
	pcIntMS=statusM
	if len(pcIntMS)<1 then
		pcIntMS=0
	end if
	'// Add if missing
	if pcIntMS=1 and InStr(pcStrScSubVersion,"Ms")=0 then
		pcStrScSubVersion=pcStrScSubVersion & "Ms"
	end if
	'// Remove if not needed
	if pcIntMS=0 and InStr(pcStrScSubVersion,"Ms")=1 then
		pcStrScSubVersion=replace(pcStrScSubVersion, "Ms", "")
	end if
	
	'// Sub-version clean up (from older versions)
		pcStrScSubVersion=replace(pcStrScSubVersion, "p2", "")
		pcStrScSubVersion=replace(pcStrScSubVersion, "p", "")
		pcStrScSubVersion=replace(pcStrScSubVersion, "g131", "")
		pcStrScSubVersion=replace(pcStrScSubVersion, "g13", "")
		pcStrScSubVersion=replace(pcStrScSubVersion, "a171", "a")
		pcStrScSubVersion=replace(pcStrScSubVersion, "a17", "a")
		pcStrScSubVersion=replace(pcStrScSubVersion, "a3", "a")
		pcStrScSubVersion=replace(pcStrScSubVersion, "aa", "a")
		pcStrScSubVersion=replace(pcStrScSubVersion, "cmcm", "cm")
		pcStrScSubVersion=replace(pcStrScSubVersion, "MsMs", "Ms")

'// END ProductCart Sub-Version	

query = "UPDATE pcStoreVersions SET pcStoreVersion_Num='"&removeReplaceSQ(pcStrScVersion)&"', pcStoreVersion_Sub='"&removeReplaceSQ(pcStrScSubVersion)&"' ,pcStoreVersion_SP="&removeReplaceSQ(pcStrScSP)&" WHERE pcStoreVersion_ID=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

set rs=nothing

query="UPDATE pcStoreSettings SET pcStoreSettings_CompanyName='"&removeReplaceSQ(pcStrCompanyName)&"', pcStoreSettings_CompanyAddress='"&removeReplaceSQ(pcStrCompanyAddress)&"', pcStoreSettings_CompanyZip='"&removeReplaceSQ(pcStrCompanyZip)&"', pcStoreSettings_CompanyCity='"&removeReplaceSQ(pcStrCompanyCity)&"', pcStoreSettings_CompanyState='"&removeReplaceSQ(pcStrCompanyState)&"', pcStoreSettings_CompanyCountry='"&removeReplaceSQ(pcStrCompanyCountry)&"', pcStoreSettings_CompanyLogo='"&removeReplaceSQ(pcStrCompanyLogo)&"', pcStoreSettings_QtyLimit="&pcIntQtyLimit&", pcStoreSettings_AddLimit="&pcIntAddLimit&", pcStoreSettings_Pre="&pcIntPre&", pcStoreSettings_CustPre="&pcIntCustPre&", pcStoreSettings_CatImages="&pcIntCatImages&", pcStoreSettings_ShowStockLmt="&pcIntShowStockLmt&", pcStoreSettings_OutOfStockPurchase="&pcIntOutOfStockPurchase&", pcStoreSettings_Cursign='"&removeReplaceSQ(pcStrCurSign)&"', pcStoreSettings_DecSign='"&removeReplaceSQ(pcStrDecSign)&"', pcStoreSettings_DivSign='"&removeReplaceSQ(pcStrDivSign)&"', pcStoreSettings_DateFrmt='"&removeReplaceSQ(pcStrDateFrmt)&"', pcStoreSettings_MinPurchase="&pcIntMinPurchase&", pcStoreSettings_WholesaleMinPurchase="&pcIntWholesaleMinPurchase&", pcStoreSettings_URLredirect='"&removeReplaceSQ(pcStrURLredirect)&"', pcStoreSettings_SSL='"&removeReplaceSQ(pcStrSSL)&"', pcStoreSettings_SSLUrl='"&removeReplaceSQ(pcStrSSLUrl)&"', pcStoreSettings_IntSSLPage='"&removeReplaceSQ(pcStrIntSSLPage)&"', pcStoreSettings_PrdRow="&pcIntPrdRow&", pcStoreSettings_PrdRowsPerPage="&pcIntPrdRowsPerPage&",  pcStoreSettings_CatRow="&pcIntCatRow&", pcStoreSettings_CatRowsPerPage="&pcIntCatRowsPerPage&", pcStoreSettings_BType='"&removeReplaceSQ(pcStrBType)&"', pcStoreSettings_StoreOff='"&removeReplaceSQ(pcStrStoreOff)&"', pcStoreSettings_StoreMsg='"&removeReplaceSQ(pcStrStoreMsg)&"', pcStoreSettings_WL="&pcIntWL&", pcStoreSettings_TF="&pcIntTF&", pcStoreSettings_orderLevel='"&removeReplaceSQ(pcStrorderLevel)&"', pcStoreSettings_DisplayStock="&pcIntDisplayStock&", pcStoreSettings_HideCategory="&pcIntHideCategory&", pcStoreSettings_AllowNews="&pcIntAllowNews&", pcStoreSettings_NewsCheckOut="&pcIntNewsCheckOut&", pcStoreSettings_NewsReg="&pcIntNewsReg&", pcStoreSettings_NewsLabel='"&removeReplaceSQ(pcStrNewsLabel)&"', pcStoreSettings_PCOrd="&pcIntPCOrd&", pcStoreSettings_HideSortPro="&pcIntHideSortPro&", pcStoreSettings_DFLabel='"&removeReplaceSQ(pcStrDFLabel)&"', pcStoreSettings_DFShow='"&removeReplaceSQ(pcStrDFShow)&"', pcStoreSettings_DFReq='"&removeReplaceSQ(pcStrDFReq)&"', pcStoreSettings_TFLabel='"&removeReplaceSQ(pcStrTFLabel)&"', pcStoreSettings_TFShow='"&removeReplaceSQ(pcStrTFShow)&"', pcStoreSettings_TFReq='"&removeReplaceSQ(pcStrTFReq)&"', pcStoreSettings_DTCheck='"&removeReplaceSQ(pcStrDTCheck)&"', pcStoreSettings_DeliveryZip='"&removeReplaceSQ(pcStrDeliveryZip)&"', pcStoreSettings_OrderName='"&removeReplaceSQ(pcStrOrderName)&"', pcStoreSettings_HideDiscField='"&removeReplaceSQ(pcStrHideDiscField)&"', pcStoreSettings_AllowSeparate='"&removeReplaceSQ(pcStrAllowSeparate)&"', pcStoreSettings_ReferLabel='"&removeReplaceSQ(pcStrReferLabel)&"', pcStoreSettings_ViewRefer="&pcIntViewRefer&", pcStoreSettings_RefNewCheckout="&pcIntRefNewCheckout&", pcStoreSettings_RefNewReg="&pcIntRefNewReg&", pcStoreSettings_BrandLogo="&pcIntBrandLogo&", pcStoreSettings_BrandPro="&pcIntBrandPro&", pcStoreSettings_RewardsActive="&pcIntRewardsActive&", pcStoreSettings_RewardsIncludeWholesale="&pcIntRewardsIncludeWholesale&", pcStoreSettings_RewardsPercent="&pcIntRewardsPercent&", pcStoreSettings_RewardsLabel='"&removeReplaceSQ(pcStrRewardsLabel)&"', pcStoreSettings_RewardsReferral="&pcIntRewardsReferral&", pcStoreSettings_RewardsFlat="&pcIntRewardsFlat&", pcStoreSettings_RewardsFlatValue="&pcIntRewardsFlatValue&", pcStoreSettings_RewardsPerc="&pcIntRewardsPerc&", pcStoreSettings_RewardsPercValue="&pcIntRewardsPercValue&", pcStoreSettings_XML='"&removeReplaceSQ(pcStrXML)&"', pcStoreSettings_QDiscounttype="&pcIntQDiscountType&", pcStoreSettings_BTODisplayType="&pcIntBTODisplayType&", pcStoreSettings_BTOOutofStockPurchase="&pcIntBTOOutofStockPurchase&", pcStoreSettings_BTOShowImage="&pcIntBTOShowImage&", pcStoreSettings_BTOQuote="&pcIntBTOQuote&", pcStoreSettings_BTOQuoteSubmit="&pcIntBTOQuoteSubmit&", pcStoreSettings_BTOQuoteSubmitOnly="&pcIntBTOQuoteSubmitOnly&", pcStoreSettings_BTODetLinkType="&pcIntBTODetLinkType&", pcStoreSettings_BTODetTxt='"&removeReplaceSQ(pcStrBTODetTxt)&"', pcStoreSettings_BTOPopWidth="&pcIntBTOPopWidth&", pcStoreSettings_BTOPopHeight="&pcIntBTOPopHeight&", pcStoreSettings_BTOPopImage="&pcIntBTOPopImage&", pcStoreSettings_ConfigPurchaseOnly="&pcIntConfigPurchaseOnly&", pcStoreSettings_Terms="&pcIntTerms&", pcStoreSettings_TermsLabel='"&removeReplaceSQ(pcStrTermsLabel)&"', pcStoreSettings_TermsCopy='"&removeReplaceSQ(pcStrTermsCopy)&"', pcStoreSettings_TermsShown="&pcIntTermsShown&", pcStoreSettings_ShowSKU="&pcIntShowSKU&", pcStoreSettings_ShowSmallImg="&pcIntShowSmallImg&", pcStoreSettings_HideRMA="&pcIntHideRMA&", pcStoreSettings_ShowHD="&pcIntShowHD&", pcStoreSettings_StoreUseToolTip="&pcIntStoreUseToolTip&", pcStoreSettings_ErrorHandler="&pcIntErrorHandler&", pcStoreSettings_AllowCheckoutWR="&pcIntAllowCheckoutWR&", pcStoreSettings_ViewPrdStyle='"&removeReplaceSQ(pcStrViewPrdStyle)&"', pcStoreSettings_CustomerIPAlert='"&removeReplaceSQ(pcStrCustomerIPAlert)&"', pcStoreSettings_CompanyPhoneNumber='"&removeReplaceSQ(pcStrCompanyPhoneNumber)&"', pcStoreSettings_CompanyFaxNumber='"&removeReplaceSQ(pcStrCompanyFaxNumber)&"', pcStoreSettings_DisableGiftRegistry="&pcIntDisableGiftRegistry&", pcStoreSettings_DisableDiscountCodes="&pcIntDisableDiscountCodes&", pcStoreSettings_SeoURLs="&pcIntSeoURLs&", pcStoreSettings_SeoURLs404='"&removeReplaceSQ(pcStrSeoURLs404)&"', pcStoreSettings_QuickBuy="&pcIntQuickBuy&", pcStoreSettings_ATCEnabled="&pcIntATCEnabled&", pcStoreSettings_RestoreCart="&pcIntRestoreCart&",pcStoreSettings_GuestCheckoutOpt=" & pcIntGuestCheckoutOpt &",pcStoreSettings_AddThisDisplay="&pcIntAddThisDisplay&",pcStoreSettings_AddThisCode='"&removeReplaceSQ(pcStrAddThisCode)&"',pcStoreSettings_PinterestDisplay="&pcIntPinterestDisplay&", pcStoreSettings_PinterestCounter='"&removeReplaceSQ(pcStrPinterestCounter)&"',pcStoreSettings_GoogleAnalytics='"&removeReplaceSQ(pcStrGoogleAnalytics)&"', pcStoreSettings_MetaTitle='"&removeReplaceSQ(pcStrMetaTitle)&"', pcStoreSettings_MetaDescription='"&removeReplaceSQ(pcStrMetaDescription)&"', pcStoreSettings_MetaKeywords='"&removeReplaceSQ(pcStrMetaKeywords)&"' WHERE (((pcStoreSettings_ID)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

set rs=nothing

Call Closedb()
	
'/////////////////////////////////////////////////////
'// Write all changes to Settings.asp file
'/////////////////////////////////////////////////////
Dim objFS
Dim objFile

Set objFS = Server.CreateObject ("Scripting.FileSystemObject")

if PPD="1" then
	pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/settings.asp")
else
	pcStrFileName=Server.Mappath ("../includes/settings.asp")
end if

Set objFile = objFS.OpenTextFile (pcStrFileName, 2, True, 0)
objFile.WriteLine CHR(60)&CHR(37)&"'// Storewide Settings //" & vbCrLf
objFile.WriteLine "private const scVersion = """&pcStrScVersion&"""" & vbCrLf
objFile.WriteLine "private const scSubVersion = """&pcStrScSubVersion&"""" & vbCrLf
objFile.WriteLine "private const scSP = """&pcStrScSP&"""" & vbCrLf
objFile.WriteLine "private const scRegistered = """&pcStrScRegistered&"""" & vbCrLf
objFile.WriteLine "private const scCompanyName = """&removeSQ(pcStrCompanyName)&"""" & vbCrLf
objFile.WriteLine "private const scCompanyAddress = """&removeSQ(pcStrCompanyAddress)&"""" & vbCrLf
objFile.WriteLine "private const scCompanyZip = """&removeSQ(pcStrCompanyZip)&"""" & vbCrLf
objFile.WriteLine "private const scCompanyCity = """&removeSQ(pcStrCompanyCity)&"""" & vbCrLf
objFile.WriteLine "private const scCompanyState = """&removeSQ(pcStrCompanyState)&"""" & vbCrLf
objFile.WriteLine "private const scCompanyCountry = """&removeSQ(pcStrCompanyCountry)&"""" & vbCrLf
objFile.WriteLine "private const scCompanyPhoneNumber = """&removeSQ(pcStrCompanyPhoneNumber)&"""" & vbCrLf
objFile.WriteLine "private const scCompanyFaxNumber = """&removeSQ(pcStrCompanyFaxNumber)&"""" & vbCrLf
objFile.WriteLine "private const scCompanyLogo = """&removeSQ(pcStrCompanyLogo)&"""" & vbCrLf
objFile.WriteLine "private const scMetaTitle = """&removeSQ(pcStrMetaTitle)&"""" & vbCrLf
objFile.WriteLine "private const scMetaDescription = """&removeSQ(pcStrMetaDescription)&"""" & vbCrLf
objFile.WriteLine "private const scMetaKeywords = """&removeSQ(pcStrMetaKeywords)&"""" & vbCrLf
objFile.WriteLine "private const scQtyLimit = "&pcIntQtyLimit&"" & vbCrLf
objFile.WriteLine "private const scAddLimit = "&pcIntAddLimit&"" & vbCrLf
objFile.WriteLine "private const scPre = "&pcIntPre&"" & vbCrLf
objFile.WriteLine "private const scCustPre = "&pcIntCustPre&"" & vbCrLf
objFile.WriteLine "private const scBTO = "&pcIntBTO&"" & vbCrLf
objFile.WriteLine "private const scAPP = "&pcIntAPP&"" & vbCrLf
objFile.WriteLine "private const scCM = "&pcIntCM&"" & vbCrLf
objFile.WriteLine "private const scMS = "&pcIntMS&"" & vbCrLf
objFile.WriteLine "private const scCatImages = "&pcIntCatImages&"" & vbCrLf
objFile.WriteLine "private const scShowStockLmt = "&pcIntShowStockLmt&"" & vbCrLf
objFile.WriteLine "private const scOutOfStockPurchase = "&pcIntOutOfStockPurchase&"" & vbCrLf
objFile.WriteLine "private const scCurSign = """&removeSQ(pcStrCurSign)&"""" & vbCrLf
objFile.WriteLine "private const scDecSign = """&removeSQ(pcStrDecSign)&"""" & vbCrLf
objFile.WriteLine "private const scDivSign = """&removeSQ(pcStrDivSign)&"""" & vbCrLf
objFile.WriteLine "private const scDateFrmt = """&removeSQ(pcStrDateFrmt)&"""" & vbCrLf
objFile.WriteLine "private const scMinPurchase = "&pcIntMinPurchase&"" & vbCrLf
objFile.WriteLine "private const scWholesaleMinPurchase = "&pcIntWholesaleMinPurchase&"" & vbCrLf
objFile.WriteLine "private const scURLredirect = """&removeSQ(pcStrURLredirect)&"""" & vbCrLf
objFile.WriteLine "private const scSSL = """&removeSQ(pcStrSSL)&"""" & vbCrLf
objFile.WriteLine "private const scSSLUrl = """&removeSQ(pcStrSSLUrl)&"""" & vbCrLf
objFile.WriteLine "private const scIntSSLPage = """&removeSQ(pcStrIntSSLPage)&"""" & vbCrLf
objFile.WriteLine "private const scPrdRow = "&pcIntPrdRow&"" & vbCrLf
objFile.WriteLine "private const scPrdRowsPerPage = "&pcIntPrdRowsPerPage&"" & vbCrLf
objFile.WriteLine "private const scCatRow = "&pcIntCatRow&"" & vbCrLf
objFile.WriteLine "private const scCatRowsPerPage = "&pcIntCatRowsPerPage&"" & vbCrLf
objFile.WriteLine "private const bType = """&removeSQ(pcStrBType)&"""" & vbCrLf
objFile.WriteLine "private const scStoreOff = """&removeSQ(pcStrStoreOff)&"""" & vbCrLf
objFile.WriteLine "private const scStoreMsg = """&removeSQ(pcStrStoreMsg)&"""" & vbCrLf
objFile.WriteLine "private const scWL = "&pcIntWL&"" & vbCrLf
objFile.WriteLine "private const scTF = "&pcIntTF&"" & vbCrLf
objFile.WriteLine "private const scorderLevel = """&removeSQ(pcStrorderLevel)&"""" & vbCrLf
objFile.WriteLine "private const scDisplayStock = "&pcIntDisplayStock&"" & vbCrLf
objFile.WriteLine "private const scHideCategory = "&pcIntHideCategory&"" & vbCrLf
objFile.WriteLine "private const AllowNews = "&pcIntAllowNews&"" & vbCrLf
objFile.WriteLine "private const NewsCheckOut = "&pcIntNewsCheckOut&"" & vbCrLf
objFile.WriteLine "private const NewsReg = "&pcIntNewsReg&"" & vbCrLf
objFile.WriteLine "private const NewsLabel = """&removeSQ(pcStrNewsLabel)&"""" & vbCrLf
objFile.WriteLine "private const PCOrd = "&pcIntPCOrd&"" & vbCrLf
objFile.WriteLine "private const HideSortPro = "&pcIntHideSortPro&"" & vbCrLf
objFile.WriteLine "private const scViewPrdStyle = """&removeSQ(pcStrViewPrdStyle)&"""" & vbCrLf
objFile.WriteLine "private const DFLabel = """&removeSQ(pcStrDFLabel)&"""" & vbCrLf
objFile.WriteLine "private const DFShow = """&removeSQ(pcStrDFShow)&"""" & vbCrLf
objFile.WriteLine "private const DFReq = """&removeSQ(pcStrDFReq)&"""" & vbCrLf
objFile.WriteLine "private const TFLabel = """&removeSQ(pcStrTFLabel)&"""" & vbCrLf
objFile.WriteLine "private const TFShow = """&removeSQ(pcStrTFShow)&"""" & vbCrLf
objFile.WriteLine "private const TFReq = """&removeSQ(pcStrTFReq)&"""" & vbCrLf
objFile.WriteLine "private const DTCheck = """&removeSQ(pcStrDTCheck)&"""" & vbCrLf
objFile.WriteLine "private const DeliveryZip = """&removeSQ(pcStrDeliveryZip)&"""" & vbCrLf
objFile.WriteLine "private const CustomerIPAlert = """&removeSQ(pcStrCustomerIPAlert)&"""" & vbCrLf
objFile.WriteLine "private const scOrderName = """&removeSQ(pcStrOrderName)&"""" & vbCrLf
objFile.WriteLine "private const scHideDiscField = """&removeSQ(pcStrHideDiscField)&"""" & vbCrLf
objFile.WriteLine "private const scAllowSeparate = """&removeSQ(pcStrAllowSeparate)&"""" & vbCrLf
objFile.WriteLine "private const ReferLabel = """&removeSQ(pcStrReferLabel)&"""" & vbCrLf
objFile.WriteLine "private const ViewRefer = "&pcIntViewRefer&"" & vbCrLf
objFile.WriteLine "private const RefNewCheckout = "&pcIntRefNewCheckout&"" & vbCrLf
objFile.WriteLine "private const RefNewReg = "&pcIntRefNewReg&"" & vbCrLf
objFile.WriteLine "private const sBrandLogo = "&pcIntBrandLogo&"" & vbCrLf
objFile.WriteLine "private const sBrandPro = "&pcIntBrandPro&"" & vbCrLf
objFile.WriteLine "private const RewardsActive = "&pcIntRewardsActive&"" & vbCrLf
objFile.WriteLine "private const RewardsIncludeWholesale = "&pcIntRewardsIncludeWholesale&"" & vbCrLf
objFile.WriteLine "private const RewardsPercent = "&pcIntRewardsPercent&"" & vbCrLf
objFile.WriteLine "private const RewardsLabel = """&removeSQ(pcStrRewardsLabel)&"""" & vbCrLf
objFile.WriteLine "private const RewardsReferral = "&pcIntRewardsReferral&"" & vbCrLf
objFile.WriteLine "private const RewardsFlat = "&pcIntRewardsFlat&"" & vbCrLf
objFile.WriteLine "private const RewardsFlatValue = "&pcIntRewardsFlatValue&"" & vbCrLf
objFile.WriteLine "private const RewardsPerc = "&pcIntRewardsPerc&"" & vbCrLf
objFile.WriteLine "private const RewardsPercValue = "&pcIntRewardsPercValue&"" & vbCrLf
objFile.WriteLine "private const pcQDiscountType = "&pcIntQDiscountType&"" & vbCrLf
objFile.WriteLine "private const iBTODisplayType = "&pcIntBTODisplayType&"" & vbCrLf
objFile.WriteLine "private const iBTOOutofStockPurchase = "&pcIntBTOOutofStockPurchase&"" & vbCrLf
objFile.WriteLine "private const iBTOShowImage = "&pcIntBTOShowImage&"" & vbCrLf
objFile.WriteLine "private const iBTOQuote = "&pcIntBTOQuote&"" & vbCrLf
objFile.WriteLine "private const iBTOQuoteSubmit = "&pcIntBTOQuoteSubmit&"" & vbCrLf
objFile.WriteLine "private const iBTOQuoteSubmitOnly = "&pcIntBTOQuoteSubmitOnly&"" & vbCrLf
objFile.WriteLine "private const iBTODetLinkType = "&pcIntBTODetLinkType&"" & vbCrLf
objFile.WriteLine "private const vBTODetTxt = """&removeSQ(pcStrBTODetTxt)&"""" & vbCrLf
objFile.WriteLine "private const iBTOPopWidth = "&pcIntBTOPopWidth&"" & vbCrLf
objFile.WriteLine "private const iBTOPopHeight = "&pcIntBTOPopHeight&"" & vbCrLf
objFile.WriteLine "private const iBTOPopImage = "&pcIntBTOPopImage&"" & vbCrLf
objFile.WriteLine "private const scConfigPurchaseOnly = "&pcIntConfigPurchaseOnly&"" & vbCrLf
objFile.WriteLine "private const scTerms = "&pcIntTerms&"" & vbCrLf
objFile.WriteLine "private const scTermsShown = "&pcIntTermsShown&"" & vbCrLf
objFile.WriteLine "private const scShowSKU = "&pcIntShowSKU&"" & vbCrLf
objFile.WriteLine "private const scShowSmallImg = "&pcIntShowSmallImg&"" & vbCrLf
objFile.WriteLine "private const scHideRMA = "&pcIntHideRMA&"" & vbCrLf
objFile.WriteLine "private const scShowHD = "&pcIntShowHD&"" & vbCrLf
objFile.WriteLine "private const scStoreUseToolTip = "&pcIntStoreUseToolTip&"" & vbCrLf
objFile.WriteLine "private const scErrorHandler = "&pcIntErrorHandler&"" & vbCrLf
objFile.WriteLine "private const scDisableGiftRegistry = """&pcIntDisableGiftRegistry&"""" & vbCrLf
objFile.WriteLine "private const scDisableDiscountCodes = """&pcIntDisableDiscountCodes&"""" & vbCrLf
objFile.WriteLine "private const scAllowCheckoutWR = "&pcIntAllowCheckoutWR&"" & vbCrLf
objFile.WriteLine "private const scSeoURLs = "&pcIntSeoURLs&"" & vbCrLf
objFile.WriteLine "private const scSeoURLs404 = """&removeSQ(pcStrSeoURLs404)&"""" & vbCrLf
objFile.WriteLine "private const scQuickBuy = "&pcIntQuickBuy&"" & vbCrLf
objFile.WriteLine "private const scATCEnabled = "&pcIntATCEnabled&"" & vbCrLf
objFile.WriteLine "private const scRestoreCart = "&pcIntRestoreCart&"" & vbCrLf
objFile.WriteLine "private const scXML = """&removeSQ(pcStrXML)&"""" & vbCrLf
objFile.WriteLine "private const scGuestCheckoutOpt = "&pcIntGuestCheckoutOpt&"" & vbCrLf
objFile.WriteLine "private const scAddThisDisplay = "&pcIntAddThisDisplay&"" & vbCrLf
objFile.WriteLine "private const scPinterestDisplay = "&pcIntPinterestDisplay&"" & vbCrLf
objFile.WriteLine "private const scPinterestCounter = """&removeSQ(pcStrPinterestCounter)&"""" & vbCrLf
objFile.WriteLine "private const scGoogleAnalytics = """&removeSQ(pcStrGoogleAnalytics)&"""" & vbCrLf
objFile.WriteLine "'// Storewide Settings // " &CHR(37)&CHR(62)& vbCrLf
objFile.Close
set objFS=nothing
set objFile=nothing
%>
