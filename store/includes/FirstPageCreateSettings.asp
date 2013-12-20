<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="adminv.asp"-->
<!--#include file="storeconstants.asp"-->
<!--#include file="opendb.asp"-->
<!--#include file="secureadminfolder.asp"-->
<!--#include file="status.inc"-->
<%
Dim PageName, Body, q, findit

' request values
q=Chr(34)
PageName="settings.asp"
findit=Server.MapPath(PageName)
If statusBTO="1" then
	tVersion="4.7b"
Else
	tVersion="4.7"
End If
tSubVersion=""
tSP="0"

if tSP&""="" then
	tSP="0"
end if

'// Enter values for version into DB
dim conntemp, rs, query

call opendb()
query="INSERT INTO pcStoreVersions (pcStoreVersion_Num, pcStoreVersion_Sub, pcStoreVersion_SP) VALUES ('"&tVersion&"', '"&tSubVersion&"', "&tSP&");"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
set rs = nothing
call closedb()

Body=CHR(60)&CHR(37)&"private const scVersion="&q&tVersion&q&CHR(10)
Body=Body & "private const scSubVersion="&q&tSubVersion&q&CHR(10)
Body=Body & "private const scSP="&q&tSP&q&CHR(10)
Body=Body & "private const scRegistered="&q&"124588"&q&CHR(10)
Body=Body & "private const scCompanyName="&q&""&q&CHR(10)
Body=Body & "private const scCompanyAddress="&q&""&q&CHR(10)
Body=Body & "private const scCompanyZip="&q&""&q&CHR(10)
Body=Body & "private const scCompanyCity="&q&""&q&CHR(10)
Body=Body & "private const scCompanyState="&q&"CA"&q&CHR(10)
Body=Body & "private const scCompanyCountry="&q&"US"&q&CHR(10)
Body=Body & "private const scCompanyLogo="&q&"yourlogohere.gif"&q&CHR(10)
Body=Body & "private const scMetaTitle="&q&""&q&CHR(10)
Body=Body & "private const scMetaDescription="&q&""&q&CHR(10)
Body=Body & "private const scMetaKeywords="&q&""&q&CHR(10)
Body=Body & "private const scQtyLimit=10"&CHR(10)
Body=Body & "private const scAddLimit=20"&CHR(10)
Body=Body & "private const scPre=0"&CHR(10)
Body=Body & "private const scCustPre=0"&CHR(10)
Body=Body & "private const scBTO="&statusBTO&CHR(10)
Body=Body & "private const scAPP=0"&CHR(10)
Body=Body & "private const scCM=0"&CHR(10)
Body=Body & "private const scMS=0"&CHR(10)
Body=Body & "private const scCatImages=1"&CHR(10)
Body=Body & "private const scShowStockLmt=-1"&CHR(10)
Body=Body & "private const scOutofstockpurchase=0"&CHR(10)
Body=Body & "private const scCurSign="&q&"$"&q&CHR(10)
Body=Body & "private const scDecSign="&q&"."&q&CHR(10)
Body=Body & "private const scDivSign="&q&","&q&CHR(10)
Body=Body & "private const scDateFrmt="&q&"MM/DD/YY"&q&CHR(10)
Body=Body & "private const scMinPurchase=0"&CHR(10)
Body=Body & "private const scWholesaleMinPurchase=0"&CHR(10)
Body=Body & "private const scURLredirect="&q&q&CHR(10)
Body=Body & "private const scSSL="&q&"0"&q&CHR(10)
Body=Body & "private const scSslURL="&q&q&CHR(10)
Body=Body & "private const scIntSSLPage="&q&q&CHR(10)
Body=Body & "private const scPrdRow=3"&CHR(10)
Body=Body & "private const scPrdRowsPerPage=2"&CHR(10)
Body=Body & "private const scCatRow=3"&CHR(10)
Body=Body & "private const scCatRowsPerPage=2"&CHR(10)
Body=Body & "private const bType="&q&"h"&q&CHR(10)
Body=Body & "private const scStoreOff="&q&"0"&q&CHR(10)
Body=Body & "private const scStoreMsg="&q&"This store has been temporarily turned off."&q&CHR(10)
Body=Body & "private const scWL=-1"&CHR(10)
Body=Body & "private const scTF=-1"&CHR(10)
Body=Body & "private const scorderlevel="&q&"0"&q&CHR(10)
Body=Body & "private const scdisplayStock=0"&CHR(10)
Body=Body & "private const schideCategory=0"&CHR(10)
Body=Body & "private const AllowNews=0" & CHR(10)
Body=Body & "private const NewsCheckOut=0" & CHR(10)
Body=Body & "private const NewsReg=0" & CHR(10)
Body=Body & "private const NewsLabel=" & q & q & CHR(10)
Body=Body & "private const PCOrd=0" & CHR(10)
Body=Body & "private const HideSortPro=0" & CHR(10)
Body=Body & "private const DFLabel=" & q & q & CHR(10)
Body=Body & "private const DFShow="&q&"0"& q & CHR(10)
Body=Body & "private const DFReq="&q&"0"& q & CHR(10)
Body=Body & "private const TFLabel=" & q & q & CHR(10)
Body=Body & "private const TFShow="&q&"0"& q & CHR(10)
Body=Body & "private const TFReq="&q&"0"& q & CHR(10)
Body=Body & "private const DTCheck="&q&"0"& q & CHR(10)
Body=Body & "private const DeliveryZip="& q & "0" & q & CHR(10)
Body=Body & "private const scOrderName=0"&CHR(10)
Body=Body & "private const scHideDiscField=0"&CHR(10)
Body=Body & "private const scAllowSeparate=" & q & "0" & q & CHR(10)
Body=Body & "private const scDisableDiscountCodes=" & q & "1" & q & CHR(10)
Body=Body & "private const ReferLabel="&q&"How did you hear about us?"&q&CHR(10)
Body=Body & "private const ViewRefer=0"&CHR(10)
Body=Body & "private const RefNewCheckout=0"&CHR(10)
Body=Body & "private const RefNewReg=0"&CHR(10)
Body=Body & "private const sBrandPro=0" & CHR(10)
Body=Body & "private const sBrandLogo=0" & CHR(10)
Body=Body & "private const RewardsActive=1"&CHR(10)
Body=Body & "private const RewardsIncludeWholesale=0 "&CHR(10)
Body=Body & "private const RewardsPercent=0"&CHR(10)
Body=Body & "private const RewardsLabel="&q&"Reward Points"&q&CHR(10)
Body=Body & "private const RewardsReferral=0"&CHR(10)
Body=Body & "private const RewardsFlat=0"&CHR(10)
Body=Body & "private const RewardsFlatValue=0"&CHR(10)
Body=Body & "private const RewardsPerc=0"&CHR(10)
Body=Body & "private const RewardsPercValue=0"&CHR(10)
Body=Body & "private const pcQDiscountType=0 "&CHR(10)
if statusBTO="1" then
	Body=Body & "private const iBTODisplayType=2"&CHR(10)
	Body=Body & "private const iBTOOutofStockPurchase=0"&CHR(10)
	Body=Body & "private const iBTOShowImage=0"&CHR(10)
	Body=Body & "private const iBTOQuote=1"&CHR(10)
	Body=Body & "private const iBTOQuoteSubmit=0"&CHR(10)
	Body=Body & "private const iBTOQuoteSubmitOnly=0"&CHR(10)
	Body=Body & "private const iBTODetLinkType=0"&CHR(10)
	Body=Body & "private const vBTODetTxt="&q&"View Details"&q&CHR(10)
	Body=Body & "private const iBTOPopWidth=400"& CHR(10)
	Body=Body & "private const iBTOPopHeight=500"& CHR(10)
	Body=Body & "private const iBTOPopImage=1"&CHR(10)
	Body=Body & "private const scConfigPurchaseOnly=0"&CHR(10)
End if
Body=Body & "private const scTerms=0"&CHR(10)
Body=Body & "private const scTermsLabel=" & q & q & CHR(10)
Body=Body & "private const scTermsShown=0"&CHR(10)
Body=Body & "private const scShowSKU=1"&CHR(10)
Body=Body & "private const scShowSmallImg=1"&CHR(10)
Body=Body & "private const scHideRMA=0"&CHR(10)
Body=Body & "private const scShowHD=1"&CHR(10)
Body=Body & "private const scStoreUseToolTip=1"&CHR(10)
Body=Body & "private const scErrorHandler=0"&CHR(10)
Body=Body & "private const scAllowCheckoutWR=1"&CHR(10)
Body=Body & "private const scSeoURLs=0"&CHR(10)
Body=Body & "private const scSeoURLs404=" & q & q & CHR(10)
Body=Body & "private const scQuickBuy=0"&CHR(10)
Body=Body & "private const scATCEnabled=0"&CHR(10)
Body=Body & "private const scRestoreCart=1"&CHR(10)
Body=Body & "private const scGuestCheckoutOpt=0"&CHR(10)
Body=Body & "private const scAddThisDisplay=0"&CHR(10)
Body=Body & "private const scAddThisCode=" & q & q & CHR(10)
Body=Body & "private const scPinterestDisplay=0"&CHR(10)
Body=Body & "private const scPinterestCounter="&q&"."&q&CHR(10)
Body=Body & "private const scGoogleAnalytics=" & q & q & CHR(10)
pcStrXML=session("XMLUse")
Body=Body & "private const scXML = """&replace(pcStrXML,"''","'")&"""" & CHR(37)&CHR(62) & vbCrLf

' create the file using the FileSystemObject
Dim fso, f
on error resume next
Set fso=server.CreateObject("Scripting.FileSystemObject")
Set f=fso.GetFile(findit)
Err.number=0
f.Delete
if Err.number>0 then
	response.redirect "../"&scAdminFolderName&"/techErr.asp?error="&Server.URLEncode("Permissions Not Set to Modify First")
end if
Set f=nothing

Set f=fso.OpenTextFile(findit, 2, True)
f.Write Body
f.Close
Set fso=nothing
Set f=nothing
response.redirect "FirstPageCreateShipFromSettings.asp"
%>