<%@LANGUAGE="VBSCRIPT"%>
<% 
On Error Resume Next
PmAdmin=19
pageTitle = "ProductCart v4.7 - Database Update Completed"
Section = "" 
%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<% dim f, mySQL, conntemp, rstemp, rs, iCnt %>
<!--#include file="Adminheader.asp"-->
<!--#include file="pcAdminRetrieveSettings.asp"-->
<%
if NOT isNULL(pcStrCompanyName) then pcStrCompanyName=replace(pcStrCompanyName,"'","''")
if NOT isNULL(pcStrCompanyAddress) then pcStrCompanyAddress=replace(pcStrCompanyAddress,"'","''")
if NOT isNULL(pcStrCompanyZip) then pcStrCompanyZip=replace(pcStrCompanyZip,"'","''")
if NOT isNULL(pcStrCompanyCity) then pcStrCompanyCity=replace(pcStrCompanyCity,"'","''")
if NOT isNULL(pcStrCompanyState) then pcStrCompanyState=replace(pcStrCompanyState,"'","''")
if NOT isNULL(pcStrCompanyCountry) then pcStrCompanyCountry=replace(pcStrCompanyCountry,"'","''")
if NOT isNULL(pcStrCompanyLogo) then pcStrCompanyLogo=replace(pcStrCompanyLogo,"'","''")
if NOT isNULL(pcStrCurSign) then pcStrCurSign=replace(pcStrCurSign,"'","''")
if NOT isNULL(pcStrDecSign) then pcStrDecSign=replace(pcStrDecSign,"'","''")
if NOT isNULL(pcStrDivSign) then pcStrDivSign=replace(pcStrDivSign,"'","''")
if NOT isNULL(pcStrDateFrmt) then pcStrDateFrmt=replace(pcStrDateFrmt,"'","''")
if NOT isNULL(pcStrURLredirect) then pcStrURLredirect=replace(pcStrURLredirect,"'","''")
if NOT isNULL(pcStrSSL) then pcStrSSL=replace(pcStrSSL,"'","''")
if NOT isNULL(pcStrSSLUrl) then pcStrSSLUrl=replace(pcStrSSLUrl,"'","''")
if NOT isNULL(pcStrIntSSLPage) then pcStrIntSSLPage=replace(pcStrIntSSLPage,"'","''")
if NOT isNULL(pcStrBType) then pcStrBType=replace(pcStrBType,"'","''")
if NOT isNULL(pcStrStoreOff) then pcStrStoreOff=replace(pcStrStoreOff,"'","''")
if NOT isNULL(pcStrStoreMsg) then pcStrStoreMsg=replace(pcStrStoreMsg,"'","''")
if NOT isNULL(pcStrorderLevel) then pcStrorderLevel=replace(pcStrorderLevel,"'","''")
if NOT isNULL(pcStrNewsLabel) then pcStrNewsLabel=replace(pcStrNewsLabel,"'","''")
if NOT isNULL(pcStrDFLabel) then pcStrDFLabel=replace(pcStrDFLabel,"'","''")
if NOT isNULL(pcStrDFShow) then pcStrDFShow=replace(pcStrDFShow,"'","''")
if NOT isNULL(pcStrDFReq) then pcStrDFReq=replace(pcStrDFReq,"'","''")
if NOT isNULL(pcStrTFLabel) then pcStrTFLabel=replace(pcStrTFLabel,"'","''")
if NOT isNULL(pcStrTFShow) then pcStrTFShow=replace(pcStrTFShow,"'","''")
if NOT isNULL(pcStrTFReq) then pcStrTFReq=replace(pcStrTFReq,"'","''")
if NOT isNULL(pcStrDTCheck) then pcStrDTCheck=replace(pcStrDTCheck,"'","''")
if NOT isNULL(pcStrDeliveryZip) then pcStrDeliveryZip=replace(pcStrDeliveryZip,"'","''")
if NOT isNULL(pcStrOrderName) then pcStrOrderName=replace(pcStrOrderName,"'","''")
if NOT isNULL(pcStrHideDiscField) then pcStrHideDiscField=replace(pcStrHideDiscField,"'","''")
if NOT isNULL(pcStrAllowSeparate) then pcStrAllowSeparate=replace(pcStrAllowSeparate,"'","''")
if NOT isNULL(pcStrReferLabel) then pcStrReferLabel=replace(pcStrReferLabel,"'","''")
if NOT isNULL(pcStrRewardsLabel) then pcStrRewardsLabel=replace(pcStrRewardsLabel,"'","''")
if NOT isNULL(pcStrXML) then pcStrXML=replace(pcStrXML,"'","''")
if NOT isNULL(pcStrBTODetTxt) then pcStrBTODetTxt=replace(pcStrBTODetTxt,"'","''")
if NOT isNULL(pcStrTermsLabel) then pcStrTermsLabel=replace(pcStrTermsLabel,"'","''")
if NOT isNULL(pcStrTermsCopy) then pcStrTermsCopy=replace(pcStrTermsCopy,"'","''")
if NOT isNULL(pcStrViewPrdStyle) then pcStrViewPrdStyle=replace(pcStrViewPrdStyle,"'","''")
if NOT isNULL(pcStrCustomerIPAlert) then pcStrCustomerIPAlert=replace(pcStrCustomerIPAlert,"'","''")
if NOT isNULL(pcStrCompanyPhoneNumber) then pcStrCompanyPhoneNumber=replace(pcStrCompanyPhoneNumber,"'","''")
if NOT isNULL(pcStrCompanyFaxNumber) then pcStrCompanyFaxNumber=replace(pcStrCompanyFaxNumber,"'","''")
if NOT isNULL(pcStrseoURLs404) then pcStrseoURLs404=replace(pcStrseoURLs404,"'","''")

'New Fields for Beta-Updaters
if len(pcIntShowSKU)<1 then
	pcIntShowSKU=0
end if
if len(pcIntShowSmallImg)<1 then
	pcIntShowSmallImg=-1
end if
if len(pcStrViewPrdStyle)<1 then
	pcStrViewPrdStyle="c"
end if

'/////////////////////////////////////////////////////
'// Update version number
'/////////////////////////////////////////////////////

'//Version 3 & UP only - change for any new version updates
if scBTO=1 then
	pcStrScVersion="4.7b"
else
	pcStrScVersion="4.7"
end if

'//Service Pack Number
pcStrScSP = "0"

'// Detection of add-ons and update of version number based on their presence
'// is performed by pcAdminSaveSettings.asp

%>
<!--#include file="pcAdminSaveSettings.asp"-->
<% 'Detect Add-on %>
<!--#include file="pcAddOnDetection.asp"-->
<style>
li {
	padding-bottom: 8px;
}
h2 {
	font-size: 12px;
}
</style>
<table class="pcCPcontent">
	<tr>
		<td>
      <div class="pcCPmessageSuccess">Your store database was successfully updated. The version number will be updated to v4.7 the next time you load any Control Panel page.</div>
      <div class="pcCPmessage">Please note: you are not done. Carefully review the information below for additional steps you need to take <strong>before re-opening your storefront</strong>. For detailed information, <a href="http://wiki.earlyimpact.com/updates/v4.7" target="_blank">click here</a>.</div>
    </td>
  </tr>
  <tr>
  	<td class="pcCPspacer"></td>
  </tr>
</table>
<!--#include file="AdminFooter.asp"-->