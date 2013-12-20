<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"--> 
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<%
if session("pc_sdsIsDropShipper")="1" then
	pcv_pageType="0"
	pcv_Table="pcSupplier"
else
	pcv_pageType="1"
	pcv_Table="pcDropShipper"
end if
%>
<%
dim f, query, conntemp, rstemp, rs

pcStrPageName="pcmodsdsB.asp"
pcStrErrPage="pcmodsdsA.asp"


'*****************************************************************	
' START: Declare Page Requirements
'*****************************************************************
pcv_sdsCompanyRequired = true
pcv_sdsFirstNameRequired = true
pcv_sdsLastNameRequired = true
pcv_sdsPhoneRequired = true
pcv_sdsEmailRequired = true
pcv_sdsURLRequired = false
pcv_sdsIsDropShipperRequired = false
if pcv_PageType="1" then
	pcv_sdsFromAddressRequired = true
	pcv_sdsFromAddress2Required = false
	pcv_sdsFromCityRequired = true
	pcv_sdsFromZipRequired = true
	pcv_sdsFromCountrycodeRequired = true		
	pcv_sdsFromState1Required = true
	pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
	if  len(pcv_strStateCodeRequired)>0 then
		pcv_sdsFromState1Required=pcv_strStateCodeRequired
	end if		
	pcv_sdsFromState2Required = false
	pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
	if  len(pcv_strProvinceCodeRequired)>0 then
		pcv_sdsFromState2Required=pcv_strProvinceCodeRequired
	end if			
	pcv_sdsUsernameRequired = true
	pcv_sdsPasswordRequired = true
else
	pcv_sdsFromAddressRequired =  false
	pcv_sdsFromAddress2Required =  false
	pcv_sdsFromCityRequired = false
	pcv_sdsFromZipRequired = false
	pcv_sdsFromCountrycodeRequired = false	
	pcv_sdsFromState1Required = false
	pcv_sdsFromState2Required = false	
	pcv_sdsUsernameRequired = false
	pcv_sdsPasswordRequired = false
end if
'*****************************************************************	
' END: Declare Page Requirements
'*****************************************************************
%>
<% 
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' START: Get the Data From the Form
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	'// Main Contact
	pcs_ValidateTextField	"pcv_sdsCompany", pcv_sdsCompanyRequired, 0
	pcs_ValidateTextField	"pcv_sdsFirstName", pcv_sdsFirstNameRequired, 0
	pcs_ValidateTextField	"pcv_sdsLastName", pcv_sdsLastNameRequired, 0
	pcs_ValidatePhoneNumber	"pcv_sdsPhone", pcv_sdsPhoneRequired, 0
	pcs_ValidateEmailField	"pcv_sdsEmail", pcv_sdsEmailRequired, 0
	pcs_ValidateTextField	"pcv_sdsURL", pcv_sdsURLRequired, 0
	pcs_ValidateTextField	"pcv_sdsIsDropShipper", pcv_sdsIsDropShipperRequired, 0
	
	'// Ship-From Address
	pcs_ValidateTextField	"pcv_sdsFromAddress", pcv_sdsFromAddressRequired, 0
	pcs_ValidateTextField	"pcv_sdsFromAddress2", pcv_sdsFromAddress2Required, 0
	pcs_ValidateTextField	"pcv_sdsFromCity", pcv_sdsFromCityRequired, 0
	pcs_ValidateTextField	"pcv_sdsFromZip", pcv_sdsFromZipRequired, 0
	pcs_ValidateTextField	"pcv_sdsFromCountrycode", pcv_sdsFromCountrycodeRequired, 0	
	pcs_ValidateTextField	"pcv_sdsFromState1", pcv_sdsFromState1Required, 0
	pcs_ValidateTextField	"pcv_sdsFromState2", pcv_sdsFromState2Required, 0
	
	'// Login Information
	pcs_ValidateTextField	"pcv_sdsUsername", pcv_sdsUsernameRequired, 0
	pcs_ValidateTextField	"pcv_sdsPassword", pcv_sdsPasswordRequired, 0
	
	'// Drop Shipper Settings
	pcs_ValidateTextField	"pcv_sdsCustNotifyUpdates", false, 0
	pcs_ValidateEmailField	"pcv_sdsNoticeEmail", false, 0
	pcs_ValidateTextField	"pcv_sdsNoticeType", false, 0
	pcs_ValidateTextField	"pcv_sdsNotifyManually", false, 0
	
	'// Billing Address
	pcs_ValidateTextField	"pcv_sdsBillingCountrycode", false, 0
	pcs_ValidateTextField	"pcv_sdsBillingCountrycode", false, 0
	pcs_ValidateTextField	"pcv_sdsBillingAddress", false, 0
	pcs_ValidateTextField	"pcv_sdsBillingAddress2", false, 0
	pcs_ValidateTextField	"pcv_sdsBillingCity", false, 0
	pcs_ValidateTextField	"pcv_sdsBillingState1", false, 0
	pcs_ValidateTextField	"pcv_sdsBillingState2", false, 0
	pcs_ValidateTextField	"pcv_sdsBillingZip", false, 0
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' END: Get the Data From the Form
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' START: Fix Data
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if (Session("pcSFpcv_sdsIsDropShipper")="") or (not Isnumeric(Session("pcSFpcv_sdsIsDropShipper"))) then
		Session("pcSFpcv_sdsIsDropShipper")=0
	end if
		
	pcv_sdsFromState1=Session("pcSFpcv_sdsFromState1")
	pcv_sdsFromState2=Session("pcSFpcv_sdsFromState2")
	if pcv_sdsFromState2<>"" then
		pcv_sdsFromState1=""
		pcv_sdsFromStateProvinceCode=pcv_sdsFromState2
	else
		pcv_sdsFromState2=""
		pcv_sdsFromStateProvinceCode=pcv_sdsFromState1
	end if
	
	if Session("pcSFpcv_sdsPassword")<>"" then
		Session("pcSFpcv_sdsPassword")=enDeCrypt(Session("pcSFpcv_sdsPassword"), scCrypPass)
	end if
	
	if (Session("pcSFpcv_sdsCustNotifyUpdates")="") or (not Isnumeric(Session("pcSFpcv_sdsCustNotifyUpdates"))) then
		Session("pcSFpcv_sdsCustNotifyUpdates")=0
	end if
	
	if (Session("pcSFpcv_sdsNoticeType")="") or (not Isnumeric(Session("pcSFpcv_sdsNoticeType"))) then
		Session("pcSFpcv_sdsNoticeType")=0
	end if
	
	if (Session("pcSFpcv_sdsNotifyManually")="") or (not Isnumeric(Session("pcSFpcv_sdsNotifyManually"))) then
		Session("pcSFpcv_sdsNotifyManually")=0
	end if
	
	pcv_sdsBillingState1=Session("pcSFpcv_sdsBillingState1")
	pcv_sdsBillingState2=Session("pcSFpcv_sdsBillingState2")

	if pcv_sdsBillingState2<>"" then
		pcv_sdsBillingState1=""
		pcv_sdsBillingStateProvinceCode=pcv_sdsBillingState2
	else
		pcv_sdsBillingState2=""
		pcv_sdsBillingStateProvinceCode=pcv_sdsBillingState1
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' END: Fix
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	query=""

call openDb()

pcv_idsds=session("pc_idsds")
if (pcv_idsds="") or (not Isnumeric(pcv_idsds)) then
	pcv_idsds=0
end if
	
If pcv_intErr>0 Then
	response.redirect pcStrErrPage&"?msg="&pcv_strGenericPageError&"&pagetype=" & pcv_pageType & "&idsds=" & pcv_idsds
Else
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' START: Update
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	query="UPDATE " & pcv_Table & "s SET " & pcv_Table & "_Username='" & Session("pcSFpcv_sdsUsername") & "'," & pcv_Table & "_Password='" & Session("pcSFpcv_sdsPassword") & "'," & pcv_Table & "_FirstName='" & Session("pcSFpcv_sdsFirstName") & "'," & pcv_Table & "_LastName='" & Session("pcSFpcv_sdsLastName") & "'," & pcv_Table & "_Company='" & Session("pcSFpcv_sdsCompany") & "'," & pcv_Table & "_Phone='" & Session("pcSFpcv_sdsPhone") & "'," & pcv_Table & "_Email='" & Session("pcSFpcv_sdsEmail") & "'," & pcv_Table & "_URL='" & Session("pcSFpcv_sdsURL") & "'," & pcv_Table & "_FromAddress='" & Session("pcSFpcv_sdsFromAddress") & "'," & pcv_Table & "_FromAddress2='" & Session("pcSFpcv_sdsFromAddress2") & "'," & pcv_Table & "_FromCity='" & Session("pcSFpcv_sdsFromCity") & "'," & pcv_Table & "_FromStateProvinceCode='" & pcv_sdsFromStateProvinceCode & "'," & pcv_Table & "_FromZip='" & Session("pcSFpcv_sdsFromZip") & "'," & pcv_Table & "_FromCountryCode='" & Session("pcSFpcv_sdsFromCountrycode") & "'," & pcv_Table & "_BillingAddress='" & Session("pcSFpcv_sdsBillingAddress") & "'," & pcv_Table & "_BillingAddress2='" & Session("pcSFpcv_sdsBillingAddress2") & "'," & pcv_Table & "_BillingCity='" & Session("pcSFpcv_sdsBillingCity") & "'," & pcv_Table & "_BillingStateProvinceCode='" & pcv_sdsBillingStateProvinceCode & "'," & pcv_Table & "_BillingZip='" & Session("pcSFpcv_sdsBillingZip") & "'," & pcv_Table & "_BillingCountryCode='" & Session("pcSFpcv_sdsBillingCountrycode") & "'," & pcv_Table & "_NoticeEmail='" & Session("pcSFpcv_sdsNoticeEmail") & "'," & pcv_Table & "_NoticeType=" & Session("pcSFpcv_sdsNoticeType") & "," & pcv_Table & "_NoticeMsg='" & Session("pcSFpcv_sdsNoticeMsg") & "'," & pcv_Table & "_NotifyManually=" & Session("pcSFpcv_sdsNotifyManually") & "," & pcv_Table & "_CustNotifyUpdates=" & Session("pcSFpcv_sdsCustNotifyUpdates")
	if pcv_pageType="0" then
		query=query & ","  & pcv_Table & "_IsDropShipper=" & Session("pcSFpcv_sdsIsDropShipper")
	end if
	query=query & " WHERE " & pcv_Table & "_ID=" & pcv_idsds
	
	'response.write query
	'response.end
	set rs=connTemp.execute(query)
	set rs=nothing
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' END: Update
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
END IF

'// Clear the sessions
	pcs_ClearAllSessions
	
set rstemp=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

call closedb()

conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing

response.redirect "sds_MainMenu.asp?msg=" & dictLanguage.Item(Session("language")&"_ModsdsB_1") %>