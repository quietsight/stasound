<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
if request("pagetype")="1" then
	pcv_PageType=1
	pcv_Title="Drop-Shipper"
	pcv_Table="pcDropShipper"
else
	pcv_PageType=0
	pcv_Title="Supplier"
	pcv_Table="pcSupplier"
end if

if request("clone")="1" then
	pcv_clone="1"
else
	pcv_clone="0"
end if

pcv_idsds=request("idsds")
if (pcv_idsds="") or (not Isnumeric(pcv_idsds)) then
	pcv_idsds="0"
end if

if pcv_idsds="0" AND request("action")<>"add" then
	'response.redirect "menu.asp"
end if

if pcv_clone="1" then
	pageTitle="Clone " & pcv_Title 
else
	pageTitle="Modify " & pcv_Title 
end if

%>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<% 
call opendb()

Dim connTemp,rs,query

pcStrPageName="sds_modify.asp"

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
pcv_sdspcv_sdsIsDropShipper_Cur = false
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
	pcv_sdsNoticeMsg = true
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
	pcv_sdsNoticeMsgRequired = false
end if
'*****************************************************************	
' END: Declare Page Requirements
'*****************************************************************
%>
<%
'*****************************************************************	
' START: Load the Form
'*****************************************************************
query="SELECT " & pcv_Table & "_ID," & pcv_Table & "_Username," & pcv_Table & "_Password," & pcv_Table & "_FirstName," & pcv_Table & "_LastName," & pcv_Table & "_Company," & pcv_Table & "_Phone," & pcv_Table & "_Email," & pcv_Table & "_URL," & pcv_Table & "_FromAddress," & pcv_Table & "_FromAddress2," & pcv_Table & "_FromCity," & pcv_Table & "_FromStateProvinceCode," & pcv_Table & "_FromZip," & pcv_Table & "_FromCountryCode," & pcv_Table & "_BillingAddress," & pcv_Table & "_BillingAddress2," & pcv_Table & "_BillingCity," & pcv_Table & "_BillingStateProvinceCode," & pcv_Table & "_BillingZip," & pcv_Table & "_BillingCountryCode," & pcv_Table & "_NoticeEmail," & pcv_Table & "_NoticeType," & pcv_Table & "_NoticeMsg," & pcv_Table & "_NotifyManually," & pcv_Table & "_CustNotifyUpdates"
	if pcv_pageType="0" then
		query=query & ","  & pcv_Table & "_IsDropShipper"
	end if	
	query=query & " FROM " & pcv_Table & "s WHERE " & pcv_Table & "_ID=" & pcv_idsds
	set rs=connTemp.execute(query)
	
	Session("pcAdminpcv_IDsds")=rs(pcv_Table & "_ID")
	Session("pcAdminpcv_sdsUsername")=rs(pcv_Table & "_Username")
	Session("pcAdminpcv_sdsPassword")=rs(pcv_Table & "_Password")
	if Session("pcAdminpcv_sdsPassword")<>"" then
		Session("pcAdminpcv_sdsPassword")=enDeCrypt(Session("pcAdminpcv_sdsPassword"), scCrypPass)
	end if
	Session("pcAdminpcv_sdsFirstName")=rs(pcv_Table & "_FirstName")
	Session("pcAdminpcv_sdsLastName")=rs(pcv_Table & "_LastName")
	Session("pcAdminpcv_sdsCompany")=rs(pcv_Table & "_Company")
	Session("pcAdminpcv_sdsPhone")=rs(pcv_Table & "_Phone")
	Session("pcAdminpcv_sdsEmail")=rs(pcv_Table & "_Email")
	Session("pcAdminpcv_sdsURL")=rs(pcv_Table & "_URL")
	Session("pcAdminpcv_sdsFromAddress")=rs(pcv_Table & "_FromAddress")
	Session("pcAdminpcv_sdsFromAddress2")=rs(pcv_Table & "_FromAddress2")
	Session("pcAdminpcv_sdsFromCity")=rs(pcv_Table & "_FromCity")
	Session("pcAdminpcv_sdsFromStateProvinceCode")=rs(pcv_Table & "_FromStateProvinceCode")
	Session("pcAdminpcv_sdsFromZip")=rs(pcv_Table & "_FromZip")
	Session("pcAdminpcv_sdsFromCountrycode")=rs(pcv_Table & "_FromCountrycode")
	Session("pcAdminpcv_sdsBillingAddress")=rs(pcv_Table & "_BillingAddress")
	Session("pcAdminpcv_sdsBillingAddress2")=rs(pcv_Table & "_BillingAddress2")
	Session("pcAdminpcv_sdsBillingCity")=rs(pcv_Table & "_BillingCity")
	Session("pcAdminpcv_sdsBillingStateProvinceCode")=rs(pcv_Table & "_BillingStateProvinceCode")
	Session("pcAdminpcv_sdsBillingZip")=rs(pcv_Table & "_BillingZip")
	Session("pcAdminpcv_sdsBillingCountrycode")=rs(pcv_Table & "_BillingCountrycode")
	Session("pcAdminpcv_sdsNoticeEmail")=rs(pcv_Table & "_NoticeEmail")
	Session("pcAdminpcv_sdsNoticeType")=rs(pcv_Table & "_NoticeType")
	if (Session("pcAdminpcv_sdsNoticeType")="") or (not Isnumeric(Session("pcAdminpcv_sdsNoticeType"))) then
		Session("pcAdminpcv_sdsNoticeType")=0
	end if
	Session("pcAdminpcv_sdsNoticeMsg")=rs(pcv_Table & "_NoticeMsg")
	Session("pcAdminpcv_sdsNotifyManually")=rs(pcv_Table & "_NotifyManually")
	if (Session("pcAdminpcv_sdsNotifyManually")="") or (not Isnumeric(Session("pcAdminpcv_sdsNotifyManually"))) then
		Session("pcAdminpcv_sdsNotifyManually")=0
	end if
	
	Session("pcAdminpcv_sdsCustNotifyUpdates")=rs(pcv_Table & "_CustNotifyUpdates")
	if (Session("pcAdminpcv_sdsCustNotifyUpdates")="") or (not Isnumeric(Session("pcAdminpcv_sdsCustNotifyUpdates"))) then
		Session("pcAdminpcv_sdsCustNotifyUpdates")=0
	end if
	
	Session("pcAdminpcv_sdsIsDropShipper")=0
	if pcv_pageType="0" then
		Session("pcAdminpcv_sdsIsDropShipper")=rs(pcv_Table & "_IsDropShipper")
		if (Session("pcAdminpcv_sdsIsDropShipper")="") or (not Isnumeric(Session("pcAdminpcv_sdsIsDropShipper"))) then
			Session("pcAdminpcv_sdsIsDropShipper")=0
		end if
	end if

	set rs=nothing	
'*****************************************************************	
' END: Load the Form
'*****************************************************************
%>
<%
'*****************************************************************	
' START: POSTBACK
'*****************************************************************
IF (request("action")="add") or (request("action")="upd") THEN

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
	pcs_ValidateTextField	"pcv_sdsIsDropShipper_Cur", pcv_sdspcv_sdsIsDropShipper_Cur, 0
	
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
	pcs_ValidateTextField	"pcv_sdsNoticeMsg", pcv_sdsNoticeMsgShipperRequired, 0
	
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
	if (Session("pcAdminpcv_sdsIsDropShipper")="") or (not Isnumeric(Session("pcAdminpcv_sdsIsDropShipper"))) then
		Session("pcAdminpcv_sdsIsDropShipper")=0
	end if
		
	pcv_sdsFromState1=Session("pcAdminpcv_sdsFromState1")
	pcv_sdsFromState2=Session("pcAdminpcv_sdsFromState2")
	if pcv_sdsFromState2<>"" then
		pcv_sdsFromStateProvinceCode=pcv_sdsFromState2
	else
		pcv_sdsFromStateProvinceCode=pcv_sdsFromState1
	end if
	
	if Session("pcAdminpcv_sdsPassword")<>"" then
		Session("pcAdminpcv_sdsPassword")=enDeCrypt(Session("pcAdminpcv_sdsPassword"), scCrypPass)
	end if
	
	if (Session("pcAdminpcv_sdsCustNotifyUpdates")="") or (not Isnumeric(Session("pcAdminpcv_sdsCustNotifyUpdates"))) then
		Session("pcAdminpcv_sdsCustNotifyUpdates")=0
	end if
	
	if (Session("pcAdminpcv_sdsNoticeType")="") or (not Isnumeric(Session("pcAdminpcv_sdsNoticeType"))) then
		Session("pcAdminpcv_sdsNoticeType")=0
	end if
	
	if (Session("pcAdminpcv_sdsNotifyManually")="") or (not Isnumeric(Session("pcAdminpcv_sdsNotifyManually"))) then
		Session("pcAdminpcv_sdsNotifyManually")=0
	end if
	
	pcv_sdsBillingState1=Session("pcAdminpcv_sdsBillingState1")
	pcv_sdsBillingState2=Session("pcAdminpcv_sdsBillingState2")
	if pcv_sdsBillingState2<>"" then
		pcv_sdsBillingStateProvinceCode=pcv_sdsBillingState2
	else
		pcv_sdsBillingStateProvinceCode=pcv_sdsBillingState1
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' END: Fix
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError&"&pagetype=" & pcv_pageType & "&idsds=" & pcv_idsds
	Else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		' START: Add OR Update
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		query=""
		'// Add a Drop Shipper
		IF request("action")="add" then
			query="INSERT INTO " & pcv_Table & "s (" & pcv_Table & "_Username," & pcv_Table & "_Password," & pcv_Table & "_FirstName," & pcv_Table & "_LastName," & pcv_Table & "_Company," & pcv_Table & "_Phone," & pcv_Table & "_Email," & pcv_Table & "_URL," & pcv_Table & "_FromAddress," & pcv_Table & "_FromAddress2," & pcv_Table & "_FromCity," & pcv_Table & "_FromStateProvinceCode," & pcv_Table & "_FromZip," & pcv_Table & "_FromCountryCode," & pcv_Table & "_BillingAddress," & pcv_Table & "_BillingAddress2," & pcv_Table & "_BillingCity," & pcv_Table & "_BillingStateProvinceCode," & pcv_Table & "_BillingZip," & pcv_Table & "_BillingCountryCode," & pcv_Table & "_NoticeEmail," & pcv_Table & "_NoticeType," & pcv_Table & "_NoticeMsg," & pcv_Table & "_NotifyManually," & pcv_Table & "_CustNotifyUpdates"
			if pcv_pageType="0" then
				query=query & ","  & pcv_Table & "_IsDropShipper"
			end if
			query=query & ") VALUES (" & "'" & Session("pcAdminpcv_sdsUsername") & "','" & Session("pcAdminpcv_sdsPassword") & "','" & Session("pcAdminpcv_sdsFirstName") & "','" & Session("pcAdminpcv_sdsLastName") & "','" & Session("pcAdminpcv_sdsCompany") & "','" & Session("pcAdminpcv_sdsPhone") & "','" & Session("pcAdminpcv_sdsEmail") & "','" & Session("pcAdminpcv_sdsURL") & "','" & Session("pcAdminpcv_sdsFromAddress") & "','" & Session("pcAdminpcv_sdsFromAddress2") & "','" & Session("pcAdminpcv_sdsFromCity") & "','" & pcv_sdsFromStateProvinceCode & "','" & Session("pcAdminpcv_sdsFromZip") & "','" & Session("pcAdminpcv_sdsFromCountrycode") & "','" & Session("pcAdminpcv_sdsBillingAddress") & "','" & Session("pcAdminpcv_sdsBillingAddress2") & "','" & Session("pcAdminpcv_sdsBillingCity") & "','" & pcv_sdsBillingStateProvinceCode & "','" & Session("pcAdminpcv_sdsBillingZip") & "','" & Session("pcAdminpcv_sdsBillingCountrycode") & "','" & Session("pcAdminpcv_sdsNoticeEmail") & "'," & Session("pcAdminpcv_sdsNoticeType") & ",'" & Session("pcAdminpcv_sdsNoticeMsg") & "'," & Session("pcAdminpcv_sdsNotifyManually") & "," & Session("pcAdminpcv_sdsCustNotifyUpdates")
			if pcv_pageType="0" then
				query=query & ","  & Session("pcAdminpcv_sdsIsDropShipper")
			end if
			query=query & ");"
		ELSE
		'// Update a Drop Shipper
			pcv_idsds=request("idsds")
			if (pcv_idsds="") or (not Isnumeric(pcv_idsds)) then
				pcv_idsds=0
			end if
			query="UPDATE " & pcv_Table & "s SET " & pcv_Table & "_Username='" & Session("pcAdminpcv_sdsUsername") & "'," & pcv_Table & "_Password='" & Session("pcAdminpcv_sdsPassword") & "'," & pcv_Table & "_FirstName='" & Session("pcAdminpcv_sdsFirstName") & "'," & pcv_Table & "_LastName='" & Session("pcAdminpcv_sdsLastName") & "'," & pcv_Table & "_Company='" & Session("pcAdminpcv_sdsCompany") & "'," & pcv_Table & "_Phone='" & Session("pcAdminpcv_sdsPhone") & "'," & pcv_Table & "_Email='" & Session("pcAdminpcv_sdsEmail") & "'," & pcv_Table & "_URL='" & Session("pcAdminpcv_sdsURL") & "'," & pcv_Table & "_FromAddress='" & Session("pcAdminpcv_sdsFromAddress") & "'," & pcv_Table & "_FromAddress2='" & Session("pcAdminpcv_sdsFromAddress2") & "'," & pcv_Table & "_FromCity='" & Session("pcAdminpcv_sdsFromCity") & "'," & pcv_Table & "_FromStateProvinceCode='" & pcv_sdsFromStateProvinceCode & "'," & pcv_Table & "_FromZip='" & Session("pcAdminpcv_sdsFromZip") & "'," & pcv_Table & "_FromCountryCode='" & Session("pcAdminpcv_sdsFromCountrycode") & "'," & pcv_Table & "_BillingAddress='" & Session("pcAdminpcv_sdsBillingAddress") & "'," & pcv_Table & "_BillingAddress2='" & Session("pcAdminpcv_sdsBillingAddress2") & "'," & pcv_Table & "_BillingCity='" & Session("pcAdminpcv_sdsBillingCity") & "'," & pcv_Table & "_BillingStateProvinceCode='" & pcv_sdsBillingStateProvinceCode & "'," & pcv_Table & "_BillingZip='" & Session("pcAdminpcv_sdsBillingZip") & "'," & pcv_Table & "_BillingCountryCode='" & Session("pcAdminpcv_sdsBillingCountrycode") & "'," & pcv_Table & "_NoticeEmail='" & Session("pcAdminpcv_sdsNoticeEmail") & "'," & pcv_Table & "_NoticeType=" & Session("pcAdminpcv_sdsNoticeType") & "," & pcv_Table & "_NoticeMsg='" & Session("pcAdminpcv_sdsNoticeMsg") & "'," & pcv_Table & "_NotifyManually=" & Session("pcAdminpcv_sdsNotifyManually") & "," & pcv_Table & "_CustNotifyUpdates=" & Session("pcAdminpcv_sdsCustNotifyUpdates")
			if pcv_pageType="0" then
				query=query & ","  & pcv_Table & "_IsDropShipper=" & Session("pcAdminpcv_sdsIsDropShipper")
			end if
			query=query & " WHERE " & pcv_Table & "_ID=" & pcv_idsds
		END IF
		'response.write query
		'response.end
		set rs=connTemp.execute(query)
		set rs=nothing
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		' END: Add OR Update
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	End If
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' START: Update Products
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if (pcv_pageType="0") and (request("action")="upd") and (Session("pcAdminpcv_sdsIsDropShipper")="0") then
		if Session("pcAdminpcv_sdsIsDropShipper")="0" AND Session("pcAdminpcv_sdsIsDropShipper_Cur")="1" then
			query="SELECT pcDropShippersSuppliers.idproduct FROM pcDropShippersSuppliers INNER JOIN products ON (pcDropShippersSuppliers.idproduct=products.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=1) WHERE products.pcDropShipper_ID=" & pcv_idsds &";"
			set rs=connTemp.execute(query)
			if not rs.eof then
				pcArr=rs.getRows()
				intCount=ubound(pcArr,2)
				set rs=nothing
				For i=0 to intCount
					query="UPDATE Products set pcProd_IsDropShipped=0,pcDropShipper_ID=0 WHERE idproduct=" & pcArr(0,i) & ";"
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
					query="DELETE FROM pcDropShippersSuppliers WHERE idproduct=" & pcArr(0,i) & ";"
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
				Next
			end if
			set rs=nothing
		end if
	else
		if (pcv_pageType="0") and (request("action")="upd") and (Session("pcAdminpcv_sdsIsDropShipper")="1") then
			if Session("pcAdminpcv_sdsIsDropShipper")="1" AND Session("pcAdminpcv_sdsIsDropShipper_Cur")="0" then
				query="UPDATE products SET pcProd_IsDropShipped=1,pcDropShipper_ID=" & pcv_idsds &" WHERE pcSupplier_ID=" & pcv_idsds &";"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="SELECT idproduct FROM products WHERE pcSupplier_ID=" & pcv_idsds &";"
				set rs=connTemp.execute(query)
				if not rs.eof then
					pcArr=rs.getRows()
					intCount=ubound(pcArr,2)
					set rs=nothing
					For i=0 to intCount
						query="DELETE FROM pcDropShippersSuppliers WHERE idproduct=" & pcArr(0,i) & ";"
						set rstemp=connTemp.execute(query)
						set rstemp=nothing
						query="INSERT INTO pcDropShippersSuppliers (idProduct,pcDS_IsDropShipper) VALUES (" & pcArr(0,i) & ",1)"
						set rstemp=connTemp.execute(query)
						set rstemp=nothing
					Next
				end if
				set rs=nothing
			end if
		end if
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' END: Update Products
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' START: Display Message
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If request("action")="add" then
		pcMessage=pcv_Title & " was added successfully!"
	else
		pcMessage=pcv_Title & " was updated successfully!"
	End if
	
	'// Clear the sessions
	pcs_ClearAllSessions
	%>

	<table class="pcCPcontent">
	<tr>
		<td width="15%">&nbsp;</td>
		<td>
			<% ' START show message, if any
			If pcMessage <> "" Then %>
			<div class="pcCPmessageSuccess">
				<%=pcMessage%>
				<br /><br />
				<a href="sds_manage.asp?pagetype=1">Manage Drop-Shippers</a>
				&nbsp;|&nbsp;
				<a href="sds_manage.asp?pagetype=0">Manage Suppliers</a>
			</div>
			<% 	end if
			' END show message %>
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	</table>
	<%
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	' END: Display Message
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
'*****************************************************************	
' END: POSTBACK
'*****************************************************************
ELSE ' if not add/edit action
%> 
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf
response.write "<!--"&vbcrlf	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf
pcs_JavaTextField	"pcv_sdsCompany", pcv_iscompanyRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pcv_sdsFirstName",pcv_sdsFirstNameRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pcv_sdsLastName",pcv_sdsLastNameRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pcv_sdsEmail", pcv_sdsEmailRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pcv_sdsPhone", pcv_sdsPhoneRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
if pcv_PageType="0" then
	response.write "if (theForm.pcv_sdsIsDropShipper.checked==true)"&vbcrlf
	response.write "{"&vbcrlf
end if
	pcs_JavaTextField	"pcv_sdsFromAddress", true, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcv_sdsFromCity", true, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcv_sdsFromZip", true, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcv_sdsFromCountrycode", true, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcv_sdsUsername", true, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcv_sdsPassword", true, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcv_sdsNoticeMsg", true, dictLanguage.Item(Session("language")&"_NewCust_3")
if pcv_PageType="0" then
	response.write "}"&vbcrlf
end if
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "//-->"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>	

<form name="form1" action="<%if pcv_clone="1" then%>sds_addnew.asp<%else%><%=pcStrPageName%><%end if%>?action=<%if pcv_clone="1" then%>add<%else%>upd<%end if%>" method="post" class="pcForms" onsubmit="return Form1_Validator(this)">
	<table class="pcCPcontent">
		
		<% 
		msg=request.QueryString("msg")
		if msg<>"" then %>
		<tr> 
			<td colspan="2">
				<div class="pcCPmessage"><%=msg%></div>  
			</td>
		</tr>   
		<% end if %>
		
		<%if pcv_clone="0" then%>
		<tr align="left">
			<td colspan="2">
            	<h2>Your are currently editing the <%=pcv_Title%>: <strong><%=Session("pcAdminpcv_sdsCompany")%> (<%=Session("pcAdminpcv_sdsFirstName") & " " & Session("pcAdminpcv_sdsLastName")%>)</strong></h2>
                <div><a href="sds_modify.asp?pagetype=<%=pcv_PageType%>&clone=1&idsds=<%=Session("pcAdminpcv_IDsds")%>">Clone this <%=pcv_Title%></a> |  <img src="images/pc_required.gif" alt="required field" width="9" height="9" hspace="5">Indicates required fields</div></td>
		</tr>
		<%end if%>
		<tr>
			<td class="pcCPspacer" colspan="2"></td>
		</tr>
		<tr>
			<th colspan="2">Main Contact</th>
		</tr>
		<tr>
			<td class="pcCPspacer" colspan="2"></td>
		</tr>
		<tr>
			<td width="20%"><p>Company:</p></td>
			<td width="80%">
				<p>
				<input type=text name="pcv_sdsCompany" size="50" value="<% =pcf_FillFormField ("pcv_sdsCompany", pcv_sdsCompanyRequired) %>">
				&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9">
				</p>
			</td>
		</tr>
		<tr>
			<td><p>First Name:<p></td>
			<td>
				<p>
				<input type=text name="pcv_sdsFirstName" size="50" value="<% =pcf_FillFormField ("pcv_sdsFirstName", pcv_sdsFirstNameRequired) %>">
				&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9">
				<p>
			</td>
		</tr>
		<tr>
			<td><p>Last Name:</p></td>
			<td>
				<p><input type=text name="pcv_sdsLastName" size="50" value="<% =pcf_FillFormField ("pcv_sdsLastName", pcv_sdsLastNameRequired) %>">
				&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9"></p>
			</td>
		</tr>
		
		<%	'// Phone Custom Error
		if session("Errpcv_sdsPhone")<>"" then %>
			<tr> 
				<td>&nbsp;</td>
				<td>
				<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%>
				</td>
			</tr>
			<% 
			session("Errpcv_sdsPhone") = ""
		end if 
		%>

		<tr>
			<td><p>Phone:</p></td>
			<td>
				<p><input type=text name="pcv_sdsPhone" size="50" value="<% =pcf_FillFormField ("pcv_sdsPhone", pcv_sdsPhoneRequired) %>">
				&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9"></p>
			</td>
		</tr>
		
		<%	'// Email Custom Error
		if session("Errpcv_sdsEmail")<>"" then %>
			<tr> 
				<td>&nbsp;</td>
				<td>
				<img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> <%=dictLanguage.Item(Session("language")&"_Custmoda_16")%>
				</td>
			</tr>
			<% 
			session("Errpcv_sdsEmail") = ""
		end if 
		%>
		
		<tr>
			<td><p>E-mail:</p></td>
			<td>
				<p><input type=text name="pcv_sdsEmail" size="50" value="<% =pcf_FillFormField ("pcv_sdsEmail", pcv_sdsEmailRequired) %>">
				&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9"></p>
			</td>
		</tr>
		<tr>
			<td><p>Website URL:</p></td>
			<td>
				<p><input type=text name="pcv_sdsURL" size="50" value="<% =pcf_FillFormField ("pcv_sdsURL", pcv_sdsURLRequired) %>"></p>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer" colspan="2"></td>
		</tr>
		
		<%if pcv_PageType="0" then%>
		<tr>
			<td align="right">
			<input type="checkbox" name="pcv_sdsIsDropShipper" value="1" onclick="javascript: isDropShipper(this.checked);" class="clearBorder" <% if Session("pcAdminpcv_sdsIsDropShipper")="1" then response.write "checked" %>>
            <input name="pcv_sdsIsDropShipper_Cur" type="hidden" value="<%=Session("pcAdminpcv_sdsIsDropShipper")%>">
			</td>
			<td>
			<p>
			Enable Drop-Shipping <a href="JavaScript:win('helpOnline.asp?ref=101')">
			<img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
			</p>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer" colspan="2"></td>
		</tr>
		<%end if%>
		<tr>
			<td colspan="2">
				<table width="100%" id="show_1" <%if (pcv_PageType="0") and (pcv_sdsIsDropShipper<>"1") then%>style="display:none"<%end if%>>
				<tr>
					<td class="pcCPspacer" colspan="2"></td>
				</tr>
				<tr>
					<th colspan="2">Shipping From Address</th>
				</tr>
				<tr>
					<td class="pcCPspacer" colspan="2"></td>
				</tr>				
				<%
				'///////////////////////////////////////////////////////////
				'// START: COUNTRY AND STATE/ PROVINCE CONFIG
				'///////////////////////////////////////////////////////////
				' 
				' 1) Place this section ABOVE the Country field
				' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
				' 3) Additional Required Info
				
				'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
				pcv_isStateCodeRequired = pcv_sdsFromState1Required '// determines if validation is performed (true or false)
				pcv_isProvinceCodeRequired = pcv_sdsFromState2Required '// determines if validation is performed (true or false)
				pcv_isCountryCodeRequired = pcv_sdsFromCountrycodeRequired '// determines if validation is performed (true or false)
				
				'// #3 Additional Required Info
				pcv_strTargetForm = "form1" '// Name of Form
				pcv_strCountryBox = "pcv_sdsFromCountrycode" '// Name of Country Dropdown
				pcv_strTargetBox = "pcv_sdsFromState1" '// Name of State Dropdown
				pcv_strProvinceBox =  "pcv_sdsFromState2" '// Name of Province Field
				
				'// Set local Country to Session
				if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
					Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
				end if
				
				'// Set local State to Session
				if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
					Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session(pcv_strSessionPrefix&"pcv_sdsFromStateProvinceCode")
				end if
				
				'// Set local Province to Session
				if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
					Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session(pcv_strSessionPrefix&"pcv_sdsFromStateProvinceCode")
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
					<td width="20%"><p>Address:</p></td>
					<td width="80%">
						<p><input type=text name="pcv_sdsFromAddress" size="50" value="<% =pcf_FillFormField ("pcv_sdsFromAddress", pcv_sdsFromAddressRequired) %>">
						&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9"></p>
					</td>
				</tr>
				<tr>
					<td><p>Address 2:</p></td>
					<td>
						<p><input type=text name="pcv_sdsFromAddress2" size="50" value="<% =pcf_FillFormField ("pcv_sdsFromAddress2", pcv_sdsFromAddress2Required) %>"></p>
					</td>
				</tr>
				<tr>
					<td><p>City:</p></td>
					<td>
						<p><input type=text name="pcv_sdsFromCity" size="50" value="<% =pcf_FillFormField ("pcv_sdsFromCity", pcv_sdsFromCityRequired) %>">
						&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9"></p>
					</td>
				</tr>
				<%
				'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
				pcs_StateProvince
				%>
				<tr>
					<td><p>Postal Code:</p></td>
					<td>
						<p><input type=text name="pcv_sdsFromZip" size="10" value="<% =pcf_FillFormField ("pcv_sdsFromZip", pcv_sdsFromZipRequired) %>">
						&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9"></p>
					</td>
				</tr>
				
				<tr>
					<td class="pcCPspacer" colspan="2"></td>
				</tr>
				<tr>
					<th colspan="2">Login Information <a href="JavaScript:win('helpOnline.asp?ref=103')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
				</tr>
				<tr>
					<td class="pcCPspacer" colspan="2"></td>
				</tr>
				<tr>
					<td><p>Username:</p></td>
					<td>

					<%if (pcv_clone<>"1") and (trim(Session("pcAdminpcv_sdsUsername"))<>"") then%>
						<p>
						<b><%=Session("pcAdminpcv_sdsUsername")%></b>
						<input type="hidden" name="pcv_sdsUsername" size="50" value="<% =pcf_FillFormField ("pcv_sdsUsername", pcv_sdsUsernameRequired) %>">
						</p>
						<%else%>
						<p>
						<input type="text" name="pcv_sdsUsername" size="50" value="<% =pcf_FillFormField ("pcv_sdsUsername", pcv_sdsUsernameRequired) %>">
						&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9">
						</p>
					<%end if%>
					</td>
				</tr>
				<tr>
					<td><p>Password:<p></td>
					<td>
						<p><input type="password" name="pcv_sdsPassword" size="50" value="<% =pcf_FillFormField ("pcv_sdsPassword", pcv_sdsPasswordRequired) %>">
						&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9"></p>
					</td>
				</tr>
				<tr>
					<td align="right">
					<input type="checkbox" name="pcv_sdsCustNotifyUpdates" value="1" class="clearBorder" <%if Session("pcAdminpcv_sdsCustNotifyUpdates")="1" then response.write "checked"%>>
					</td>
					<td>
					<p>
					Notify customer when order is updated <a href="JavaScript:win('helpOnline.asp?ref=104')">
					<img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
					</p>
					</td>
				</tr>
				<tr>
					<td class="pcCPspacer" colspan="2"></td>
				</tr>

				<tr>
					<td class="pcCPspacer" colspan="2"></td>
				</tr>
				
				<tr>
					<th colspan="2">Drop-Shipper Settings</th>
				</tr>
				
				<tr>
					<td class="pcCPspacer" colspan="2"></td>
				</tr>
				<tr>
					<td nowrap="nowrap"><p>Order Notification E-mail: <a href="JavaScript:win('helpOnline.asp?ref=106')">
					<img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></p></td>
					<td>
						<p><input type=text name="pcv_sdsNoticeEmail" size="50" value="<% =pcf_FillFormField ("pcv_sdsNoticeEmail", pcv_sdsNoticeEmailRequired) %>"></p>
					</td>
				</tr>
				<tr valign="top">
					<td nowrap="nowrap"><p>Order Notification Content: <a href="JavaScript:win('helpOnline.asp?ref=107')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></p></td>
					<td>
						<p><input type="radio" name="pcv_sdsNoticeType" value="0" <%if Session("pcAdminpcv_sdsNoticeType")="0" or Session("pcAdminpcv_sdsNoticeType")="" then response.write "checked"%> class="clearBorder"> Products + Customer shipping information<br>
						<input type="radio" name="pcv_sdsNoticeType" value="1" <%if Session("pcAdminpcv_sdsNoticeType")="1" then response.write "checked"%> class="clearBorder"> Products Only <i>(products are shipped to the store)</i></p>
					</td>
				</tr>
				<tr valign="top">
					<td><p>Order Notification Message:<br><br>
					<span class="pcSmallText">Please refer to the User Guide for detailed information on how to use the dynamic tags shown in the default message to automatically include order information.</span></p>
					</td>
					<td><p><textarea name="pcv_sdsNoticeMsg" rows="15" cols="50"><%=Session("pcAdminpcv_sdsNoticeMsg")%></textarea>&nbsp;<img src="images/pc_required.gif" alt="required field" width="9" height="9" align="top"></p></td>
				</tr>
				<tr>
					<td align="right"><input type="checkbox" name="pcv_sdsNotifyManually" value="1" <%if Session("pcAdminpcv_sdsNotifyManually")="1" then response.write "checked"%> class="clearBorder"></td>
					<td><p>Only Notify Manually <a href="JavaScript:win('helpOnline.asp?ref=105')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></p></td>
				</tr>
				</table>
			</td>
		</tr>
		
		<tr>
			<td class="pcCPspacer" colspan="2"></td>
		</tr>

		<tr>
			<th colspan="2">Billing Address</th>
		</tr>
		<tr>
			<td class="pcCPspacer" colspan="2"></td>
		</tr>
		<%
		'///////////////////////////////////////////////////////////
		'// START: COUNTRY AND STATE/ PROVINCE CONFIG
		'///////////////////////////////////////////////////////////
		' 
		' 1) Place this section ABOVE the Country field
		' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
		' 3) Additional Required Info
		
		'// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
		pcv_isStateCodeRequired = false '// determines if validation is performed (true or false)
		pcv_isProvinceCodeRequired = false '// determines if validation is performed (true or false)
		pcv_isCountryCodeRequired = false '// determines if validation is performed (true or false)
		
		'// #3 Additional Required Info
		pcv_strTargetForm = "form1" '// Name of Form
		pcv_strCountryBox = "pcv_sdsBillingCountrycode" '// Name of Country Dropdown
		pcv_strTargetBox = "pcv_sdsBillingState1" '// Name of State Dropdown
		pcv_strProvinceBox =  "pcv_sdsBillingState2" '// Name of Province Field
		
		'// Set local Country to Session
		if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
			Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
		end if
		
		'// Set local State to Session
		if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
			Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session(pcv_strSessionPrefix&"pcv_sdsBillingStateProvinceCode")
		end if
		
		'// Set local Province to Session
		if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
			Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session(pcv_strSessionPrefix&"pcv_sdsBillingStateProvinceCode")
		end if

		'// Declare the instance number if greater than 1
		pcv_strFormInstance = "2"  

		'///////////////////////////////////////////////////////////
		'// END: COUNTRY AND STATE/ PROVINCE CONFIG
		'///////////////////////////////////////////////////////////
		%>
		
		<%
		'// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
		pcs_CountryDropdown
		%>	
		<tr>
			<td width="20%"><p>Address:</p></td>
			<td>
				<p><input type=text name="pcv_sdsBillingAddress" size="50" value="<% =pcf_FillFormField ("pcv_sdsBillingAddress", false) %>"></p>
			</td>
		</tr>
		<tr>
			<td><p>Address 2:</p></td>
			<td>
				<p><input type=text name="pcv_sdsBillingAddress2" size="50" value="<% =pcf_FillFormField ("pcv_sdsBillingAddress2", false) %>"></p>
			</td>
		</tr>
		<tr>
			<td><p>City:</p></td>
			<td>
				<p><input type=text name="pcv_sdsBillingCity" size="50" value="<% =pcf_FillFormField ("pcv_sdsBillingCity", false) %>"></p>
			</td>
		</tr>
		<%
		'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
		pcs_StateProvince
		%>
		<tr>
			<td><p>Postal Code:</p></td>
			<td>
				<p><input type=text name="pcv_sdsBillingZip" size="10" value="<% =pcf_FillFormField ("pcv_sdsBillingZip", false) %>"></p>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer" colspan="2"></td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"> 
			<input type="submit" name="modify" value="<%if pcv_clone="1" then%>Add<%else%>Update<%end if%>&nbsp;<%=pcv_Title%>" class="submit2">
            &nbsp;
            <input type="button" name="back" value="Back" onClick="javascript:history.back()" class="ibtnGrey"> 
			</td>
		</tr>
	</table>
	<input type=hidden name="pagetype" value="<%=pcv_PageType%>">
	<%if pcv_clone="0" then%>
		<input type=hidden name="idsds" value="<%=pcv_IDsds%>">
	<%end if%>
</form>
<script>
function isDropShipper(a) 
{
	if (a==true) {
		document.getElementById('show_1').style.display='';
		} else {
		document.getElementById('show_1').style.display='none';
	}
}
<% if pcv_PageType <> 1 then %>
var b = document.form1.pcv_sdsIsDropShipper.checked;
if (b==true) 
{
document.getElementById('show_1').style.display='';
}	
<% end if %>	
</script>
<%END IF ' if not add/edit action%>
<%call closedb()%><!--#include file="AdminFooter.asp"-->