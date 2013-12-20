<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="View &amp; Modify Customer" 
pageIcon="pcv4_icon_people.png"
section="mngAcc" 
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/encrypt.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<link href="../includes/spry/SpryTabbedPanels-PP.css" rel="stylesheet" type="text/css" />
<script src="../includes/spry/SpryTabbedPanels.js" type="text/javascript"></script>
<script src="../includes/spry/SpryURLUtils.js" type="text/javascript"></script>
<script type="text/javascript"> var params = Spry.Utils.getLocationParamsAsObject(); </script>
<!--#include file="../includes/stringfunctions.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<!--#include file="../includes/MailUpFunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<% 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: PAGE CONFIG
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
pidcustomer=trim(request("idcustomer"))

If Not validNum(pidcustomer) then
	response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query.")
	else
	Session("adminidcustomer")=pidcustomer
End If

dim conntemp, query, rs

call openDb()

'// Set Page Name
pcStrPageName = "modCusta.asp"

'MailUp-S

Dim MaxRequestTime,StopHTTPRequests

'maximum seconds for each HTTP request time
MaxRequestTime=5

StopHTTPRequests=0

'MailUp-E

'MAILUP-S

	tmp_setup=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	call opendb()
	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess FROM pcMailUpSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcMailUpSett_APIUser=rs("pcMailUpSett_APIUser")
		session("CP_MU_APIUser")=pcMailUpSett_APIUser
		pcMailUpSett_APIPassword=enDeCrypt(rs("pcMailUpSett_APIPassword"), scCrypPass)
		session("CP_MU_APIPassword")=pcMailUpSett_APIPassword
		pcMailUpSett_URL=rs("pcMailUpSett_URL")
		session("CP_MU_URL")=pcMailUpSett_URL
		tmp_Auto=rs("pcMailUpSett_AutoReg")
		if IsNull(tmp_Auto) or tmp_Auto="" then
			tmp_Auto=0
		end if
		session("CP_MU_Auto")=tmp_Auto
		tmp_setup=rs("pcMailUpSett_RegSuccess")
		if IsNull(tmp_setup) or tmp_setup="" then
			tmp_setup=0
		end if
		session("CP_MU_Setup")=tmp_setup
	end if
	set rs=nothing
	call closedb()

'MAILUP-E

call openDb()

'// Set Required Fields

'// Vat Settings
pcv_ShowVatId = false
pcv_isVatIdRequired = false
pcv_ShowSSN = false
pcv_isSSNRequired = false
if pshowVatID="1" then pcv_ShowVatId = true
if pVatIdReq="1" then pcv_isVatIdRequired = true
if pshowSSN="1" then pcv_ShowSSN = true
if pSSNReq="1" then pcv_isSSNRequired = true

'// General Info
pcv_isBillingFirstNameRequired = true
pcv_isBillingLastNameRequired = true
pcv_isBillingCompanyRequired = false
pcv_isBillingPhoneRequired = true
pcv_isBillingFaxRequired = false
pcv_isBillingEmailRequired = true

'// Billing
pcv_isBillingAddressRequired = true
pcv_isBillingPostalCodeRequired = true
pcv_isBillingCityRequired = true
pcv_isBillingCountryCodeRequired = true
pcv_isBillingAddress2Required = false

pcv_isBillingStateCodeRequired = true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isBillingStateCodeRequired=pcv_strStateCodeRequired
end if

pcv_isBillingProvinceRequired = false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isBillingProvinceRequired=pcv_strProvinceCodeRequired
end if

'// Shipping
pcv_isShipCompanyRequired=False
pcv_isShipAddressRequired=False
pcv_isShipCityRequired=False
pcv_isShipStateCodeRequired=False
pcv_isShipProvinceCodeRequired=False
pcv_isShipZipRequired=False
pcv_isShipCountryCodeRequired=False
pcv_isShipPhoneRequired=False
pcv_isShipFaxRequired=False
pcv_isShipEmailRequired=False

pcv_ispasswordRequired= False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: PAGE CONFIG
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ONLOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
' PRV41 begin
query="SELECT customers.idcustomer, customers.pcCust_VATID, customers.pcCust_SSN, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.email, customers.fax, customers.password, customers.address,customers.address2, customers.zip, customers.stateCode, customers.state, customers.city, customers.countryCode, customers.shippingCompany, customers.shippingaddress, customers.shippingAddress2, customers.shippingcity, customers.shippingStateCode, customers.shippingState, customers.shippingCountryCode, customers.shippingZip, customers.customerType, customers.TotalOrders, customers.TotalSales, customers.iRewardPointsAccrued, customers.iRewardPointsUsed,IDRefer,customers.RecvNews,customers.suspend,customers.idcustomerCategory,customers.pcCust_Locked, customers.pcCust_DateCreated, customers.ShippingEmail, customers.ShippingPhone, customers.pcCust_Notes, customers.pcCust_AllowReviewEmails FROM customers WHERE idcustomer="&pidcustomer&";"
' PRV41 end
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "viewCusta.asp?msg="& Server.Urlencode("The customer could not be found in the database. It might have been previously removed. Use the search filters below to locate another customer account.") 
end If	

if err.number <> 0 then
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
end If

' General Information
Session("pcAdminpcBillingFirstName")= pcf_ResetFormField(Session("pcAdminpcBillingFirstName"), rs("name"))
Session("pcAdminpcBillingLastName")= pcf_ResetFormField(Session("pcAdminpcBillingLastName"), rs("lastName"))
Session("pcAdminpcBillingCompany")= pcf_ResetFormField(Session("pcAdminpcBillingCompany"), rs("customerCompany"))
Session("pcAdminpcBillingPhone")= pcf_ResetFormField(Session("pcAdminpcBillingPhone"), rs("phone"))
Session("pcAdminpcBillingEmail")= pcf_ResetFormField(Session("pcAdminpcBillingEmail"), rs("email"))
	pcv_CustomerEmail=Session("pcAdminpcBillingEmail")
Session("pcAdminpcBillingFax")= pcf_ResetFormField(Session("pcAdminpcBillingFax"), rs("fax"))
mpassword=encrypt(enDeCrypt(rs("password"), scCrypPass), 9286803311968)
Session("pcAdminpassword")= pcf_ResetFormField(Session("pcAdminPassword"), "")

'// Billing
Session("pcAdminpcBillingAddress")= pcf_ResetFormField(Session("pcAdminpcBillingAddress"), rs("address"))
Session("pcAdminpcBillingAddress2")= pcf_ResetFormField(Session("pcAdminpcBillingAddress2"), rs("address2"))
Session("pcAdminpcBillingPostalCode")= pcf_ResetFormField(Session("pcAdminpcBillingPostalCode"), rs("zip"))
Session("pcAdminpcBillingStateCode")= pcf_ResetFormField(Session("pcAdminpcBillingStateCode"), rs("stateCode"))
Session("pcAdminpcBillingProvince")= pcf_ResetFormField(Session("pcAdminpcBillingProvince"), rs("state"))
Session("pcAdminpcBillingCity")= pcf_ResetFormField(Session("pcAdminpcBillingCity"), rs("city"))
Session("pcAdminpcBillingCountryCode")= pcf_ResetFormField(Session("pcAdminpcBillingCountryCode"), rs("countryCode"))
Session("pcAdminpcBillingVATID")= pcf_ResetFormField(Session("pcAdminpcBillingVATID"), rs("pcCust_VATID"))
Session("pcAdminpcBillingSSN")= pcf_ResetFormField(Session("pcAdminpcBillingSSN"), rs("pcCust_SSN"))


'// Shipping
Session("pcAdminShipCompany") = pcf_ResetFormField(Session("pcAdminShipCompany"), rs("shippingCompany"))
Session("pcAdminShipAddress") = pcf_ResetFormField(Session("pcAdminShipAddress"), rs("shippingaddress"))
Session("pcAdminShipAddress2") = pcf_ResetFormField(Session("pcAdminShipAddress2"), rs("shippingaddress2"))
Session("pcAdminShipCity") = pcf_ResetFormField(Session("pcAdminShipCity"), rs("shippingcity"))	
Session("pcAdminShipStateCode") = pcf_ResetFormField(Session("pcAdminShipStateCode"), rs("shippingstateCode"))
Session("pcAdminShipState") = pcf_ResetFormField(Session("pcAdminShipState"),rs("shippingstate"))
Session("pcAdminShipCountryCode") = pcf_ResetFormField(Session("pcAdminShipCountryCode"), rs("shippingcountryCode"))	
Session("pcAdminShipZip") = pcf_ResetFormField(Session("pcAdminShipZip"), rs("shippingzip"))
Session("pcAdminShipEmail") = pcf_ResetFormField(Session("pcAdminShipEmail"), rs("shippingEmail"))
Session("pcAdminShipPhone") = pcf_ResetFormField(Session("pcAdminShipPhone"), rs("shippingPhone"))

'// Customer Type
Session("pcAdmincustomertype")=pcf_ResetFormField(Session("pcAdmincustomerType"), rs("customerType"))
Session("pcAdminCustLocked")=pcf_ResetFormField(Session("pcAdminCustLocked"), rs("pcCust_Locked"))


'// Misc.
Session("pcAdmintotalorders")=rs("TotalOrders")
Session("pcAdmintotalsales")=rs("TotalSales")
pcvCust_DateCreated=rs("pcCust_DateCreated")
' PRV41 begin
Session("pcAllowReviewEmails")=rs("pcCust_AllowReviewEmails")
' PRV41 end

'// Suspend Account
Session("pcAdminsuspend")=pcf_ResetFormField(Session("pcAdminsuspend"), rs("suspend"))

intIdcustomerCategory=rs("idcustomerCategory")
if isNULL(intIdcustomerCategory) OR intIdcustomerCategory="" OR intIdcustomerCategory=0 then
else
	Session("pcAdmincustomertype")="CC_"&intIdcustomerCategory
end if

' Reward Points
If RewardsActive <> 0 then
	Session("pcAdminiRewardPointsAccrued")=Int(rs("iRewardPointsAccrued")}
	if len(int(Session("pcAdminiRewardPointsAccrued"))&"A")=1 then
		Session("pcAdminiRewardPointsAccrued")=0
	end if
	
	Session("pcAdminiRewardPointsUsed")=Int(trim(rs("iRewardPointsUsed")))
	if len(int(Session("pcAdminiRewardPointsUsed"))&"A")=1 then
		Session("pcAdminiRewardPointsUsed")=0
	end if

	Session("pcAdminiBalance")=int(Session("pcAdminiRewardPointsAccrued")-Session("pcAdminiRewardPointsUsed"))
end if
Session("IDRefer")=rs("IDRefer")	
Session("pcAdminCRecvNews")=pcf_ResetFormField(Session("pcAdminCRecvNews"), rs("RecvNews"))	
Session("pcAdminCustNotes")=pcf_ResetFormField(Session("pcAdminCustNotes"), rs("pcCust_Notes"))	
set rs=nothing


'// Start Special Customer Fields
session("cp_nc_custfields")=""
session("cp_nc_custfields_exists")=""
query="SELECT pcCField_ID,pcCField_Name,pcCField_FieldType,pcCField_Value,pcCField_Length,pcCField_Maximum,pcCField_Required,pcCField_PricingCategories,pcCField_ShowOnReg,pcCField_ShowOnCheckout,'' FROM pcCustomerFields ORDER BY pcCField_Order ASC, pcCField_Name ASC;"
set rs=connTemp.execute(query)
if not rs.eof then
	session("cp_nc_custfields")=rs.GetRows()
	session("cp_nc_custfields_exists")="YES"
end if
set rs=nothing

if session("cp_nc_custfields_exists")="YES" then
	pcArr=session("cp_nc_custfields")
	For k=0 to ubound(pcArr,2)
		pcArr(10,k)=""
		query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & pidcustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
		set rs=connTemp.execute(query)
		if not rs.eof then
			pcArr(10,k)=rs("pcCFV_Value")
		end if
		set rs=nothing
	Next
	session("cp_nc_custfields")=pcArr
end if	
'/  End of Special Customer Fields	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ONLOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IF request.Form("Modify")<>"" THEN
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	
	'// General Info
	pcs_ValidateTextField "pcBillingFirstName", pcv_isBillingFirstNameRequired, 0
	pcs_ValidateTextField "pcBillingLastName", pcv_isBillingLastNameRequired, 0
	pcs_ValidateTextField "pcBillingCompany", pcv_isBillingCompanyRequired, 0
	pcs_ValidatePhoneNumber "pcBillingPhone", pcv_isBillingPhoneRequired, 0
	pcs_ValidatePhoneNumber "pcBillingFax", pcv_isBillingFaxRequired, 0
	pcs_ValidateEmailField "pcBillingEmail", pcv_isBillingEmailRequired, 0
	
	'// Billing
	pcs_ValidateTextField "pcBillingAddress", pcv_isBillingAddressRequired, 0
	pcs_ValidateTextField "pcBillingPostalCode", pcv_isBillingPostalCodeRequired, 0
	pcs_ValidateTextField "pcBillingStateCode", pcv_isBillingStateCodeRequired, 0
	pcs_ValidateTextField "pcBillingProvince", pcv_isBillingProvinceRequired, 0
	pcs_ValidateTextField "pcBillingCity", pcv_isBillingCityRequired, 0
	pcs_ValidateTextField "pcBillingCountryCode", pcv_isBillingCountryCodeRequired, 0
	pcs_ValidateTextField "pcBillingAddress2", pcv_isBillingAddress2Required, 0
	
	'// VATID
	If pcv_ShowVatId = True Then
		pcs_ValidateVATIDField "pcBillingVATID", pcv_isVATIDRequired, getUserInput(request("pcBillingCountryCode"),0)
	End If		
	
	'// SSN
	If pcv_ShowSSN = True Then
		pcs_ValidateSSNField "pcBillingSSN", pcv_isSSNRequired, getUserInput(request("pcBillingCountryCode"),0)
	End If

	'// Shipping
	pcs_ValidateTextField "ShipCompany", pcv_isShipCompanyRequired, 0
	pcs_ValidateTextField "ShipAddress", pcv_isShipAddressRequired, 0
	pcs_ValidateTextField "ShipAddress2", pcv_isShipAddressRequired, 0
	pcs_ValidateTextField "ShipCity", pcv_isShipCityRequired, 0
	pcs_ValidateTextField "ShipState", pcv_isShipProvinceCodeRequired, 0
	pcs_ValidateTextField "ShipStateCode", pcv_isShipStateCodeRequired, 0
	pcs_ValidateTextField "ShipZip", pcv_isShipZipRequired, 0
	pcs_ValidateTextField "ShipCountryCode", pcv_isShipCountryCodeRequired, 0	
    pcs_ValidateEmailField "ShipEmail", pcv_isShipEmailRequired, 0	
	pcs_ValidateTextField "ShipPhone", pcv_isShipPhoneRequired, 0	
	
	'// Misc.
	pcs_ValidateTextField "password", pcv_ispasswordRequired, 0
	pcs_ValidateTextField "suspend", false, 0
	pcs_ValidateTextField "ResetPass", false, 0
	
	'// Customer Type
	pcs_ValidateTextField "customerType", true, 0
	if request.form("lock")="1" then
		Session("pcAdminCustLocked")="1"
	else
		Session("pcAdminCustLocked")="0"
	end if
	
	'// News
	pcs_ValidateTextField "CRecvNews", false, 0
	'MAILUP-S
	IF session("CP_MU_Setup")="1" THEN
	Session("pcAdminCRecvNews")="0"
	Session("pcAdminpcNewsListCount")=""
	tmpNewsListCount=getUserInput(request("newslistcount"),0)
	if tmpNewsListCount<>"" then
		Session("pcAdminpcNewsListCount")=tmpNewsListCount
		For j=0 to tmpNewsListCount
			Session("pcAdminpcNewsList" & j)=getUserInput(request("newslist" & j),0)
			if Session("pcAdminpcNewsList" & j)<>"" then
				Session("pcAdminCRecvNews")="1"
			end if
		Next
	end if
	END IF
	'MAILUP-E

	'// Adjustment
	pcs_ValidateTextField "iAdjustment", false, 0
	
	'// Rewared Points
	pcs_ValidateTextField "iRewardPointsAccrued", false, 0
	
	'// Comments
	pcs_ValidateTextField "custnotes", false, 0

    ' PRV41 begin
	'// Misc
	if request.form("allowreviewemails")="1" then
		Session("pcAllowReviewEmails")="1"
	else
		Session("pcAllowReviewEmails")="0"
	end If
	' PRV41 end
		
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError&"&idcustomer="&pidcustomer
	Else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' News	
		if Session("pcAdminCRecvNews")="" then
			Session("pcAdminCRecvNews")=0
		end if
		
		'// Suspend		
		if Session("pcAdminsuspend")="" then
			Session("pcAdminsuspend")="0"
		end if
		
		'//Reset Password
		if Session("pcAdminResetPass")="" then
			Session("pcAdminResetPass")="0"
		end if
		
		'// Password
		if Session("pcAdminResetPass")="1" AND Session("pcAdminpassword")<>"" then
			Session("pcAdminpassword")	= enDeCrypt(Session("pcAdminpassword"), scCrypPass)
		end if
		
		' Email Already in Database
		' IF customer is a registered customer (not a guest) or a guest converting to a registered customer enforce unique e-mail
		if pcf_GetCustType(pidcustomer)=0 then queryTemp= "pcCust_Guest=0 AND "
		if pcf_GetCustType(pidcustomer)=0 or (pcf_GetCustType(pidcustomer)<>0 and Session("pcAdminpassword")<>"") then
			query="SELECT email FROM customers WHERE " & queryTemp & "email='"&Session("pcAdminpcBillingEmail")&"' AND idCustomer<>"&Session("adminidcustomer")&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if NOT rs.eof then
					response.redirect pcStrPageName&"?msg=" & Server.URLEncode("The email you have chosen is already in use by another customer. If you still wish to use this e-mail for this customer account, <a href='viewcustb.asp?key4=" & Session("pcAdminpcBillingEmail") & "'>search</a> for all customers with the same e-mail address, consolidate their accounts into one using the corresponding feature (orders are moved to the consolidated account), and then remove the accounts that are no longer needed.") & "&idcustomer=" & pidcustomer
			end if	
		end if
		
		'// Rewared Points		
		If Session("pcAdminiRewardPointsAccrued")="" then
			Session("pcAdminiRewardPointsAccrued")="0"
		End If	

		
		'If Customer type is not retail or wholesale (0 or 1) then we must find customer type from database and find if wholesale priv is active.
		intidcustomerCategory=0
		If instr(Session("pcAdminCustomerType"),"CC") then
			intidcustomerCategory=replace(Session("pcAdminCustomerType"),"CC_","")
			intidcustomerCategory=int(intidcustomerCategory)
			query="SELECT pcCustomerCategories.idcustomerCategory, pcCustomerCategories.pcCC_WholesalePriv FROM pcCustomerCategories WHERE (((pcCustomerCategories.idcustomerCategory)="&intidcustomerCategory&"));"	
			SET rs=Server.CreateObject("ADODB.RecordSet")		
			SET rs=conntemp.execute(query)
			intpcCC_WholesalePriv=rs("pcCC_WholesalePriv")
			if intpcCC_WholesalePriv=1 then
				Session("pcAdminCustomerType")=1
			else
				Session("pcAdminCustomerType")=0
			end if
		end if
			   
		' update customer record
		query="UPDATE customers SET pcCust_VATID='" &Session("pcAdminpcBillingVATID")& "', pcCust_SSN='" &Session("pcAdminpcBillingSSN")& "', name='" &Session("pcAdminpcBillingFirstName")& "', lastName='" &Session("pcAdminpcBillingLastName")& "', email='" &Session("pcAdminpcBillingEmail")& "', fax='" &Session("pcAdminpcBillingFax")& "', city='" &Session("pcAdminpcBillingCity")& "', zip='" &Session("pcAdminpcBillingPostalCode")& "', countryCode='" &Session("pcAdminpcBillingCountryCode")& "', state='" &Session("pcAdminpcBillingProvince")& "', stateCode='" &Session("pcAdminpcBillingStateCode")& "', shippingcity='" &Session("pcAdminShipCity")& "', shippingzip='" &Session("pcAdminShipZip")& "', shippingcountryCode='" &Session("pcAdminShipCountryCode")& "', shippingstate='" &Session("pcAdminShipState")& "', shippingstateCode='" &Session("pcAdminShipStateCode")& "', phone='" &Session("pcAdminpcBillingPhone")& "', address='" &Session("pcAdminpcBillingAddress")& "', address2='" &Session("pcAdminpcBillingAddress2")& "', shippingCompany='" &Session("pcAdminShipCompany")& "', shippingaddress='" &Session("pcAdminShipAddress")& "', shippingaddress2='" &Session("pcAdminShipAddress2")& "', customercompany='" &Session("pcAdminpcBillingCompany")& "', customerType=" &Session("pcAdminCustomerType")& ", shippingEmail='" &Session("pcAdminShipEmail")& "', shippingPhone='" &Session("pcAdminShipPhone")& "'"
		
		if RewardsActive <> 0 then			
			if Session("pcAdminiAdjustment") <> "" and isNumeric(Session("pcAdminiAdjustment")) then
				Session("pcAdminiRewardPointsAccrued")=int(Session("pcAdminiRewardPointsAccrued"))+int(Session("pcAdminiAdjustment"))
			end if
			query=query&", iRewardPointsAccrued="&Session("pcAdminiRewardPointsAccrued")
		end if
		
		'//Add New Password
		if Session("pcAdminResetPass")="1" AND Session("pcAdminpassword")<>"" then
			query=query&", [password]='" &Session("pcAdminpassword")&"', pcCust_Guest=0"
		end if
			
	    ' PRV41 begin
		query=query&", RecvNews="&Session("pcAdminCRecvNews")&", suspend="&Session("pcAdminsuspend")&", idcustomerCategory=" & intidcustomerCategory&", pcCust_Locked=" & Session("pcAdminCustLocked") & ", pcCust_Notes='" & Session("pcAdminCustNotes") & "', pcCust_AllowReviewEmails=" & Session("pcAllowReviewEmails") & " WHERE idCustomer="&Session("adminidcustomer")
		' PRV41 end

		set rs=conntemp.execute(query)
		
		if err.number <> 0 then
			set rs=nothing	
			call closedb()					
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
		end if
		
		call updCustEditedDate(Session("adminidcustomer"))
		
		'Start Special Customer Fields
			if session("cp_nc_custfields_exists")="YES" then
				pcArr=session("cp_nc_custfields")
				For k=0 to ubound(pcArr,2)
					tmp_cf=""
					tmp_cf=request.form("custfield_" & pcArr(0,k))
					if not IsNull(tmp_cf) then
						tmp_cf=replace(tmp_cf,"'","''")
					end if
					pcArr(3,k)=tmp_cf
				Next
		
				pcv_IDCustomer=Session("adminidcustomer")
		
				For k=0 to ubound(pcArr,2)
					query="SELECT pcCField_ID FROM pcCustomerFieldsValues WHERE idcustomer=" & pcv_IDCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
					set rs=connTemp.execute(query)
					if not rs.eof then
						query="UPDATE pcCustomerFieldsValues SET pcCFV_Value='" & pcArr(3,k) & "' WHERE idcustomer=" & pcv_IDCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
					else
						query="INSERT INTO pcCustomerFieldsValues (idcustomer,pcCField_ID,pcCFV_Value) VALUES (" & pcv_IDCustomer & "," & pcArr(0,k) & ",'" & pcArr(3,k) & "');"
					end if
					set rs=connTemp.execute(query)
					set rs=nothing
				Next
				
				session("cp_nc_custfields")=""
			end if
		'End of Special Customer Fields
		
		pcv_IDCustomer=Session("adminidcustomer")
		'MAILUP-S
			MUResult=1
			tmpNewsListCount=Session("pcAdminpcNewsListCount")
			if tmpNewsListCount<>"" then
				For j=0 to tmpNewsListCount
					if Session("pcAdminpcNewsList" & j)<>"" then
							query="SELECT pcMailUpLists_ListID,pcMailUpLists_ListGuid FROM pcMailUpLists WHERE pcMailUpLists_ID=" & Session("pcAdminpcNewsList" & j) & ";"
							set rs=connTemp.execute(query)
							ListID=rs("pcMailUpLists_ListID")
							ListGuid=rs("pcMailUpLists_ListGuid")
							tmpMUResult=UpdUserReg(pcv_IDCustomer,Session("pcAdminpcBillingEmail"),ListID,ListGuid,session("CP_MU_URL"),session("CP_MU_Auto"))
							if tmpMUResult=0 then
								MUResult=0
							end if
					end if
				Next
				query="SELECT pcMailUpLists_ID FROM pcMailUpSubs WHERE idCustomer=" & pcv_IDCustomer & ";"
				set rs=connTemp.execute(query)
				if not rs.eof then
					tmpArr=rs.getRows()
					intCount=ubound(tmpArr,2)
					For j=0 to intCount
						tmpRmv=1
						For k=0 to tmpNewsListCount
							if Session("pcAdminpcNewsList" & k)<>"" then
								if Clng(Session("pcAdminpcNewsList" & k))=Clng(tmpArr(0,j)) then
									tmpRmv=0
									exit for
								end if
							end if
						Next
						if tmpRmv=1 then
							query="SELECT pcMailUpLists_ListID,pcMailUpLists_ListGuid FROM pcMailUpLists WHERE pcMailUpLists_ID=" & tmpArr(0,j) & ";"
							set rs=connTemp.execute(query)
							ListID=rs("pcMailUpLists_ListID")
							ListGuid=rs("pcMailUpLists_ListGuid")
							tmpMUResult=UnsubUser(pcv_IDCustomer,Session("pcAdminpcBillingEmail"),ListID,ListGuid,session("CP_MU_URL"),session("CP_MU_Auto"))
							if tmpMUResult=0 then
								MUResult=0
							end if
						end if
					Next
				end if
				set rs=nothing
			end if
		'MAILUP-E
		set rs=nothing	
		call closedb()	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Clear the sessions
		pcs_ClearAllSessions
		Session("pcAdminCustLocked")=""
		
		'// Redirect
		If session("CP_MU_Auto")="0" then
			MUResult=1
		end if
		response.redirect pcStrPageName&"?idcustomer="&Session("adminidcustomer") & "&s=1&msg=Customer data has been updated successfully!&mailup=" & MUResult
		
	End If	
END IF	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<script language="JavaScript">
<!--	
function Form1_Validator(theForm)
{
<%'Start Special Customer Fields
	if session("cp_nc_custfields_exists")="YES" then
		pcArr=session("cp_nc_custfields")
		For k=0 to ubound(pcArr,2)
			if pcArr(6,k)="1" then%>
			if (theForm.custfield_<%=pcArr(0,k)%>.value == "")
		  	{
				<%if pcArr(0,k)="1" then%>
					alert("Please select the option.");
				<%else%>
					alert("Please enter a value for this field.");
				<%end if%>
			    theForm.custfield_<%=pcArr(0,k)%>.focus();
			    return (false);
			}
			<%end if
		Next
	end if
'End of Special Customer Fields%>
	
return (true);
}
//-->
</script>

<!--#include file="pcv4_showMessage.asp"-->

<div class="cpOtherLinks" style="margin: 0 10px 0 7px;"><% if pcf_GetCustType(pidcustomer)=0 then %><% mpassword=Server.URLEncode(mpassword) %><a href="adminPlaceOrder.asp?idcustomer=<%=pidcustomer%>" target="_blank" class="pcPageNav">Place an Order</a> | <% end if %><a href="viewCustOrders.asp?idcustomer=<%=Session("adminidcustomer")%>">View Orders</a> | <a href="adminviewallmsgs.asp?idcustomer=<%=Session("adminidcustomer")%>">View Help Desk Messages</a> | <a href="ggg_manageGRs.asp?idcustomer=<%=Session("adminidcustomer")%>">View Gift Registries</a> | <a href="pushOrdersA.asp?idcustomer=<%=Session("adminidcustomer")%>">Consolidate Accounts</a> | <a href="manageCustFields.asp" target="_blank">Special Fields</a> | <a href="AdminCustomerCategory.asp" target="_blank">Pricing Categories</a></div>

<script>var tmpNListChecked=0;</script>
<form method="post" name="modCust" action="<%=pcStrPageName%>?idcustomer=<%=pidcustomer%>" onSubmit="<%if session("CP_MU_Auto")="1" then%>if (tmpNListChecked) pcf_Open_MailUp();<%end if%> return Form1_Validator(this)" class="pcForms">
		<%
		'// TABBED PANELS - MAIN DIV START
		%>
	  <div id="TabbedPanels1" class="VTabbedPanels">
		
		<%
		'// TABBED PANELS - START NAVIGATION
		%>
			<ul class="TabbedPanelsTabGroup">
				<li class="TabbedPanelsTab" tabindex="100">General Information</li>
                <li class="TabbedPanelsTab" tabindex="100">Default Billing Address</li>
                <li class="TabbedPanelsTab" tabindex="100">Default Shipping Address</li>
                <li class="TabbedPanelsTab" tabindex="100">Other Settings</li>
				<li class="TabbedPanelsTabButtons" tabindex="1200">
                    <input type="submit" name="Modify" value="Modify" class="submit2"> 
                    <div style="padding-top: 6px">
                    <input type="button" name="back" value="Back" onClick="JavaScript:history.go(-1);">&nbsp;
                    <input type="button" name="delete" value="Delete" onClick="javascript:if (confirm('Please note: the following will occur if you click on OK. If there ARE NO orders associated with this account, the customer account will be permanently deleted from the database. If the customer had initiated, but not completed one or more orders (i.e. incomplete orders), they will also be deleted. If there ARE orders associated with this account, you will be prompted to either Remove the customer account, but keep the associated orders in the database, OR Delete the customer account and all of the associated orders. Are you sure to want to continue?')) location='delCustomer.asp?idcustomer=<%=Session("adminidcustomer")%>'">
                    </div>
                    <div style="padding-top: 12px;">
                    <input type="button" name="Search" value="Locate Another" onClick="document.location.href='viewcusta.asp';">&nbsp;
                    <input type="button" name="AddNew" value="Add" onClick="document.location.href='instcusta.asp';">
                    </div>
                </li>
            </ul>
            
		<%
		'// TABBED PANELS - END NAVIGATION
		
		'// TABBED PANELS - START PANELS
		%>
		
			<div class="TabbedPanelsContentGroup">
			
			<%
			'// =========================================
			'// FIRST PANEL - START - Name, SKU, descriptions
			'// =========================================
			%>
				<div class="TabbedPanelsContent">
				
					<table class="pcCPcontent">				
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
                        <tr> 
                            <th colspan="2">You are editing: 
                            <strong>
							<% response.write Session("pcAdminpcBillingFirstName") & " " & Session("pcAdminpcBillingLastName")%>
                            <% if trim(Session("pcAdminpcBillingCompany"))<>"" then response.write " - " & Session("pcAdminpcBillingCompany") %>
                            </strong>
                            </th>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <% 
						' Calculate customer number using sccustpre constant
						Dim pcCustomerNumber
						pcCustomerNumber = (sccustpre + int(pidcustomer))
						if sccustpre > 0 then
                        %>
                        <tr> 
                            <td><p>Customer Number:</p></td>
                            <td><p><%=pcCustomerNumber%> <span style="margin-left: 50px">Customer ID (<em>database</em>): <%=pidcustomer%></span></p></td>
                        </tr>
                        <% else %>
                        <tr> 
                            <td><p>Customer ID:</p></td>
                            <td><p><%=pidcustomer%></p></td>
                        </tr>
                        <% end if %>
                        
                        
                        <tr> 
                            <td><p>Customer Type:</p></td>
                            <td><p>
                                <select name="customerType">
                                    <option value='0' 
                                    <% if Session("pcAdmincustomertype")="0" then 
                                         response.write "selected"
                                    end if%>
                                    >Retail Customer</option>
                                    <option value='1'
                                    <%if Session("pcAdmincustomertype")="1" then 
                                        response.write "selected"
                                    end if%>
                                    >Wholesale Customer</option>
                
                                    <% 'START CT ADD %>
                                    <% 'if there are PBP customer type categories - List them here 
                                    
                                    query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
                                    SET rs=Server.CreateObject("ADODB.RecordSet")
                                    SET rs=conntemp.execute(query)
                                    if NOT rs.eof then 
                                        do until rs.eof 
                                            intIdcustomerCategory=rs("idcustomerCategory")
                                            strpcCC_Name=rs("pcCC_Name")
                                            %>
                                            <option value='CC_<%=intIdcustomerCategory%>'
                                            <%if Session("pcAdmincustomertype")="CC_"&intIdcustomerCategory then 
                                                response.write "selected"
                                            end if%>
                                            ><%=strpcCC_Name%></option>
                                            <% rs.moveNext
                                        loop
                                    end if
                                    SET rs=nothing
                                    
                                    'END CT ADD %>
                                </select>
                                </p>
                            </td>
                        </tr>
                        
                        <%if pcf_GetCustType(pidcustomer)=0 then%>
						<tr>
                        	<td><p>Customer Status</p></td>
                            <td><p><strong>Registered</strong> - Customer registered an account (saved password)</p></td>
                        </tr>
                        <% else %>
						<tr>
                        	<td valign="top"><p>Customer Status</p></td>
                            <td valign="top"><p><strong>Guest</strong> - Customer did not register an account (did not save a password)</p><p>The <em>Place Order</em> feature is hidden as you cannot use it for a <em>Guest</em>.</td>
                        </tr>			
						<% end if %>
                        
                        <% if pcvCust_DateCreated<>"" then %>
						<tr>
                        	<td><p>Created on</p></td>
                            <td><p>This customer account was first saved to the database on <strong><%=ShowDateFrmt(pcvCust_DateCreated)%></strong></p></td>
                        </tr>
                        <% end if %>
                        
                        <tr>
                        	<td colspan="2"><hr></td>
                        </tr>
                        <tr>
                            <td><p>
                                <%response.write dictLanguage.Item(Session("language")&"_order_C")%>
                            </p></td>
                            <td><p>
                                <input type="text" name="pcBillingFirstName" value="<% =pcf_FillFormField ("pcBillingFirstName", pcv_isBillingFirstNameRequired) %>" size="20" />
                                <%pcs_RequiredImageTag "pcBillingFirstName", pcv_isBillingFirstNameRequired %>
                            </p></td>
                        </tr>
                        <tr>
                            <td><p>
                                <%response.write dictLanguage.Item(Session("language")&"_order_D")%>
                            </p></td>
                            <td><p>
                                <input type="text" name="pcBillingLastName" value="<% =pcf_FillFormField ("pcBillingLastName", pcv_isBillingLastNameRequired) %>" size="20" />
                                <%pcs_RequiredImageTag "pcBillingLastName", pcv_isBillingLastNameRequired %>
                            </p></td>
                        </tr>
                        <tr>
                            <td><p>
                                <%response.write dictLanguage.Item(Session("language")&"_order_E")%>
                            </p></td>
                            <td><p>
                                <input type="text" name="pcBillingCompany" value="<% =pcf_FillFormField ("pcBillingCompany", pcv_isBillingCompanyRequired) %>" size="30" />
                                <%pcs_RequiredImageTag "pcBillingCompany", pcv_isBillingCompanyRequired %>
                            </p></td>
                        </tr>
                
                        <% if pcv_ShowVatId = True then %>
                            <%
                            if session("ErrpcBillingVATID")<>"" then %>
                                <tr> 
                                    <td></td>
                                    <td><p><img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> <%=dictLanguage.Item(Session("language")&"_Custmoda_27")%></p></td>
                                </tr>
                                <% session("ErrpcBillingVATID") = ""
                            end if
                            %>									
                                    
                            <tr>
                                <td><p><%=dictLanguage.Item(Session("language")&"_Custmoda_26")%></p></td>
                                <td><p><input type="text" name="pcBillingVATID" value="<%=pcf_FillFormField ("pcBillingVATID", pcv_isVatIdRequired) %>" ID="Text1">
                                <% pcs_RequiredImageTag "pcBillingVATID", pcv_isVatIdRequired  %></p>
                                </td>
                            </tr>
                        <% end if %>	
                    
                    
                        <% if pcv_ShowSSN = True then %>	
                            <% 									
                            if session("ErrpcBillingSSN")<>"" then %>
                                <tr> 
                                    <td>&nbsp;</td>
                                    <td><p><img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> <%=dictLanguage.Item(Session("language")&"_Custmoda_25")%></p></td>
                                </tr>
                                <% session("ErrpcBillingSSN") = ""
                            end if 
                            %>	
                            <tr>
                                <td><p><%=dictLanguage.Item(Session("language")&"_Custmoda_24")%></p></td>
                                <td><p id="spanValueSSN"><input type="text" name="pcBillingSSN" value="<%=pcf_FillFormField ("pcBillingSSN", pcv_isSSNRequired) %>" ID="Text2">
                                <% pcs_RequiredImageTag "pcBillingSSN", pcv_isSSNRequired %>
                                </p></td>
                            </tr>
                        <% end if %>
                
                
                        <%	'// Phone Custom Error
                        if session("ErrpcBillingPhone")<>"" then %>
                        <tr>
                            <td>&nbsp;</td>
                            <td><img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%></td>
                        </tr>
                        <% 
                        session("ErrpcBillingPhone")=""
                        end if 
                        %>
                                    
                        <tr>
                            <td><p>
                                <%response.write dictLanguage.Item(Session("language")&"_order_F")%>
                            </p></td>
                            <td><p>
                                <input type="text" name="pcBillingPhone" value="<% =pcf_FillFormField ("pcBillingPhone", pcv_isBillingPhoneRequired) %>" size="15" />
                                <%pcs_RequiredImageTag "pcBillingPhone", pcv_isBillingPhoneRequired %>
                            </p></td>
                        </tr>
                        <tr>
                            <td><p>
                                <%response.write dictLanguage.Item(Session("language")&"_order_AA")%>
                            </p></td>
                            <td><p>
                                <input type="text" name="pcBillingFax" value="<% =pcf_FillFormField ("pcBillingFax", pcv_isBillingFaxRequired) %>" size="15" />
                                <%pcs_RequiredImageTag "pcBillingFax", pcv_isBillingFaxRequired %>
                            </p></td>
                        </tr>
                
                        <% if Session("pcAdminErremail")<>"" then %>
                            <tr> 
                                <td>&nbsp;</td>
                                <td><img src="images/next.gif" width="10" height="10"> <%=Session("pcAdminErremail")%></td>
                            </tr>
                        <% end if %>
                        <tr> 
                            <td><p>E-mail:</p></td>
                            <td>
                            <p>
                                <input type="text" name="pcBillingEmail" value="<% =pcf_FillFormField ("pcBillingEmail", pcv_isBillingEmailRequired) %>" size="25" maxlength="150">
                                <%pcs_RequiredImageTag "pcBillingEmail", pcv_isBillingEmailRequired %>
                            </p>
                            </td>
                        </tr>
                        
                        <%
						'// Reset or Add Password
						'// Customer is a registered customer = show password reset
						if pcf_GetCustType(pidcustomer)=0 then %>
                        <tr> 
                            <td><p>Password:</p></td>
                            <td><p><i>The password is not shown for privacy reasons.</i></p></td>
                        </tr>
                        <% 
						else 
						%>
                        <tr> 
                            <td><p>Password:</p></td>
                            <td><p><i>This is a <strong>Guest Account</strong>. Add a password to turn it into a Registered Customer.</i></p></td>
                        </tr>
                        <%
						end if
						%>
                        <tr>
                            <td></td>
                            <td  valign="top">
                                <p>
								<input type="checkbox" name="ResetPass" value="1" <% if Session("pcAdminResetPass")="1" then%>checked<% end if %> onclick="javascript: if (this.checked) {document.modCust.password.disabled=false} else {document.modCust.password.disabled=true; document.modCust.password.value=''};" class="clearBorder">
								<% if pcf_GetCustType(pidcustomer)=0 then %>Reset Password<% else %>Add Password<%end if%>:&nbsp;<input type="password" name="password" <% if Session("pcAdminResetPass")<>"1" then%>disabled<% end if %> value="<% if Session("pcAdminResetPass")="1" then%><% =pcf_FillFormField ("password", pcv_ispasswordRequired) %><%end if%>" size="25" maxlength="50">
                                </p>
                            </td>
                        </tr>
                        <tr>
                        	<td colspan="2"><hr></td>
                        </tr>
						<tr>
							<td valign="top">Administrator Comments:</td>
							<td><textarea name="custnotes" rows="6" cols="60"><%=Session("pcAdminCustNotes")%></textarea></td>
						</tr>                   
					</table>
					
				</div>
			<%
			'// =========================================
			'// FIRST PANEL - END
			'// =========================================
			
			'// =========================================
			'// SECOND PANEL - START - Default Billing
			'// =========================================
			%>
				<div class="TabbedPanelsContent">
				
					<table class="pcCPcontent">
					
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <th colspan="2">Billing Address</th>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
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
                        pcv_isStateCodeRequired = pcv_isBillingStateCodeRequired '// determines if validation is performed (true or false)
                        pcv_isProvinceCodeRequired = pcv_isBillingProvinceRequired '// determines if validation is performed (true or false)
                        pcv_isCountryCodeRequired = pcv_isBillingCountryCodeRequired '// determines if validation is performed (true or false)					
                        
                        '// #3 Additional Required Info
                        pcv_strTargetForm = "modCust" '// Name of Form
                        pcv_strCountryBox = "pcBillingCountryCode" '// Name of Country Dropdown
                        pcv_strTargetBox = "pcBillingStateCode" '// Name of State Dropdown
                        pcv_strProvinceBox =  "pcBillingProvince" '// Name of Province Field
                        
                        '// Set local Country to Session
                        if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
                            Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
                        end if
                        
                        '// Set local State to Session
                        if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
                            Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session(pcv_strSessionPrefix&pcv_strTargetBox)
                        end if
                        
                        '// Set local Province to Session
                        if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
                            Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session(pcv_strSessionPrefix&pcv_strProvinceBox)
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
                            <td><p>
                                <%response.write dictLanguage.Item(Session("language")&"_order_K")%>
                            </p></td>
                            <td><p>
                                <input type="text" name="pcBillingAddress" value="<% =pcf_FillFormField ("pcBillingAddress", pcv_isBillingAddressRequired) %>" size="30" />
                                <%pcs_RequiredImageTag "pcBillingAddress", pcv_isBillingAddressRequired %>
                            </p></td>
                        </tr>
                        <tr>
                            <td><p>&nbsp;</p></td>
                            <td><p>
                                <input type="text" name="pcBillingAddress2" value="<% =pcf_FillFormField ("pcBillingAddress2", pcv_isBillingAddress2Required) %>" size="30" />
                                <%pcs_RequiredImageTag "pcBillingAddress2", pcv_isBillingAddress2Required %>
                            </p></td>
                        </tr>
                        <tr>
                            <td><p>
                                <%response.write dictLanguage.Item(Session("language")&"_order_L")%>
                            </p></td>
                            <td><p>
                                <input type="text" name="pcBillingCity" value="<% =pcf_FillFormField ("pcBillingCity", pcv_isBillingCityRequired) %>" size="30" />
                                <%pcs_RequiredImageTag "pcBillingCity", pcv_isBillingCityRequired %>
                            </p></td>
                        </tr>
                        
                        <%
                        '// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
                        pcs_StateProvince
                        %>
                        
                        <tr>
                            <td><p>
                                <%response.write dictLanguage.Item(Session("language")&"_order_O")%>
                            </p></td>
                            <td><p>
                                <input type="text" name="pcBillingPostalCode" value="<% =pcf_FillFormField ("pcBillingPostalCode", pcv_isBillingPostalCodeRequired) %>" size="10" />
                                <%pcs_RequiredImageTag "pcBillingPostalCode", pcv_isBillingPostalCodeRequired %>
                            </p></td>
                        </tr>
					</table>
					
				</div>
			<%
			'// =========================================
			'// SECOND PANEL - END
			'// =========================================
			
			'// =========================================
			'// THIRD PANEL - START - Default Shipping
			'// =========================================
			%>
				<div class="TabbedPanelsContent">
				
					<table class="pcCPcontent">
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <th colspan="2">Default Shipping Address</th>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td colspan="2"><p>If the customer's shipping address is the same as their billing address, leave the shipping address information blank. Customers can have more than one shipping addresses and can add/edit addresses by logging into their account. What is shown here is the &quot;default&quot; shipping address, if different from the billing address.</p>
                            </td>
                        </tr>
                        <tr> 
                            <td colspan="2"><p><strong>The following fields marked with a star (<img src="<%=pcv_strRequiredIcon%>">) are only required if there is a &quot;default&quot; shipping address. The &quot;default&quot; shipping address is optional.</strong></p>
                            </td>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <td><p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_9")%></p></td>
                            <td width="75%">
                                <p>
                                <input type="text" name="ShipCompany" id="ShipCompany" size="20" value="<% =pcf_FillFormField ("ShipCompany", pcv_iscompanyRequired) %>">
                                </p>
                            </td>
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
                        pcv_isStateCodeRequired = True '// determines if validation is performed (true or false)
                        pcv_isProvinceCodeRequired = pcv_isShipProvinceCodeRequired '// determines if validation is performed (true or false)
                        pcv_isCountryCodeRequired = True '// determines if validation is performed (true or false)					
                        
                        '// #3 Additional Required Info
                        pcv_strTargetForm = "modCust" '// Name of Form
                        pcv_strCountryBox = "ShipCountryCode" '// Name of Country Dropdown
                        pcv_strTargetBox = "ShipStateCode" '// Name of State Dropdown
                        pcv_strProvinceBox =  "ShipState" '// Name of Province Field
                        
                        '// Set local Country to Session
                        if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
                            Session(pcv_strSessionPrefix&pcv_strCountryBox) = Session(pcv_strSessionPrefix&pcv_strCountryBox)
                        end if
                        
                        '// Set local State to Session
                        if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
                            Session(pcv_strSessionPrefix&pcv_strTargetBox) = Session(pcv_strSessionPrefix&pcv_strTargetBox)
                        end if
                        
                        '// Set local Province to Session
                        if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
                            Session(pcv_strSessionPrefix&pcv_strProvinceBox) = Session(pcv_strSessionPrefix&pcv_strProvinceBox)
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
                            <td>
                            <p>
                            <%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_3")%></p></td>
                            <td>
                                <p>
                                <input type="text" name="ShipAddress" id="ShipAddress" size="20" value="<% =pcf_FillFormField ("ShipAddress", True) %>">
                                <% pcs_RequiredImageTag "ShipAddress", True %>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td>&nbsp;</td>
                            <td>
                                <p>
                                <input type="text" name="ShipAddress2" id="ShipAddress2" size="20" value="<% =pcf_FillFormField ("ShipAddress2", false) %>">
                                </p>
                            </td>
                        </tr>
                        <tr> 
                            <td> 
                                <p>
                                <%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_4")%>
                                </p>
                            </td>
                            <td>
                                <p>
                                <input type="text" name="ShipCity" id="ShipCity" size="20" value="<% =pcf_FillFormField ("ShipCity", True) %>">
                                <% pcs_RequiredImageTag "ShipCity", True %>
                                </p>
                            </td>
                        </tr>
                        
                        <%
                        '// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
                        pcs_StateProvince
                        %>
                
                        <tr> 
                            <td> 
                            <p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_7")%></p></td>
                            <td>
                                <p><input type="text" name="ShipZip" id="ShipZip" size="20" value="<% =pcf_FillFormField ("ShipZip", True) %>">
                                <% pcs_RequiredImageTag "ShipZip", True %></p>
                            </td>
                        </tr>
                        
                        <tr>	
                        <td>
                            <p>
                            <%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_15")%></p></td>
                            <td>
                                <p>
                                <input type="text" name="ShipEmail" id="ShipEmail" size="20" value="<% =pcf_FillFormField ("ShipEmail", false) %>">
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td><p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_10")%></p></td>
                            <td>
                                <p>
                                <input type="text" name="ShipPhone" id="ShipPhone" size="20" value="<% =pcf_FillFormField ("ShipPhone", false) %>">
                                </p>
                            </td>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
					</table>
					
				</div>
			<%
			'// =========================================
			'// THIRD PANEL - END
			'// =========================================
			
			'// =========================================
			'// FOURTH PANEL - START - Other Information
			'// =========================================
			%>
				<div class="TabbedPanelsContent">
				
					<table class="pcCPcontent">
						<tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <th colspan="2">Other Information</th>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <%
                        'Start Special Customer Fields
                        if session("cp_nc_custfields_exists")="YES" then
                        pcArr=session("cp_nc_custfields")
                        For k=0 to ubound(pcArr,2)
                        %>
                        <tr> 
                            <td><p><%=pcArr(1,k)%>:</p></td>
                            <td>
                            <p>
                                <%if pcArr(2,k)="1" then%>
                                    <input type="checkbox" name="custfield_<%=pcArr(0,k)%>" <%if pcArr(10,k)<>"" then%>value="<%=pcArr(10,k)%>"<%else%><%if pcArr(3,k)<>"" then%>value="<%=pcArr(3,k)%>"<%else%>value="1"<%end if%><%end if%> <%if pcArr(10,k)<>"" then%>checked<%end if%> class="clearBorder">
                                <%else%>
                                    <input type="text" name="custfield_<%=pcArr(0,k)%>" value="<%=replace(pcArr(10,k),"""","&quot;")%>" size="<%=pcArr(4,k)%>" <%if pcArr(5,k)>"0" then%>maxlength="<%=pcArr(5,k)%>"<%end if%>>
                                <%end if%>
                                <%if pcArr(6,k)="1" then%>
                                    <img src="images/pc_required.gif" width="9" height="9">
                                <%end if%>
                            </p>
                            </td>
                        </tr>
                        <%
                        Next
                        end if
                        'End of Special Customer Fields
                        %>
                        <%if (session("IDRefer")<>"0") and (session("IDRefer")<>"") then
                            %>
                            <tr> 
                                <td><p>Referrer Info:</p></td>
                                <td>
                                <p>
                                <%query="select [name] from Referrer where IDRefer=" & session("IDRefer")
                                set rs=server.CreateObject("ADODB.RecordSet")
                                set rs=connTemp.execute(query)
                                if not rs.eof then%>
                                    <%=rs("name")%> 
                                <%end if
                                set rs=nothing
                                %>
                                </p>
                                </td>
                            </tr>
                        <%end if%>
					<% 'MAILUP-S: MailUp Lists, show it for new customer and when existing customers edit their account
					IF session("CP_MU_Setup")="1" THEN
					call opendb()
					query="SELECT pcMailUpLists_ID,pcMailUpLists_ListID,pcMailUpLists_ListGuid,pcMailUpLists_ListName,pcMailUpLists_ListDesc,0 FROM pcMailUpLists WHERE pcMailUpLists_Active>0 and pcMailUpLists_Removed=0;"
					set rs=connTemp.execute(query)
					if not rs.eof then
						tmpArr=rs.getRows()
						set rs=nothing
						intCount=ubound(tmpArr,2)
						tmpNListChecked=0
						pcv_MUSynError=0
						if pidcustomer<>0 then
						'Synchronizing
						For j=0 to intCount
							tmpResult=CheckUserStatus(pidcustomer,pcv_CustomerEmail,tmpArr(1,j),tmpArr(2,j),session("CP_MU_URL"),session("CP_MU_Auto"))
							if tmpResult="-1" then
								pcv_MUSynError=1
								exit for
							else
								if tmpResult="2" then
									query="SELECT pcMailUpSubs_ID FROM pcMailUpSubs WHERE idCustomer=" & pidcustomer & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
									set rs=connTemp.execute(query)
									dtTodaysDate=Date()
									if SQL_Format="1" then
										dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
									else
										dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
									end if
									if not rs.eof then
										if scDB="SQL" then
											query="UPDATE pcMailUpSubs SET idCustomer=" & pidcustomer & ",pcMailUpLists_ID=" & tmpArr(0,j) & ",pcMailUpSubs_LastSave='" & dtTodaysDate & "',pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & pidcustomer & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
										else
											query="UPDATE pcMailUpSubs SET idCustomer=" & pidcustomer & ",pcMailUpLists_ID=" & tmpArr(0,j) & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & pidcustomer & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
										end if
									else
										if scDB="SQL" then
											query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & pidcustomer & "," & tmpArr(0,j) & ",'" & dtTodaysDate & "',0,0);"
										else
											query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & pidcustomer & "," & tmpArr(0,j) & ",#" & dtTodaysDate & "#,0,0);"
										end if
									end if
									set rs=nothing
									set rs=connTemp.execute(query)
									set rs=nothing
								end if
								if tmpResult="1" or tmpResult="3" then
									query="DELETE FROM pcMailUpSubs WHERE idCustomer=" & pidcustomer & " AND pcMailUpLists_ID=" & tmpArr(0,j) & ";"
									set rs=connTemp.execute(query)
									set rs=nothing
								end if
							end if
						Next
						For j=0 to intCount
							query="SELECT idcustomer FROM pcMailUpSubs WHERE idcustomer=" & pidcustomer & " AND pcMailUpLists_ID=" & tmpArr(0,j) & " AND pcMailUpSubs_Optout=0;"
							set rs=connTemp.execute(query)
							tmpOptedIn=0
							if not rs.eof then
								tmpOptedIn=1
								tmpNListChecked=1
							end if
							set rs=nothing
							tmpArr(5,j)=tmpOptedIn
						Next
						end if%>
						<%if pcv_MUSynError=1 then%>
						<tr> 
							<td colspan="2">
								<div class="pcCPmessage">
									<%response.write dictLanguage.Item(Session("language")&"_MailUp_SynNote1")%>
								</div>
							</td>
						</tr>
						<%end if%>
						<tr> 
							<td colspan="2" class="pcSpacer"><script>tmpNListChecked=<%=tmpNListChecked%>;</script><input type="hidden" name="newslistcount" value="<%=intCount%>"></td>
						</tr>
						<tr> 
							<td colspan="2"><p><%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel")%></p></td>
						</tr>
						<%For j=0 to intCount%>
						<tr> 
							<td align="right" valign="top"><input type="checkbox" onclick="javascript: tmpNListChecked=1;" value="<%=tmpArr(0,j)%>" name="newslist<%=j%>" <%if tmpArr(5,j)="1" then%>checked<%end if%> class="clearBorder" /></td>
							<td valign="top"><b><%=tmpArr(3,j)%></b><%if tmpArr(4,j)<>"" then%><br><%=tmpArr(4,j)%><%end if%></td>
						</tr>
						<%Next
					end if
					set rs=nothing	
					'End If MailUp Lists
					ELSE%>
                        <%if AllowNews="1" then%>
                            <tr> 
                                <td align="right"><p><%=NewsLabel%></p></td>
                                <td><p><input type="checkbox" name="CRecvNews" value="1" size="20" <%if Session("pcAdminCRecvNews")="1" then%>checked<%end if%> class="clearBorder"></p></td>
                            </tr>
                        <%end if%>
					<%END IF
					'MAILUP-E%>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
                        <tr> 
                            <th colspan="2">Security Settings</th>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td colspan="2">
                                <input type="checkbox" name="lock" value="1" <% if Session("pcAdmincustomertype")="3" OR Session("pcAdminCustLocked")="1" then%>checked<% end if %> class="clearBorder"> <strong>Lock</strong> this customer&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=456')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
                            </td>
                        </tr>
                        <tr> 
                            <td colspan="2">
                                <input type="checkbox" name="suspend" value="1" <% if Session("pcAdminsuspend")="1" then%>checked<% end if %> class="clearBorder"> <strong>Suspend</strong> this customer&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=457')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
                            </td>
                        </tr>
						<% 
						'// REWARD POINTS - Start
						If RewardsActive <> 0 then 
						%>
						<tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <th colspan="2"><%=RewardsLabel%></th>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td colspan="2">Current Balance: <%=Session("pcAdminiBalance")%></td>
                        </tr>
                        <tr> 
                            <td colspan="2">Redeemed to Date: <%=Session("pcAdminiRewardPointsUsed")%></td>
                        </tr>
                        <tr> 
                            <td colspan="2">
                                Add/Deduct <%=RewardsLabel%>: <input name="iAdjustment" type="text" value="<%=Session("pcAdminiAdjustment")%>" size="10">
                                &nbsp;This can be a negative number.
                                <input name="iRewardPointsAccrued" type="hidden" value="<%=Session("pcAdminiRewardPointsAccrued")%>"> 
                            </td>
                        </tr>
                        <%
						end if 
						'// REWARD POINTS - End

						'// TAX ZONES EXCEPTIONS - Start
						query="SELECT customers.idcustomer, pcTaxZoneRates.pcTaxZoneRate_ID, pcTaxZoneRates.pcTaxZoneRate_Name, pcTaxZoneRates.pcTaxZoneRate_Rate, pcTaxZoneDescriptions.pcTaxZoneDesc FROM pcTaxZoneDescriptions INNER JOIN (((pcTaxEptCust INNER JOIN customers ON pcTaxEptCust.idCustomer = customers.idcustomer) INNER JOIN pcTaxZoneRates ON pcTaxEptCust.pcTaxZoneRate_ID = pcTaxZoneRates.pcTaxZoneRate_ID) INNER JOIN pcTaxZonesGroups ON pcTaxZoneRates.pcTaxZoneRate_ID = pcTaxZonesGroups.pcTaxZoneRate_ID) ON pcTaxZoneDescriptions.pcTaxZoneDesc_ID = pcTaxZonesGroups.pcTaxZoneDesc_ID WHERE customers.idcustomer="&pidcustomer&";"
                        set rs=server.CreateObject("ADODB.RecordSet")
                        set rs=conntemp.execute(query)
                        if NOT rs.eof then %>
                            <tr>
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr> 
                                <th colspan="2">Tax Zone Rule Exemptions</th>
                            </tr>
                            <tr>
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            
                            <% do until rs.eof
                                intTaxZoneRateID=rs("pcTaxZoneRate_ID")
                                strTaxZoneRateName=rs("pcTaxZoneRate_Name")
                                dblTaxZoneRate=rs("pcTaxZoneRate_Rate")
                                if dblTaxZoneRate<>0 then
                                    dblTaxZoneRate=(dblTaxZoneRate*100)
                                end if
                                strTaxZoneDesc=rs("pcTaxZoneDesc") %>
                                <tr> 
                                    <td colspan="2">
                                    <%=strTaxZoneRateName&" ("&strTaxZoneDesc&")" %>&nbsp;&nbsp;&nbsp;<%=dblTaxZoneRate%>%&nbsp;&nbsp;&nbsp;<a href="manageTaxEptCust.asp?ZoneRateID=<%=intTaxZoneRateID%>&mode=view&referback=modCusta.asp?idcustomer=<%=pidcustomer%>">Edit</a>
                                    </td>
                                </tr>
                                <% rs.moveNext
                            loop %>
                            <tr>
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                        <% else
                            '// See if there are zones set - if so alert admin that this customer is not currently exempt
                            query="SELECT pcTaxZoneRate_ID FROM pcTaxZoneRates;"
                            set rsTEMP=server.CreateObject("ADODB.RecordSet")
                            set rsTEMP=conntemp.execute(query)
                            if NOT rsTEMP.eof then %>
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>
                                <tr> 
                                    <th colspan="2">Tax Zone Rule Exemptions</th>
                                </tr>
                                <tr>
                                    <td colspan="2" class="pcCPspacer"></td>
                                </tr>
                                <tr> 
                                    <td colspan="2"><p>This customer is not currently associated with a tax zone rule exemption. To exempt this customer, <a href="viewTax.asp">edit an existing tax zone</a> (or create a new one) and click on &quot;View/Edit Customer Exemptions for this Tax&quot; on the &quot;Add/Edit Tax Rate by Zone&quot; window.</p></td>
                                </tr>
                            <% 
							end if
                        end if
                        set rs=nothing
						'// TAX ZONES EXCEPTIONS - End


						' PRV41 start 
						' Check to see if Product Reviews are active
						query = "SELECT pcRS_Active FROM pcRevSettings;"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						pcv_Active=rs("pcRS_Active")
						if isNull(pcv_Active) or pcv_Active="" then
							pcv_Active="0"
						end if
						Set rs=Nothing
						if pcv_Active<>"0" then
						%>
						<tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <th colspan="2">Miscellaneous</th>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td colspan="2">
                                <input type="checkbox" name="allowreviewemails" value="1" <% if Session("pcAllowReviewEmails")&""="1" then%>checked<% end if %> class="clearBorder"> 
								Customer wants Product Review reminders 
							<%
							response.write "&nbsp;&nbsp;|&nbsp;<a href=""prv_ManageReviews.asp?idProduct=0&nav=2&idcustomer=" & pidcustomer & """>See reviews written by this customer</a>"
							' PRV41 end
							%>
                            </td>
                        </tr>
						<%
						end if
						' PRV41 end
						%>

                        
					</table>
				
				</div>
				
			<%
			'// FOURTH PANEL - END
			%>
			
			</div>
		
	  </div>
		<%
		'// TABBED PANELS - MAIN DIV END
		%>

	<div style="clear: both;">&nbsp;</div>
  <script type="text/javascript">
		<!--
		var TabbedPanels1 = new Spry.Widget.TabbedPanels("TabbedPanels1", {defaultTab: params.tab ? params.tab : 0});
		//-->
  </script>
  
</form>
<%Response.write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_MailUp_SynNote2"),"MailUp", 300))%>
<% 
call closedb()
'// Clear the sessions
pcs_ClearAllSessions
'// Clear the Sessions not Auto-Cleared
Session("pcAdminCRecvNews")=""
Session("pcAdmincustomertype")=""
Session("pcAdminiRewardPointsAccrued")="" 
Session("pcAdminiRewardPointsUsed")=""
Session("pcAdminSuspend")=""
Session("pcAdminiBalance")=""
Session("IDRefer")=""
Session("pcAdminCustLocked")=""
Session("pcAdminCustNotes")=""
%><!--#include file="AdminFooter.asp"-->