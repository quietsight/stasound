<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Customer" %>
<% section="mngAcc"%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<!--#include file="../includes/MailUpFunctions.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: PAGE CONFIG
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
dim conntemp, query, rstemp, rstemp2, rs
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

'// Set Page Name
pcStrPageName = "instCusta.asp"

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

pcv_ispasswordRequired= true
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: PAGE CONFIG
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ONLOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
'// Start Special Customer Fields
session("cp_nc_custfields")=""
session("cp_nc_custfields_exists")=""
query="SELECT pcCField_ID,pcCField_Name,pcCField_FieldType,pcCField_Value,pcCField_Length,pcCField_Maximum,pcCField_Required,pcCField_PricingCategories,pcCField_ShowOnReg,pcCField_ShowOnCheckout,'' FROM pcCustomerFields;"
set rs=connTemp.execute(query)
if not rs.eof then
	session("cp_nc_custfields")=rs.GetRows()
	session("cp_nc_custfields_exists")="YES"
end if
set rs=nothing
'/  End of Special Customer Fields	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: ONLOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IF request.Form("Modify")<>"" THEN

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
		session("cp_nc_custfields")=pcArr
	end if
	'End of Special Customer Fields
	
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
	
	'// Customer Type
	pcs_ValidateTextField "customerType", true, 0
	
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
		
		'// Customer Type	
		If instr(Session("pcAdmincustomertype"),"CC") then
			intidcustomerCategory=replace(Session("pcAdmincustomertype"),"CC_","")
			intidcustomerCategory=int(intidcustomerCategory)
			query="SELECT pcCustomerCategories.idcustomerCategory, pcCustomerCategories.pcCC_WholesalePriv FROM pcCustomerCategories WHERE (((pcCustomerCategories.idcustomerCategory)="&intidcustomerCategory&"));"	
			SET rs=Server.CreateObject("ADODB.RecordSet")
	
			SET rs=conntemp.execute(query)
			intpcCC_WholesalePriv=rs("pcCC_WholesalePriv")
			if intpcCC_WholesalePriv=1 then
				Session("pcAdmincustomertype")=1
			else
				Session("pcAdmincustomertype")=0
			end if
		end if	
		

		' Email Already in Database and NOT a guest customer
		query="SELECT idcustomer, email FROM customers WHERE pcCust_Guest=0 AND email='"&trim(Session("pcAdminpcBillingEmail"))&"';"	
		set rstemp=conntemp.execute(query)
		if err.number <> 0 then
			response.redirect "techErr.asp?error="& Server.Urlencode("Error occurred while checking for duplicate emails: "&Err.Description) 
		end if		
		if NOT rstemp.eof then
				response.redirect pcStrPageName&"?msg=The email you have chosen is already in use by another customer.&idcustomer=" & pidcustomer
		end if
		
		'// Password
		Session("pcAdminpassword")	= enDeCrypt(Session("pcAdminpassword"), scCrypPass)	
		
		'// Rewared Points		
		If Session("pcAdminiRewardPointsAccrued")="" then
			Session("pcAdminiRewardPointsAccrued")="0"
		End If	

		'// Id Refer
		pIdRefer=replace(Request("idRefer"),"'","''")
		if pIdRefer="" then
			pIdRefer="0"
		end If
		
		' PRV41 begin
		If request.Form("allowreviewemails")<>"1" Then
		   pcAllowReviewEmails = 0
		Else
		   pcAllowReviewEmails = 1
		End if
		' PRV41 end
		
		if intidcustomerCategory = "" then
			intidcustomerCategory=0
		end if
		
		dtTodaysDate=Date()
		if SQL_Format="1" then
			dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
		else
			dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
		end if

		'// Insert Customer
		if scDB="SQL" then
			query="INSERT INTO customers (pcCust_VATID, pcCust_SSN, name, lastName, email, fax, [password],city,zip,CountryCode, state, stateCode,shippingcity,shippingZip,shippingCountryCode, shippingState, shippingStateCode, phone, address, shippingAddress, customercompany, customerType,IDRefer,address2,shippingCompany, shippingAddress2,RecvNews,idcustomerCategory,pcCust_DateCreated, shippingEmail, shippingPhone,pcCust_AllowReviewEmails) VALUES ('" &Session("pcAdminpcBillingVATID")& "', '" &Session("pcAdminpcBillingSSN")& "','" &Session("pcAdminpcBillingFirstName")& "', '" &Session("pcAdminpcBillingLastName")& "', '" &Session("pcAdminpcBillingEmail")& "', '" &Session("pcAdminpcBillingFax")& "' , '" &Session("pcAdminpassword")&"','" &Session("pcAdminpcBillingCity")& "','" &Session("pcAdminpcBillingPostalCode")& "','" &Session("pcAdminpcBillingCountryCode")& "', '" &Session("pcAdminpcBillingProvince")& "', '" &Session("pcAdminpcBillingStateCode")& "','" &Session("pcAdminShipCity")& "','" &Session("pcAdminShipZip")& "','" &Session("pcAdminShipCountryCode")& "', '" &Session("pcAdminShipState")& "', '" &Session("pcAdminShipStateCode")& "', '" &Session("pcAdminpcBillingPhone")& "', '" &Session("pcAdminpcBillingAddress")& "', '" &Session("pcAdminShipAddress")& "', '" &Session("pcAdminpcBillingCompany")& "', " &Session("pcAdminCustomerType")& ","&pIdRefer&",'" &Session("pcAdminpcBillingAddress2")& "','" &Session("pcAdminShipCompany")& "','" &Session("pcAdminShipAddress2")& "',"&Session("pcAdminCRecvNews")&"," & intidcustomerCategory & ",'" & dtTodaysDate & "','" &Session("pcAdminShipEmail")& "','"&Session("pcAdminShipPhone")&"'," & pcAllowReviewEmails & ")"
		else
			query="INSERT INTO customers (pcCust_VATID, pcCust_SSN, name, lastName, email, fax, [password],city,zip,CountryCode, state, stateCode,shippingcity,shippingZip,shippingCountryCode, shippingState, shippingStateCode, phone, address, shippingAddress, customercompany, customerType,IDRefer,address2,shippingCompany, shippingAddress2,RecvNews,idcustomerCategory,pcCust_DateCreated, shippingEmail, ShippingPhone,pcCust_AllowReviewEmails) VALUES ('" &Session("pcAdminpcBillingVATID")& "', '" &Session("pcAdminpcBillingSSN")& "','" &Session("pcAdminpcBillingFirstName")& "', '" &Session("pcAdminpcBillingLastName")& "', '" &Session("pcAdminpcBillingEmail")& "', '" &Session("pcAdminpcBillingFax")& "' , '" &Session("pcAdminpassword")&"','" &Session("pcAdminpcBillingCity")& "','" &Session("pcAdminpcBillingPostalCode")& "','" &Session("pcAdminpcBillingCountryCode")& "', '" &Session("pcAdminpcBillingProvince")& "', '" &Session("pcAdminpcBillingStateCode")& "','" &Session("pcAdminShipCity")& "','" &Session("pcAdminShipZip")& "','" &Session("pcAdminShipCountryCode")& "', '" &Session("pcAdminShipState")& "', '" &Session("pcAdminShipStateCode")& "', '" &Session("pcAdminpcBillingPhone")& "', '" &Session("pcAdminpcBillingAddress")& "', '" &Session("pcAdminShipAddress")& "', '" &Session("pcAdminpcBillingCompany")& "', " &Session("pcAdminCustomerType")& ","&pIdRefer&",'" &Session("pcAdminpcBillingAddress2")& "','" &Session("pcAdminShipCompany")& "','" &Session("pcAdminShipAddress2")& "',"&Session("pcAdminCRecvNews")&"," & intidcustomerCategory & ",#" & dtTodaysDate & "#,'" &Session("pcAdminShipEmail")& "','"&Session("pcAdminShipPhone")&"'," & pcAllowReviewEmails & ")"
		end If
		' PRV41 end
		
		set rstemp=conntemp.execute(query)	
		if err.number <> 0 then
			set rstemp=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error occurred while adding new customer into database: "&Err.Description) 
		end if
		set rstemp=nothing
		
		query="SELECT idcustomer FROM Customers WHERE [email]='"&trim(Session("pcAdminpcBillingEmail"))&"';"
		set rs=connTemp.execute(query)
		pcv_IDCustomer=rs("idcustomer")
		set rs=nothing 
		
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
			end if
		'MAILUP-E
		
		'Start Special Customer Fields
		if session("cp_nc_custfields_exists")="YES" then
			set rs=nothing 
			pcArr=session("cp_nc_custfields")
			For k=0 to ubound(pcArr,2)
				query="INSERT INTO pcCustomerFieldsValues (idcustomer,pcCField_ID,pcCFV_Value) VALUES (" & pcv_IDCustomer & "," & pcArr(0,k) & ",'" & pcArr(3,k) & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			Next
			session("cp_nc_custfields")=""
		end if
		'End of Special Customer Fields
		
		set rs=nothing	
		call closedb()	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Clear the sessions
		pcs_ClearAllSessions
		
		'// Redirect
		If session("CP_MU_Auto")="0" then
			MUResult=1
		end if
		response.redirect "viewCusta.asp?action=added&idCustomer="&pcv_IDCustomer&"&mailup=" & MUResult
		
	End If	
END IF	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

msg=request.querystring("msg") 


if Session("adminCountryCode")="" then
	Session("adminCountryCode")=scShipFromPostalCountry
end if

if Session("adminshippingCountryCode")="" then
	Session("adminshippingCountryCode")=scShipFromPostalCountry
end if

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

<% if msg<>"" then %>
	<br>
	<div class="pcCPmessage"> 
		<% if request.querystring("s")=1 then %>
			<img src="images/pcadmin_successful.gif" width="18" height="18"> 
		<% else %>
			<img src="images/pcadmin_note.gif" width="20" height="20"> 
		<% end if %>
		<%=msg%>
	</div>
	<br>
<% end if %>
<script>var tmpNListChecked=0;</script>
<form method="post" name="modCust" action="<%=pcStrPageName%>" onSubmit="<%if session("CP_MU_Auto")="1" then%>if (tmpNListChecked) pcf_Open_MailUp();<%end if%> return Form1_Validator(this)" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">General Information</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td nowrap><p>Customer Type:</p></td>
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
		
		<tr> 
			<td><p>Password:</p></td>
			<td><p>
				<input type="password" name="password" value="<% =pcf_FillFormField ("password", pcv_ispasswordRequired) %>" size="25" maxlength="50">
				<%pcs_RequiredImageTag "password", pcv_ispasswordRequired %>
			</p>
			</td>
		</tr>
		
		<%
		'Start Special Customer Fields
			if session("cp_nc_custfields_exists")="YES" then
				pcArr=session("cp_nc_custfields")
				For k=0 to ubound(pcArr,2)
		%>
				<tr> 
					<td><p><%=pcArr(1,k)%></p></td>
					<td><p>
						<%if pcArr(2,k)="1" then%>
							<input type="checkbox" name="custfield_<%=pcArr(0,k)%>" <%if pcArr(3,k)<>"" then%>value="<%=pcArr(3,k)%>"<%else%>value="1"<%end if%> <%if pcArr(6,k)="1" then%>checked<%end if%> class="clearBorder">
						<%else%>
							<input type="text" name="custfield_<%=pcArr(0,k)%>" value="<%=pcArr(3,k)%>" size="<%=pcArr(4,k)%>" <%if pcArr(5,k)>"0" then%>maxlength="<%=pcArr(5,k)%>"<%end if%>>
						<%end if%>
						<%if pcArr(6,k)="1" then%>
							<img src="images/pc_required.gif" width="9" height="9">
						<%end if%></p>
					</td>
				</tr>
		<%
				Next
			end if
		'End of Special Customer Fields
		%>
				
		<%if (RefNewReg="1" OR RefNewCheckout="1") AND ReferLabel<>"" then 
			query="SELECT idRefer,[Name],sortOrder FROM Referrer ORDER BY sortOrder;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if NOT rs.eof then %>
		<tr> 
			<td><p><%=ReferLabel%></p></td>
			<td>			
				<p><select name="idRefer">
				<% do until rs.eof
					idRefer=rs("idRefer")
					pName=rs("Name") 
					if session("adminidRefer")=pName then %>
						<option value="<%=idRefer%>" selected><%=pName%></option>
					<% else %>
						<option value="<%=idRefer%>"><%=pName%></option>
					<% end if  %>
					<% rs.movenext
				loop
				set rs=nothing %>
				</select></p>
			</td>
		</tr>
		<% end if 
		end if%>
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
						pcv_MUSynError=0%>
						<tr> 
							<td colspan="2" class="pcSpacer"><script>tmpNListChecked=<%=tmpNListChecked%>;</script><input type="hidden" name="newslistcount" value="<%=intCount%>"></td>
						</tr>
						<tr> 
							<td><p><%response.write dictLanguage.Item(Session("language")&"_MailUp_RegisterLabel")%></p></td>
							<td>&nbsp;</td>
						</tr>
						<%For j=0 to intCount%>
						<tr> 
							<td valign="top" align="right"><input type="checkbox" onclick="javascript: tmpNListChecked=1;" value="<%=tmpArr(0,j)%>" name="newslist<%=j%>" <%if tmpArr(5,j)="1" OR Session("pcAdminpcNewsList" & j)&""=tmpArr(0,j)&"" then%>checked<%end if%> class="clearBorder" /></td>
							<td valign="top"><b><%=tmpArr(3,j)%></b><%if tmpArr(4,j)<>"" then%><br><%=tmpArr(4,j)%><%end if%></td>
						</tr>
						<%Next
					end if
					set rs=nothing	
					'End If MailUp Lists
					ELSE%> 
						<%if AllowNews="1" then%>
							<tr> 
								<td align="right"><input type="checkbox" name="CRecvNews" value="1" size="20" class="clearBorder"></td>
								<td><p><%=NewsLabel%></p></td>
							</tr>
						<%end if
					END IF
					'MAILUP-E %> 
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
		<% 
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
                   <input type="checkbox" name="allowreviewemails" value="1" class="clearBorder"> 
                   Customer wants Product Review reminders
               </td>
            </tr>
		<% 
		end if
		' PRV41 end 
		%>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"> 
				<input type="submit" name="Modify" value="Add Customer" class="submit2">&nbsp;
				<input type="button" name="Search" value="Locate a Customer" onClick="location.href='viewCusta.asp'">&nbsp;
				<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
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
%>
<!--#include file="AdminFooter.asp"-->