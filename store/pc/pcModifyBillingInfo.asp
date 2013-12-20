<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/languages_Ship.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp" --> 
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/validation.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="pcStartSession.asp" -->
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<% 
'// Check if store is turned off and return message to customer
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If 

dim conntemp, query, rs

'// extract real idorder (without prefix)
pTrueOrderId=(int(session("GWOrderId"))-scpre)

If Not validNum(pTrueOrderId) then
	response.redirect "msg.asp?message=10"
End If

'verify that this order doesn't alreay exists and that the idCustomer is only that of the customer logged in.
if session("idCustomer")="" OR session("idCustomer")=0 then
	response.redirect "viewCart.asp"
else
	'// Open the database
	call opendb()
	query="SELECT idCustomer From orders WHERE idOrder="&pTrueOrderId
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if NOT rs.eof then
		pcv_tempID=rs("idCustomer")
	end if
	set rs=nothing
	call closedb()
	
	if isNumeric(pcv_tempID) AND pcv_tempID<>session("idCustomer") then
		response.redirect "msg.asp?message=211"     
	end if
end if


'// Set Page Name
pcStrPageName = "pcModifyBillingInfo.asp"

'// Set Required Fields
pcv_isBillingFirstNameRequired = true
pcv_isBillingLastNameRequired = true
pcv_isBillingCompanyRequired = false
pcv_isBillingPhoneRequired = true
pcv_isBillingAddressRequired = true
pcv_isBillingPostalCodeRequired = true
pcv_isBillingCityRequired = true
pcv_isBillingCountryCodeRequired = true
pcv_isBillingAddress2Required = false
pcv_isBillingFaxRequired = false

'// Use the Request object to toggle State (based of Country selection)
pcv_isBillingStateCodeRequired = true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isBillingStateCodeRequired=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
pcv_isBillingProvinceRequired = false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isBillingProvinceRequired=pcv_strProvinceCodeRequired
end if



'// If the Form was submitted to modify
If request("Modified")="YES" then

	'//set error to zero
	pcv_intErr=0
	
	'//generic error for page
	pcv_strGenericPageError = server.URLEncode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	
	'//validate the forms
	pcs_ValidateTextField "idCustomer", false, 0
	pcs_ValidateTextField "pcBillingFirstName", pcv_isBillingFirstNameRequired, 0
	pcs_ValidateTextField "pcBillingLastName", pcv_isBillingLastNameRequired, 0
	pcs_ValidateTextField "pcBillingCompany", pcv_isBillingCompanyRequired, 0
	pcs_ValidatePhoneNumber "pcBillingPhone", pcv_isBillingPhoneRequired, 0
	pcs_ValidateTextField "pcBillingAddress", pcv_isBillingAddressRequired, 0
	pcs_ValidateTextField "pcBillingPostalCode", pcv_isBillingPostalCodeRequired, 10
	pcs_ValidateTextField "pcBillingStateCode", pcv_isBillingStateCodeRequired, 0
	pcs_ValidateTextField "pcBillingProvince", pcv_isBillingProvinceRequired, 0
	pcs_ValidateTextField "pcBillingCity", pcv_isBillingCityRequired, 0
	pcs_ValidateTextField "pcBillingCountryCode", pcv_isBillingCountryCodeRequired, 0
	pcs_ValidateTextField "pcBillingAddress2", pcv_isBillingAddress2Required, 0
	pcs_ValidatePhoneNumber "pcBillingFax", pcv_isBillingFaxRequired, 0
	
	'// Set to Local Variables
	pcIntIdCustomer=Session("pcSFidCustomer")
	pcStrBillingFirstName=Session("pcSFpcBillingFirstName")
	pcStrBillingLastName=Session("pcSFpcBillingLastName")
	pcStrBillingCompany=Session("pcSFpcBillingCompany")
	pcStrBillingPhone=Session("pcSFpcBillingPhone")
	pcStrBillingAddress=Session("pcSFpcBillingAddress")
	pcStrBillingPostalCode=Session("pcSFpcBillingPostalCode")
	pcStrBillingStateCode=Session("pcSFpcBillingStateCode")
	pcStrBillingProvince=Session("pcSFpcBillingProvince")
	pcStrBillingCity=Session("pcSFpcBillingCity")
	pcStrBillingCountryCode=Session("pcSFpcBillingCountryCode")
	pcStrBillingAddress2=Session("pcSFpcBillingAddress2")
	pcStrBillingFax=Session("pcSFpcBillingFax")
	
	if pcStrBillingProvince<>"" then
		pcStrBillingStateCode=""
	end if
	
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg=" & pcv_strGenericPageError
	Else
		'// Open the database
		call opendb()
		
		'save customer info to "customers" table
		query="UPDATE customers SET name='"&pcStrBillingFirstName&"', lastName='"&pcStrBillingLastName&"', customerCompany='"&pcStrBillingCompany&"', phone='"&pcStrBillingPhone&"', address='"&pcStrBillingAddress&"', zip='"&pcStrBillingPostalCode&"', stateCode='"&pcStrBillingStateCode&"', state='"&pcStrBillingProvince&"', city='"&pcStrBillingCity&"', countryCode='"&pcStrBillingCountryCode&"', address2='"&pcStrBillingAddress2&"', fax='"&pcStrBillingFax&"' WHERE ((idCustomer="&pcIntIdCustomer&"));"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		'save customer info to "orders" table
		query="UPDATE orders SET address='"&pcStrBillingAddress&"', zip='"&pcStrBillingPostalCode&"', stateCode='"&pcStrBillingStateCode&"', state='"&pcStrBillingProvince&"', city='"&pcStrBillingCity&"', countryCode='"&pcStrBillingCountryCode&"', address2='"&pcStrBillingAddress2&"' WHERE ((idOrder="&pTrueOrderId&"));"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		'if store only allows shipments to the billing address, we must update the shipping address with this infomation also.
		if (scAlwAltShipAddress="1") OR (session("pcShipOpt")="-1") then
			if (scAlwAltShipAddress="1") then
			query="UPDATE customers SET shippingAddress='"&pcStrBillingAddress&"', shippingCity='"&pcStrBillingCity&"', shippingState='"&pcStrBillingProvince&"', shippingStateCode='"&pcStrBillingStateCode&"', shippingZip='"&pcStrBillingPostalCode&"', shippingCountryCode='"&pcStrBillingCountryCode&"', shippingCompany='"&pcStrBillingCompany&"', shippingAddress2='"&pcStrBillingAddress2&"' WHERE ((idCustomer="&pcIntIdCustomer&"));"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
			end if
		
			'save customer info to "orders" table
			query="UPDATE orders SET shippingAddress='"&pcStrBillingAddress&"', shippingZip='"&pcStrBillingPostalCode&"', shippingState='"&pcStrBillingProvince&"', shippingStateCode='"&pcStrBillingStateCode&"', shippingCity='"&pcStrBillingCity&"', shippingCountryCode='"&pcStrBillingCountryCode&"', pcOrd_shippingPhone='"&pcStrBillingPhone&"', shippingFullName='"&pcStrBillingFirstName&" "&pcStrBillingLastName&"', shippingCompany='"&pcStrBillingCompany&"', shippingAddress2='"&pcStrBillingAddress2&"' WHERE ((idOrder="&pTrueOrderId&"));"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		end if
	
		set rs=nothing
		
		'//Close database before redirecting 		
		call closedb()
		
		'// Clear all sessions (related to the form)
		pcs_ClearAllSessions()
		
		'redirect customer to redirect page
		response.Redirect( session("redirectPage") )
		response.end
		
	End IF
	
Else

	'Get customer info from order id that's in session 
	call opendb()
	
	query="SELECT customers.idcustomer, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.address, customers.zip, customers.stateCode, customers.state, customers.city, customers.countryCode, customers.address2, customers.fax FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pTrueOrderId&"));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	pcIntIdcustomer= pcf_ResetFormField(Session("pcSFidcustomer"), rs("idcustomer"))
	pcStrBillingFirstName= pcf_ResetFormField(Session("pcSFpcBillingFirstName"), rs("name"))
	pcStrBillingLastName= pcf_ResetFormField(Session("pcSFpcBillingLastName"), rs("lastName"))
	pcStrBillingCompany= pcf_ResetFormField(Session("pcSFpcBillingCompany"), rs("customerCompany"))
	pcStrBillingPhone= pcf_ResetFormField(Session("pcSFpcBillingPhone"), rs("phone"))
	pcStrBillingAddress= pcf_ResetFormField(Session("pcSFpcBillingAddress"), rs("address"))
	pcStrBillingPostalCode= pcf_ResetFormField(Session("pcSFpcBillingPostalCode"), rs("zip"))
	pcStrBillingStateCode= pcf_ResetFormField(Session("pcSFpcBillingStateCode"), rs("stateCode"))
	pcStrBillingProvince= pcf_ResetFormField(Session("pcSFpcBillingProvince"), rs("state"))
	pcStrBillingCity= pcf_ResetFormField(Session("pcSFpcBillingCity"), rs("city"))
	pcStrBillingCountryCode= pcf_ResetFormField(Session("pcSFpcBillingCountryCode"), rs("countryCode"))
	pcStrBillingAddress2= pcf_ResetFormField(Session("pcSFpcBillingAddress2"), rs("address2"))
	pcStrBillingFax= pcf_ResetFormField(Session("pcSFpcBillingFax"), rs("fax"))

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Config Client-Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	response.write "<script language=""JavaScript"">"&vbcrlf
	response.write "<!--"&vbcrlf	
	response.write "function Form1_Validator(theForm)"&vbcrlf
	response.write "{"&vbcrlf
	pcs_JavaTextField	"pcBillingFirstName", pcv_isBillingFirstNameRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcBillingLastName", pcv_isBillingLastNameRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcBillingCompany", pcv_isBillingCompanyRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcBillingPhone", pcv_isBillingPhoneRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcBillingAddress", pcv_isBillingAddressRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcBillingPostalCode", pcv_isBillingPostalCodeRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcBillingCity", pcv_isBillingCityRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcBillingCountryCode", pcv_isBillingCountryCodeRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcBillingAddress2", pcv_isBillingAddress2Required, dictLanguage.Item(Session("language")&"_NewCust_3")
	pcs_JavaTextField	"pcBillingFax", pcv_isBillingFaxRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
	response.write "return (true);"&vbcrlf
	response.write "}"&vbcrlf
	response.write "//-->"&vbcrlf
	response.write "</script>"&vbcrlf
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: FORM VALIDATION
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
	<div id="pcMain">
	<form name="modifybillingform" action="<%=pcStrPageName%>" method="post" onsubmit="return Form1_Validator(this)" class="pcForms">
	<input type="hidden" name="Modified" value="YES">
	<input type="hidden" name="idCustomer" value="<%=pcIntIdcustomer%>">
	<table class="pcMainTable">
		<% msg=getUserInput(request.querystring("msg"),0)
		If msg<>"" then %>
			<tr>
				<td colspan="2"><div class="pcErrorMessage"><%=msg%></div></td>
			</tr>
		<% end if %> 
		<tr>
			<td colspan="2"><h1><%response.write dictLanguage.Item(Session("language")&"_order_J")%></h1></td>
		</tr>
		<tr>
			<td>
            	<table class="pcShowContent">
                	<tr>
                    	<td>
                        <p>
                            <%response.write dictLanguage.Item(Session("language")&"_order_C")%>
                        </p></td>
                        <td><p>
                            <input type="text" name="pcBillingFirstName" value="<%=pcStrBillingFirstName%>" size="20" />
                            <%pcs_RequiredImageTag "pcBillingFirstName", pcv_isBillingFirstNameRequired %>
                        </p></td>
                    </tr>
                    <tr>
                        <td><p>
                            <%response.write dictLanguage.Item(Session("language")&"_order_D")%>
                        </p></td>
                        <td><p>
                            <input type="text" name="pcBillingLastName" value="<%=pcStrBillingLastName%>" size="20" />
                            <%pcs_RequiredImageTag "pcBillingLastName", pcv_isBillingLastNameRequired %>
                        </p></td>
                    </tr>
                    <tr>
                        <td><p>
                            <%response.write dictLanguage.Item(Session("language")&"_order_E")%>
                        </p></td>
                        <td><p>
                            <input type="text" name="pcBillingCompany" value="<%=pcStrBillingCompany%>" size="30" />
                            <%pcs_RequiredImageTag "pcBillingCompany", pcv_isBillingCompanyRequired %>
                        </p></td>
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
                    pcv_strTargetForm = "modifybillingform" '// Name of Form
                    pcv_strCountryBox = "pcBillingCountryCode" '// Name of Country Dropdown
                    pcv_strTargetBox = "pcBillingStateCode" '// Name of State Dropdown
                    pcv_strProvinceBox =  "pcBillingProvince" '// Name of Province Field
                    
                    '// Set local Country to Session
                    if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
                        Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrBillingCountryCode
                    end if
                    
                    '// Set local State to Session
                    if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
                        Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrBillingStateCode
                    end if
                    
                    '// Set local Province to Session
                    if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
                        Session(pcv_strSessionPrefix&pcv_strProvinceBox) = pcStrBillingProvince
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
                            <input type="text" name="pcBillingAddress" value="<%=pcStrBillingAddress%>" size="30" />
                            <%pcs_RequiredImageTag "pcBillingAddress", pcv_isBillingAddressRequired %>
                        </p></td>
                    </tr>
                    <tr>
                        <td><p>&nbsp;</p></td>
                        <td><p>
                            <input type="text" name="pcBillingAddress2" value="<%=pcStrBillingAddress2%>" size="30" />
                            <%pcs_RequiredImageTag "pcBillingAddress2", pcv_isBillingAddress2Required %>
                        </p></td>
                    </tr>
                    <tr>
                        <td><p>
                            <%response.write dictLanguage.Item(Session("language")&"_order_L")%>
                        </p></td>
                        <td><p>
                            <input type="text" name="pcBillingCity" value="<%=pcStrBillingCity%>" size="30" />
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
                            <input name="pcBillingPostalCode" type="text" value="<%=pcStrBillingPostalCode%>" size="10" maxlength="10" />
                            <%pcs_RequiredImageTag "pcBillingPostalCode", pcv_isBillingPostalCodeRequired %>
                        </p></td>
                    </tr>
            
                    <%	'// Phone Custom Error
                    if session("ErrpcBillingPhone")<>"" then %>
                    <tr>
                        <td>&nbsp;</td>
                        <td><img src="<%=pcf_GenerateIconURL(rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%></td>
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
                            <input type="text" name="pcBillingPhone" value="<%=pcStrBillingPhone%>" size="15" />
                            <%pcs_RequiredImageTag "pcBillingPhone", pcv_isBillingPhoneRequired %>
                        </p></td>
                    </tr>
                    <tr>
                        <td><p>
                            <%response.write dictLanguage.Item(Session("language")&"_order_AA")%>
                        </p></td>
                        <td><p>
                            <input type="text" name="pcBillingFax" value="<%=pcStrBillingFax%>" size="15" />
                            <%pcs_RequiredImageTag "pcBillingFax", pcv_isBillingFaxRequired %>
                        </p></td>
                    </tr>
                    <tr>
                        <td colspan="2"><hr></td>
                    </tr>
                    <tr>
                        <td><input type="image" id="Submit" src="<%=rslayout("submit")%>" name="Submit" value="Save and Continue" class="submit"></td>
                        <td>&nbsp;</td>
                    </tr>
                </table>
			</td>
		</tr>
	</table>
	</form>
</div>
<% End If %>
<!--#include file="footer.asp"-->