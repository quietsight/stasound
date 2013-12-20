<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="header.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<% '// Check if store is turned off and return message to customer
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If 
dim query, conntemp, rs

'// Get recipient ID
reID=getUserInput(request("reID"),0)
if not reID<>"" then
	response.redirect "CustSAmanage.asp"
end if

'// Page Name
pcStrPageName="CustModShip.asp"

pcv_isShipFirstNameRequired=True
pcv_isShipLastNameRequired=True
pcv_isShipNickNameRequired=False
pcv_isShipCompanyRequired=False
pcv_isShipAddressRequired=True
pcv_isShipCityRequired=True
'// Use the Request object to toggle State (based of Country selection)
pcv_isShipStateCodeRequired=True
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isShipStateCodeRequired=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
pcv_isShipProvinceCodeRequired=False
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isShipProvinceCodeRequired=pcv_strProvinceCodeRequired
end if
pcv_isShipZipRequired=True
pcv_isShipCountryCodeRequired=True
pcv_isShipPhoneRequired=True
pcv_isShipFaxRequired=False
pcv_isShipEmailRequired=False

if request.form("updatemode")="1" then
	'//set error to zero
	pcv_intErr=0
	
	'//generic error for page
	pcv_strGenericPageError = server.URLEncode(dictLanguage.Item(Session("language")&"_Custmoda_18"))
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	IF reID<>"0" then
		pcs_ValidateTextField "shipFirstName", pcv_isShipFirstNameRequired, 0
		pcs_ValidateTextField "shipLastName", pcv_isShipLastNameRequired, 0
		pcs_ValidateTextField "shipNickName", pcv_isShipNickNameRequired, 0
		pcs_ValidatePhoneNumber "ShipFax", pcv_isShipFaxRequired, 14
	End If
	pcs_ValidatePhoneNumber "ShipPhone", pcv_isShipPhoneRequired, 14
	pcs_ValidateEmailField "ShipEmail", pcv_isShipEmailRequired, 0
	pcs_ValidateTextField "ShipCompany", false, 0
	pcs_ValidateTextField "ShipAddress", pcv_isShipAddressRequired, 0
	pcs_ValidateTextField "ShipAddress2", false, 0
	pcs_ValidateTextField "ShipCity", pcv_isShipCityRequired, 0
	pcs_ValidateTextField "ShipState", pcv_isShipProvinceCodeRequired, 0
	pcs_ValidateTextField "ShipStateCode", pcv_isShipStateCodeRequired, 0
	pcs_ValidateTextField "ShipZip", pcv_isShipZipRequired, 0
	pcs_ValidateTextField "ShipCountryCode", pcv_isShipCountryCodeRequired, 0
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Server Side Validation
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START: Set Local Variables for recipient
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	IF reID<>"0" then
		pcStrShipFirstName = Session("pcSFshipFirstName")
		pcStrShipLastName = Session("pcSFshipLastName")
		pcStrShipNickName = Session("pcSFshipNickName")
	end if
	pcStrShipCompany = Session("pcSFShipCompany")
	pcStrShipAddress = Session("pcSFShipAddress")
	pcStrShipAddress2 = Session("pcSFShipAddress2")
	pcStrShipCity = Session("pcSFShipCity")
	pcStrShipState = Session("pcSFShipState")
	pcStrShipStateCode = Session("pcSFShipStateCode")
	pcStrShipZip = Session("pcSFShipZip")
	pcStrShipCountryCode = Session("pcSFShipCountryCode")
	pcStrShipEmail = Session("pcSFShipEmail")
	pcStrShipPhone = Session("pcSFShipPhone")
	IF reID<>"0" then
		pcStrShipFax = Session("pcSFShipFax")
		pcStrShipFullName=pcStrShipFirstName&" "&pcStrShipLastName
	end if
	
	if len(pcStrShipNickName)<1 then
		pcStrShipNickName=pcStrShipFullName
	end if
	
	If pcStrShipState<>"" then
		pcStrShipStateCode = ""
	End If
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END: Set Local Variables for recipient
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Check for Validation Errors. Do not proceed if there are errors.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	If pcv_intErr>0 Then	
		response.redirect pcStrPageName&"?reID="&reID&"&msg=" & pcv_strGenericPageError
	Else
		
		'//Open database to update data		
		call openDb()

		IF reID="0" then
			query="UPDATE customers SET shippingAddress='" & pcStrShipAddress & "', shippingCity='" & pcStrShipCity & "', shippingState='" & pcStrShipState & "', shippingStateCode='" & pcStrShipStateCode & "', shippingZip='" & pcStrShipZip & "', shippingCountryCode='" & pcStrShipCountryCode & "', shippingCompany='" & pcStrShipCompany & "', shippingAddress2='" & pcStrShipAddress2 & "', shippingPhone='" & pcStrShipPhone & "', shippingEmail='" & pcStrShipEmail & "' WHERE IDCustomer=" & session("idCustomer") &";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		ELSE
		
			'// Check the Nickname
			pcStrShipNickNameTaken=0
			query="SELECT recipients.idRecipient FROM recipients WHERE recipient_NickName='"&pcStrShipNickName&"' AND idCustomer="&session("idCustomer")&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if NOT rs.eof then
				pcv_stridRecipient = rs("idRecipient")
				if (pcv_stridRecipient=cint(reID))=False then
					'// Nickname in use already
					pcStrShipNickNameTaken=1
				end if
			end if
			set rs=nothing
			
			'// If Nickname is in use, redirect with a message.
			if pcStrShipNickNameTaken=1 then
				'// Alert that this address is already existing in the database.	
				response.redirect pcStrPageName&"?reID="&reID&"&msg=" & dictLanguage.Item(Session("language")&"_CustSAmanage_14")
			else
				query="update recipients set recipient_FullName='" & pcStrShipFullName & "',recipient_Address='" & pcStrShipAddress & "',recipient_City='" & pcStrShipCity & "',recipient_StateCode='" & pcStrShipStateCode & "',recipient_State='" & pcStrShipState & "',recipient_Zip='" & pcStrShipZip & "',recipient_CountryCode='" & pcStrShipCountryCode & "',recipient_Company='" & pcStrShipCompany & "',recipient_Address2='" & pcStrShipAddress2 & "', recipient_NickName='" & pcStrShipNickName & "', recipient_FirstName='" & pcStrShipFirstName & "', recipient_LastName='" & pcStrShipLastName & "', recipient_Phone='" & pcStrShipPhone & "', recipient_Fax='" & pcStrShipFax & "', recipient_Email='" & pcStrShipEmail & "' where idRecipient=" & reID & " and IDCustomer=" & session("idCustomer")
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
			end if
			
		END if
	
		set rs=nothing
		'//Close database before redirecting 		
		call closedb()
		
		'// Clear all sessions
		pcs_ClearAllSessions()
		
		response.redirect "CustSAmanage.asp?msg=2"
	End If
end if




'//Open database to retrieve data from database
call opendb()

IF reID="0" then
	
	query="SELECT Address, City, State, Statecode, Zip, CountryCode, customerCompany, phone, email, shippingAddress, shippingCity, shippingState, shippingStateCode, shippingZip, shippingCountryCode, shippingCompany, shippingAddress2, shippingPhone, shippingEmail FROM customers WHERE idCustomer=" &session("idCustomer")& ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	
	if not rs.eof then
		pcStrAddress=rs("Address")
		pcStrCity=rs("City")
		pcStrState=rs("State")
		pcStrStatecode=rs("Statecode")
		pcStrZip=rs("Zip")
		pcStrCountryCode=rs("CountryCode")
		pcStrcustomerCompany=rs("customerCompany")
		pcStrphone=rs("phone")
		pcStremail=rs("email")
		pcStrShipAddress=rs("shippingAddress")
		pcStrShipCity=rs("shippingCity")
		pcStrShipState=rs("shippingState")
		pcStrShipStateCode=rs("shippingStateCode") 
		pcStrShipZip=rs("shippingZip")
		pcStrShipCountryCode=rs("shippingCountryCode")
		pcStrShipCompany=rs("shippingCompany")
		pcStrShipAddress2=rs("shippingAddress2")
		pcStrShipPhone=rs("shippingPhone")
		pcStrShipEmail=rs("shippingEmail")
		if rs("shippingAddress")<>"" then
		else
			pcStrShipAddress=pcStrAddress
			pcStrShipZip=pcStrZip
			pcStrShipState=pcStrState
			pcStrShipStateCode=pcStrStatecode
			pcStrShipCity=pcStrCity
			pcStrShipCountryCode=pcStrCountryCode
			pcStrShipCompany=pcStrcustomerCompany
			pcStrShipPhone=pcStrphone
			pcStrShipEmail=pcStremail
		end if
		set rs=nothing
	else
		set rs=nothing
		call closeDb()
		response.redirect "CustSAmanage.asp"
	end if	
ELSE
	query="SELECT recipient_FullName, recipient_Address, recipient_City, recipient_StateCode, recipient_State, recipient_Zip, recipient_CountryCode, recipient_Company, recipient_Address2, recipient_NickName, recipient_FirstName, recipient_LastName, recipient_Phone, recipient_Fax, recipient_Email FROM recipients WHERE (((idRecipient)=" & reID & ") AND ((idCustomer)=" & session("idCustomer") & "));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rs.eof then
		set rs=nothing
		call closeDb()
		response.redirect "CustSAmanage.asp"
	else
		pcStrShipFullName=rs("recipient_FullName")
		pcStrShipAddress=rs("recipient_Address")
		pcStrShipCity=rs("recipient_City")
		pcStrShipStateCode=rs("recipient_StateCode")
		pcStrShipState=rs("recipient_State")
		pcTempShipZip=rs("recipient_Zip")
		pcTempSplitZip=split(pcTempShipZip,"||")
		if ubound(pcTempSplitZip)>-1 then
			pcStrShipZip=pcTempSplitZip(0)
			if ubound(pcTempSplitZip)>0 then
				pcStrShipPhone=pcTempSplitZip(1)
			end if
		end if
		pcStrShipCountryCode=rs("recipient_CountryCode")
		pcStrShipCompany=rs("recipient_Company")
		pcStrShipAddress2=rs("recipient_Address2")
		pcStrShipNickName=rs("recipient_NickName")
		pcStrShipFirstName=rs("recipient_FirstName")
		pcStrShipLastName=rs("recipient_LastName")
		pcStrShipPhone=rs("recipient_Phone")
		pcStrShipFax=rs("recipient_Fax")
		pcStrShipEmail=rs("recipient_Email")
		set rs=nothing
		
		'//If First and Last Names are not present, parse FullName
		If len(pcStrShipFirstName)<1 AND len(pcStrShipLastName)<1 AND len(pcStrShipFullName)>0 then
			pcStrShipFullNameArray=split(pcStrShipFullName, " ")
			pcStrShipFirstName=pcStrShipFullNameArray(0)
			if ubound(pcStrShipFullNameArray)>0 then
				pcStrShipLastName=pcStrShipFullNameArray(1)
			end if
		end if
		If len(pcStrShipNickname)<1 then
			pcStrShipNickName=pcStrShipFirstName&" "&pcStrShipLastName
		End if
		
	end if
END IF



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Re-Set the Variables
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


IF reID<>"0" then	
	pcStrShipFirstName = pcf_ResetFormField(Session("pcSFshipFirstName"), pcStrShipFirstName)	
	pcStrShipLastName = pcf_ResetFormField(Session("pcSFshipLastName"), pcStrShipLastName)	
	pcStrShipNickName = pcf_ResetFormField(Session("pcSFshipNickName"), pcStrShipNickName)	
end if
pcStrShipCompany = pcf_ResetFormField(Session("pcSFShipCompany"), pcStrShipCompany)
pcStrShipAddress = pcf_ResetFormField(Session("pcSFShipAddress"), pcStrShipAddress)
pcStrShipAddress2 = pcf_ResetFormField(Session("pcSFShipAddress2"), pcStrShipAddress2)
pcStrShipCity = pcf_ResetFormField(Session("pcSFShipCity"), pcStrShipCity)
pcStrShipState = pcf_ResetFormField(Session("pcSFShipState"), pcStrShipState)
pcStrShipStateCode = pcf_ResetFormField(Session("pcSFShipStateCode"), pcStrShipStateCode)
pcStrShipZip = pcf_ResetFormField(Session("pcSFShipZip"), pcStrShipZip)
pcStrShipCountryCode = pcf_ResetFormField(Session("pcSFShipCountryCode"), pcStrShipCountryCode)
pcStrShipPhone = pcf_ResetFormField(Session("pcSFShipPhone"), pcStrShipPhone)
pcStrShipEmail = pcf_ResetFormField(Session("pcSFShipEmail"), pcStrShipEmail)
IF reID<>"0" then
	pcStrShipFax = pcf_ResetFormField(Session("pcSFShipFax"), pcStrShipFax)
	pcStrShipFullName=pcStrShipFirstName&" "&pcStrShipLastName
end if
if len(pcStrShipNickName)<1 then
	pcStrShipNickName=pcStrShipFullName
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Re-Set the Variables
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<div id="pcMain">		
	<table class="pcMainTable">
		<tr>
			<td>
				<h1><%response.write dictLanguage.Item(Session("language")&"_CustSAmanage_1")%></h1>
			</td>
		</tr>
		<tr>
			<td>
				<h2><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_1")%></h2>
			</td>
		</tr>
		<% msg=getUserInput(request.querystring("msg"),0)
		If msg<>"" then %>
			<tr>
				<td><div class="pcErrorMessage"><%=msg%></div></td>
			</tr>
		<% end if %> 
		<tr>
			<td>
			<form action="<%=pcStrPageName%>" method="post" name="shippingform" class="pcForms">
				<input type="hidden" name="updatemode" value="1">
				<input type=hidden name="reID" value="<%=ReID%>">                   
				<table class="pcShowContent">
					<% if reID<>"0" then %>
						<tr> 
							<td width="25%">
							<p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_16")%></p></td>
							<td width="75%">
							<p>
							<input type="text" name="shipNickName" id="shipNickName" size="20" value="<%=pcStrShipNickName %>">
							<% pcs_RequiredImageTag "shipNickName", pcv_isShipNickNameRequired %>
							</p>
							</td>
						</tr>
						<tr> 
							<td>
								<p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_12")%></p></td>
							<td>
							<p>
							<input type="text" name="shipFirstName" id="shipFirstName" size="20" value="<%=pcStrShipFirstName %>">
						 	<% pcs_RequiredImageTag "shipFirstName", pcv_isShipFirstNameRequired %>
							</p>
							</td>
						</tr>
						<tr> 
							<td>
								<p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_13")%></p></td>
							<td>
								<p>
								<input type="text" name="shipLastName" id="shipLastName" size="20" value="<%=pcStrShipLastName %>">
								<% pcs_RequiredImageTag "shipLastName", pcv_isShipLastNameRequired %>
								</p>
							</td>
						</tr>
					<% end if %>
					<tr>
						<td><p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_9")%></p></td>
						<td width="75%">
							<p>
							<input type="text" name="ShipCompany" id="ShipCompany" size="20" value="<% =pcStrShipCompany %>">
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
					pcv_isStateCodeRequired = pcv_isShipStateCodeRequired '// determines if validation is performed (true or false)
					pcv_isProvinceCodeRequired = pcv_isShipProvinceCodeRequired '// determines if validation is performed (true or false)
					pcv_isCountryCodeRequired = pcv_isShipCountryCodeRequired '// determines if validation is performed (true or false)					
					
					'// #3 Additional Required Info
					pcv_strTargetForm = "shippingform" '// Name of Form
					pcv_strCountryBox = "ShipCountryCode" '// Name of Country Dropdown
					pcv_strTargetBox = "ShipStateCode" '// Name of State Dropdown
					pcv_strProvinceBox =  "ShipState" '// Name of Province Field
					
					'// Set local Country to Session
					if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
						Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrShipCountryCode
					end if
					
					'// Set local State to Session
					if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
						Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrShipStateCode
					end if
					
					'// Set local Province to Session
					if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
						Session(pcv_strSessionPrefix&pcv_strProvinceBox) = pcStrShipState
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
						<td>
						<p>
						<%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_3")%></p></td>
						<td>
							<p>
							<input type="text" name="ShipAddress" id="ShipAddress" size="20" value="<% =pcStrShipAddress %>">
							<% pcs_RequiredImageTag "ShipAddress", pcv_isShipAddressRequired %>
							</p>
						</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>
							<p>
							<input type="text" name="ShipAddress2" id="ShipAddress2" size="20" value="<% =pcStrShipAddress2 %>">
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
							<input type="text" name="ShipCity" id="ShipCity" size="20" value="<% =pcStrShipCity %>">
							<% pcs_RequiredImageTag "ShipCity", pcv_isShipCityRequired %>
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
							<p>
							<input type="text" name="ShipZip" id="ShipZip" size="20" value="<% =pcStrShipZip %>">
							<% pcs_RequiredImageTag "ShipZip", pcv_isShipZipRequired %>
							<span class="pcSmallText"><%response.write dictLanguage.Item(Session("language")&"_checkout_12")%></span>
							</p>
						</td>
					</tr>

					<%	'// Phone Custom Error
					if session("ErrShipPhone")<>"" then %>
						<tr> 
							<td>&nbsp;</td>
							<td><img src="<%=pcf_GenerateIconURL(rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%></td>
						</tr>
						<% session("ErrShipPhone") = ""
					end if %>
					<tr> 
						<td>
							<p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_10")%></p></td>
						<td>
							<p>
							<input type="text" name="ShipPhone" id="ShipPhone" size="20" value="<% =pcStrShipPhone %>">
							<% pcs_RequiredImageTag "ShipPhone", pcv_isShipPhoneRequired %>
							</p>
						</td>
					</tr>
					<% if reID<>"0" then %>
					<%	'// Phone Custom Error
					if session("ErrShipFax")<>"" then %>
					<tr>
						<td>&nbsp;</td>
						<td><img src="<%=pcf_GenerateIconURL(rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_15")%></td>
					</tr>
					<% end if %>
					<tr> 
						<td>
						<p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_14")%></p></td>
						<td>
							<p>
							<input type="text" name="ShipFax" id="ShipFax" size="20" value="<% =pcStrShipFax %>">
							<% pcs_RequiredImageTag "ShipFax", pcv_isShipFaxRequired %>
							</p>
						</td>
					</tr>
				<%end if%>
				<%	'// Email Custom Error
                if session("ErrShipEmail")<>"" then %>
                <tr>
                    <td>&nbsp;</td>
                    <td><img src="<%=pcf_GenerateIconURL(rsIconObj("nexticon"))%>"> <%=dictLanguage.Item(Session("language")&"_Custmoda_16")%></td>
                 </tr>
                <% end if %>
                <tr> 
                    <td>
                    <p><%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_15")%></p></td>
                    <td>
                        <p>
                        <input type="text" name="ShipEmail" id="ShipEmail" size="20" value="<% =pcStrShipEmail %>">
                        <% pcs_RequiredImageTag "ShipEmail", pcv_isShipEmailRequired %>
                        </p>
                    </td>
                </tr>
				<tr>
					<td colspan="2" class="pcSpacer"></td>
				</tr>
				<tr> 
					<td colspan="2">
						<p><input name="submitship" type="image" id="submit" value="<%response.write dictLanguage.Item(Session("language")&"_CustAddModShip_11")%>" src="<%=rslayout("submit")%>">&nbsp;
					  <a href="javascript:location='CustSAmanage.asp'"><img src='<%=rslayout("back")%>' alt='Cancel' border=0></a></p>
					</td>
				</tr>
			</table>
		</form>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->