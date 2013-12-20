<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp" -->
<!--#include file="../includes/securitysettings.asp" -->
<%
if session("idCustomer")="" OR session("idCustomer")=0 then
	response.write "<SCRIPT LANGUAGE=JAVASCRIPT><!--"&vbCrlf&vbCrlf	
	response.write "opener.document.location = 'Checkout.asp?cmode=1';"&vbCrlf
	response.write "self.close();"&vbCrlf
	response.write "//--></SCRIPT>"&vbCrlf
	response.end
end if


'// PostBack
if Request("Dcnt")<>"" then
	Dcnt = request("Dcnt")
	if Dcnt=0 then
		Dcnt=""
	end if

	pcStrShippingPhone=pcf_SanitizeApostrophe(Request("pcABPhone"&Dcnt))
	pcv_PayerBusiness=pcf_SanitizeApostrophe(Request("pcABCompany"&Dcnt))
	pcv_FirstName=pcf_SanitizeApostrophe(Request("pcABFirstName"&Dcnt))
	pcv_LastName=pcf_SanitizeApostrophe(Request("pcABLastName"&Dcnt))
	pcv_ShipToName=pcf_SanitizeApostrophe(Request("pcABNickName"&Dcnt))
	pcv_Street1=pcf_SanitizeApostrophe(Request("pcABAddress"&Dcnt))
	pcv_Street2=pcf_SanitizeApostrophe(Request("pcABAddress2"&Dcnt))
	pcv_CityName=pcf_SanitizeApostrophe(Request("pcABCity"&Dcnt))
	pcv_StateOrProvince=pcf_SanitizeApostrophe(Request("pcABProvince"&Dcnt))
	pcv_StateCode=pcf_SanitizeApostrophe(Request("pcABStateCode"&Dcnt))
	pcv_Country=pcf_SanitizeApostrophe(Request("pcABCountryCode"&Dcnt))
	pcv_PostalCode=pcf_SanitizeApostrophe(Request("pcABPostalCode"&Dcnt))
	pcv_Email=pcf_SanitizeApostrophe(Request("pcABEmail"&Dcnt))
	response.Write(pcv_PayerBusiness) & "<br />"
	response.Write(pcv_Payer) & "<br />"
	response.Write(pcv_ShipToName) & "<br />"
	response.Write(pcv_Street1) & "<br />"
	response.Write(pcv_Street2) & "<br />"
	response.Write(pcv_CityName) & "<br />"
	response.Write(pcv_StateOrProvince) & "<br />"
	response.Write(pcv_Country) & "<br />"
	response.Write(pcv_PostalCode) & "<br />"	
	response.Write(pcStrShippingPhone) & "<br />"
	response.Write(pcv_Email) & "<br />"
	'response.End()		
	
	call opendb()
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Update Customer Sessions
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	query="UPDATE pcCustomerSessions SET idCustomer="&session("idCustomer")&", "
	if pcv_ShipToName<>"" then query=query&"pcCustSession_ShippingNickName='"&pcv_ShipToName&"', "
	if pcv_FirstName<>"" then query=query&"pcCustSession_ShippingFirstName='"&pcv_FirstName&"', "
	if pcv_LastName<>"" then query=query&"pcCustSession_ShippingLastName='"&pcv_LastName&"', "
	if pcv_PayerBusiness<>"" then query=query&"pcCustSession_ShippingCompany='"&pcv_PayerBusiness&"', "
	if pcStrShippingPhone<>"" then query=query&"pcCustSession_ShippingPhone='"&pcStrShippingPhone&"',  "
	if pcv_Email<>"" then query=query&"pcCustsession_ShippingEmail='"&pcv_Email&"',  "
	query=query&"pcCustSession_ShippingAddress='"&pcv_Street1&"', "
	query=query&"pcCustSession_ShippingPostalCode='"&pcv_PostalCode&"', "
	if pcv_StateCode<>"" then query=query&"pcCustSession_ShippingStateCode='"&pcv_StateCode&"', "
	if pcv_StateOrProvince<>"" then query=query&"pcCustSession_ShippingProvince='"&pcv_StateOrProvince&"', "
	query=query&"pcCustSession_ShippingCity='"&pcv_CityName&"', "
	query=query&"pcCustSession_ShippingCountryCode='"&pcv_Country&"', "
	query=query&"pcCustSession_ShippingAddress2='"&pcv_Street2&"' WHERE (((idDbSession)="&session("pcSFIdDbSession")&") AND ((randomKey)="&session("pcSFRandomKey")&"));"
	set rs=server.CreateObject("ADODB.RecordSet")
	'response.write query
	'response.end
	set rs=conntemp.execute(query)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Update Customer Sessions
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	

	call closedb()	

	response.write "<SCRIPT LANGUAGE=JAVASCRIPT><!--"&vbCrlf&vbCrlf	
	response.write "opener.document.location = 'login.asp';"&vbCrlf
	response.write "self.close();"&vbCrlf
	response.write "//--></SCRIPT>"&vbCrlf
end if

%>
<head>
<title><%response.write dictLanguage.Item(Session("language")&"_AddressBook_1")%></title>
<SCRIPT LANGUAGE="JavaScript"><!--

function toggle(id)
{
	var tr = document.getElementById(id);
	if (tr==null) { return; }
	var bExpand = tr.style.display == '';
	tr.style.display = (bExpand ? 'none' : '');
	//tr.showhide.value = ('Hide Address');
	var img = document.getElementById('img'+id);
	if (img!=null)
	{
		if (!bExpand)
			img.src = 'minus.gif';
		else
			img.src = 'plus.gif';
	}
}
//--></SCRIPT>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body style="margin: 5px;">
<div id="pcMain">
		<table class="pcMainTable">
			<tr>
			 	<td><h2><%response.write dictLanguage.Item(Session("language")&"_AddressBook_6")%></h2></td>
			</tr>
			<tr>
				<td>
				<%
				dim rs, query, conntemp
				call opendb()
				query="SELECT name, lastName, phone, address, address2, zip, city, state, stateCode, countryCode, shippingaddress, shippingcity, shippingState, shippingStateCode, shippingCountryCode, shippingZip, shippingCompany, shippingAddress2, shippingPhone, shippingEmail FROM customers WHERE (((idcustomer)="&session("idCustomer")&"));"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				pcDefaultName=rs("name")
				pcDefaultLastName=rs("lastName")
				pcDefaultPhone=rs("phone")
				pcDefaultAddress=rs("address")
				pcDefaultAddress2=rs("address2")
				pcDefaultZip=rs("zip")
				pcDefaultCity=rs("city")
				pcDefaultState=rs("state")
				pcDefaultStateCode=rs("stateCode")
				pcDefaultCountryCode=rs("countryCode")				
				pcStrDefaultShipAddress=rs("shippingAddress")
				pcStrDefaultShipCity=rs("shippingCity")
				pcStrDefaultShipState=rs("shippingState")
				pcStrDefaultShipStateCode=rs("shippingStateCode") 
				pcStrShippingCountryCode=rs("shippingCountryCode")
				pcStrShippingZip=rs("shippingZip")
				pcStrDefaultShippingCompany=rs("shippingCompany")
				pcStrShippingAddress2=rs("shippingAddress2")
				pcStrShippingPhone=rs("shippingPhone")
				pcStrShippingEmail=rs("shippingEmail")
				If pcStrDefaultShipAddress="" OR isNULL(pcStrDefaultShipAddress)=True then
					pcStrDefaultShipAddress=pcDefaultAddress
					pcStrDefaultShipCity=pcDefaultCity
					pcStrDefaultShipState=pcDefaultState
					pcStrDefaultShipStateCode=pcDefaultStateCode
					pcStrShippingCountryCode=pcDefaultCountryCode
					pcStrShippingZip=pcDefaultZip					
					pcStrShippingAddress2=pcDefaultAddress2
				End if			
				set rs=nothing
				
				'// Express Checkout			
				if session("ExpressCheckoutPayment")="YES" then
					pcv_strFormAction="method=""post"" action=""pcCustomerAddressBook.asp?Dcnt=0"""
					pcv_strButtonAction=""
				else
					pcv_strFormAction="onSubmit=""return setForm();"""
					pcv_strButtonAction="onsubmit=""return setForm();"""
				end if
				%>
				<FORM NAME="inputForm" <%=pcv_strFormAction%> class="pcForms">
					<table class="pcShowContent">
							<tr>
								<td><p><strong><%=dictLanguage.Item(Session("language")&"_CustSAmanage_10")%></strong></p></td>
								<td width="47%" align="right">	
									<input type="hidden" name="pcABCompany" value="<%=pcStrDefaultShippingCompany%>" />	
									<input type="hidden" name="pcABAddress" value="<%=pcStrDefaultShipAddress%>" />
									<input type="hidden" name="pcABAddress2" value="<%=pcStrShippingAddress2%>" />
									<input type="hidden" name="pcABCity" value="<%=pcStrDefaultShipCity%>" />
									<input type="hidden" name="pcABStateCode" value="<%=pcStrDefaultShipStateCode%>" />
									<input type="hidden" name="pcABProvince" value="<%=pcStrDefaultShipState%>" />
									<input type="hidden" name="pcABPostalCode" value="<%=pcStrShippingZip%>" />
									<input type="hidden" name="pcABCountryCode" value="<%=pcStrShippingCountryCode%>" />
									<input type="hidden" name="pcABPhone" value="<%=pcStrShippingPhone%>" />
									<input type="hidden" name="pcABEmail" value="<%=pcStrShippingEmail%>" />
								<a href="#" onClick="javascript:toggle('Row')">
								<input type="button" name="showhide" value="<%response.write dictLanguage.Item(Session("language")&"_AddressBook_2")%>"></a>
								&nbsp;
								<input type="submit" name="UPD" value="<%response.write dictLanguage.Item(Session("language")&"_AddressBook_3")%>" <%=pcv_strButtonAction%> class="submit2" />
								</td>
							</tr>
							<tr>
								<td colspan="2" class="pcSpacer"></td>
							</tr>								
							<tr id="Row<%=DCnt%>" style="display: none;" class="pcPageDesc">
								<td colspan="2" style="padding:5px;">
								<%=pcStrDefaultShipAddress%><br />
								<% if pcStrShippingAddress2<>"" and isNULL(pcStrShippingAddress2)=False then %>
								<%=pcStrShippingAddress2%><br />
								<% end if %>
								<%=pcStrDefaultShipCity%>,&nbsp;<% if pcStrDefaultShipStateCode <> "" then response.write pcStrDefaultShipStateCode else response.write pcStrDefaultShipState%>&nbsp;<%=pcStrShippingCountryCode%>&nbsp;<%=pcStrShippingZip%>
								</td>
							</tr>
					</table>
				</FORM>
				<%
				response.write "<SCRIPT LANGUAGE=JAVASCRIPT><!--"&vbCrlf&vbCrlf			
				response.write "function setForm"&i&"() {"&vbCrlf			
				response.write "opener.document.loginform.pcShippingCompany.value = document.inputForm.pcABCompany.value;"&vbCrlf
				response.write "opener.document.loginform.pcShippingAddress.value = document.inputForm.pcABAddress.value;"&vbCrlf
				response.write "opener.document.loginform.pcShippingAddress2.value = document.inputForm.pcABAddress2.value;"&vbCrlf
				response.write "opener.document.loginform.pcShippingCity.value = document.inputForm.pcABCity.value;"&vbCrlf
				response.write "opener.document.loginform.pcShippingPhone.value = document.inputForm.pcABPhone.value;"&vbCrlf
				response.write "opener.document.loginform.pcShippingEmail.value = document.inputForm.pcABEmail.value;"&vbCrlf
				response.write "opener.document.loginform.pcShippingStateCode.value = document.inputForm.pcABStateCode.value;"&vbCrlf
				response.write "opener.document.loginform.pcShippingProvince.value = document.inputForm.pcABProvince.value;"&vbCrlf
				response.write "opener.document.loginform.pcShippingPostalCode.value = document.inputForm.pcABPostalCode.value;"&vbCrlf
				response.write "opener.document.loginform.pcShippingCountryCode.value = document.inputForm.pcABCountryCode.value;"&vbCrlf			
				response.write "opener.SelectState('pcShippingCountryCode', 'pcShippingStateCode', 'pcShippingProvince', '"& pcStrDefaultShipStateCode &"', '2');"&vbCrlf
				response.write "self.close();"&vbCrlf
				response.write "return false;"&vbCrlf
				response.write "}"&vbCrlf
				response.write "//--></SCRIPT>"&vbCrlf
				%>
			</td>
		</tr>		
		<tr>
			<td>
		<% 		
		query="SELECT recipients.idRecipient, recipients.recipient_FullName, recipients.recipient_Address, recipients.recipient_City, recipients.recipient_StateCode, recipients.recipient_State, recipients.recipient_Zip, recipients.recipient_CountryCode, recipients.recipient_Company, recipients.recipient_Address2, recipients.recipient_NickName, recipients.recipient_FirstName, recipients.recipient_LastName, recipients.recipient_phone, recipient_Fax, recipient_Email FROM recipients WHERE (((recipients.idCustomer)="&session("idCustomer")&"));" 
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if rs.eof then
	
		end if
			
		DCnt=0 
		i=0
		do until rs.eof
			DCnt=DCnt+1
			i=i+1
			pcIntIdRecipient=rs("idRecipient")
			pcStrABFullName=rs("recipient_FullName")
			pcStrABAddress=rs("recipient_Address")
			pcStrABCity=rs("recipient_City")
			pcStrABStateCode=rs("recipient_StateCode")
			pcStrABProvince=rs("recipient_State")
			pcStrABPostalCode=rs("recipient_Zip")
			if instr(pcStrABPostalCode,"||") then
				pcStrABPostalCodeSplit=split(pcStrABPostalCode,"||")
				pcStrABPostalCode=pcStrABPostalCodeSplit(0)
			end if
			pcStrABCountryCode=rs("recipient_CountryCode")
			pcStrABCompany=rs("recipient_Company")
			pcStrABAddress2=rs("recipient_Address2")
			pcStrABNickName=rs("recipient_NickName")
			pcStrABFirstName=rs("recipient_FirstName")
			pcStrABLastName=rs("recipient_LastName")
			pcStrABPhone=rs("recipient_phone")
			pcStrABFax=rs("recipient_Fax")
			pcStrABEmail=rs("recipient_Email")
			if pcStrABFirstName="" AND pcStrABLastName="" then
				if pcStrABFullName<>"" then
					pcStrABFullNameSplit=split(pcStrABFullName," ")
					if ubound(pcStrABFullNameSplit)>0 then
						pcStrABFirstName=pcStrABFullNameSplit(0)
						pcStrABLastName=pcStrABFullNameSplit(1)
					else
						pcStrABFirstName=pcStrABFullNameSplit(0)
					end if
				end if
			end if
			pcShowName=pcStrABFullName&""
			if pcStrABNickName<>"" then
				pcShowName=pcStrABNickName&""
			end if
			if len(pcShowName)=0 then
				pcShowName=pcStrABNickName
				pcShowName="No Shipping Name"
			end if
			
			'// Express Checkout			
			if session("ExpressCheckoutPayment")="YES" then
				pcv_strFormAction="method=""post"" action=""pcCustomerAddressBook.asp?Dcnt="&DCnt&""""
				pcv_strButtonAction=""
			else
				pcv_strFormAction="onSubmit=""return setForm"&DCnt&"();"""
				pcv_strButtonAction="onsubmit=""return setForm();"""
			end if
			%>
			<FORM NAME="inputForm<%=DCnt%>" <%=pcv_strFormAction%> class="pcForms">
				<table class="pcShowContent">
						<tr>
							<td><p><strong><%=pcShowName%></strong></p></td>
							<td width="47%" align="right">
								
								<input type="hidden" name="pcABIdRecipient<%=DCnt%>" value="<%=pcIntIdRecipient%>" />
								<input type="hidden" name="pcABNickName<%=DCnt%>" value="<%=pcStrABNickName%>" />
								<input type="hidden" name="pcABCompany<%=DCnt%>" value="<%=pcStrABCompany%>" />
								<input type="hidden" name="pcABFirstName<%=DCnt%>" value="<%=pcStrABFirstName%>" />
								<input type="hidden" name="pcABLastName<%=DCnt%>" value="<%=pcStrABLastName%>" />
								<input type="hidden" name="pcABAddress<%=DCnt%>" value="<%=pcStrABAddress%>" />
								<input type="hidden" name="pcABAddress2<%=DCnt%>" value="<%=pcStrABAddress2%>" />
								<input type="hidden" name="pcABCity<%=DCnt%>" value="<%=pcStrABCity%>" />
								<input type="hidden" name="pcABStateCode<%=DCnt%>" value="<%=pcStrABStateCode%>" />
								<input type="hidden" name="pcABProvince<%=DCnt%>" value="<%=pcStrABProvince%>" />
								<input type="hidden" name="pcABPostalCode<%=DCnt%>" value="<%=pcStrABPostalCode%>" />
								<input type="hidden" name="pcABCountryCode<%=DCnt%>" value="<%=pcStrABCountryCode%>" />
								<input type="hidden" name="pcABPhone<%=DCnt%>" value="<%=pcStrABPhone%>" />
                                <input type="hidden" name="pcABFax<%=DCnt%>" value="<%=pcStrABFax%>" />
                                <input type="hidden" name="pcABEmail<%=DCnt%>" value="<%=pcStrABEmail%>" />
								<a href="#" onClick="javascript:toggle('Row<%=DCnt%>')"><input type="button" name="showhide<%=DCnt%>" value="<%response.write dictLanguage.Item(Session("language")&"_AddressBook_2")%>"></a>&nbsp;<input type="submit" name="UPD<%=DCnt%>" value="<%response.write dictLanguage.Item(Session("language")&"_AddressBook_3")%>" <%=pcv_strButtonAction%> class="submit2" />
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
							
						<tr id="Row<%=DCnt%>" style="display: none;" class="pcPageDesc">
							<td colspan="2" style="padding:5px;">
							<%=pcStrABAddress%><br />
							<% if pcStrABAddress2<>"" and isNULL(pcStrABAddress2)=False then %>
							<%=pcStrABAddress2%><br />
							<% end if %>
							<%=pcStrABCity%>,&nbsp;<% if pcStrABStateCode <> "" then response.write pcStrABStateCode else response.write pcStrABProvince%>&nbsp;<%=pcStrABCountryCode%>&nbsp;<%=pcStrABPostalCode%>
							</td>
						</tr>
				</table>
			</FORM>
			<%
			response.write "<SCRIPT LANGUAGE=JAVASCRIPT><!--"&vbCrlf&vbCrlf			
			response.write "function setForm"&i&"() {"&vbCrlf
			response.write "opener.document.loginform.pcShippingReferenceId.value = document.inputForm"&i&".pcABIdRecipient"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingNickName.value = document.inputForm"&i&".pcABNickName"&i&".value;"&vbCrlf	
			response.write "opener.document.loginform.pcShippingCompany.value = document.inputForm"&i&".pcABCompany"&i&".value;"&vbCrlf	
			response.write "opener.document.loginform.pcShippingFirstName.value = document.inputForm"&i&".pcABFirstName"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingLastName.value = document.inputForm"&i&".pcABLastName"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingAddress.value = document.inputForm"&i&".pcABAddress"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingAddress2.value = document.inputForm"&i&".pcABAddress2"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingCity.value = document.inputForm"&i&".pcABCity"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingStateCode.value = document.inputForm"&i&".pcABStateCode"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingProvince.value = document.inputForm"&i&".pcABProvince"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingPostalCode.value = document.inputForm"&i&".pcABPostalCode"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingCountryCode.value = document.inputForm"&i&".pcABCountryCode"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingPhone.value = document.inputForm"&i&".pcABPhone"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingFax.value = document.inputForm"&i&".pcABFax"&i&".value;"&vbCrlf
			response.write "opener.document.loginform.pcShippingEmail.value = document.inputForm"&i&".pcABEmail"&i&".value;"&vbCrlf
			response.write "opener.SelectState('pcShippingCountryCode', 'pcShippingStateCode', 'pcShippingProvince', '"&pcStrABStateCode&"', '2');"&vbCrlf
			response.write "self.close();"&vbCrlf
			response.write "return false;"&vbCrlf
			response.write "}"&vbCrlf
			response.write "//--></SCRIPT>"&vbCrlf
			rs.moveNext
		loop
	
		set rs=nothing
		call closedb() %>
		
		<% 		
		response.write "<SCRIPT LANGUAGE=JAVASCRIPT><!--"&vbCrlf&vbCrlf

			response.write "function addnew() {"&vbCrlf
			response.write "opener.document.location = 'CustAddShip.asp?mode=1';"&vbCrlf
			response.write "self.close();"&vbCrlf
			response.write "return false;"&vbCrlf
			response.write "}"&vbCrlf

		response.write "//--></SCRIPT>"&vbCrlf
		%>
		</td>
	</tr>
	<tr>
		<td><hr></td>
	</tr>
	<tr>
		<td align="center">
					<form class="pcForms">
						<input type="button" class="submit2" name="Add New Address" value="<%response.write dictLanguage.Item(Session("language")&"_AddressBook_4")%>" onClick="javascript:addnew();" />
						&nbsp;&nbsp;
						<input type="button" value="<%response.write dictLanguage.Item(Session("language")&"_AddressBook_5")%>" onClick="javascript:window.close();">
					</form>
		</td>
	</tr>
	</table>
	</div>
</body>
</html>
