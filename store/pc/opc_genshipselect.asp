<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<!--#include file="opc_contentType.asp" -->
<% On Error Resume Next

Call SetContentType()

if session("idCustomer")=0 OR session("idCustomer")="" then
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
end if

dim rs,connTemp,query
call openDb()

pcShipArr=""

IF session("idCustomer")>0 then

	query="SELECT [name], lastName, email, phone, fax, customerCompany, address,address2, city, state, stateCode, zip, countryCode, shippingCompany, shippingAddress, shippingAddress2, shippingCity, shippingState, shippingStateCode,shippingZip,shippingCountryCode,shippingPhone,shippingEmail, pcCust_Residential,shippingFax FROM customers WHERE idcustomer="&session("idCustomer")&" AND pcCust_Guest=" & session("CustomerGuest") & ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	pcStrDefaultShipAddress=rs("shippingAddress")

	If len(pcStrDefaultShipAddress)<1 Then
	
		If NOT session("CustomerGuest")>0 Then
				pcShipArr="0|*|Default shipping address|*|" & rs("name") & "|*|" & rs("lastName") & "|*|" & rs("email") & "|*|" & rs("phone") & "|*|" & rs("fax") & "|*|" & rs("customerCompany") & "|*|" & rs("address") & "|*|" & rs("address2") & "|*|" & rs("city") & "|*|" & rs("state") & "|*|" & rs("stateCode") & "|*|" & rs("zip") & "|*|" & rs("countryCode") & "|*|" & rs("pcCust_Residential") & "|$|"
		End If	
			
	Else
	
		pcStrDefaultShipEmail=rs("shippingEmail")
		if pcStrDefaultShipEmail="" then
			pcStrDefaultShipEmail=rs("email")
		end if
		pcStrDefaultShipPhone=rs("shippingPhone")
		if pcStrDefaultShipPhone="" then
			pcStrDefaultShipPhone=rs("phone")
		end if
		pcStrDefaultShipCompany=rs("shippingCompany")
		if pcStrDefaultShipCompany="" then
			pcStrDefaultShipCompany=rs("customerCompany")
		end if
		pcStrDefaultShipFax=rs("shippingFax")
		if pcStrDefaultShipFax="" then
			pcStrDefaultShipFax=rs("fax")
		end if
		
		If (NOT session("CustomerGuest")=1 OR session("CustomerGuest")=2) AND (len(pcStrDefaultShipEmail)>0) Then
				pcShipArr="0|*|Default shipping address|*|" & rs("name") & "|*|" & rs("lastName") & "|*|" & pcStrDefaultShipEmail & "|*|" & pcStrDefaultShipPhone & "|*|" & pcStrDefaultShipFax & "|*|" & pcStrDefaultShipCompany & "|*|" & rs("shippingAddress") & "|*|" & rs("shippingAddress2") & "|*|" & rs("shippingCity") & "|*|" & rs("shippingState") & "|*|" & rs("shippingStateCode") & "|*|" & rs("shippingZip") & "|*|" & rs("shippingCountryCode") & "|*|" & rs("pcCust_Residential") & "|$|"
		End If
	
	End if
	set rs=nothing


	query="SELECT idRecipient, recipient_NickName, recipient_FirstName, recipient_LastName, recipient_Email, recipient_Phone, recipient_Fax, recipient_Company, recipient_Address, recipient_Address2, recipient_City, recipient_State, recipient_StateCode, recipient_Zip, recipient_CountryCode, Recipient_Residential FROM recipients WHERE idCustomer="&session("idCustomer")&";"
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)				
	do while not rs.eof
		reNickName=trim(rs("recipient_NickName"))
		if NOT len(reNickName)>1 OR isNULL(reNickName) then
			reNickName=dictLanguage.Item(Session("language")&"_CustSAmanage_12")
		end if
		pcShipArr=pcShipArr & rs("idRecipient") & "|*|" & reNickName & "|*|" & rs("recipient_FirstName") & "|*|" & rs("recipient_LastName") & "|*|" & rs("recipient_Email") & "|*|" & rs("recipient_Phone") & "|*|" & rs("recipient_Fax") & "|*|" & rs("recipient_Company") & "|*|" & rs("recipient_Address") & "|*|" & rs("recipient_Address2") & "|*|" & rs("recipient_City") & "|*|" & rs("recipient_State") & "|*|" & rs("recipient_StateCode") & "|*|" & rs("recipient_Zip") & "|*|" & rs("recipient_CountryCode") & "|*|" & rs("Recipient_Residential") & "|$|"
		rs.movenext
	loop
	set rs=nothing

	response.write session("pcShipOpt") & "|$|" & pcShipArr
END IF
%>		