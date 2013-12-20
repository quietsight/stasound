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
Dim connTemp, query, rs
call openDb()

if session("CustomerGuest")<>"1" OR session("idCustomer")=0 OR session("idCustomer")="" then
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
end if

if session("idCustomer")>"0" AND (session("CustomerGuest")="0" OR session("CustomerGuest")="2") then
	response.clear
	Call SetContentType()
	if session("CustomerGuest")="2" then
		response.write "REGA"
	else
		response.write "REG"
	end if
	response.End
end if

pcErrMsg=""

pcStrNewPass1=getUserInput(request("newPass1"),250)
pcStrNewPass2=getUserInput(request("newPass2"),0)

if pcStrNewPass1="" then
	pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_5") & "</li>"
end if

if pcStrNewPass2="" then
	pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_4") & "</li>"
end if

if pcStrNewPass1<>pcStrNewPass2 then
	pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_3") & "</li>"
end if

if pcErrMsg="" then
	query="SELECT [email] FROM Customers WHERE idCustomer=" & session("idCustomer") & ";"
	set rs=connTemp.execute(query)
	if rs.eof then
		pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_opc_createacc_2") & "</li>"
	else
		pEmail=rs("email")
	end if
	set rs=nothing
end if

newCustomerGuest=0

if pcErrMsg="" then

	query="SELECT idCustomer FROM Customers WHERE email LIKE '" & replace(pEmail,"'","''") & "' AND idCustomer<>" & session("idCustomer") & " AND pcCust_Guest=0;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		newCustomerGuest=2
	end if	
	set rs=nothing

	pcStrNewPass1=enDeCrypt(pcStrNewPass1, scCrypPass)	
	query="UPDATE Customers SET password='" & pcStrNewPass1 & "',pcCust_Guest=" & newCustomerGuest & " WHERE idCustomer=" & session("idCustomer") & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
	session("CustomerGuest")=newCustomerGuest
	if session("CustomerGuest")="2" then
		OKmsg="OKA"
	else
		OKmsg="OK"
	end if
	
	If session("CustomerGuest")="0" Then
		'// Send New Customer Email
		query="SELECT [name], lastName, email FROM Customers WHERE idCustomer=" & session("idCustomer") & ";"
		set rs=connTemp.execute(query)
		if not rs.eof then
			pcStrBillingFirstName = rs("name")
			pcStrBillingLastName = rs("lastName")
			pcStrCustomerEmail = rs("email")
		end if	
		set rs=nothing
		pcv_strNewCustEmail="1" '// Send to Customer
		%> <!--#include file="adminNewCustEmail.asp"--> <%
	End If
end if

if pcErrMsg<>"" then
	pcErrMsg= dictLanguage.Item(Session("language")&"_opc_createacc_1") & "<ul>" & pcErrMsg & "</ul>"
	response.write pcErrMsg
else
	response.write OKmsg
end if
call closeDb()
%>