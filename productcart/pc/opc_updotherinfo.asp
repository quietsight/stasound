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

Call SetContentType()

if session("idCustomer")=0 OR session("idCustomer")="" then
	response.clear
	Call SetContentType()
	response.write "SECURITY"
	response.End
end if

pcErrMsg=""

pcStrOrderNickName=URLDecode(getUserInput(request("OrderNickName"),250))
pcStrOrderComments=URLDecode(getUserInput(request("OrderComments"),0))

pcStrGcReName=URLDecode(getUserInput(request("GcReName"),250))
pcStrGcReEmail=URLDecode(getUserInput(request("GcReEmail"),250))
pcStrGcReMsg=URLDecode(getUserInput(request("GcReMsg"),0))

if pcStrGcReEmail<>"" then
	pcStrGcReEmail=replace(pcStrGcReEmail," ","")
	if instr(pcStrGcReEmail,"@")=0 or instr(pcStrGcReEmail,".")=0 then
		pcErrMsg=pcErrMsg & "<li>"&dictLanguage.Item(Session("language")&"_opc_70")&"</li>"
	end if
end if

if pcErrMsg="" then
	query="UPDATE pcCustomerSessions SET pcCustSession_OrderName='" & pcStrOrderNickName & "',pcCustSession_Comment='" & pcStrOrderComments & "',pcCustSession_GcReName='" & pcStrGcReName & "',pcCustSession_GcReEmail='" & pcStrGcReEmail & "',pcCustSession_GcReMsg='" & pcStrGcReMsg & "' WHERE pcCustomerSessions.idDbSession="&session("pcSFIdDbSession")&" AND pcCustomerSessions.randomKey="&session("pcSFRandomKey")&" AND pcCustomerSessions.idCustomer="&session("idCustomer")&";"
	set rs=connTemp.execute(query)
	set rs=nothing
	OKmsg="OK"
end if

if pcErrMsg<>"" then
	pcErrMsg=dictLanguage.Item(Session("language")&"_opc_71")&"<br><ul>" & pcErrMsg & "</ul>"
	response.write pcErrMsg
else
	response.write OKmsg
end if
%>


