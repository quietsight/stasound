<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/productcartFolder.asp"-->
<%
'Get Path Info
pcv_filePath=Request.ServerVariables("PATH_INFO")
do while instr(pcv_filePath,"/")>0
	pcv_filePath=mid(pcv_filePath,instr(pcv_filePath,"/")+1,len(pcv_filePath))
loop

pcv_Query=Request.ServerVariables("QUERY_STRING")

if pcv_Query<>"" then
	pcv_filePath=pcv_filePath & "?" & pcv_Query
end if

if request("redirectUrl")<>"" then
	Session("SFStrRedirectUrl")=getUserInput(request("redirectUrl"),0)
else
	Session("SFStrRedirectUrl")=pcv_filePath
end if

'// If the Customer is a Guest and the page does not allow guests... redirect them somewhere
if ((Session("idCustomer")=0) AND ((session("REGidCustomer")<"1") OR (AllowGuestAccess=0))) OR (Session("idCustomer")>"0" AND session("CustomerGuest")>"0" AND AllowGuestAccess=0) then

	'// If the Customer is trying to access a second account, then ask them to consolidate.
	if Session("idCustomer")>"0" AND session("CustomerGuest")="2" then
		response.redirect "CustConsolidate.asp"
	'// Otherwise, ask them to login
	else
		if session("pcCartIndex")<"0" then
			Session("idCustomer")=""
		else
			Session("idCustomer")=0
		end if
		session("CustomerGuest")=0
		session("CustomerType")=0
		session("customerCategory")=0
		dim strRedirectSSL
		strRedirectSSL="Checkout.asp?cmode=1"
		if scSSL="1" AND scIntSSLPage="1" then
			strRedirectSSL=replace((scSslURL&"/"&scPcFolder&"/pc/Checkout.asp?cmode=1"),"//","/")
			strRedirectSSL=replace(strRedirectSSL,"https:/","https://")
			strRedirectSSL=replace(strRedirectSSL,"http:/","http://")
		end if
		response.redirect strRedirectSSL
	end if
	
end if
%>