<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<%
ClearCartURL=session("SFClearCartURL")
if session("admin")<>0 then
	session("idcustomer")=""
	session("pcStrCustName")=""
	session("customerCategory")=""
	session("customerType")=""
	session("ATBCustomer")= Cint(0)
	session("ATBPercentOff")= Cint(0)
	session("customerCategoryType")=""
	session("CustomerGuest")=""
end if

' clear the saved cart cookie
Response.Cookies("SavedCartGUID") = ""

' clear cart data
dim pcCartArray2(100,45)
Session("pcCartSession")=pcCartArray2
Session("pcCartIndex")=Cint(0)

Dim tempURL
tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/default.asp"),"//","/")
tempURL=replace(tempURL,"//","/")
tempURL=replace(tempURL,"https:/","https://")
tempURL=replace(tempURL,"http:/","http://")
response.redirect tempURL & "?ClearCartURL=" & Server.URLEncode(ClearCartURL)
%>