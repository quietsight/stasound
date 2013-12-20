<%@ Language=VBScript %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/productcartinc.asp"--> 
<!--#INCLUDE FILE="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->

<% 
if session("idCustomer")="" OR session("idCustomer")=0 then
 	response.redirect "msg.asp?message=211"     
end if

dim conntemp, rs, query
call opendb()

' extract real idorder (without prefix)
pTrueOrderId=(int(request("idOrder"))-scpre)
pcv_tempID=0
pcv_tempOrderstatus=0

'verify that this order doesn't alreay exists and that the idCustomer is only that of the customer logged in.
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

%>
<% psslurl=request("psslurl") %>
<HTML>
<HEAD>
</HEAD>
<body onLoad="document.frmSSL.submit();">
<% 
'----------------------------------------------------
' START - Free order, skip payment pages
'----------------------------------------------------
if request.QueryString("purchasemode")="FREE" then
	if scSSL="1" then
		tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwReturn.asp"),"//","/")
	else
		tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwReturn.asp"),"//","/")
	end if
	tempURL=replace(tempURL,"//","/")
	tempURL=replace(tempURL,"http:/","http://")
	tempURL=replace(tempURL,"https:/","https://") %>
	<form name="frmSSL" method="POST" action="<%=tempURL%>">
		<input type="hidden" name="purchasemode" value="FREE">
		<input type="hidden" name="idOrder" value="<%=request("idOrder")%>"><br>
        <input type="hidden" name="pcIsSubscription" value="<%=session("pcIsSubscription")%>"><br>
        <input type="hidden" name="pcIsSubTrial" value="<%=session("pcIsSubTrial")%>"><br>
	</form>
	<% 
    '----------------------------------------------------
    ' END - Free order
    '----------------------------------------------------
else
    '----------------------------------------------------
    ' START - Redirect to payment page
    '					Check for use of SSL certificate
    '----------------------------------------------------
	if scSSL="1" then 
		tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/"&psslurl),"//","/")
		tempURL=replace(tempURL,"http:/","http://")
		tempURL=replace(tempURL,"https:/","https://")%>
		<form name="frmSSL" method="POST" action="<%=tempURL%>">
	<% else
		tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"&psslurl),"//","/")
		tempURL=replace(tempURL,"//","/")
		tempURL=replace(tempURL,"http:/","http://")
		tempURL=replace(tempURL,"https:/","https://") %>
		<form name="frmSSL" method="POST" action="<%=tempURL%>">
	<% end if	%>
	<input type="hidden" name="idPayment" value="<%=request("idPayment")%>"><br>
	<input type="hidden" name="idOrder" value="<%=request("idOrder")%>"><br>
    <input type="hidden" name="pcIsSubscription" value="<%=session("pcIsSubscription")%>"><br>
    <input type="hidden" name="pcIsSubTrial" value="<%=session("pcIsSubTrial")%>"><br>
</form>
<%
'----------------------------------------------------
' END - Redirect to payment page
'----------------------------------------------------
end if
%>