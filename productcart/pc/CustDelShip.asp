<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<%
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If 

reID=getUserInput(request("reID"),0)
if not reID<>"" then
	response.redirect "CustSAmanage.asp"
end if

dim query, conntemp, rs

call openDb()

query="delete from recipients where idRecipient=" & reID & " and IDCustomer=" & session("idCustomer")
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

set rs=nothing

call closedb()

if err.number<>0 then
	response.redirect "CustSAmanage.asp"
else
	response.redirect "CustSAmanage.asp?msg=3"
end if
%>