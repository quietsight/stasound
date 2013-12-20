<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% On Error Resume Next
Dim connTemp, query, rs
call openDb()

tmpEmail=getUserInput(request("billemail"),250)

if tmpEmail="" then
	pcTestMsg="false"
end if

if pcTestMsg="" then
	query="SELECT [email] FROM Customers WHERE [email] like '" & tmpEmail & "';"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcTestMsg="false"
	else
		pcTestMsg="true"
	end if
	set rs=nothing
end if

Call SetContentType()
response.write pcTestMsg
call closeDb()
%>