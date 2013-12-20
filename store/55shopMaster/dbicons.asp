<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 
dim query, conntemp, rs
call openDb()

erroricon=Request.QueryString("File1")
requiredicon=Request.QueryString("File2")
errorfieldicon=Request.QueryString("File3")
previousicon=Request.QueryString("File4")
nexticon=Request.QueryString("File5")
zoom=Request.QueryString("File6")
discount=Request.QueryString("File7")
arrowUp=Request.QueryString("File8")
arrowDown=Request.QueryString("File9")
	
cma="0"	
query="UPDATE icons SET "
If erroricon <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	erroricon2="images/pc/"&erroricon
	query=query &"erroricon='"& erroricon2 &"'"
End If
If requiredicon <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	requiredicon2="images/pc/"&requiredicon
	query=query &"requiredicon='"& requiredicon2 &"'"
End If
If errorfieldicon <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	errorfieldicon2="images/pc/"&errorfieldicon
	query=query &"errorfieldicon='"& errorfieldicon2 &"'"
End If
If previousicon <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	previousicon2="images/pc/"&previousicon
	query=query &"previousicon='"& previousicon2 &"'"
End If
If nexticon <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	nexticon2="images/pc/"&nexticon
	query=query &"nexticon='"& nexticon2 &"'"
End If
If zoom <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	zoom2="images/pc/"&zoom
	query=query &"zoom='"& zoom2 &"'"
End If
If discount <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	discount2="images/pc/"&discount
	query=query &"discount='"& discount2 &"'"
End If
If arrowUp <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	arrowUp2="images/pc/"&arrowUp
	query=query &"arrowUp='"& arrowUp2 &"'"
End If
If arrowDown <> "" Then
	If cma="0" Then
	Else
		query=query & ","
	End If
	cma="1"
	arrowDown2="images/pc/"&arrowDown
	query=query &"arrowDown='"& arrowDown2 &"'"
End If


query=query &" WHERE id=1"

set rs=Server.CreateObject("ADODB.Recordset")     
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	call closedb()
	response.write "Error: "&Err.Description
end If 

set rs=nothing
call closedb()

s=request.querystring("s")
msg=request.querystring("msg")
response.redirect "AdminIcons.asp?msg="&Server.URLEncode(msg)&"&s="&s 
%>