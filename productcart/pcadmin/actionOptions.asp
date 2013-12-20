<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<%
Dim intidoption, intidoptiongrp, strOptionName

If Request.QueryString("delete") <> "" Then
	intidoption=Request.QueryString("delete")
	intidoptiongrp=Request.QueryString("idOptionGroup")
	
	Dim rs, connTemp, query
	call openDb()
	
	query="Delete From options_optionsGroups WHERE idoption="&intidoption&" AND idOptionGroup="&intidoptiongrp&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	query="Delete From optGrps WHERE idoption="&intidoption&" AND idOptionGroup="&intidoptiongrp&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	set rs=nothing
	call closeDB()
	Response.Redirect "modOptGrpa.asp?s=1&idOptionGroup="&intidoptiongrp&"&msg="&server.URLencode("Option attribute successfully deleted.")
End If
%>