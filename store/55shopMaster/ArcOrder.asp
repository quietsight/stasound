<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->

<% 
	dim connTemp, query, rs, qry_ID
 
	qry_ID=request("id")
	if Not IsNumeric(qry_ID) OR qry_ID="0" OR qry_ID="" then
		response.redirect "menu.asp"
	end if
	ptype=request("t")
	if ptype="" then
		ptype=1
	end if	
	
	call openDb()
	if ptype="1" then
		query="UPDATE orders SET pcOrd_Archived=1 WHERE idOrder=" & qry_ID & ";"
	else
		query="UPDATE orders SET pcOrd_Archived=0 WHERE idOrder=" & qry_ID & ";"
	end if			
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=connTemp.execute(query)
	set rs=nothing
	call closedb()
	
	response.redirect "Orddetails.asp?id=" & qry_ID
%>