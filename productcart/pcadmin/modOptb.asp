<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<% 
dim query, conntemp, rstemp

pidOption=request.form("idOption")
poptionDescrip=replace(request.form("optionDescrip"),"'","''")
pidOptionGroup=request.form("idOptionGroup")
predirectURL=request.form("redirectURL")
pmode=request.form("mode")
pboton=request.form("modify")

call openDb()

	query="UPDATE options SET optionDescrip='" &poptionDescrip& "' WHERE idOption=" &pidOption
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		set rstemp=nothing
		call closeDb()
	  	response.redirect "techErr.asp?error="& Server.Urlencode("Error renaming option attribute on modOptb.asp") 
	end If
	
	set rstemp=nothing
	call closeDb()
	
	if predirectURL<>"" then 
		response.redirect predirectURL&"&mode="&pmode
	else
		response.redirect "modOptGrpa.asp?idOptionGroup="&pidOptionGroup&"&s=1&msg="&Server.URLEncode("Attribute successfully renamed.")
	end if
%>