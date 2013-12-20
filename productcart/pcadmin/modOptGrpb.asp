<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 

dim f,g,z, randomnum, query, conntemp, rstemp

pidOptionGroup=request.form("idOptionGroup")
poptionGroupDesc=trim(replace(request.form("optionGroupDesc"),"'","''"))
poptionGroupDesc=replace(poptionGroupDesc,"""","&quot;")
if poptionGroupDesc = "" then
	response.redirect "modOptGrpa.asp?idOptionGroup="&pidOptionGroup&"&s=1&msg="&Server.URLEncode("Please enter a description for this option group")
end if

	call openDb()
	'ensure that the new name does not already exist
	query="SELECT * FROM optionsGroups WHERE optionGroupDesc='"&poptionGroupDesc&"'"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	nomatch="0"
	do until rstemp.eof
		comparestring=StrComp(poptionGroupDesc, rstemp("optionGroupDesc"), 1) 
		if comparestring=-1 then
			nomatch="1"
		end if
	rstemp.movenext
	loop
	
	if nomatch="1" then
		set rstemp=nothing
		call closeDb()
		response.redirect "modOptGrpa.asp?msg="&Server.Urlencode("Unable to rename group. There is already a group that uses this name.")&"&idOptionGroup="&pidOptionGroup
		response.end
	end if
	
	query="UPDATE optionsGroups SET optionGroupDesc='" &poptionGroupDesc& "' WHERE idOptionGroup=" &pidOptionGroup
	set rstemp=conntemp.execute(query)
	
	if err.number <> 0 then
		set rstemp=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error modifying the option group name on modOptGrpb.asp") 
	end If
	
	set rstemp=nothing
	call closeDb()
	response.redirect "modOptGrpa.asp?idOptionGroup="&pidOptionGroup&"&s=1&msg="&Server.URLEncode("Option Group successfully renamed.")
%>