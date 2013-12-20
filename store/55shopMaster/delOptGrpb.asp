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

' form parameter 
pidOptionGroup=request.QueryString("idOptionGroup")

call openDb()

' Verifies product assignments for integrity reasons
query="SELECT * FROM options_optionsGroups WHERE idOptionGroup=" &pidOptionGroup
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	pcErrDescription=Err.Description
	set rstemp=nothing
	call closeDb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error in delOptGrpb: "&pcErrDescription) 
end If

if not rstemp.eof then
	set rstemp=nothing
	call closeDb()
	response.redirect "msgb.asp?message="&Server.Urlencode("<b>The selected Option Group cannot be deleted</b><br><br>This option group is currently assigned to one or more products in your store. Once you remove it from all products to which it has been assigned, you will be able to delete it.<br><br><a href=editGrpOptions.asp?idOptionGroup="&pidOptionGroup&">View/remove product assignments</a>.<br><br><a href=manageOptions.asp>Back</a>")
end if

' deleted the option group
query="DELETE FROM optionsGroups WHERE idOptionGroup="&pidOptionGroup
set rstemp=conntemp.execute(query)
if err.number <> 0 then
	pcErrDescription=Err.Description
	set rstemp=nothing
	call closeDb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error in delOptGrpb: "&pcErrDescription) 
end If

' delete all of its options
query="DELETE FROM options_optionsGroups WHERE idOptionGroup="&pidOptionGroup
set rstemp=conntemp.execute(query)
if err.number <> 0 then
	pcErrDescription=Err.Description
	set rstemp=nothing
	call closeDb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error in delOptGrpb: "&pcErrDescription) 
end If
set rstemp=nothing
call closeDb()
response.redirect "manageOptions.asp?s=1&msg="& Server.Urlencode("Option Group deleted successfully.")
%>