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
dim f,g,z, query, conntemp, rstemp, prdFrom

poptionGroupDesc=request.querystring("optionGroupDesc")
poptionGroupDesc = replace(poptionGroupDesc,"'","''")

'get product ID of the referring product page, if any
prdFrom = request.QueryString("prdFrom")
if not validNum(prdFrom) then prdFrom = 0
AssignID = request.QueryString("AssignID")

if trim(poptionGroupDesc)="" then
   response.redirect "msg.asp?message=16"
end if

call openDb()

' insert in to db new option group
query="INSERT INTO optionsGroups (optionGroupDesc) VALUES ('" &poptionGroupDesc& "')"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in addoptiongroupexec: "&Err.Description) 
end If

query="SELECT Count(*) As TotalGrp FROM optionsGroups WHERE optionGroupDesc like '" &poptionGroupDesc& "';"
set rstemp=connTemp.execute(query)

Count=0
if not rstemp.eof then
	Count=rstemp("TotalGrp")
	if Count<>"" then
	else
		Count=0
	end if
	Count=Cint(Count)
end if
set rstemp=nothing

tmpStr=""
if Count>1 then
	tmpStr="<br /><br />Note: Option Group(s) with the same name already exist."
end if

' if the referring product exist, go back to that page, otherwise go to Manage Options
if prdFrom = 0 then
  response.redirect "ManageOptions.asp?s=1&msg="&Server.Urlencode("Successfully added new Option Group." & tmpStr)
 else
  response.redirect "modPrdOpta1.asp?AssignID="& AssignID &"&idproduct="& prdFrom
end if
%>
<!--#include file="AdminFooter.asp"-->