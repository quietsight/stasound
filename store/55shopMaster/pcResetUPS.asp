<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=8%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle="Delete UPS registration" %>
<% 
dim query, conntemp, rs

call openDb()

query="UPDATE ups_license SET  ups_UserId='', ups_Password='', ups_AccessLicense='';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

query="UPDATE ShipmentTypes SET active=0, userId='', [password]='', AccessLicense='' WHERE idShipment=3;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='01';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='02';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='03';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='07';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='08';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='11';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='12';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='13';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='14';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='54';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='59';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
query="UPDATE shipService SET serviceActive=0 WHERE serviceCode='65';"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
set rs=nothing
call closeDb()
response.redirect "viewShippingOptions.asp"
%>
