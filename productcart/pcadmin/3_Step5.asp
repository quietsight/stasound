<% response.Buffer=true %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEX Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" --><!--#include file="AdminHeader.asp"-->	
<% 
FedEX_ServiceStr=Session("ship_FEDEX_SERVICE")
FedEX_freeshipStr=Session("ship_FedEX_freeshipStr")
FedEX_handlingStr=Session("ship_FedEX_handlingStr")
FedEX_EMPackage=Session("ship_FEDEX_FEDEX_PACKAGE")
FEDEX_HEIGHT=Session("ship_FEDEX_HEIGHT")
FEDEX_WIDTH=Session("ship_FEDEX_WIDTH")
FEDEX_LENGTH=Session("ship_FEDEX_LENGTH")
FedEX_HandlingFee=Session("ship_FedEX_HandlingFee")
Dim connTemp, mySQL, rs
call openDb()
set rs=Server.CreateObject("ADODB.Recordset")
mySQL="UPDATE ShipmentTypes SET active=-1 WHERE idShipment=1;"
set rs=connTemp.execute(mySQL)
'clear all informatin out of shipService for FedEX
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='111';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='222';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='333';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='444';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='555';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='666';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='777';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='888';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='999';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='i111';"
set rs=connTemp.execute(mySQL)
mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='i222';"
set rs=connTemp.execute(mySQL)

Dim i
shipServiceArray=split(FedEX_ServiceStr,", ")
for i=0 to ubound(shipServiceArray)
	mySQL="UPDATE shipService SET serviceActive=-1 WHERE serviceCode='"&shipServiceArray(i)&"';"
	'response.write mySQL
	set rs=connTemp.execute(mySQL)
next

freeshipStrArray=split(FedEX_freeshipStr,",")
for i=0 to (ubound(freeshipStrArray)-1)
	freeoveramt=split(freeshipStrArray(i),"|")
	mySQL="UPDATE shipService SET serviceFreeOverAmt="&freeoveramt(1)&" WHERE serviceCode='"&freeoveramt(0)&"';"
	'response.write mySQL
	set rs=connTemp.execute(mySQL)
next

handlingStrArray=split(FedEX_handlingStr,",")
for i=0 to (ubound(handlingStrArray)-1)
	shiphandamt=split(handlingStrArray(i),"|")
	mySQL="UPDATE shipService SET serviceHandlingFee="&shiphandamt(1)&", serviceShowHandlingFee="&shiphandamt(2)&" WHERE serviceCode='"&shiphandamt(0)&"';"
	'response.write mySQL
	set rs=connTemp.execute(mySQL)
next

set rs=nothing
call closeDb()
response.redirect "../includes/PageCreateFedEXConstants.asp"
%>
<!--#include file="AdminFooter.asp"-->