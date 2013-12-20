<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="USPS Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
		<% 		
		USPS_Server=Session("ship_USPS_Server")
		USPS_LabelServer=Session("ship_USPS_LabelServer")
		USPS_userID=Session("ship_USPS_ID")
		USPS_ServiceStr=Session("ship_USPS_Service")
		USPS_freeshipStr=Session("ship_USPS_freeshipStr")
		USPS_handlingStr=Session("ship_USPS_handlingStr")
		USPS_EM_PACKAGE=Session("ship_USPS_EM_PACKAGE")
		USPS_PM_PACKAGE=Session("ship_USPS_PM_PACKAGE")
		USPS_HEIGHT=Session("ship_USPS_HEIGHT")
		USPS_WIDTH=Session("ship_USPS_WIDTH")
		USPS_LENGTH=Session("ship_USPS_LENGTH")
		USPS_HandlingFee=Session("ship_USPS_HandlingFee")
		Dim connTemp, mySQL, rs
		call openDb()
		set rs=Server.CreateObject("ADODB.Recordset")
		mySQL="UPDATE ShipmentTypes SET shipServer='"&USPS_Server&"', userID='"&USPS_userID&"', AccessLicense='"&USPS_LabelServer&"', active=-1 WHERE idShipment=4;"
		set rs=connTemp.execute(mySQL)
		'clear all informatin out of shipService for USPS
		mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9901';"
		set rs=connTemp.execute(mySQL)
		mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9902';"
		set rs=connTemp.execute(mySQL)
		mySQL="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='9903';"
		set rs=connTemp.execute(mySQL)
		Dim i
		shipServiceArray=split(USPS_ServiceStr,", ")
		for i=0 to ubound(shipServiceArray)
			mySQL="UPDATE shipService SET serviceActive=-1 WHERE serviceCode='"&shipServiceArray(i)&"';"
			'response.write mySQL
			set rs=connTemp.execute(mySQL)
		next
		
		freeshipStrArray=split(USPS_freeshipStr,",")
		for i=0 to (ubound(freeshipStrArray)-1)
			freeoveramt=split(freeshipStrArray(i),"|")
			mySQL="UPDATE shipService SET serviceFreeOverAmt="&freeoveramt(1)&" WHERE serviceCode='"&freeoveramt(0)&"';"
			'response.write mySQL
			set rs=connTemp.execute(mySQL)
		next
		
		handlingStrArray=split(USPS_handlingStr,",")
		for i=0 to (ubound(handlingStrArray)-1)
			shiphandamt=split(handlingStrArray(i),"|")
			mySQL="UPDATE shipService SET serviceHandlingFee="&shiphandamt(1)&", serviceShowHandlingFee="&shiphandamt(2)&" WHERE serviceCode='"&shiphandamt(0)&"';"
			'response.write mySQL
			set rs=connTemp.execute(mySQL)
		next
		set rs=nothing
		call closeDb()
		response.redirect "../includes/PageCreateUSPSConstants.asp"
		%>
<!--#include file="AdminFooter.asp"-->