<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="UPS OnLine&reg; Tools Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<% response.Buffer=true %>
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr>
		<td class="normal">
		<% 
		UPS_AccessLicense=Session("ship_UPS_AccessLicense")
		UPS_userID=Session("ship_UPS_userID")
		UPS_Password=Session("ship_UPS_Password")
		
		UPS_ServiceStr=Session("ship_UPS_ServiceStr")
		UPS_freeshipStr=Session("ship_UPS_freeshipStr")
		UPS_handlingStr=Session("ship_UPS_handlingStr")
		
		Dim connTemp, query, rs
		call openDb()
		set rs=Server.CreateObject("ADODB.Recordset")
		query="UPDATE ShipmentTypes SET AccessLicense='"&UPS_AccessLicense&"', userID='"&UPS_userID&"', [password]='"&UPS_Password&"', active=-1 WHERE idShipment=3;"
		set rs=connTemp.execute(query)
		'clear all informatin out of shipService for UPS
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='01';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='02';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='03';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='07';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='08';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='11';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='12';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='13';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='14';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='54';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='59';"
		set rs=connTemp.execute(query)
		query="UPDATE shipService SET serviceActive=0, servicePriority=0, serviceFree=0, serviceFreeOverAmt=0 WHERE serviceCode='65';"
		set rs=connTemp.execute(query)
		Dim i
		shipServiceArray=split(UPS_ServiceStr,", ")
		for i=0 to ubound(shipServiceArray)
			query="UPDATE shipService SET serviceActive=-1 WHERE serviceCode='"&shipServiceArray(i)&"';"
			'response.write query
			set rs=connTemp.execute(query)
		next
		
		freeshipStrArray=split(UPS_freeshipStr,",")
		for i=0 to (ubound(freeshipStrArray)-1)
			freeoveramt=split(freeshipStrArray(i),"|")
			query="UPDATE shipService SET serviceFreeOverAmt="&freeoveramt(1)&" WHERE serviceCode='"&freeoveramt(0)&"';"
			'response.write query
			set rs=connTemp.execute(query)
		next
		
		handlingStrArray=split(UPS_handlingStr,",")
		for i=0 to (ubound(handlingStrArray)-1)
			shiphandamt=split(handlingStrArray(i),"|")
			query="UPDATE shipService SET serviceHandlingFee="&shiphandamt(1)&", serviceShowHandlingFee="&shiphandamt(2)&" WHERE serviceCode='"&shiphandamt(0)&"';"
			'response.write query
			set rs=connTemp.execute(query)
		next
		
			'/////////////////////////////////////////////////////
			'// Set Local Variables for Setting
			'/////////////////////////////////////////////////////

			pcStrPickupType = removeSQ(Session("pcAdminPickupType"))
			pcStrPackageType = removeSQ(Session("pcAdminPackageType"))
			pcStrPackageHeight = removeSQ(Session("pcAdminPackageHeight"))
			pcStrPackageWidth = removeSQ(Session("pcAdminPackageWidth"))
			pcStrPackageLength = removeSQ(Session("pcAdminPackageLength"))
			pcStrPackageDimUnit = removeSQ(Session("pcAdminPackageDimUnit"))
			pcStrShipperCompanyName = removeSQ(Session("pcAdminShipperCompanyName"))
			pcStrShipperAttentionName = removeSQ(Session("pcAdminShipperAttentionName"))
			pcStrShipperAddress1 = removeSQ(Session("pcAdminShipperAddress1"))
			pcStrShipperAddress2 = removeSQ(Session("pcAdminShipperAddress2"))
			pcStrShipperAddress3 = removeSQ(Session("pcAdminShipperAddress3"))
			pcStrShipperCity = removeSQ(Session("pcAdminShipperCity"))
			pcStrShipperState = removeSQ(Session("pcAdminShipperState"))
			pcStrShipperPostalCode = removeSQ(Session("pcAdminShipperPostalCode"))
			pcStrShipperCountryCode = removeSQ(Session("pcAdminShipperCountryCode"))
			pcStrShipperPhone = removeSQ(Session("pcAdminShipperPhone"))
			pcStrShipperFax = removeSQ(Session("pcAdminShipperFax"))
			pcCurInsuredValue =  UPS_INSUREDVALUE
			pcStrDynamicInsuredValue =  UPS_DYNAMICINSUREDVALUE
			pcStrUseNegotiatedRates = UPS_USENEGOTIATEDRATES 
			pcStrShipperNumber =  UPS_SHIPPERNUM
			%>
		<!--#include file="pcAdminSaveUPSConstants.asp"-->
		<% 
		set rs = nothing
		call closedb()
		response.redirect "1_Step6.asp" 
		%>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->