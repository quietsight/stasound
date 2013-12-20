<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="UPS OnLine&reg; Tools Shipping Configuration" %>
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
Session("ship_UPS_AccessLicense")=""
Session("ship_UPS_userID")=""
Session("ship_UPS_Password")=""
Session("ship_UPS_ServiceStr")=""
Session("ship_UPS_freeshipStr")=""
Session("pcAdminPickupType")=""
Session("pcAdminPackageType")=""
Session("pcAdminClassificationType")=""
Session("pcAdminPackageHeight")=""
Session("pcAdminPackageWidth")=""
Session("pcAdminPackageLength")=""
Session("pcAdminPackageDimUnit")=""
Session("pcAdminShipperCompanyName")=""
Session("pcAdminShipperAttentionName")=""
Session("pcAdminShipperAddress1")=""
Session("pcAdminShipperAddress2")=""
Session("pcAdminShipperAddress3")=""
Session("pcAdminShipperCity")=""
Session("pcAdminShipperState")=""
Session("pcAdminShipperPostalCode")=""
Session("pcAdminShipperCountryCode")=""
Session("pcAdminShipperPhone")=""
Session("pcAdminShipperFax")=""
%>
<div class="pcCPmessageSuccess">UPS OnLine&reg; Tools Configuration is complete. <a href="viewshippingoptions.asp">View/Modify 
existing shipping options</a></div>
<!--#include file="AdminFooter.asp"-->