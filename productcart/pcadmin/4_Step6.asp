<% response.Buffer=true %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<% Session("ship_CP_Server")=""
		Session("ship_CP_ID")=""
		Session("ship_CP_Password")=""
		Session("ship_CP_Service")=""
		Session("ship_CP_freeshipStr")=""
		Session("ship_CP_EMPackage")=""
		Session("ship_CP_PMPackage")=""
		Session("ship_CP_Height")=""
		Session("ship_CP_Width")=""
		Session("ship_CP_Length")=""
		Session("ship_CP_HandlingFee")=""
		%>
<div class="pcCPmessageSuccess">Canada Post Configuration is complete. <a href="viewshippingoptions.asp">Add or View/Modify existing options</a></div>
<!--#include file="AdminFooter.asp"-->