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
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->	
<% 
Session("ship_FEDEX_SERVICE")=""
Session("ship_FedEX_freeshipStr")=""
Session("ship_FEDEX_FEDEX_PACKAGE")=""
Session("ship_FEDEX_HEIGHT")=""
Session("ship_FEDEX_WIDTH")=""
Session("ship_FEDEX_LENGTH")=""
Session("ship_FedEX_HandlingFee")=""
%>
<div class="pcCPmessageSuccess">You have successfully configured FedEx. <a href="viewshippingoptions.asp">Add or View/Modify existing options</a>.</div>
<!--#include file="AdminFooter.asp"-->