<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="ProductCart Online Help - Troubleshooting Utility" %>
<% Section="" %>
<%PmAdmin=0%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<%
pcPageName="pcTSUtility.asp"
Dim conntemp, query, rs, rs2, Obj
err.clear
err.number=0

if request.Form("submit")<>"" then %>
	<!--#include file="pcAdminRetrieveSettings.asp"-->
	<%
	pcStrXML = request("ChangeXMLParser")

	'/////////////////////////////////////////////////////
	'// Update database with new Settings
	'/////////////////////////////////////////////////////
	%>
	<!--#include file="pcAdminSaveSettings.asp"-->
	<% response.redirect pcPageName
end if
%>