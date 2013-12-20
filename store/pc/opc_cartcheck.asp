<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "opc_cartCheck.asp"
' This page displays the items in the cart.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/bto_language.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"--> 
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include FILE="../includes/ErrorHandler.asp"-->
<!--#include FILE="../includes/pcProductOptionsCode.asp"-->  
<!--#include file="inc_PrdCatTip.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
<%
response.expires=-1
Response.Buffer = True
Dim query, conntemp, rs
'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************

'*****************************************************************************************************
' START: PAGE ON LOAD
'*****************************************************************************************************

'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************

if countCartRows(pcCartArray, pcCartIndex)=0 then
 	response.end
end if

call opendb()

Dim strCCSLCheck
strCCSLcheck = checkCartStockLevels(pcCartArray, pcCartIndex, aryBadItems)
response.clear
If Len(Trim(strCCSLCheck))=0 Then
   response.write "OK"
'else
'   response.write "<div class=""pcErrorMessage"">"
'   response.write dictLanguage.Item(Session("language")&"_alert_19") & strCCSLcheck
'   response.write "</div>"
End If

call closedb()
%>