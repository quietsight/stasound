<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard - Print Label" %>
<% Section="mngAcc" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<%
Const iPageSize=5

Dim iPageCurrent, conntemp, rs, varFlagIncomplete, uery, strORD, pcv_intOrderID



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// SET PAGE NAMES
pcPageName = "FedExWS_ManageShipmentsPrinting.asp"
ErrPageName = "FedExWS_ManageShipmentsPrinting.asp"

'// OPEN DATABASE
call openDb()

'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExClass

'// TABLE SIZE
pTableHeight=((96/200)*1900) '// 11 - (.075*2) = 9.5 inches
pTableWidth=((96/200)*1400) '// 8.5 - (.075*2) = 7 inches

'// SPACER SIZE
pVertical=((96/200)*900) '//  4.5 inches
pHorizontal=((96/200)*1400) '// 8.5 - (.075*2) = 7 inches

'// LABEL SIZE
iVertical=((96/200)*950) '//  4.75 inches
iHorizontal=((96/200)*1400) '// 8.5 - (.075*2) = 7 inches

'// GET PAGE NUMBER
if request.querystring("path")="" then
	pCurrentPath=""
else
	pCurrentPath=Request.QueryString("path")
end if
%>
<table height="<%=pTableHeight%>" width="<%=pTableWidth%>" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td valign="top" align="center"><img src="<%=pCurrentPath%>" height="<%=iVertical%>" width="<%=iHorizontal%>" border="0" /></td>
  </tr>
  <tr>
	<td valign="top" align="center"><img src="images/spacer.gif" height="<%=pVertical%>" width="<%=pHorizontal%>" border="0" /></td>
  </tr>
</table>
<%
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>