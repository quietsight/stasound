<%@ LANGUAGE = VBScript.Encode %>
<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<% Response.Buffer = True %>
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/dimensionsformatinc.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/languages_ship.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/FedEXconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/FedEXWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="../includes/CPconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/pcProductOptionsCode.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/pcAffConstants.asp"-->
<%
Dim query, conntemp, rstemp, rs
Dim objFEDEXXmlDoc, objFedExStream, strFileName, GraphicXML
Dim fedex_postdata, objFedExClass, objOutputXMLDoc, srvFEDEXXmlHttp, FEDEX_result, FEDEX_URL, pcv_strErrorMsg, pcv_strAction

Dim pcv_strAccountNameWS, pcv_strMeterNumberWS, pcv_strCarrierCodeWS
Dim pcv_strMethodNameWS, pcv_strMethodReplyWS, fedex_postdataWS, objFedExWSClass, objOutputXMLDocWS, srvFEDEXWSXmlHttp, FEDEXWS_result, FEDEXWS_URL, pcv_strErrorMsgWS


Dim pcStrCustomerRefKey, f
Dim pcCartIndex
Dim pSubTotal
Dim pShipSubTotal
Dim pShipWeight
Dim intUniversalWeight
Dim pCartQuantity
Dim pCartShipQuantity
Dim pCartTotalWeight, paymentTotal
Dim pcv_strAffiliateID

Dim pcCartArray(100,45)

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

call openDb()
%>
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<!--#include file="../includes/dimensionsformatinc.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/USPSCountry.asp"-->
<!--#include file="pcPay_GoogleCheckout_Global.asp"-->
<!--#include file="pcPay_GoogleCheckout_Handler.asp"-->
<%
Dim xmlResponse
Dim xmlAcknowledgment
Dim biData
biData = Request.BinaryRead(Request.TotalBytes)
Dim nIndex
For nIndex = 1 to LenB(biData)
	xmlResponse = xmlResponse & Chr(AscB(MidB(biData,nIndex,1)))
Next
processXmlData(xmlResponse)

conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing
%>

