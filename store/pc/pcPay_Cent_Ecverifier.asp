<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%@ LANGUAGE="VBSCRIPT" %>
<% 'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'=====================================================================================
'= Cardinal Commerce (http://www.cardinalcommerce.com)
'= pcPay_Cent_Ecverifier.asp
'= Purpose
'=		This page represents the pcPay_Cent_TermURL passed on the pcPay_Cent_Ecauth.asp page. 
'=		The Card Issuerwill post the results of the authentication to this page. This page will 
'=		Pass the PARes to the MAPS for validation of the PARes and will return the XID, CAVV,  
'=		ECI, Authentication Status and Signature values. 
'=
'=		Checking these values will determine what the next step in the flow should be. If the 
'=		authentication is successful then the CAVV, ECI, and XID values should be passed
'=		on the authorization message. If the authentication was unsuccessful, or resulted
'=		in an error the consumer should be prompted for another form of payment.
'=====================================================================================
%>
<%response.Expires=-1%>
<%response.Buffer=true%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<% dim connTemp, rs
call openDB()
 
Set rs=connTemp.execute("Select * From layout Where layout.ID=2")
strBack=rs("back")
%>
<!-- #Include File="pcPay_Cent_XMLFunctions.asp"-->
<% call closedb()

Dim pares, merchantData, redirectPage 
Dim oDOM, oErrorNo, oErrorDesc, oSignature, oAuthStatus, oCavv, oEci, oXid
Dim xmlRequest, xmlResponse

'=====================================================================================
' Retrieve the PaRes and MD values from the Card Issuer's Form POST to this Term URL page.
' If you like, the MD data passed to the Card Issuer could contain the TransactionId
' that would enable you to reestablish the transaction session. This would be the 
' alternative to using the Client Session Cookies
'=====================================================================================

pares = request.Form("PaRes")
merchantData = request.Form("MD")

dim tempURL
If scSSL="0" or scSSL="" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If

If (pares <> "") Then
	xmlRequest = "<CardinalMPI>"
	xmlRequest = xmlRequest & generateXMLTag("Version", Cstr(pcPay_Cent_MessageVersion))
	xmlRequest = xmlRequest & generateXMLTag("MsgType", "cmpi_authenticate")
	xmlRequest = xmlRequest & generateXMLTag("ProcessorId", Cstr(pcPay_Cent_ProcessorId))
	xmlRequest = xmlRequest & generateXMLTag("TransactionType", "C")
	xmlRequest = xmlRequest & generateXMLTag("MerchantId", Cstr(pcPay_Cent_MerchantId))
	xmlRequest = xmlRequest & generateXMLTag("TransactionId", Session("Centinel_TransactionId"))
	xmlRequest = xmlRequest & generateXMLTag("TransactionPwd", pcPay_Cent_TransactionPwd)
	xmlRequest = xmlRequest & generateXMLTag("PAResPayload", Cstr(pares))
	
	xmlRequest = xmlRequest & "</CardinalMPI>"
	
	xmlResponse = sendMessage(xmlRequest, "gwUSAePay.asp")

	'=====================================================================================
	' Retrieve the Elements from the XML Response Message
	'=====================================================================================

	Set oDOM = CreateObject("Msxml2.DOMDocument"&scXML)
	oDOM.async = False 
	oDOM.LoadXML xmlResponse

	Set oErrorNo = oDOM.documentElement.selectSingleNode("ErrorNo")
	Set oErrorDesc = oDOM.documentElement.selectSingleNode("ErrorDesc")
	Set oSignature = oDOM.documentElement.selectSingleNode("SignatureVerification")
	Set oAuthStatus = oDOM.documentElement.selectSingleNode("PAResStatus")
	Set oCavv = oDOM.documentElement.selectSingleNode("Cavv")
	Set oEci = oDOM.documentElement.selectSingleNode("EciFlag")
	Set oXid = oDOM.documentElement.selectSingleNode("Xid")

	Session("Centinel_PAResStatus") = oAuthStatus.text
	Session("Centinel_SignatureVerification") = oSignature.text
	Session("Centinel_ErrorNo") = oErrorNo.text
	Session("Centinel_ErrorDesc") = oErrorDesc.text
	Session("Centinel_XID") = oXid.text
	Session("Centinel_CAVV") = oCavv.text
	Session("Centinel_ECI") = oEci.text
	
	intDebug=0
	if intDebug=1 then
		response.write "error: "&Session("Centinel_ErrorNo")&"<BR>"
		response.write "PAResStatus: "&Session("Centinel_PAResStatus")&"<BR>"
		response.write "Signature: "&Session("Centinel_SignatureVerification")&"<BR>"
		response.end
	end if
	
		' Handle Payer Authentication Specific Logic
		If Session("Centinel_ErrorNo") = "0" AND Session("Centinel_SignatureVerification") = "N" Then
			redirectPage = "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>:&nbsp;Your Credit Card Issuer has indicated that the transaction was not able to be authenticated. To protect against unauthorized use, this card cannot be used to complete your purchase. Please provide another form of payment to complete your purchase.<br><br><a href="""&tempURL&"?psslurl=gwAuthorizeAIM.asp&idCustomer="&session("reqCustomerID")&"&idOrder="&session("reqOrderID")&"&ordertotal="&session("AIMOrderTotal")&"&billingFirstName="&session("reqBillFirstName")&"&billingLastName="&session("reqBillLastName")&"&billingAddress="&session("reqBillAddress")&"&billingZip="&session("reqBillZip")&"&billingEmail="&session("reqBillEmail")&"&idDbSession="&pIdDbSession&"&randomKey="&pRandomKey&"&billingCompany="&session("reqBillCompany")&"&password=&billingPhone="&session("reqBillPhone")&"&billingStateCode="&session("reqBillState")&"&billingState="&session("reqBillState")&"&billingCity="&session("reqBillCity")&"&billingCountryCode="&session("reqBillCountry")&"&shippingFullName=" &session("reqShipFirstName")&" "&session("reqShipLastName")&"&shippingCompany=" &x_ship_to_company & "&shippingAddress=" &session("reqShipAddress")& "&shippingcity="&session("reqShipCity")& "&shippingState="&session("reqShipState")&"&shippingStateCode=" &session("reqShipState")& "&shippingZip=" &session("reqShipZip")& "&shippingCountryCode="&session("reqShipCountry")&"""><img src="""&strBack&"""></a>")

		ElseIf Session("Centinel_ErrorNo") = "0" AND Session("Centinel_SignatureVerification") = "Y" AND Session("Centinel_PAResStatus") = "Y" Then
			redirectPage=session("redirectPage")&"?centinel=Y"
		
		ElseIf Session("Centinel_ErrorNo") = "0" AND Session("Centinel_SignatureVerification") = "Y" AND Session("Centinel_PAResStatus") = "A" Then
		
			redirectPage=session("redirectPage")&"?centinel=Y"
		
		ElseIf Session("Centinel_ErrorNo") = "0" AND Session("Centinel_SignatureVerification") = "Y" AND Session("Centinel_PAResStatus") = "U" Then
		
			redirectPage=session("redirectPage")&"?centinel=Y"
		
		ElseIf Session("Centinel_ErrorNo") = "0" AND Session("Centinel_SignatureVerification") = "Y" AND Session("Centinel_PAResStatus") = "N" Then
			redirectPage = "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>:&nbsp;Your Credit Card Issuer has indicated that the transaction was not able to be authenticated. To protect against unauthorized use, this card cannot be used to complete your purchase. Please provide another form of payment to complete your purchase.<br><br><a href="""&tempURL&"?psslurl=gwAuthorizeAIM.asp&idCustomer="&session("reqCustomerID")&"&idOrder="&session("reqOrderID")&"&ordertotal="&session("AIMOrderTotal")&"&billingFirstName="&session("reqBillFirstName")&"&billingLastName="&session("reqBillLastName")&"&billingAddress="&session("reqBillAddress")&"&billingZip="&session("reqBillZip")&"&billingEmail="&session("reqBillEmail")&"&idDbSession="&pIdDbSession&"&randomKey="&pRandomKey&"&billingCompany="&session("reqBillCompany")&"&password=&billingPhone="&session("reqBillPhone")&"&billingStateCode="&session("reqBillState")&"&billingState="&session("reqBillState")&"&billingCity="&session("reqBillCity")&"&billingCountryCode="&session("reqBillCountry")&"&shippingFullName=" &session("reqShipFirstName")&" "&session("reqShipLastName")&"&shippingCompany=" &x_ship_to_company & "&shippingAddress=" &session("reqShipAddress")& "&shippingcity="&session("reqShipCity")& "&shippingState="&session("reqShipState")&"&shippingStateCode=" &session("reqShipState")& "&shippingZip=" &session("reqShipZip")& "&shippingCountryCode="&session("reqShipCountry")&"""><img src="""&strBack&"""></a>")
		Else
			redirectPage = "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>:</font>&nbsp;Your Credit Card transaction was not able to be authenticated. Please provide another form of payment.<br><br><a href="""&tempURL&"?psslurl=gwAuthorizeAIM.asp&idCustomer="&session("reqCustomerID")&"&idOrder="&session("reqOrderID")&"&ordertotal="&session("AIMOrderTotal")&"&billingFirstName="&session("reqBillFirstName")&"&billingLastName="&session("reqBillLastName")&"&billingAddress="&session("reqBillAddress")&"&billingZip="&session("reqBillZip")&"&billingEmail="&session("reqBillEmail")&"&idDbSession="&pIdDbSession&"&randomKey="&pRandomKey&"&billingCompany="&session("reqBillCompany")&"&password=&billingPhone="&session("reqBillPhone")&"&billingStateCode="&session("reqBillState")&"&billingState="&session("reqBillState")&"&billingCity="&session("reqBillCity")&"&billingCountryCode="&session("reqBillCountry")&"&shippingFullName=" &session("reqShipFirstName")&" "&session("reqShipLastName")&"&shippingCompany=" &x_ship_to_company & "&shippingAddress=" &session("reqShipAddress")& "&shippingcity="&session("reqShipCity")& "&shippingState="&session("reqShipState")&"&shippingStateCode=" &session("reqShipState")& "&shippingZip=" &session("reqShipZip")& "&shippingCountryCode="&session("reqShipCountry")&"""><img src="""&strBack&"""></a>")
		End If
End If
%>
<HTML>
<HEAD>
<SCRIPT Language="Javascript">
<!--
	function onLoadHandler(){
		document.frmResultsPage.submit();
	}
//-->
</Script>
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</HEAD>
<body onLoad="onLoadHandler();">
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<FORM name="frmResultsPage" target="_top" method="Post" action="<%=redirectPage%>">
				<noscript>
					<center>
					<h1>Processing Your Transaction</h1>
					<h2>JavaScript is currently disabled or is not supported by your browser.<br></h2>
					<h3>Please click Submit to continue	the processing of your transaction.</h3>
					<input type="submit" value="Submit">
					</center>
				</noscript>
				</FORM>
			</td>
		</tr>
	</table>
</div>
</BODY>
</HTML>