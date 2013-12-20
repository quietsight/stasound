<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'=====================================================================================
'= Easy Connect - Cardinal Commerce (http://www.cardinalcommerce.com)
'=
'= Usage
'=		Form used to POST the payer authentication request to the Card Issuer Servers.
'=		The Card Issuer Servers will in turn display the payer authentication window
'=		to the consumer within this location.
'=
'=		Note that the form field names below are case sensitive. For additional information
'=		please see the merchant integration documentation.
'=
'=====================================================================================
%>
<%response.Expires=-1%>
<%response.Buffer=true%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->

<% dim query, conntemp, rs
call opendb()
query="SELECT pcPay_Cent_TransactionURL,pcPay_Cent_ProcessorId,pcPay_Cent_MerchantID, pcPay_Cent_Active FROM pcPay_Centinel WHERE pcPay_Cent_ID=1;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	pcPay_Cent_TransactionURL=rs("pcPay_Cent_TransactionURL")
	pcPay_Cent_ProcessorId=rs("pcPay_Cent_ProcessorId")
	pcPay_Cent_MerchantID=rs("pcPay_Cent_MerchantID")
	pcPay_Cent_Active=rs("pcPay_Cent_Active")
set rs=nothing
if scSSL="0" OR scSSL="" then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/pcPay_Cent_Ecverifier.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
else
	tempURL=replace(( scSslURL&"/"&scPcFolder&"/pc/pcPay_Cent_Ecverifier.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
end if
pcPay_Cent_TermURL = tempURL

%>
<HTML>
<HEAD>
<SCRIPT Language="Javascript">
<!--
	function onLoadHandler(){
		document.frmLaunchACS.submit();
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
				<center>
				<FORM name="frmLaunchACS" method="Post" action="<%=Session("Centinel_ACSURL")%>">
				<noscript>
					<br><br>
					<center>
					<h1>Processing your Payer Authentication Transaction</h1>
					<h2>JavaScript is currently disabled or is not supported by your browser.</h2>
					<h3>Please click Submit to continue	the processing of your transaction.</h3>
					<input type="submit" value="Submit">
					</center>
				</noscript>
				<input type=hidden name="PaReq" value="<%=Session("Centinel_PAYLOAD")%>">
				<input type=hidden name="TermUrl" value="<%=Cstr(pcPay_Cent_TermURL)%>">
				<input type=hidden name="MD" value="Session Cookies Used">
				</FORM>
				</center>
			</td>
		</tr>
	</table>
</div>
</BODY>
</HTML>