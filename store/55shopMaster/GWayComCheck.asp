<%@Language="VBScript"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'*******************************************************************
'The following don't use a local component:
' 2Checkout
' Authorize.Net - SIM mode
' BluePay
' BofA
' CBN (ChecksByNet)
' CyberCash
' Concord EBiz
' EWay
' FastTransact
' ITransact
' viaKLIX
' LinkPoint Connect
' Moneris
' PayFlow Link
' PayPal
' PsiGate
' SecurePay
' TrustCommerce
' WorldPay
' Linkpoint API

'The following require an installed class:
' Authorize.Net - AIM mode requires: Microsoft.XMLHTTP
' CyberSource requires: CyberSourceWS.MerchantConfig, CyberSourceWS.Hashtable, CyberSourceWS.Client
' Echo requires: "Msxml2.serverXmlHttp"&scXML
' NetBill requires: "Msxml2.ServerXMLHTTP"&scXML
' PaymentTech requires: Paymentech.Transaction, "Msxml2.serverXmlHttp"&scXML
' or PayPal.Payments.Communication.PayflowNETAPI
' USAePay requires: USAePayXChargeCom2.XChargeCom2
' SkipJack requires: SJComAPI.TransactionInfo.1
%>

<%PmAdmin=1%>
<!-- #include file="adminv.asp" -->
<!-- #include file="../includes/settings.asp" -->

<%
'Simple array for component names, 2-D array for class names (multiple possible classes per component)
Dim strComponent(7)
Dim strClass(7,3)

'The component names
strComponent(0) = "CyberSource"
strComponent(1) = "USAePay"
strComponent(2) = "PayJunction"
strComponent(3) = "LinkPoint API"
strComponent(4) = "ACHDirect"
strComponent(5) = "FastCharge"
strComponent(6) = "ParaData"
strComponent(7) = "PSIGate"

'The component class names
strClass(0,0) = "CyberSourceWS.MerchantConfig"
strClass(0,1) = "CyberSourceWS.Hashtable"
strClass(0,2) = "CyberSourceWS.Client"
strClass(1,0) = "USAePayXChargeCom2.XChargeCom2"
strClass(2,0) = "WinHttp.WinHttpRequest.5"
strClass(2,1) = "WinHttp.WinHttpRequest.5.1"
strClass(2,3) = "1"
strClass(3,0) = "LpiCom_6_0.LPOrderPart"
strClass(3,1) = "LpiCom_6_0.LinkPointTxn"
strClass(4,0) = "SendPmt.clsSendPmt"
strClass(5,0) = "ATS.SecurePost"
strClass(6,0) = "Paygateway.EClient.1"
strClass(7,0) = "MyServer.PsiGate"

Function IsObjInstalled(intClassNum)
On Error Resume Next
'This function tests the classes for the component indicated by the passed-in number.  Uses elements in the strClasses array above, correlated with the component names in the strComponent array.  Returns a string with non-present class names if any are missing, otherwise returns a ZLS.

	'Increase this constant to reflect the ubound of the classes array if a component requires more classes.
	CONST CLASSBOUND = 2

	Dim objTest, j
	Dim strError
	
	'init
	strError = ""

	'Test up to possible number of classes for each component...
	for j = 0 to CLASSBOUND

		'depending on whether or not there's a class name present in the array.

		   If Not IsEmpty(strClass(intClassNum, j)) Then 		
				
				Set objTest = Server.CreateObject(strClass(intClassNum, j))						
				If Err.Number = 0 Then
					Set objTest = Nothing
					If Not isEmpty(strClass(intClassNum, 3)) Then
					strError = ""
					End if 
				Else
					'If the class test failed, create or append to an error string for reporting results.
					If IsObject(objTest) Then Set objTest = Nothing
					If strError = "" Then
						strError = strClass(intClassNum, j)
					Else
					    strError = strError & ",<br>" & strClass(intClassNum, j)
					End If
				End If
			Else
			   ' check to see if we multi mutliples components 3 
				If Not isEmpty(strClass(intClassNum, 3)) Then
				  errArray = split(strError, ",<br>", -1)
				  if ubound(errArray) = 1 then
					 strError = errArray(0)
					 strError = strError & "<BR> or " & errArray(1)
				  else
					 strError = ""
				  end if 
				End if  				
			
			End If	
	Next

	'Report any resulting errors.
	IsObjInstalled = strError

End Function
%>

<HTML>
<HEAD>
<TITLE>Payment Gateway Component Test</TITLE>
</HEAD>
<body bgcolor="#ffffff" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" marginwidth="0" marginheight="0" stylesrc="1bg_info.htm">
<table border="0" width="100%" cellspacing="0" cellpadding="3">
  <tr> 
    <td width="100%" bgcolor="#000080"><font size="2" color="mintcream"><b><font face="Arial, Helvetica, sans-serif">Payment Gateway Component Test</font></b></font></td>
  </tr>
</table>

<br>
<br>
<table border=0 cellspacing=1 cellpadding=4 align="center">
	<% Dim i, strErr
	
	For i=0 to UBound(strComponent) %>
		<tr> 
			<td bgcolor="#FFFFFF" valign="top" A0B0E0"" right""> 
			<div align="right"><font face="" verdana,="Verdana," arial,="Arial," helvetica="Helvetica" size="2" 2="2"><strong> 
			<font face="Arial, Helvetica, sans-serif"><%= strComponent(i)%>:</font>&nbsp;</strong></font></div>
			</td>
			<td bgcolor="#FFFFFF" valign="top" A0B0E0"" center""> <font size="2"> <font face="Arial, Helvetica, sans-serif">
			<% strErr = IsObjInstalled(i)
			If strErr <> "" Then %>
				</font></font> 
				<div align="left"><font size="2" face="Arial, Helvetica, sans-serif">Not Available:<br><%=strErr%>
			<% Else %>
				<strong>Available</strong> 
			<% End If %>
			</font></div>
			</td>
		</tr>
	<% Next %>
</table>
<p align="center"><a href=# onClick="self.close();"><font size=2 face="Verdana,Helvetica,Arial,sans-serif"><b>Close 
  Window</b></font></a> 
</BODY>
</HTML>