<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/pcPayPalClass.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<%
Dim connTemp,query

pcv_strAdminPrefix="1"

TransactionID=Request("transactionID")
if TransactionID="" then
	response.redirect "menu.asp"
end if

call opendb()
	
	set objPayPalClass = New pcPayPalClass
	
	'// Query our PayPal Settings
	objPayPalClass.pcs_SetAllVariables()

	'// Add the required NVP’s
	nvpstr="" '// clear
	objPayPalClass.AddNVP "TRANSACTIONID", trim(TransactionID)
	
	'// Post to PayPal by calling .hash_call
	Set resArray = objPayPalClass.hash_call("gettransactionDetails",nvpstr)

	'// Set Response
	Set Session("nvpResArray")=resArray

	'// Check for success
	ack = UCase(resArray("ACK"))
	
	'// Check for code errors
	if err.number <> 0 then 
		'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
		Dim pcv_PayPalErrMessage
		%><!--#include file="../includes/pcPayPalErrors.asp"--><%                                             
	end if

	If ack="SUCCESS" Then
	else              
		'// append the user friendly errors to API errors
		objPayPalClass.GenerateErrorReport()
	end if
call closedb()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>PayPal - Transaction Details</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image:none;">
<table class="pcCPcontent" width="100%">
<tr>
	<td><div class="title">PayPal - Transaction Details</div></td>
</tr>
<tr>
	<td class="pcSpacer">&nbsp;</td>
</tr>
</table>
<%if pcv_PayPalErrMessage<>"" then%>
<table class="pcCPcontent" width="100%">
	<tr>
		<td>
			<div class="pcCPmessage">
				PayPal Gateway Transaction Error<br>
				<%=replace(replace(pcv_PayPalErrMessage,"</div>",""),"<div align=""left"">","")%>
			</div>
		</td>
	</tr>
</table>
<%else%>
<table class="pcCPcontent" width="100%">
<tr>
	<td>Payer ID:</td>
	<td><%=resArray("PAYERID")%></td>
</tr>
<tr>
	<td>Payer Name:</td>
	<td><%=resArray("FIRSTNAME")%>&nbsp;<%=resArray("LASTNAME")%></td>
</tr>
<tr>
	<td>Payer Status:</td>
	<td><%=resArray("PAYERSTATUS")%></td>
</tr>
<tr>
	<td>Transaction ID:</td>
	<td><%=resArray("TRANSACTIONID")%></td>
</tr>
<tr>
	<td>Amount:</td>
	<td><%=resArray("AMT")%></td>
</tr>
<%
	set objPayPalClass=nothing
	Set resArray=nothing
	Set Session("nvpResArray")=nothing
	
	set objPayPalClass = New pcPayPalClass
	
	'// Query our PayPal Settings
	objPayPalClass.pcs_SetAllVariables()

	'// Add the required NVP’s
	nvpstr="" '// clear
	objPayPalClass.AddNVP "STARTDATE", "2005-01-01T00:00:00Z"
	objPayPalClass.AddNVP "TRANSACTIONID", TransactionID
	objPayPalClass.AddNVP "TRXTYPE", "Q"
	
	'// Post to PayPal by calling .hash_call
	Set resArray = objPayPalClass.hash_call("TransactionSearch",nvpstr)

	'// Set Response
	Set Session("nvpResArray")=resArray

	'// Check for success
	ack = UCase(resArray("ACK"))
	
	'// Check for code errors
	if err.number <> 0 then 
		'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
		pcv_PayPalErrMessage=""
		%><!--#include file="../includes/pcPayPalErrors.asp"--><%                                             
	end if

	If ack="SUCCESS" Then
		For resindex = 0 To resArray.Count - 1
			TrxnID="L_TRANSACTIONID"&resindex
			If resArray(TrxnID) = TransactionID AND resArray("L_STATUS"&resindex)<>"..2e" Then
				PaymentStatus=resArray("L_STATUS"&resindex)
				exit for
			End if
		Next%>
	<tr>
	<td>Date/Time:</td>
	<td>
		<%Timestamp="L_TIMESTAMP"
		If Not resArray(Timestamp&resindex) = "" Then
			Timestamp=split(resArray(Timestamp&resindex),"T")
			Timestamp1=split(Timestamp(0),"-")
			Timestamp2=replace(Timestamp(1),"Z"," ")%>
			<%=Timestamp1(1) & "/" & Timestamp1(2) & "/" & Timestamp1(0) & " " & Timestamp2%>
		<%End if%>
	</td>
</tr>
	<tr>
		<td>Payment Status:</td>
		<td>
			<b><%=PaymentStatus%></b>
		</td>
	</tr>
	<%End if%>
</table>
<%end if%>
<table class="pcCPcontent" width="100%">
<tr>
	<td class="pcSpacer">&nbsp;</td>
</tr>
<tr>
	<td>
		<input type="button" name="back" value=" Back " onClick="javascript:history.back();" class="ibtnGrey">
		&nbsp;
		<input type="button" name="back1" value=" New Search " onClick="location='popup_PayPalTransSearch.asp';" class="ibtnGrey">
		&nbsp;
		<input type="button" name="close" value=" Close window " onClick="javascript:window.close();" class="ibtnGrey">
	</td>
</tr>
</table>
</body>
</html>