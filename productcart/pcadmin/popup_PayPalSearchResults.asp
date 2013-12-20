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

if request("action")<>"src" then
	response.redirect "menu.asp"
end if

call opendb()

	PayPalFromDate=Request("startDate")
	PayPalToDate=Request("endDate")
	if PayPalFromDate="" then
		PayPalFromDate=Date()-6
	end if
	PayPalFromDate1=PayPalFromDate
	PayPalToDate1=PayPalToDate

	Yearno=Year(PayPalFromDate)
	monthno=Month(PayPalFromDate)
	dayno=Day(PayPalFromDate)
	PayPalFromDate=yearno &"-"& monthno &"-"& dayno & "T00:00:00Z"

	if PayPalToDate<>"" then
		PayPalToDate=CDate(PayPalToDate)+1
		Yearno=Year(PayPalToDate)
		monthno=Month(PayPalToDate)
		dayno=Day(PayPalToDate)
		PayPalToDate=yearno &"-"& monthno &"-"& dayno & "T00:00:00Z"
	end if
	TransactionID=Request("transactionID")
	
	set objPayPalClass = New pcPayPalClass
	
	'// Query our PayPal Settings
	objPayPalClass.pcs_SetAllVariables()

	'// Add the required NVP’s
	nvpstr="" '// clear
	if PayPalFromDate<>"" then
		objPayPalClass.AddNVP "STARTDATE", PayPalFromDate
	end if
	if PayPalToDate<>"" then
		objPayPalClass.AddNVP "ENDDATE", PayPalToDate
	end if
	if TRANSACTIONID<>"" then
		objPayPalClass.AddNVP "TRANSACTIONID", trim(TransactionID)
	end if
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
<title>PayPal - Transaction Search Results</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image: none;">
<table class="pcCPcontent" width="100%">
<tr>
	<td><div class="title">PayPal - Transaction Search Results</div></td>
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
	<td colspan="6">
	<% 
	CntOfTrxn=0

	For resindex = 0 To resArray.Count - 1 
		TrxnID="L_TRANSACTIONID"&resindex
		If Not resArray(TrxnID) = "" Then
			CntOfTrxn= CntOfTrxn +1
		End if
	Next
	%>
	<%if PayPalFromDate1<>"" then%>From Date: <%=PayPalFromDate%>&nbsp;<%end if%><%if PayPalToDate1<>"" then%>To Date: <%=PayPalToDate1%>&nbsp;<%end if%><br><br>
	Found <%=CntOfTrxn%> results
	</td>
</tr>
<%if CntOfTrxn>0 then%>
<tr>
	<th>&nbsp;</th>
	<th nowrap>Transaction ID</th>
    <th nowrap>Date/Time</th>
    <th nowrap>Status</th>
	<th nowrap>Payer Name</th>
    <th nowrap>Amount</th>
</tr>
<% 
IndexOfTrxn=0
reskey = resArray.Keys
resitem = resArray.items
For resindex = 0 To resArray.Count - 1 %>
<%TrxnID="L_TRANSACTIONID"&resindex
If Not resArray(TrxnID) = "" Then
IndexOfTrxn= IndexOfTrxn+1%>
<tr>
	<td>
		<%=IndexOfTrxn%>
	</td>
	<td nowrap>
		<a href="popup_TransactionDetails.asp?transactionID=<%=resArray(TrxnID)%>"><%=resArray(TrxnID)%></a>
	</td>
	<td nowrap>
		<%Timestamp="L_TIMESTAMP"&resindex
		If Not resArray(Timestamp) = "" Then
			Timestamp=split(resArray(Timestamp),"T")
			Timestamp1=split(Timestamp(0),"-")
			Timestamp2=replace(Timestamp(1),"Z"," ")%>
			<%=Timestamp1(1) & "/" & Timestamp1(2) & "/" & Timestamp1(0) & " " & Timestamp2%>
		<%End if%>
	</td>
	<td  nowrap>
		<%Status="L_STATUS"&resindex%>
		<%=resArray(Status)%>
	</td>
	<td nowrap>
		<%Name="L_NAME"&resindex%>
		<%=resArray(Name)%>
	</td>
	<td nowrap>
		<%Amt="L_AMT"&resindex%>
		<%=scCurSign & money(resArray(Amt))%>
	</td>
</tr>
<%End if
next
End if %>
</table>
<%end if%>
<table class="pcCPcontent" width="100%">
<tr>
	<td class="pcSpacer">&nbsp;</td>
</tr>
<tr>
	<td>
		<input type="button" name="back" value=" New Search " onClick="location='popup_PayPalTransSearch.asp';" class="ibtnGrey">
		&nbsp;
		<input type="button" name="close" value=" Close window " onClick="javascript:window.close();" class="ibtnGrey">
	</td>
</tr>
</table>
</body>
</html>