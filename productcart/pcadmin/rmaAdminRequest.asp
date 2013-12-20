<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Modify RMA Information" %>
<% section="mngRma"%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/SQLFormat.txt" -->
<%
dim conntemp, rstemp, query

call openDb()

pRmaReturnReason=replace(request.form("rmaReturnReason"),"'","''")			
pIdOrder=replace(request.form("idOrder"),"'","''")
pIdProduct=replace(request.form("idProduct"),"'","''")
pRmaApproved=request.form("rmaApproved")
pSendEmail=request.Form("sendEmail")

function makePassword(byVal maxLen)
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize

	For intCounter = 1 To maxLen
		whatsNext = Int((1 - 0 + 1) * Rnd + 0)
		If whatsNext = 0 Then
			'character
			upper = 90
			lower = 65
		Else
			upper = 57
			lower = 48
		End If
		strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	makePassword = strNewPass
end function

pRmaNumber = request("pRmaNumber")
if pRmaNumber<>"" then
	pRmaNumber=replace(pRmaNumber,"'","''")
else
	pRmaNumber = makePassword(16)
end if
pRMADate=now()
if SQL_Format = "1" then pRMADate = day(pRMADate) & "/" & month(pRMADate) & "/" & year(pRMADate)
if scDB="SQL" then
	query="INSERT INTO PCReturns (rmaNumber,rmaReturnReason,rmaDateTime,idOrder,rmaIdProducts,rmaApproved) VALUES ('" &pRmaNumber& "','" &pRmaReturnReason& "','"&pRMADate&"',"&pIdOrder&",'"&pIdProduct&"',"&pRmaApproved&")"
else
	query="INSERT INTO PCReturns (rmaNumber,rmaReturnReason,rmaDateTime,idOrder,rmaIdProducts,rmaApproved) VALUES ('" &pRmaNumber& "','" &pRmaReturnReason& "',#"&pRMADate&"#,"&pIdOrder&",'"&pIdProduct&"',"&pRmaApproved&")"
end if

set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query) 

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
	response.redirect "/"&scPcFolder&"/pc/techErr.asp?error="& Server.Urlencode("Error in rmaAdminRequest.asp: "&Err.Description) 
end if

'send out email
if pSendEmail="1" then
	query="SELECT orders.idOrder, orders.idCustomer, customers.name, customers.lastName, customers.email FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pIdOrder&"));"
	set rsTemp=conntemp.execute(query)
	pFirstName=rsTemp("name")
	pLastName=rsTemp("lastName")
	pRcpt=rsTemp("email")
	MsgTitle= dictLanguage.Item(Session("language")&"_sendMail_rma_1") & scCompanyName
	MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_2") & VBCRLF
	'If the RMA has been approved:
	MsgBody=MsgBody & "" & VBCRLF
	MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_3") & VBCRLF
	MsgBody=MsgBody & "" & VBCRLF
	MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_4") & pRmaNumber & VBCRLF
	MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_5") & pRMADate & VBCRLF
	MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_sendMail_rma_6") & pRmaReturnReason & VBCRLF
	MsgBody=MsgBody & "" & VBCRLF
	
	call sendmail (scCompanyName,scFrmEmail,pRcpt,MsgTitle,MsgBody)
end if

set rstemp=nothing
call closeDb()
session("pRmaNumber")=pRmaNumber
response.redirect "rmaAdminThankyou.asp?idOrder="&pIdOrder
%>