<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/validation.asp"--> 
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp" -->
<% On Error Resume Next
Dim query,rs,connTemp
call openDb()

if session("idCustomer")=0 OR session("idCustomer")="" then
	response.write "SECURITY"
	response.End
end if

'// Select the email
query="SELECT [email] FROM Customers WHERE idcustomer=" & session("idCustomer") 
set rs=connTemp.execute(query)
if rs.eof then
	response.write "SECURITY"
	response.End
else
	pCustomerEmail=rs("email")
end if
set rs=nothing

'// Select Name and ID from Target
if Session("CustomerGuest")="2" then
	query="SELECT idCustomer, [name], lastname FROM Customers WHERE idcustomer=" & session("idCustomer")
else
	query="SELECT idCustomer, [name], lastname FROM Customers WHERE [email]='" & pCustomerEmail & "' AND pcCust_Guest=0;"
end if
set rs=connTemp.execute(query)
if rs.eof then
	response.write "SECURITY"
	response.End
else
	pcv_idtarget=rs("idCustomer")
	pCustomerName=rs("name") & " " & rs("lastname")
end if
set rs=nothing

tmpConStr=""

'// Get a Unique Key
TestedCustKey=0
do while (TestedCustKey=0)
	tmpConStr=generateCode(100)
	query="SELECT idCustomer FROM Customers WHERE pcCust_ConsolidateStr like '" & tmpConStr & "';"
	set rs=connTemp.execute(query)
	if rs.eof then
		TestedCustKey=1
	end if
	set rs=nothing
loop

'// Save Key with Target Customer
query="UPDATE Customers SET pcCust_ConsolidateStr='" & tmpConStr & "' WHERE idcustomer=" & pcv_idtarget
set rs=connTemp.execute(query)
set rs=nothing

'// Create Email
SPath1=Request.ServerVariables("PATH_INFO")
mycount1=0
do while mycount1<2
	if mid(SPath1,len(SPath1),1)="/" then
		mycount1=mycount1+1
	end if
	if mycount1<2 then
		SPath1=mid(SPath1,1,len(SPath1)-1)
	end if
loop
SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1

if Right(SPathInfo,1)="/" then
	pfileURL=SPathInfo & "pc/CustConsolidateConfirm.asp?e=" & pCustomerEmail & "&c=" & tmpConStr					
else
	pfileURL=SPathInfo & "/pc/CustConsolidateConfirm.asp?e=" & pCustomerEmail & "&c=" & tmpConStr
end if

customerEmail=""
customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_opc_conmail_1") & pCustomerName & "," & VbCrLf
customerEmail=customerEmail & dictLanguage.Item(Session("language")&"_opc_conmail_2") & VbCrLf & VbCrLf
customerEmail=customerEmail & pfileURL & vbCrLf & vbCrLf
customerEmail=customerEmail & scCompanyName & vbCrLf & vbCrLf

pcv_strSubject = dictLanguage.Item(Session("language")&"_opc_conmail_3")
call sendmail (scCompanyName, scEmail, pCustomerEmail, pcv_strSubject, replace(customerEmail, "&quot;", chr(34)))

'pcErrMsg=customerEmail


call closedb()

if pcErrMsg<>"" then
	response.write pcErrMsg
else
	response.write "OK"
end if

Function generateCode(keyLength)
	Dim sDefaultChars
	Dim iCounter
	Dim sMyKeys
	Dim iPickedChar
	Dim iDefaultCharactersLength
	Dim ikeyLength
	
	sDefaultChars="ABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
	ikeyLength=keyLength
	iDefaultCharactersLength = Len(sDefaultChars) 
	Randomize
	For iCounter = 1 To ikeyLength
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1) 
		sMyKeys = sMyKeys & Mid(sDefaultChars,iPickedChar,1)
	Next 
	generateCode = sMyKeys
End Function
%>