<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>

<% 
 
response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwTotalWeb.asp"
	
'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
': Open Connection to the DB
dim connTemp, rs 'DELETE FOR HARD CODED VARS
call openDb() 'DELETE FOR HARD CODED VARS
'==================================================================================

'======================================================================================
'// Retrieve ALL required gateway specific data from database or hard-code the variables
'// Alter the query string replacing "fields", "table" and "idField" to reference your 
'// new database table.
'//
'// If you are going to hard-code these variables, delete all lines that are
'// commented as 'DELETE FOR HARD CODED VARS' and then set variables starting
'// at line 101 below.
'======================================================================================
query="SELECT pcPay_TW_MerchantID,pcPay_TW_CurCode,pcPay_TW_TestMode FROM pcPay_TotalWeb Where pcPay_TW_ID=1;"
			
			
'ALTER :: DELETE FOR HARD CODED VARS
'======================================================================================
'// End custom query
'======================================================================================

': Create recordset and execute query
set rs=server.CreateObject("ADODB.RecordSet") 'DELETE FOR HARD CODED VARS
set rs=connTemp.execute(query) 'DELETE FOR HARD CODED VARS

': Capture any errors
if err.number<>0 then 'DELETE FOR HARD CODED VARS
	call LogErrorToDatabase() 'DELETE FOR HARD CODED VARS
	set rs=nothing 'DELETE FOR HARD CODED VARS
	call closedb() 'DELETE FOR HARD CODED VARS
	response.redirect "techErr.asp?err="&pcStrCustRefID 'DELETE FOR HARD CODED VARS
end if 'DELETE FOR HARD CODED VARS

'======================================================================================
'// Set gateway specific variables - These can be your "hard coded variables" or 
'// Variables retrieved from the database.
'======================================================================================
	pcPay_TW_MerchantID=rs("pcPay_TW_MerchantID")
	pcPay_TW_CurCode = rs("pcPay_TW_CurCode")	
	pcPay_TW_TestMode=rs("pcPay_TW_TestMode")
	pcPay_TW_MerchantID=enDeCrypt(pcPay_TW_MerchantID, scCrypPass)
'======================================================================================
'// End gateway specific variables
'======================================================================================

': Clear recordset and close db connection
set rs=nothing 'DELETE FOR HARD CODED VARS
call closedb() 'DELETE FOR HARD CODED VARS


strData = "CustomerID="&pcPay_TW_MerchantID&"&Notes="& request("idorder")
 
if pcPay_TW_TestMode = "1" Then 
    StrUrl="https://testsecure.totalwebsecure.com/paypage/confirm.asp"
Else
    strURL="https://secure.totalwebsecure.com/paypage/confirm.asp"
end if 

Set XmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	err.clear
	XmlHttp.open "POST", StrUrl & "", false
	XMLHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
	XmlHttp.send(strData)
	if err.number<>0 then
		response.write "ERROR: "&err.description
		response.end
	end if
	 StrResult = XmlHttp.responseText
	 
     if instr(StrResult,"SUCCESS") Then 
    		 response.Redirect "gwReturn.asp?s=true&gw=TotalWeb"
	  		 response.end 
	 Else
			 strDeclinedRedirect = "msgb.asp?message="&server.URLEncode("<b>Error&nbsp;</b>: The transaction was declined <br><br><a href="""&tempURL&"?psslurl=gwTotalWeb.asp&idCustomer="&session("idCustomer")&"&idOrder="&session("GWOrderId")&"""><img src="""&rslayout("back")&"""></a>")
			 response.redirect strDeclinedRedirect
			 response.end 	 
	 End if 
	%>