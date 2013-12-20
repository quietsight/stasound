<%@ LANGUAGE="VBSCRIPT" %>
<% 'OPTION EXPLICIT %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "CustwlRmv.asp"
' This page removes wishlist items when requested by Custquotesview.asp
'
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="CustLIv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/currencyformatinc.asp" --> 
<!--#include file="../includes/rc4.asp" --> 
<!--#include FILE="../includes/ErrorHandler.asp"--> 
<%
Response.Buffer = True

Dim conntemp, query, rstemp, pcv_strIdQuote  


'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Page On-Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

call openDb()

pIdCustomer=session("idcustomer")
pIdProduct=server.HTMLEncode(request.querystring("idProduct"))
pcv_strIdQuote=server.HTMLEncode(request.querystring("IdQuote"))

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Page On-Load
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' check if that item exists
query="SELECT idproduct FROM wishlist WHERE idcustomer=" &pIdCustomer& " AND IdQuote=" &pcv_strIdQuote
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)
if err.number<>0 then
'//Logs error to the database
call LogErrorToDatabase()
'//clear any objects
set rstemp=nothing
'//close any connections
call closedb()
'//redirect to error page
response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then
	set rstemp=nothing
	call closeDb()
	response.redirect "msg.asp?message=37"
end if
set rstemp=nothing

query="DELETE FROM wishlist WHERE idcustomer=" &pIdCustomer& " AND IdQuote=" &pcv_strIdQuote
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	'//Logs error to the database
	call LogErrorToDatabase()
	'//clear any objects
	set rstemp=nothing
	'//close any connections
	call closedb()
	'//redirect to error page
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in CustwlRmv: "&Err.Description) 
end if


set rstemp=nothing
call closeDb()

response.redirect "Custquotesview.asp?msg="&server.urlencode(dictLanguage.Item(Session("language")&"_CustwlRmv_2"))
%>