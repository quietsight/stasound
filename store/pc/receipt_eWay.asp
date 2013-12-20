<%@ Language=VBScript %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="header.asp"-->
<%
result = Request("ewayTrxnStatus")
trxnReference  = Request("ewayTrxnReference") 
transaction_number = Request("ewayTrxnNumber")
if transaction_number = "" then
	transaction_number  = "NOT DEFINED"
end if
option1 = Request("ewayoption1")
option2 = Request("ewayoption2") 
option3 = Request("ewayoption3")
if UCASE(result)="TRUE" then
	session("eWayOrdNum")=option1
	session("eWaytrxnReference")=trxnReference
	session("eWaytransaction_number")=transaction_number
	session("eWayoption2")=option2
	Response.redirect "gwReturn.asp?s=true&gw=eWay"
	RESPONSE.END
else
	response.write "Error: "&Request("ewayTrxnError ")
	response.write "<br><br><br><a href=""javascript: history.back(-1)""><img src="""&rslayout("back")&"""></a>"
end if %>
<!--#include file="footer.asp"-->