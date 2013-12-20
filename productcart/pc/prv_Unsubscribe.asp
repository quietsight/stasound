<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
' PRV41 start
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<% 

Dim pIDOrder, pIDCustomer, pCustGuest, pcv_UniqueID

' Check if the store is on. If store is turned off display store message
If scStoreOff="1" then
	response.redirect "msg.asp?message=83"
End If

Dim connTemp,rs
call opendb()
%>
<!--#include file="prv_getsettings.asp"-->
<%

pcv_UniqueID=GetUserInput(request("UID"),0)
if len(pcv_UniqueID)<>36 then
	call closedb()
	response.redirect "msg.asp?message=210"
end If

On Error goto 0

query="SELECT pcRN_idCustomer FROM pcReviewNotifications WHERE pcRN_UniqueID='" & pcv_UniqueID & "'"
set rs=connTemp.execute(query)

if rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "msg.asp?message=210"
end If

pIdCustomer = rs("pcRN_idCustomer")
'session("idCustomer") = pIdCustomer

query = "SELECT name, pcCust_Guest FROM customers WHERE idCustomer=" & pIDCustomer
Set rs = connTemp.execute(query)
If rs.eof Then
	set rs=nothing
	call closedb()
	response.redirect "msg.asp?message=210" ' Give the generic message to discourage script-kiddies
else
   pCustName = rs("name")
   pCustGuest = CLng(rs("pcCust_Guest"))
End If
rs.close

%>
<!--#include file="header.asp"-->

<div id="pcMain">
	<table class="pcMainTable">   
	<tr>
		<td>
		   <p>
		   <%
		   connTemp.execute "UPDATE Customers SET pcCust_AllowReviewEmails=0 WHERE idCustomer=" & pIdCustomer
		   response.write dictLanguage.Item(Session("language")&"_prv_33")
		   %>
		</td>
    </tr>
	</table>

</div>
<% call closedb() %>
<!--#include file="footer.asp"-->
<% 'PRV41 end %>