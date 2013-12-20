<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Return Authorization Request" %>
<% section="mngRma"%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
pcIntOrderID=request.querystring("idOrder")
%>
<table class="pcCPcontent">
	<tr>
		<td>
        The RMA number was successfully created.
        <br /><br />
        RMA Number: <strong><%=session("pRmaNumber")%></strong>
        </td>
	</tr>
	<tr>
		<td><hr></td>
	</tr>
    <%
	if validNum(pcIntOrderID) then
	%>
	<tr> 
		<td><a href="Orddetails.asp?id=<%=pcIntOrderID%>">Back to Order Details</a></td>
	</tr>
	<%
	end if
	%>
</table>
<!--#include file="AdminFooter.asp"-->