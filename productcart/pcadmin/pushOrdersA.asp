<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="Consolidate Customer Accounts" 
pageIcon="pcv4_icon_people.png"
section="mngAcc" 
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="AdminHeader.asp"-->

<%' Get customers ID
pidCustomer = request("idcustomer")
if not validNum(pidCustomer) then response.Redirect("viewCusta.asp")

dim connTemp,query
call opendb()
query="SELECT email,[name],lastname FROM customers WHERE idcustomer=" & pidCustomer
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if not rs.eof then
	pcv_Email = rs("email")
	pcv_Name = rs("name") & " " & rs("lastname")
end if
set rs=nothing
call closedb()

%>
	
<form method="post" name="listCust" action="pushOrdersB.asp" class="pcForms">
	<table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
        <tr>
        	<th colspan="2">Move orders from <strong><%=pcv_Name%></strong> to another customer</th>
        </tr>
        <tr>
        	<td colspan="2" class="pcCPspacer"></td>
        </tr>
		<tr>
			<td colspan="2">
					Customer E-mail: <b><%=pcv_Email%></b>&nbsp;|&nbsp;Customer name: <b><%=pcv_Name%></b>
			</td>
		</tr>
		<tr>
			<td colspan="2">You can remove all orders from this account and reassign them to a different customer account so that they are consolidated under this second account. To proceed, first locate the customer that you wish to reassign the orders to.</td>
		</tr>
		<tr> 
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td width="31%" align="right">Search by <strong>name</strong>, <strong>company</strong>, or <strong>e-mail</strong>:</td>
			<td width="69%">  
			<input type="text" name="customerName" value="" size="30">
			<input type="hidden" name="idcustomer" value="<%=pidCustomer%>">
			</td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
        	<td></td>
			<td>
			<input type="submit" name="srcView" value="Search" class="submit2">&nbsp;
			<input type="button" name="Button" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->