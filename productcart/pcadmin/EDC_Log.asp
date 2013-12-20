<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Endicia Postage Label Services - View Transaction Details" %>
<% response.Buffer=true %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/EndiciaFunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="AdminHeader.asp"-->
<%Dim connTemp,rs,query

private const MaxRecords=25 'Max Records to display per page

pcPageName="EDC_Log.asp"

TransID=request("id")

If TransID="" OR (Not IsNumeric(TransID)) then
	response.redirect "EDC_manage.asp"
End if

iPageCurrent=request("iPageCurrent")

call opendb()


query="SELECT pcELog_Request,pcELog_Response FROM pcEDCLogs WHERE pcET_ID=" & TransID & ";"
set rs=connTemp.execute(query)
If rs.eof then
	msg="Transaction Log not found!"
	msgType=0
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<%Else%>
	<table class="pcCPcontent">
	<tr>
		<td><font size="3"><b>Transaction # <%=TransID%></b></font></td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Request XML</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<textarea cols="85" rows="15"><%=rs("pcELog_Request")%></textarea>
		</td>
	</tr>
	<tr>
	<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Response XML</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<textarea cols="85" rows="15"><%=rs("pcELog_Response")%></textarea>
		</td>
	</tr>
	<tr>
	<td class="pcCPspacer"></td>
	</tr>
</table>
<%End if
set rs=nothing%>
<table class="pcCPcontent">
<tr>
	<td>&nbsp;</td>
	<td><input type="button" name="back" value=" Back to Endicia Transactions " onclick="javascript:location='EDC_Trans.asp?iPageCurrent=<%=iPageCurrent%>';" class="submit2"></td>
</tr>
</table>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->