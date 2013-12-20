<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="XML Tools - Transaction Details" %>
<% section="layout" %>
<%PmAdmin=19%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
pidPartner=trim(request("idPartner"))
tmpIDLog=trim(request("idxml"))

If Not IsNumeric(pidPartner) then
	response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query.")
End If

If Not IsNumeric(tmpIDLog) then
	response.redirect "techErr.asp?error="&Server.URLEncode("An error occurred when submitting your query.")
End If

Dim conntemp, query, rs

%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<%call opendb()

	query="SELECT pcXL_RequestKey,pcXL_Status,pcXL_RequestXML,pcXL_ResponseXML FROM pcXMLLogs WHERE pcXP_ID=" & pidPartner & " AND pcXL_ID=" & tmpIDLog & ";"
	set rs=connTemp.execute(query)
	If rs.eof Then
		set rs=nothing%>
	<tr> 
		<td align="center" colspan="2">
			<div class="pcCPmessage">
				XML Transaction not found!
			</div>
		</td>
	</tr>
	<%Else
		xt_Key=rs("pcXL_RequestKey")
		xt_Status=rs("pcXL_Status")
		if xt_Status<>"" then
		else
			xt_Status=0
		end if
		xt_RequestXML=trim(rs("pcXL_RequestXML"))
		xt_ResponseXML=trim(rs("pcXL_ResponseXML"))
		set rs=nothing
	%>
	<tr>
		<td>Transaction Key:</td>
		<td><%=xt_Key%></td>
	</tr>
	<tr>
		<td>Status:</td>
		<td><%Select Case xt_Status
				Case 0:%><b>Errors</b>
				<%Case 1: %>Successful
				<%Case 2: %>Successful, with some errors
			<%End Select%>
		</td>
	</tr>
	<%if xt_RequestXML<>"" then%>
	<tr>
		<th colspan="2">XML Request</th>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<textarea cols="70" rows="13"><%=xt_RequestXML%></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer">&nbsp;</td>
	</tr>
	<%end if%>
	<%if xt_ResponseXML<>"" then%>
	<tr>
		<th colspan="2">XML Response</th>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<textarea cols="70" rows="13"><%=xt_ResponseXML%></textarea>
		</td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2" class="pcSpacer">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2"><input type="button" name="Back" value="Back to XML Logs" onclick="location='viewXMLPartnerLogs.asp?idPartner=<%=pidPartner%>';" class="ibtnGrey"></td>
	</tr>
	</table>
	<%End if%>
<%call closedb()%><!--#include file="AdminFooter.asp"-->