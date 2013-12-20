<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="XML Tools - General Settings" %>
<% section="layout" %>
<%PmAdmin=19%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
Dim rs, connTemp, query
call openDB()

IF request("action")="upd" THEN

	LogTurnOn=request("LogTurnOn")
	if LogTurnOn="" then
		LogTurnOn=0
	end if
	if Not IsNumeric(LogTurnOn) then
		LogTurnOn=0
	end if
	LogErrors=request("LogErrors")
	if LogErrors="" then
		LogErrors=0
	end if
	if Not IsNumeric(LogErrors) then
		LogErrors=0
	end if
	CaptRequest=request("CaptRequest")
	if CaptRequest="" then
		CaptRequest=0
	end if
	if Not IsNumeric(CaptRequest) then
		CaptRequest=0
	end if
	CaptResponse=request("CaptResponse")
	if CaptResponse="" then
		CaptResponse=0
	end if
	if Not IsNumeric(CaptResponse) then
		CaptResponse=0
	end if
	EnforceHTTPs=request("EnforceHTTPs")
	if EnforceHTTPs="" then
		EnforceHTTPs=0
	end if
	if Not IsNumeric(EnforceHTTPs) then
		EnforceHTTPs=0
	end if

	
	query="SELECT pcXMLSet_Log,pcXMLSet_LogErrors,pcXMLSet_CaptureRequest,pcXMLSet_CaptureResponse,pcXMLSet_EnforceHTTPs FROM pcXMLSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		query="UPDATE pcXMLSettings SET pcXMLSet_Log=" & LogTurnOn & ",pcXMLSet_LogErrors=" & LogErrors & ",pcXMLSet_CaptureRequest=" & CaptRequest & ",pcXMLSet_CaptureResponse=" & CaptResponse & ",pcXMLSet_EnforceHTTPs=" & EnforceHTTPs & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
	else
		query="INSERT INTO pcXMLSettings (pcXMLSet_Log,pcXMLSet_LogErrors,pcXMLSet_CaptureRequest,pcXMLSet_CaptureResponse,pcXMLSet_EnforceHTTPs) VALUES (" & LogTurnOn & "," & LogErrors & "," & CaptRequest & "," & CaptResponse & "," & EnforceHTTPs & ");"
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	
	msg="updated"
	
END IF	

%>
<!--#include file="AdminHeader.asp"-->
	
	<% If msg<>"" Then %>
			<div class="pcCPmessageSuccess">
			<%if msg="updated" then%>
				XML Settings have been updated successfully!
			<%end if%>
			</div>
	<% End If 
	
	
	LogTurnOn=0
	LogErrors=0
	CaptRequest=0
	CaptResponse=0
	EnforceHTTPs=0
	
	query="SELECT pcXMLSet_Log,pcXMLSet_LogErrors,pcXMLSet_CaptureRequest,pcXMLSet_CaptureResponse,pcXMLSet_EnforceHTTPs FROM pcXMLSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		LogTurnOn=rs("pcXMLSet_Log")
		LogErrors=rs("pcXMLSet_LogErrors")
		CaptRequest=rs("pcXMLSet_CaptureRequest")
		CaptResponse=rs("pcXMLSet_CaptureResponse")
		EnforceHTTPs=rs("pcXMLSet_EnforceHTTPs")
	end if
	set rs=nothing
	%>

			<form name="Form1" action="AdminXMLSettings.asp?action=upd" method="post" class="pcForms">
			<table class="pcCPcontent">
				<tr>
					<th colspan="3">Security Settings</th>
				</tr>
                <tr>
                    <td colspan="3" class="pcCPspacer"></td>
                </tr>
				<tr>
					<td width="5%" align="right"><input type="checkbox" name="EnforceHTTPs" value="1" <%if EnforceHTTPs="1" then%>checked<%end if%> class="clearBorder"></td>
					<td colspan="2">Enforce HTTPs (only transactions over the HTTPS protocol will be allowed) </td>
				</tr>
                <tr>
                    <td colspan="3" class="pcCPspacer"></td>
                </tr>
				<tr>
					<th colspan="3">Log Settings</th>
				</tr>
                <tr>
                    <td colspan="3" class="pcCPspacer"></td>
                </tr>
				<tr>
					<td colspan="3">Use this feature to capture all XML transactions between your store and third party applications.</td>
				</tr>
				<tr>
					<td width="5%" align="right"><input type="checkbox" name="LogTurnOn" value="1" <%if LogTurnOn="1" then%>checked<%end if%> onclick="if (this.checked==false) {document.Form1.LogErrors.checked=false;document.Form1.CaptRequest.checked=false;document.Form1.CaptResponse.checked=false;}" class="clearBorder"></td>
					<td colspan="2">Log XML Transactions</td>
				</tr>
				<tr>
					<td>&nbsp;</td><td width="5%" align="right"><input type="checkbox" name="LogErrors" value="1" <%if LogErrors="1" then%>checked<%end if%> onclick="if (this.checked) document.Form1.LogTurnOn.checked=true;" class="clearBorder"></td>
					<td align="left" width="90%">Include transactions that resulted in an error </td>
				</tr>
				<tr>
					<td>&nbsp;</td><td width="5%" align="right"><input type="checkbox" name="CaptRequest" value="1" <%if CaptRequest="1" then%>checked<%end if%> onclick="if (this.checked) document.Form1.LogTurnOn.checked=true;" class="clearBorder"></td>
					<td align="left" width="90%">Capture XML Requests</td>
				</tr>
				<tr>
					<td>&nbsp;</td><td width="5%" align="right"><input type="checkbox" name="CaptResponse" value="1" <%if CaptResponse="1" then%>checked<%end if%> onclick="if (this.checked) document.Form1.LogTurnOn.checked=true;" class="clearBorder"></td>
					<td align="left" width="90%">Capture XML Responses</td>
				</tr>
				<tr>
					<td colspan="3" class="pcSpacer">&nbsp;</td>
				</tr>
				<tr>
					<td>&nbsp;</td><td colspan="2"><input type="submit" name="Submit1" value="Update Settings" class="submit2">&nbsp;
					<input type="button" name="Back" value="XML Tools Manager" onclick="location='XMLToolsManager.asp';" class="ibtnGrey"></td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
</table>
<%call closedb()%><!--#include file="AdminFooter.asp"-->