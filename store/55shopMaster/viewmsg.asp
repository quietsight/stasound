<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<%
dim rstemp, conntemp, query
call opendb()
%>
<%'Start SDBA
pcv_pageType=request("pagetype")
'End SDBA%>
<% pageTitle="Sent Message Details" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->

<table class="pcCPcontent">
	<tr>
	<td align="center">
			<%query="select idnews, FromDate, fromname, fromemail, title, msgbody, msgtype, CustTotal, CustFile FROM News where idNews=" & request("idnews")
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=connTemp.execute(query)
			
			if not rstemp.eof then
				intIdNews=rstemp("idnews")
				dtFromDate=rstemp("FromDate")
				strFromName=rstemp("fromname")
				strFromEmail=rstemp("fromemail")
				strTitle=rstemp("title")
				strMsgBody=rstemp("msgbody")
				intMsgType=rstemp("msgtype")
				strCustTotal=rstemp("CustTotal")
				strCustFile=rstemp("CustFile")
				
				%>
				<table class="pcCPcontent">
					<tr>
						<td width="106" valign="top" height="26">Message ID#:</td>
						<td width="348" valign="top" height="26"><%=intIdNews%></td>
					</tr>
					<tr>
						<td width="106" valign="top" height="26">Sent Date:</td>
						<td width="348" valign="top" height="26"><%=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)%></td>
					</tr>  
					<tr>
						<td width="106" valign="top" height="26">From Name:</td>
						<td width="348" valign="top" height="26"><%=strFromName%></td>
					</tr>
					<tr>
						<td width="106" valign="top" height="26">From Email:</td>
						<td width="348" valign="top" height="26"><%=strFromEmail%></td>
					</tr>
					<tr>
						<td width="106" valign="top" height="26">Subject:</td>
						<td width="348" valign="top" height="26"><%=strTitle%></td>
					</tr>
					<tr>
						<td width="106" valign="top" height="26">Message:</td>
						<td width="348" valign="top" height="26">
						<textarea rows="10" cols="80" name="details"><%=strMsgBody%></textarea></td>
					</tr>
					<tr>
						<td width="106" valign="top" height="26">Message Type:</td>
						<td width="348" valign="top" height="26"><%if intMsgType<>"1" then%>Plain Text<%else%>HTML<%end if%></td>
					</tr>
					<tr>
						<td width="106" valign="top" height="26">Numbers of recipients:</td>
						<td width="348" valign="top" height="26"><%=strCustTotal%></td>
					</tr>  
					<tr>
						<td width="106" valign="top" height="26">List of recipients:</td>
						<td width="348" valign="top" height="26">
						<%
						findit = Server.MapPath("newslists/" & strCustFile)
						Set fso = server.CreateObject("Scripting.FileSystemObject")
						Set f = fso.OpenTextFile(findit, 1)
						Do While not f.AtEndofStream
							TempStr=TempStr & f.Readline & vbcrlf
						Loop
						f.close
						Set fso = nothing
						Set f = nothing 
						%>
						<textarea rows="10" cols="60" name="recipientslist"><%=TempStr%></textarea></td>
					</tr>
				</table>
			<%end if
			set rstemp=nothing
			call closedb() 
			%>
	</td>
</tr>
<tr>
	<td align="center">&nbsp;</td>
</tr>
<tr>
	<td align="center">
		<input type="button" name=buttonnew value="Send to New List" onClick="location='germsg.asp?pagetype=<%=pcv_pageType%>&idnews=<%=intIdNews%>'" class="ibtnGrey">&nbsp;
		<input type="button" name=back value="Back" onClick="location='managenews.asp?pagetype=<%=pcv_pageType%>'" class="ibtnGrey">
	</td>
</tr>
<tr>
	<td align="center">&nbsp;</td>
</tr>
</table>

<!--#include file="AdminFooter.asp"-->