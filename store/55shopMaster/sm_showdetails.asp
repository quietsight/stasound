<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials"
pageIcon="pcv4_icon_salesManager.png"
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<% 
dim query, conntemp, rsTemp, pIdProduct

function ShowDateTimeFrmt(datestring)
Dim tmp1,tmp2
	tmp1=split(datestring," ")
	if scDateFrmt="DD/MM/YY" then
		tmp2=day(tmp1(0))&"/"&month(tmp1(0))&"/"&year(tmp1(0))
	else
		tmp2=month(tmp1(0))&"/"&day(tmp1(0))&"/"&year(tmp1(0))
	end if
	if instr(datestring," ") then
		tmp2=tmp2 & " " & tmp1(1) & tmp1(2)
	end if
	ShowDateTimeFrmt=tmp2
end function

pcSCID=getUserInput(request.QueryString("id"),0)
if not validNum(pcSCID) then
   response.redirect "menu.asp"
end if

call opendb()

	query="SELECT pcSC_Status,pcSC_StartedDate,pcSC_StoppedDate,pcSC_BUTotal,pcSC_SaveName,pcSC_SaveDesc,pcSC_SaveTech,pcSC_Archived,pcSales_ID FROM pcSales_Completed WHERE pcSC_ID=" & pcSCID & ";"
	set rs=connTemp.execute(query)
	if rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "sm_sales.asp"
	else
		tmpSaleStatus=rs("pcSC_Status")
		tmpStartedDate=rs("pcSC_StartedDate")
		tmpStoppedDate=rs("pcSC_StoppedDate")
		tmpUPrds=rs("pcSC_BUTotal")
		tmpSaleName=rs("pcSC_SaveName")
		tmpSaleDesc=rs("pcSC_SaveDesc")
		tmpSaleDetails=rs("pcSC_SaveTech")
		tmpArchived=rs("pcSC_Archived")
		if tmpArchived="" then
			tmpArchived="0"
		end if
		pcSaleID=rs("pcSales_ID")
		set rs=nothing
%> 
<html>
<head>
<title>Sale Details</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
<script>
	function WinResize()
	{
	var showScroll=0;
		if (/Firefox[\/\s](\d+\.\d+)/.test(navigator.userAgent)){
			wH=document.body.scrollHeight+100;
			wW=document.body.scrollWidth+20;
		}
			else
		{
			wH=document.body.scrollHeight+80;
			wW=document.body.scrollWidth+20;
		}
	if (wH>550)
	{
		showScroll=1;
		wH=550;
	}
	if (wW>650)
	{
		showScroll=1;
		wW=650;
	}
	
	window.resizeTo(wW,wH);
	if (showScroll==1) document.body.scroll="yes";
		
	}
</script>
<style>
table.pcCPcontent tr td {
	border-bottom: 1px dashed #CCC;
	padding-bottom: 5px;
}
</style>
</head>
<body style="margin: 0;" onload="javascript:WinResize()">

<table class="pcCPcontent">
	<tr>
		<th colspan="2">Details on this Sale</td>
	</tr>
	<tr valign="top">
		<td width="25%">Sale Name:</td>
		<td width="75%"><b><%=tmpSaleName%></b></td>
	</tr>
	<tr valign="top">
		<td nowrap>Sale Description:</td>
		<td><%=tmpSaleDesc%></td>
	</tr>
	<tr valign="top">
		<td>Sale Status:</td>
		<td><b>
			<%Select Case tmpSaleStatus
			Case "1": response.write "Started"
			Case "2": response.write "Live"
			Case "3": response.write "Stopped"
			Case "4": response.write "Completed"
			Case Else: response.write "N/A"
			End Select%>
			</b>
		</td>
	</tr>
	<tr valign="top">
		<td>Products:</td>
		<td>
			<%if tmpArchived="1" then%>
				<i>NOTE: this sales has been archived. The system no longer has detailed information on the products that were included in the sale</i>
			<%else%>
				<a href="sm_addedit_S1.asp?c=new&a=rev&id=<%=pcSaleID%>&b=9&scid=<%=pcSCID%>">Click here</a> to view list products included in this sale
			<%end if%>
		</td>
	</tr>
	<tr valign="top">
		<td nowrap>Sale Details:</td>
		<td><%=tmpSaleDetails%></td>
	</tr>
	<tr>
    	<td>Products:</td>
		<td>The sale affected <b><%=tmpUPrds%></b> product(s).</td>
	</tr>
	<tr>
    	<td>Start Date:</td>
		<td><%=ShowDateTimeFrmt(tmpStartedDate)%></td>
	</tr>
	<%if tmpStoppedDate<>"" AND (not IsNull(tmpStoppedDate)) then%>
	<tr>
    	<td>End Date:</td>
		<td><%=ShowDateTimeFrmt(tmpStoppedDate)%></td>
	</tr>
	<%end if%>
</table>
<div style="width: 100%; text-align: right;"><span style="padding: 15px;"><input type="image" src="../pc/images/close.gif" onClick="self.close()" alt="Close Window"></span></div>
</body>
</html>
<%end if
call closeDb()
%>