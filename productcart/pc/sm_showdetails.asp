<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<% 
dim query, conntemp, rsTemp, pIdProduct

pcSCID=getUserInput(request.QueryString("id"),0)
if not validNum(pcSCID) then
   response.redirect "default.asp"
end if

call opendb()

	query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveDesc FROM pcSales_Completed WHERE pcSales_Completed.pcSC_ID=" & pcSCID & ";"
	set rsS=Server.CreateObject("ADODB.Recordset")
	set rsS=conntemp.execute(query)
					
			if not rsS.eof then
				pcSCID=rsS("pcSC_ID")
				pcSCName=rsS("pcSC_SaveName")
				pcSCDesc=rsS("pcSC_SaveDesc")
			end if
			set rsS=nothing
%> 
<html>
<head>
<title>Sale Details</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
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
</head>
<body style="margin: 0;" onload="javascript:WinResize()">
<div id="pcMain">
<table class="pcMainTable">
	<tr>
		<td>
		<h2><span style="padding-left: 5px;"><%=dictLanguage.Item(Session("language")&"_Sale_1") & pcSCName%></span></h2>
		</td>
	</tr>
	<tr> 
		<th style="padding-left: 10px;"><%=dictLanguage.Item(Session("language")&"_Sale_2")%></th>
	</tr>
	<tr>
		<td style="padding: 5px 10px;">
		<%=pcSCDesc%>
		</td>
	</tr>
	<tr> 
        <td align="right" style="padding: 10px;">
        <input type="image" src="images/close.gif" onClick="self.close()" alt="<%=dictLanguage.Item(Session("language")&"_AddressBook_5")%>">
    	</td>
	</tr>
</table>
</div>
</body>
</html>
<%
call closeDb()
%>