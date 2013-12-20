<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<%
Dim rs, connTemp, mySQL

call openDB()
%>
<html>
<head>
<title><%response.write dictLanguage.Item(Session("language")&"_catering_15")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body style="margin: 0;">
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td colspan="2">
				<h2><%response.write dictLanguage.Item(Session("language")&"_catering_15")%></h2>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="padding: 4px;">
				<%response.write dictLanguage.Item(Session("language")&"_catering_16")%>
			</td>
		</tr>
		<tr>    
			<td colspan="2" class="pcSpacer"></td>
		</tr>   
		<tr>                 	
			<th nowrap width="20%">
				<%response.write dictLanguage.Item(Session("language")&"_catering_17")%>
			</th>
			<th nowrap width="60%">
				<%response.write dictLanguage.Item(Session("language")&"_catering_18")%>
			</th>
		</tr>  			  
		<%
		mySQL="select * from Blackout order by Blackout_Date asc"
		set rstemp=connTemp.execute(mySQL)
		If rstemp.eof Then
		%>
		<tr>          
			<td colspan="2" style="padding: 4px;">
				<%response.write dictLanguage.Item(Session("language")&"_catering_19")%>
			</td>
		</tr>
              
		<%
			Else 
			Dim strCol
			strCol="#E1E1E1"
			Do While NOT rstemp.EOF
			Blackout_Date=rstemp("Blackout_Date")
			Blackout_Message=rstemp("Blackout_Message")
				If strCol <> "#FFFFFF" Then
					strCol="#FFFFFF"
				Else 
					strCol="#E1E1E1"
				End If
		%>          
		<tr> 
			<td bgcolor="<%= strCol %>" style="padding: 4px;"><%=Blackout_Date%></td>
			<td bgcolor="<%= strCol %>" style="padding: 4px;"><%=Blackout_Message%></td>
		</tr>
              
		<%
			rstemp.MoveNext
			Loop
			End If
		%>         
		<tr> 
			<td colspan="2" class="pcSpacer"></td>
		</tr>
	</table>
</div>
</body>
</html>
<% 
set rstemp = nothing
call closeDB()
%>