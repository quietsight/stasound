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
<title><%response.write dictLanguage.Item(Session("language")&"_catering_9")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body style="margin: 0;">
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td colspan="2">
			<h2><%response.write dictLanguage.Item(Session("language")&"_catering_9")%></h2>
			</td>
		</tr> 
		<tr>
			<td>
				<p><%response.write dictLanguage.Item(Session("language")&"_catering_10")%></font></p>
			</td>
		</tr>
		<tr>    
			<td><hr></td>
		</tr>      
		<%
		query="select * from ZipCodeValidation order by zipcode asc"
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=connTemp.execute(query)
		If rstemp.eof Then %>
			<tr>
				<td>
					<div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_catering_11")%></div>
				</td>
			</tr>
								
	<% Else 
			Dim strCol
			strCol="#E1E1E1"
			Do While NOT rstemp.EOF
			zipcode=rstemp("zipcode")
				If strCol <> "#FFFFFF" Then
					strCol="#FFFFFF"
				Else 
					strCol="#E1E1E1"
				End If %>
		 
			<tr>
				<td bgcolor="<%=strCol%>">
					<p><%=zipcode%></p>
				</td>
			</tr>
								
			<% rstemp.MoveNext
				 Loop
				 End If
			%>
	</table>
</div>
</body>
</html>
<% 
set rstemp = nothing
call closeDB()
%>