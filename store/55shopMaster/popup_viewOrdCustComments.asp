<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<%Dim connTemp,rs,query
	idOrder=request("idorder")
	if idOrder="" then
		idOrder=0
	end if
	if (not IsNumeric(idOrder)) or (idOrder="0") then
		response.redirect "menu.asp"
	end if
%>
<html>
<head>
<title>Order #<%=int(idOrder)+scpre%> Comments</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:450px; background-image: none;">
<%call opendb()
	query="SELECT comments,admincomments FROM Orders WHERE idorder=" & idOrder & ";"
	set rs=connTemp.execute(query)
	pcv_comments=""
	pcv_admcomments=""
	if not rs.eof then
		pcv_comments=trim(rs("comments"))
		pcv_admcomments=trim(rs("admincomments"))
	end if
	set rs=nothing
call closedb()%>
<form name="form1" method="post" class="pcForms">
<table class="pcCPcontent">
	<%IF pcv_comments<>"" THEN%>
  	<tr>
    	<th>When placing order #<%=int(idOrder)+scpre%>, the customer wrote:</th>
	</tr>
  	<tr>
    	<td>
			<%=pcv_comments%>
		</td>
	</tr>
	<tr>
    	<td class="pcSpacer">&nbsp;</td>
	</tr>
	<%END IF%>
	<%IF pcv_admcomments<>"" THEN%>
  	<tr>
    	<th>Admin comments</th>
	</tr>
  	<tr>
    	<td>
			<%=pcv_admcomments%>
		</td>
	</tr>
	<tr>
    	<td class="pcSpacer">&nbsp;</td>
	</tr>
	<%END IF%>
	<tr>
		<td>
			<div align="right">
				<br>
				<br>
				<br>
				<input type="button" name="Back" value="Close Window" onClick="self.close();">
			</div>
		</td>
	</tr>
</table>
</form>
</div>
</body>
</html>