<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=0%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp" -->
<%Dim connTemp,rs,query%>
<html>
<head>
<title>MailUp - Customers are waiting for synchronization</title>
<link href="pcControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:450px;">
<%call opendb()
	intCount=-1
	query="SELECT DISTINCT customers.idcustomer,customers.name,customers.lastname,customers.customerCompany,customers.email FROM customers INNER JOIN pcMailUpSubs ON customers.idcustomer=pcMailUpSubs.idcustomer WHERE pcMailUpSubs.pcMailUpSubs_SyncNeeded<>0 ORDER BY customers.name ASC,customers.lastname ASC;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcArr=rs.getRows()
		intCount=ubound(pcArr,2)
	end if
	set rs=nothing
call closedb()%>
<form name="form1" method="post" class="pcForms">
<table class="pcCPcontent">
	<%IF intCount="-1" THEN%>
  	<tr>
    	<td colspan="4" >
			<div class="pcCPmessage">
				No customers found!
			</div>
		</td>
	</tr>
	<%ELSE%>
	<tr>
		<th nowrap>ID#</th>
		<th nowrap>Customer Name</th>
		<th nowrap>Company</th>
		<th nowrap>E-mail</th>
	</tr>		
	<%For i=0 to intCount%>
  	<tr>
    	<td nowrap><a href="modCusta.asp?idcustomer=<%=pcArr(0,i)%>" target="_blank"><%=pcArr(0,i)%></a></td>
		<td nowrap><a href="modCusta.asp?idcustomer=<%=pcArr(0,i)%>" target="_blank"><%=pcArr(1,i) & " " & pcArr(2,i)%></a></td>
		<td nowrap><a href="modCusta.asp?idcustomer=<%=pcArr(0,i)%>" target="_blank"><%=pcArr(3,i)%></a></td>
		<td nowrap><a href="mailto:<%=pcArr(4,i)%>"><%=pcArr(4,i)%></a></td>
	</tr>
	<%Next%>
	<%END IF%>
	<tr>
		<td>
			<div align="right">
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