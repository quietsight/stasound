<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true %>
<% pageTitle="Sales Reports" %>
<% Section="genRpts" %>
<%PmAdmin=10%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"--> 
<% 
dim f, query, conntemp, rstemp, counter

counter=0

call openDb()

' count statistic registers
viewyear=clng(request("year"))

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>ProductCart shopping cart software - Control Panel - Sales Summary</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="ProductCart asp shopping cart software is published by NetSource Commerce. ProductCart's Control Panel allows you to manage every aspect of your ecommerce store. For more information and for technical support, please visit NetSource Commerce at http://www.earlyimpact.com">
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image: none;">

<!--#include file="pcCharts.asp"-->

	<table class="pcCPcontent" style="width:100%;">
	<tr>
		<th>Quick Summary: Monthly Sales for <%=viewyear%></th>
	</tr>
	<tr>
	<tr>
		<td>
			<div id="chartMonthlySales" style="height:250px; "></div>
			<%Dim pcv_YearTotal
			pcv_YearTotal=0
			call pcs_MonthlySalesChart("chartMonthlySales",viewyear,1,0)%>
		</td>
	<tr>
		<td  class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td align="left" class="pcCPsectionTitle">Year Total: <%=scCurSign & money(pcv_YearTotal)%></td>
	</tr>
	<tr>
		<td align="right">
			<form class="pcForms">
				<input type="button" name="close" value="Close" onClick="window.close()">
			</form>
		</td>
	</tr>
	</table>
<%
	set rs = nothing
	set rstemp = nothing
	call closeDb()
%>
</body>
</html>