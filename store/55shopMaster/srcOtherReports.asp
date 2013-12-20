<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true %>
<% 
pageTitle="Other Reports" 
pageIcon="pcv4_icon_sales.gif"
%>
<% Section="genRpts" %>
<%PmAdmin=10%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"--> 
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/javascripts/pcDateFunctions.js"-->
<script language="JavaScript">
<!--
	function Validate_Dates(theForm)
	{
	
		if (theForm.FromDate.value == "")
		{
			alert("Please enter From Date and try again.");
			theForm.FromDate.focus();
			return (false);
		}
		
		if (theForm.ToDate.value == "")
		{
			alert("Please enter To Date and try again.");
			theForm.ToDate.focus();
			return (false);
		}
		
		if (isDate(theForm.FromDate.value,theForm.DateFormat.value,"From Date")==false)
		{
			theForm.FromDate.focus()
			return false
		}
		
		if (isDate(theForm.ToDate.value,theForm.DateFormat.value,"To Date")==false)
		{
			theForm.ToDate.focus()
			return false
		}
		
		if (CompareDates(theForm.FromDate,theForm.ToDate,"From < To")==false)
		{
			alert("From Date should be less than To Date.")
			theForm.ToDate.focus()
			return false
		}
	return (true);
	}
//-->
</script>

<% 
dim strDateFormat
strDateFormat="mm/dd/yyyy"
if scDateFrmt="DD/MM/YY" then
	strDateFormat="dd/mm/yyyy"
end if
%>

<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>View Return Merchandise Authorization (RMA) Requests</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
        	Specify a date range to view all requests for merchandise returns submitted in that period (RMAs). <br /><u>Note</u>: You must enter both dates in the format <%=strDateFormat%>
            <br /><br />
			<form action="RMAReport.asp" name="rmarpt" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
			<% todayDate=Date() %>
			<% Dim varMonth, varDay, varYear
			varMonth=Month(Date)
			varDay=Day(Date)
			varYear=Year(Date) 
			dim dtInputStrStart, dtInputStr
			dtInputStrStart=(varMonth&"/01/"&varYear)
			if scDateFrmt="DD/MM/YY" then
				dtInputStrStart=("01/"&varMonth&"/"&varYear)
			end if
			dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
			if scDateFrmt="DD/MM/YY" then
				dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
			end if
			%>
			From: <input name="FromDate" type="text" class="textbox" value="<%=dtInputStrStart%>" size="10">
			To: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">
			<input class="textbox" type="hidden" size="10" value="<%=strDateFormat%>" name="DateFormat">
			<br /><br />
			Base on:&nbsp;
			<select name="basedon">
			<option value="1" selected>Ordered Date</option>
			<option value="2">Processed Date</option>
			<option value="3">Shipped Date</option>
			</select>
			<br /><br />
			RMA Status:&nbsp;
			<select name="RMAStatus">
			<option value="-1" selected>All</option>
			<option value="0">Requested</option>
			<option value="1">Approved</option>
			<option value="2">Denied</option>
			</select>
			<br /><br />
			<input type="submit" value="Search" name="submit" class="submit2">
			</form>
			</p>
		</td>
	</tr>
<tr>

<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Product Reviews Report</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
        	Specify a date range to view all product review notifications sent to customers in that period. <br /><u>Note</u>: You must enter both dates in the format <%=strDateFormat%>
            <br /><br />
			<form action="RRReport.asp" name="rrReport" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
			<% todayDate=Date() %>
			<%
			varMonth=Month(Date)
			varDay=Day(Date)
			varYear=Year(Date) 
			dtInputStrStart=(varMonth&"/01/"&varYear)
			if scDateFrmt="DD/MM/YY" then
				dtInputStrStart=("01/"&varMonth&"/"&varYear)
			end if
			dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
			if scDateFrmt="DD/MM/YY" then
				dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
			end if
			%>
			From: <input name="FromDate" type="text" class="textbox" value="<%=dtInputStrStart%>" size="10">
			To: <input class="textbox" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">
			<input class="textbox" type="hidden" size="10" value="<%=strDateFormat%>" name="DateFormat">
			<br /><br />
			<input type="submit" value="Search" name="submit" class="submit2">
			</form>
			</p>
		</td>
	</tr>
<tr>
</table>
<!--#include file="AdminFooter.asp"-->