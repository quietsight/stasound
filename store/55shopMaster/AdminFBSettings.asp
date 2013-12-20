<meta http-equiv="Content-Language" content="en-us">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 
Dim pageTitle, Section
pageTitle="Manage Help Desk"
pageIcon="pcv4_icon_helpDesk.png"
Section="layout" 
%>
<!-- #Include File="Adminheader.asp" -->
<table class="pcCPcontent">
<tr>
	<td class="pcCPspacer"></td>
</tr>
<tr>
	<th>Help Desk Settings</th>
</tr>
<tr>
	<td class="pcCPspacer"></td>
</tr>
<tr>
	<td>
		<p>When customers view information about previous orders, they have the ability to contact you if the Help Desk is turned on. Here you can configure settings that apply to the way Help Desk messages are posted and viewed.</p>
		<ul class="pcListIcon">
            <li><a href="AdminSettings.asp?tab=4">Turn Help Desk On/Off</a></li>
            <li><a href="AdminFBTypeManager.asp">Message Type Settings</a></li>
            <li><a href="AdminFBStatusManager.asp">Message Status Settings</a></li>
            <li><a href="AdminFBPriorityManager.asp">Message Priority Settings</a></li>
		</ul>
	</td>
</tr>
<tr>
	<td class="pcCPspacer"></td>
</tr>
<tr>
	<th>View Postings</th>
</tr>
<tr>
	<td class="pcCPspacer"></td>
</tr>
<tr>
	<td>
		<p>Use the form below to <strong>view messages related to all orders</strong>. Or you can view postings specific to an order from the order details page. To do so, <a href="invoicing.asp">locate an order</a>, select &quot;View &amp; Process&quot;, and click on &quot;View/Post Messages and Files&quot;.</p>
		
			<%
			todayDate=Date()
			Dim varMonth, varDay, varYear
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
		
		<form method="post" name="filter" action="adminviewallmsgs.asp" class="pcForms">
			<input type=hidden name="Type" value="1">
			<p style="padding-top:10px;">Date From: <input type="text" name="FromDate" size="10" value="<%=dtInputStrStart%>"> To: <input type="text" name="ToDate" size="10" value="<%=dtInputStr%>"></p>
			<p style="padding-top:10px;">Show <input type="text" name="FBPerpage" size="5" value="25"> postings per page</p>
			<p style="padding-top:10px;"><input type="submit" name="submit" value="View Messages" class="submit2">&nbsp;
			<input type="button" value="Write New Message" onClick="location.href='adminaddfeedback.asp?type=1'"></p>
		</form>
		</td>
	</tr>
</table>
<!-- #Include File="Adminfooter.asp" -->