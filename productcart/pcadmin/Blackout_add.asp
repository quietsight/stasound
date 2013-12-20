<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Add New Blackout Date"
pageIcon="pcv4_icon_calendar.png"
section="layout"
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<%
Dim rs, connTemp, query
if request("action")="add" then
	call openDB()
	Blackout_Date=getUserInput(request("Blackout_Date"),0)
	Blackout_Message=getUserInput(request("Blackout_Message"),1400)
	if scDateFrmt = "DD/MM/YY" AND SQL_Format="0" then
		Blackout_DateArry=split(Blackout_Date,"/")
		Blackout_Date=Blackout_DateArry(1)&"/"&Blackout_DateArry(0)&"/"&Blackout_DateArry(2)
	end if
	query="select * from Blackout where Blackout_Date="
	if scDB="SQL" then
		query=query&"'" & Blackout_Date  & "'"
	else
		query=query&"#" & Blackout_Date  & "#"
	end if
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			call closeDb()
			response.redirect "Blackout_add.asp?r=1&msg=This Blackout Date is already in use"
		end if

	query="insert into Blackout (Blackout_Date,Blackout_Message) values ("
	if scDB="SQL" then
		query=query&"'" & Blackout_Date  & "'"
	else
		query=query&"#" & Blackout_Date  & "#"
	end if
	query = query & ",'" & Blackout_Message & "')"
	set rs=connTemp.execute(query)
	set rs=nothing
	call closeDb()
	response.redirect "Blackout_main.asp?s=1&msg=New Blackout Date was added successfully!"
end if

%>
	
<!--#include file="AdminHeader.asp"-->
<script language="javascript">
function CalPop(sInputName)
{
	window.open('../Calendar/Calendar.asp?N=' + escape(sInputName) + '&DT=' + escape(window.eval(sInputName).value), 'CalPop','toolbar=0,width=378,height=225' );
}
</script>
<form name="addnew" method="post" action="Blackout_add.asp?action=add" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="2" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr> 
		<td width="20%">Blackout Date:</td>
    <td width="80%"><input type="text" name="Blackout_Date" size="20">&nbsp;<a href="javascript:CalPop('document.addnew.Blackout_Date');"><img SRC="../Calendar/icon_Cal.gif" border="0"></a>
		</td>
	</tr>
	<tr> 
		<td valign="top" nowrap="nowrap">Blackout Message:</td>
    <td><textarea cols="60" rows="6" name="Blackout_Message"></textarea>
	</td>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" align="center">
		<input type="submit" name="submit" value="Add New" class="submit2">
		&nbsp;
		<input type="button" name="back" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->