<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials"
pageIcon="pcv4_icon_salesManager.png"
pageTitle="Sales Manager - View &amp; Edit Pending Sales"
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="sm_check.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<% 
Dim query, conntemp, rs

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

call openDB()

%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<table class="pcCPcontent">
<tr>
	<td colspan="6"><strong>Pending Sales</strong> are Sales that have been created and saved to the database, but that are not currently running. They are in a &quot;stand-by&quot; mode. Prices are not affected until these sales are started.</td>
</tr>
<%if request("m")<>"" then%>
<tr>
	<td colspan="6">
	<%Select Case request("m")
	Case "1","2", "3":%>
		<div class="pcCPmessageSuccess">
	<%Case Else:%>
		<div class="pcCPmessage">
	<%End Select%>
	<%Select Case request("m")
	Case "1","3": 
	if request("m")="3" then
		tmpN="cloned"
	else
		tmpN="created"
	end if
	response.write "The Sale has been <strong>" & tmpN & " successfully</strong> and is now listed below. The sale is now in a &quot;stand-by&quot; mode: no prices are altered until the sale is started."
	Case "2": response.write "The Sale has been <strong>updated successfully</strong>. The sale is now in a &quot;stand-by&quot; mode: no prices are altered until the sale is started."
	End Select%>
	<%if (request("id")<>"") AND (request("id")<>"0") then
	query="SELECT TOP 1 pcSC_ID FROM pcSales_Completed WHERE pcSales_ID=" & request("id") & " AND ((pcSC_Status=1) OR (pcSC_Status=2));"
		set rs=connTemp.execute(query)
		CanStart=1
		if not rs.eof then
			CanStart=0
		end if
		set rs=nothing
		if CanStart=1 then%>
			<div style="padding-top:10px;"><a href="sm_start.asp?a=start&id=<%=request("id")%>"><strong>Start this sale NOW &gt;&gt;</strong></a></div>
		<%end if%>
	<%end if%>
	</div>	
	</td>
</tr>
<%end if%>
<tr>
	<td colspan="6" class="pcCPspacer"></td>
</tr>
<%query="SELECT pcSales_ID,pcSales_Name,pcSales_CreatedDate,pcSales_EditedDate,pcSales_Param1,pcSales_Param2,pcSales_Tech FROM pcSales WHERE pcSales_Removed=0 AND (pcSales.pcSales_ID NOT IN (SELECT DISTINCT pcSales_ID FROM pcSales_Completed))ORDER BY pcSales_ID ASC;"
set rs=connTemp.execute(query)
if rs.eof then
	set rs=nothing%>
	<tr>
	<td colspan="6">
		<div class="pcCPmessage">
			No Sales have been found.
		</div>	
	</td>
	</tr>
	<tr>
		<td colspan="6" class="pcCPspacer"></td>
	</tr>
	<tr>
	<td colspan="6">
		<input type="button" name="Go" value=" Create New Sale " onclick="location='sm_addedit_S1.asp?a=new';" class="submit2">	
	</td>
	</tr>
<%else
	pcArr=rs.getRows()
	intCount=ubound(pcArr,2)
	%>
	<tr>
		<th width="40%">Sale Name</th>
		<th width="10%" nowrap>Item(s)</th>
		<th width="20%" nowrap>Saved On</th>
		<th width="20%" nowrap>Last Edited</th>
		<th width="10%" colspan="2">&nbsp;</th>
	</tr>
    <tr>
        <td colspan="6" class="pcCPspacer"></td>
    </tr>
	<%For i=0 to intCount%>
	<tr valign="top" onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist">
		<td>
			<a href="JavaScript:;" onClick="document.getElementById('pcSalesFullDetails<%=i%>').style.display='';"><%=pcArr(1,i)%></a>
        </td>
		<%tmpPrdCount="N/A"
		Set cn = Server.CreateObject("ADODB.Connection")
		Set cmd = Server.CreateObject("ADODB.Command")
		cn.Open scDSN
		Set cmd.ActiveConnection = cn
		cmd.CommandText = "uspGetPrdCount"
		cmd.CommandType = adCmdStoredProc

		cmd.Parameters.Refresh
		cmd.Parameters("@Param1") = pcArr(4,i)
		cmd.Parameters("@Param2") = pcArr(5,i)
		cmd.Execute
	
		tmpPrdCount=cmd.Parameters("@SMCOUNT")
		session("sm_PrdCount")=tmpPrdCount
		Set cmd=nothing
		set cn=nothing%>
		<td><%=tmpPrdCount%></td>
		<td><%=ShowDateTimeFrmt(pcArr(2,i))%></td>
		<td>
			<%if Not IsNull(pcArr(3,i)) then%>
				<%=ShowDateTimeFrmt(pcArr(3,i))%>
			<%else%>
				<%=ShowDateTimeFrmt(pcArr(2,i))%>
			<%end if%>
		</td>
		<td colspan="2" nowrap align="right"><a href="sm_addedit_S5.asp?a=new&id=<%=pcArr(0,i)%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit" title="Edit this saved Sale"></a>&nbsp;<a href="javascript:if (confirm('You are about to remove this SALE. Are you sure you want to complete this action?')) location='sm_remove.asp?id=<%=pcArr(0,i)%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Remove" title="Remove this saved Sale"></a>
		<a href="sm_start.asp?a=start&id=<%=pcArr(0,i)%>"><img src="images/pcIconStart.jpg" width="12" height="12" alt="Start" title="START this saved Sale"></a></td>
	</tr>
    <tr id="pcSalesFullDetails<%=i%>" style="display: none;">
    	<td colspan="6"><div class="pcCPmessageInfo"><div style="float: right;"><a href="javaScript:;" onClick="document.getElementById('pcSalesFullDetails<%=i%>').style.display='none';">Hide</a></div><%=pcArr(6,i)%><a href="sm_addedit_S1.asp?c=new&a=rev&id=<%=pcArr(0,i)%>&b=3">Click here</a> to view/edit list products included in this sale</div></td>
    </tr>
	<%Next%>
<tr>
	<td colspan="6" class="pcCPspacer"></td>
</tr>
<tr align="center">
	<td colspan="6">
		<input type="button" name="Go" value=" Create New Sale " onclick="location='sm_addedit_S1.asp?a=new';" class="submit2">
	</td>
</tr>
<%end if%>
</table>
<% 
call closeDb()
set rs= nothing
%>
<!--#include file="AdminFooter.asp"-->