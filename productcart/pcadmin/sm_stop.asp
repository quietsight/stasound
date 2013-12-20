<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials"
pageIcon="pcv4_icon_salesManager.png"
IF request("a")="stop" THEN
	pageTitle="Sales Manager - Stop A Sale - Preview"
ELSE
	pageTitle="Sales Manager - Stop A Sale"
END IF
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

call opendb()

IF (request("a")="stop") OR (request("a")="run") THEN
	pcSCID=request("id")
	if pcSCID="" then
		call closedb()
		response.redirect "sm_manage.asp"
	else
		if Not (IsNumeric(pcSCID)) then
			call closedb()
			response.redirect "sm_manage.asp"
		end if
	end if
END IF

query="SELECT pcSC_ID,pcSC_SaveName,pcSC_StartedDate FROM pcSales_Completed WHERE (pcSC_Status=1) OR (pcSC_Status=2);"
set rs=connTemp.execute(query)
if rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "sm_manage.asp"
else
	pcArr=rs.getRows()
	intCount=ubound(pcArr,2)
	set rs=nothing
end if

%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%IF request("a")="stop" THEN
	query="SELECT pcSC_SaveName,pcSC_StartedDate,pcSC_SaveDesc,pcSC_SaveTech,pcSales_ID FROM pcSales_Completed WHERE pcSC_ID=" & pcSCID & ";"
	set rs=connTemp.execute(query)
	if rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "sm_manage.asp"
	else
		tmpSaleDetails=rs("pcSC_SaveTech")
		tmpSaleName=rs("pcSC_SaveName")
		tmpStartedDate=rs("pcSC_StartedDate")
		tmpSaleDesc=rs("pcSC_SaveDesc")
		pcSaleID=rs("pcSales_ID")
		set rs=nothing
		
		Set cn = Server.CreateObject("ADODB.Connection")
		Set cmd = Server.CreateObject("ADODB.Command")
		cn.Open scDSN
		Set cmd.ActiveConnection = cn
		cmd.CommandText = "uspGetUpdatedPrdCount"
		cmd.CommandType = adCmdStoredProc

		cmd.Parameters.Refresh
		cmd.Parameters("@SCID") = pcSCID
		cmd.Execute
	
		tmpPrdCount=cmd.Parameters("@SMCOUNT")

		Set cmd=nothing
		set cn=nothing
		
	%>
	<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td width="15%">Sale Name:</td>
		<td> <b><%=tmpSaleName%></b></td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr valign="top">
		<td>Sale Description:</td>
		<td><%=tmpSaleDesc%></td>
	</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr valign="top">
		<td>Products:</td>
		<td><a href="sm_addedit_S1.asp?c=new&a=rev&id=<%=pcSaleID%>&b=8&scid=<%=pcSCID%>">Click here</a> to view list products included in this sale</td>
	</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr valign="top">
		<td>Details:</td>
		<td><%=tmpSaleDetails%></td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">The sale was started on <b><%=ShowDateTimeFrmt(tmpStartedDate)%></td>
	</tr>
	<tr>
		<td colspan="2">It affected <b><%=tmpPrdCount%></b> product(s) in your store.</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<input type="button" name="Run" value=" Stop Sale Now " onclick="javascript:if (confirm('You are about to stop the sale: <%=replace(tmpSaleName,"'","\'")%>. Are you sure you want to complete this action?')) location='sm_stop.asp?a=run&id=<%=pcSCID%>';" class="submit2">
		</td>
	</tr>
	</table>	
	<%end if
ELSE
IF request("a")="run" THEN

	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	end if
	query="UPDATE pcSales_Completed SET pcSC_StoppedDate='" & dtTodaysDate & "',pcSC_Status=3 WHERE pcSC_ID=" & pcSCID & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	'Get Affected Product Count
	Set cn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	cn.Open scDSN
	Set cmd.ActiveConnection = cn
	cmd.CommandText = "uspGetUpdatedPrdCount"
	cmd.CommandType = adCmdStoredProc

	cmd.Parameters.Refresh
	cmd.Parameters("@SCID") = pcSCID
	cmd.Execute
	
	tmpPrdCount=cmd.Parameters("@SMCOUNT")

	Set cmd=nothing
	set cn=nothing%>
	<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Processing Status</th>
	</tr>
	<tr>
	<td>
	<%
	'Inactivate Prds to restore prices
	Set cn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	cn.Open scDSN
	Set cmd.ActiveConnection = cn
	cmd.CommandText = "uspInActivatePrds"
	cmd.CommandType = adCmdStoredProc

	cmd.Parameters.Refresh
	cmd.Parameters("@SCID") = pcSCID
	cmd.Execute
	
	tmpPrdCount1=cmd.Parameters("@SMCOUNT")

	Set cmd=nothing
	set cn=nothing	

	'Restore products
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	end if
		
	query="UPDATE pcSales_Completed SET pcSC_REStartedDate='" & dtTodaysDate & "' WHERE pcSC_ID=" & pcSCID & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
		
	Set cn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	cn.Open scDSN
	Set cmd.ActiveConnection = cn
	cmd.CommandText = "uspRestorePrices"
	cmd.CommandType = adCmdStoredProc

	cmd.Parameters.Refresh
	cmd.Parameters("@SCID") = pcSCID
	cmd.Execute
	
	tmpREPrdCount=cmd.Parameters("@SMCOUNT")

	Set cmd=nothing
	set cn=nothing
	
	tmpREPrdCount=tmpPrdCount
		
	tmpSCStatus=1
	if Clng(tmpREPrdCount)<>Clng(tmpPrdCount) then
		tmpSCStatus=3
	end if
		
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	end if
		
	query="UPDATE pcSales_Completed SET pcSC_REComDate='" & dtTodaysDate & "',pcSC_RETotal=" & tmpREPrdCount & " WHERE pcSC_ID=" & pcSCID & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	if tmpSCStatus=3 then
		dtTodaysDate=Date()
		if SQL_Format="1" then
			dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
		else
			dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
		end if
		query="UPDATE pcSales_Completed SET pcSC_StoppedDate='" & dtTodaysDate & "',pcSC_Status=3 WHERE pcSC_ID=" & pcSCID & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
		
	if Clng(tmpREPrdCount)<>Clng(tmpPrdCount) then%>
		<b><%=tmpREPrdCount%></b> product(s) have been restored.<br><br>
		<div class="pcCPmessage">
			The Sale cannot be stopped because the number of restored product(s) are different than the number of affected product(s)
		</div>
		<br><br>
	<%end if
	
	IF tmpSCStatus=1 THEN
	
	'Activate Prds and remove pcSC_ID from table Products
	Set cn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	cn.Open scDSN
	Set cmd.ActiveConnection = cn
	cmd.CommandText = "uspActivatePrds"
	cmd.CommandType = adCmdStoredProc

	cmd.Parameters.Refresh
	cmd.Parameters("@SCID") = pcSCID
	cmd.Execute
	
	tmpPrdCount1=cmd.Parameters("@SMCOUNT")

	Set cmd=nothing
	set cn=nothing
	
	'Removed backed-up records
	Set cn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	cn.Open scDSN
	Set cmd.ActiveConnection = cn
	cmd.CommandText = "uspRmvBackedUpRecords"
	cmd.CommandType = adCmdStoredProc

	cmd.Parameters.Refresh
	cmd.Parameters("@SCID") = pcSCID
	cmd.Execute
	
	tmpPrdCount1=cmd.Parameters("@SMCOUNT")

	Set cmd=nothing
	set cn=nothing
			
		if Clng(tmpREPrdCount)=Clng(tmpPrdCount) then
			dtTodaysDate=Date()
			if SQL_Format="1" then
				dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
			else
				dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
			end if
			query="UPDATE pcSales_Completed SET pcSC_ComDate='" & dtTodaysDate & "',pcSC_Status=4 WHERE pcSC_ID=" & pcSCID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		
			%>
			<div class="pcCPmessageSuccess">
				The Sale has been successfully stopped!
			</div>
			<br><br>
		<%end if
		
	END IF%>
	</td>
	</tr>
	<tr>
	<td class="pcCPspacer"></td>
	</tr>
	</table>
<%ELSE%>
<form name="Form1" action="sm_stop.asp?a=stop" method="post" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2">Please select the Sale that you would like to stop:</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td width="10%" nowrap>Sales currently running:</td>
	<td width="90%">
		<select name="id" id="id">
			<%For i=0 to intCount%>
				<option value="<%=pcArr(0,i)%>"><%=pcArr(1,i)%> - started on: <%=ShowDateTimeFrmt(pcArr(2,i))%></option>
			<%Next%>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>
<tr>
	<td width="10%">&nbsp;</td>
	<td>
		<input type="submit" name="Preview" value=" Continue " class="submit2">
	</td>
</tr>
</table>
</form>
<%END IF
END IF%>
<% 
call closeDb()
set rs= nothing
%>
<!--#include file="AdminFooter.asp"-->