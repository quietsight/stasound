<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials"
pageIcon="pcv4_icon_salesManager.png"
SMArchived=request("arc")
if IsNull(SMArchived) then
	SMArchived=0
end if
if SMArchived=1 then
	pageTitle="Sales Manager - Archived Sales"
else
	pageTitle="Sales Manager - Currently Running &amp; Completed Sales"
end if
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

if request("a")="clone" then
	pcSaleID=request("id")
	if not IsNull(pcSaleID) then
	if IsNumeric(pcSaleID) AND pcSaleID>"0" then
		query="SELECT pcSales_TargetPrice,pcSales_Type,pcSales_Relative,pcSales_Amount,pcSales_Round,pcSales_Name,pcSales_Desc,pcSales_ImgURL,pcSales_CreatedDate,pcSales_Param1,pcSales_Param2,pcSales_Tech FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
		set rs=connTemp.execute(query)
		if not rs.eof then
			pcArr=rs.getRows()
			set rs=nothing
			
			dim dtTodaysDate
			dtTodaysDate=Date()
			if SQL_Format="1" then
				dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
			else
				dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
			end if
			
			For i=0 to 11
				if pcArr(i,0)<>"" then
					pcArr(i,0)=replace(pcArr(i,0),"'","''")
				end if
			Next
			
			i=0
			
			query="INSERT INTO pcSales (pcSales_TargetPrice,pcSales_Type,pcSales_Relative,pcSales_Amount,pcSales_Round,pcSales_Name,pcSales_Desc,pcSales_ImgURL,pcSales_CreatedDate,pcSales_Param1,pcSales_Param2,pcSales_Tech) VALUES (" & pcArr(0,i) & "," & pcArr(1,i) & "," & pcArr(2,i) & "," & pcArr(3,i) & "," & pcArr(4,i) & ",'" & pcArr(5,i) & " (Cloned)" & "','" & pcArr(6,i) & "','" & pcArr(7,i) & "','" & dtTodaysDate & "','" & pcArr(9,i) & "','" & pcArr(10,i) & "','" & pcArr(11,i) & "');"
			set rs=connTemp.execute(query)
			set rs=nothing
			
			query="SELECT TOP 1 pcSales_ID FROM pcSales ORDER BY pcSales_ID DESC;"
			set rs=connTemp.execute(query)
			if not rs.eof then
				pcSaleID1=rs("pcSales_ID")
				if IsNull(pcSaleID1) then
					pcSaleID1=0
				end if
			end if
			set rs=nothing
			
			query="INSERT INTO pcSales_Pending (IDProduct,pcSales_ID) SELECT DISTINCT IDProduct," & pcSaleID1 & " FROM pcSales_Pending WHERE pcSales_ID=" & pcSaleID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			
			query="UPDATE pcSales SET pcSales_Param2='pcSales_Pending.pcSales_ID=" & pcSaleID1 & "' WHERE pcSales_ID=" & pcSaleID1 & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			
			call closedb()
			response.Clear()
			response.redirect "sm_manage.asp?m=3&id=" & pcSaleID1
		end if
	end if
	end if
end if
%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<table class="pcCPcontent">
<tr>
	<td colspan="5" class="pcCPspacer"></td>
</tr>
<%if request("msg")<>"" then%>
<tr>
		<td colspan="5">
			<div class="pcCPmessageSuccess">
				<%if request("msg")="1" then%>
				The selected products you want to remove are equal to the total products of the Sale, it's also stopped.<br>The Sale has been successfully stopped!
				<%end if%>
			</div>
		</td>
</tr>
<%end if%>
<%if request("a")="arc" then
	pcSCID=request("id")
	pcSaleID=request("parent")
	if not IsNull(pcSCID) then
	if IsNumeric(pcSCID) AND pcSCID>"0" then
		query="DELETE FROM pcSales_Pending WHERE pcSales_ID=" & pcSaleID & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
		query="UPDATE pcSales_Completed SET pcSC_Archived=1 WHERE pcSC_ID=" & pcSCID & ";"
		set rs=connTemp.execute(query)
		set rs=nothing%>
		<tr>
		<td colspan="5">
			<div class="pcCPmessageSuccess">
				The Sale has been archived successfully!
			</div>
		</td>
		</tr>
	<%end if
	end if
end if%>
<%
if SMArchived=1 then
	tmp1=" AND pcSC_Archived=1 "
else
	tmp1=" AND pcSC_Archived=0 "
end if
query="SELECT pcSC_ID,pcSC_SaveName,pcSC_StartedDate,pcSC_StoppedDate,pcSC_Status,pcSales_ID,pcSC_Archived FROM pcSales_Completed WHERE pcSC_Status>=0 " & tmp1 & " ORDER BY pcSC_ID DESC;"
set rs=connTemp.execute(query)
if rs.eof then
	set rs=nothing%>
	<tr>
	<td colspan="5">
		<div class="pcCPmessage">
			<%if SMArchived=1 then%>
			No Archived Sales have been found.
			<%else%>
			No Sales have been found.
			<%end if%>
		</div>	
	</td>
	</tr>
	<tr>
		<td colspan="5" class="pcCPspacer"></td>
	</tr>
	<tr>
	<td colspan="5">
		<input type="button" name="Go" value=" Start A New Sale " onclick="location='sm_start.asp';" class="submit2">	
	</td>
	</tr>
<%else
	pcArr=rs.getRows()
	intCount=ubound(pcArr,2)
	%>
	<tr>
		<th width="36%">Name</th>
		<th width="22%">Started On</th>
		<th width="22%">Stopped On</th>
		<th width="10%">Status</th>
		<th width="10%">&nbsp;</th>
	</tr>
	<%For i=0 to intCount%>
	<tr valign="top" onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist">
		<td><%=pcArr(1,i)%></td>
		<td><%=ShowDateTimeFrmt(pcArr(2,i))%></td>
		<td>
			<%if pcArr(3,i)<>"" then%>
				<%=ShowDateTimeFrmt(pcArr(3,i))%>
			<%else%>N/A<%end if%>
		</td>
		<td>
			<%Select Case pcArr(4,i)
			Case "1": response.write "Started"
			Case "2": response.write "Live"
			Case "3": response.write "Stopped"
			Case "4": response.write "Completed"
			Case Else: response.write "N/A"
			End Select%>
		</td>
		<td align="right" nowrap><%if ((pcArr(4,i)="2") OR (pcArr(4,i)="4")) AND (pcArr(6,i)<>"1") then%><a href="sm_addedit_S1.asp?c=new&a=rev&id=<%=pcArr(5,i)%>&b=6&scid=<%=pcArr(0,i)%>"><img src="images/pcIconList.jpg" alt="Products" title="List products included in the sale"></a>&nbsp;<%end if%><a href="sm_saledetails.asp?id=<%=pcArr(0,i)%><%if pcArr(4,i)="2" then%>&e=1<%end if%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Details" title="View<%if pcArr(4,i)="2" then%>/Edit<%end if%> details"></a>
		  <%if (pcArr(4,i)="1") OR (pcArr(4,i)="2") then%><a href="sm_stop.asp?a=stop&id=<%=pcArr(0,i)%>"><img src="images/pcIconStop.jpg" width="12" height="12" alt="Stop" title="Stop the sale"></a><%end if%>
		  <%if (pcArr(4,i)="4") AND (pcArr(6,i)<>"1") then%>
		  		<a href="sm_sales.asp?a=clone&id=<%=pcArr(5,i)%>"><img src="images/pcIconClone.jpg" alt="Clone" title="Clone the sale to create a similar one"></a>
				<a href="javascript:if (confirm('You are about to ARCHIVE this sale. Product information (list of products that the sale applied to) will be removed from the sales details. Are you sure you want to complete this action?')) location='sm_sales.asp?a=arc&id=<%=pcArr(0,i)%>&parent=<%=pcArr(5,i)%>&arc=1';"><img src="images/pcIconMinus.jpg" alt="Archive" title="Archive"></a>
		  <%end if%>
		  <a href="sm_salereport.asp?id=<%=pcArr(5,i)%>&sub=<%=pcArr(0,i)%>" target="_blank"><img src="images/pcIconChart.jpg" width="12" height="12" alt="Report" title="View orders that included products on sale"></a>
		</td>
	</tr>
	<%Next%>
<tr>
	<td colspan="5" class="pcCPspacer"></td>
</tr>
<tr align="center">
	<td colspan="5">
		<input type="button" name="Go" value="Create New Sale" onclick="location='sm_addedit_S1.asp?a=new';" class="submit2">
        &nbsp;
		<%if SMArchived<>1 then
        query="SELECT pcSC_ID,pcSC_SaveName,pcSC_StartedDate,pcSC_StoppedDate,pcSC_Status,pcSales_ID FROM pcSales_Completed WHERE pcSC_Status>=0 AND pcSC_Archived=1 ORDER BY pcSC_ID DESC;"
        set rs=connTemp.execute(query)
        if not rs.eof then%>
        <input type="button" value="Show Archived Sales" onClick="location='sm_sales.asp?arc=1'">
        <%end if
        set rs=nothing
        else%>
        <input type="button" value="Show Running &amp; Completed Sales" onClick="location='sm_sales.asp'">
        <%end if%>
        
        
	</td>
</tr>
<%end if%>
</table>
<% 
call closeDb()
set rs= nothing
%>
<!--#include file="AdminFooter.asp"-->