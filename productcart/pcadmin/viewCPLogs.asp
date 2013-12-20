<%@LANGUAGE="VBSCRIPT"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View Control Panel Logs" %>
<%PmAdmin=19%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="AdminHeader.asp"-->
<% 
if request("submit")<>"" then
	dim pcv_LogName,pcv_LogDateFrom,pcv_LogDateTo,pcv_LogTimeFrom,pcv_LogTimeTo,pcv_LogAdminUserID,pcv_LogAdminAccessLevel,pcv_LogPageAccessed,pcv_LogIP,pcv_LogSessionID
	
	pcv_LogName=request("LogName")
	pcv_LogDateFrom=request("LogDateFrom")
	pcv_LogDateTo=request("LogDateTo")
	pcv_LogTimeFrom=request("LogTimeFrom")
	pcv_LogTimeTo=request("LogTimeTo")
	pcv_LogAdminUserID=request("LogAdminUserID")
	pcv_LogAdminAccessLevel=request("LogAdminAccessLevel")
	pcv_LogPageAccessed=UCase(request("LogPageAccessed"))
	pcv_LogIP=request("LogIP")
	pcv_LogSessionID=request("LogSessionID")

	dim fso, pcStrCPLogFileName, ts
	
	set fso = server.CreateObject("Scripting.FileSystemObject") 
	
	if pcv_LogDateFrom<>"" and pcv_LogDateTo<>"" then
		'Grab all text files from date to date
		dim iDateDiffCnt
		iDateDiffCnt=DateDiff("d", pcv_LogDateFrom, pcv_LogDateTo)
		iFileCnt=iDateDiffCnt+1
	else
		'count is one
		iFileCnt=1
	end if
	
	dim strDisplayRecords, iRecordCnt
	strDisplayRecords=""
	iRecordCnt=0
	
	for n=1 to iFileCnt
		if iFileCnt>1 then
			pcv_LogName=CDate(pcv_LogDateFrom)+(n-1)
			pcv_LogName=replace(pcv_LogName,"/","")
			pcv_LogName=pcv_LogName&".txt"
		end if
		pcStrCPLogFileName=Server.Mappath ("CPLogs/"&pcv_LogName)
		'pcStrCPLogFileName=Server.Mappath ("CPLogs/"&replace(Date,"/","")&".txt")
		set f = fso.GetFile(pcStrCPLogFileName) 
		set ts = f.OpenAsTextStream(1, -2) 
		
		Do While not ts.AtEndOfStream
			pcv_TempTextStream=ts.ReadLine
			'parse the stream
			if instr(pcv_TempTextStream, ",") then
				pcv_CP_TRClass=""
				if request.form("HLLogout")<>"" then
					if instr(lcase(pcv_TempTextStream), "logoff.asp") then
						strHLLogoutColor=request.Form("HLLogoutColor")
						if strHLLogoutColor="" then
							strHLLogoutColor="#FFFFFF"
						end if
						pcv_CP_TRClass=" bgcolor='"&strHLLogoutColor&"'"
					end if
				end if
				if request.form("HLAllLogin")<>"" then
					if instr(lcase(pcv_TempTextStream), "login_1.asp") then
						strHLAllLoginColor=request.Form("HLAllLoginColor")
						if strHLAllLoginColor="" then
							strHLAllLoginColor="#FFFFFF"
						end if
						pcv_CP_TRClass=" bgcolor='"&strHLAllLoginColor&"'"
					end if
				end if 
				if request.form("HLSuccessLogin")<>"" then
					if instr(lcase(pcv_TempTextStream), "login_1.asp") AND instr(lcase(pcv_TempTextStream), "menu.asp") then
						strHLSuccessLoginColor=request.Form("HLSuccessLoginColor")
						if strHLSuccessLoginColor="" then
							strHLSuccessLoginColor="#FFFFFF"
						end if
						pcv_CP_TRClass=" bgcolor='"&strHLSuccessLoginColor&"'"
					end if
				end if 
				pcv_StreamArray=split(pcv_TempTextStream,",")
				pcv_LoggedIn="N0"
				if pcv_StreamArray(0)="-1" then
					pcv_LoggedIn="Yes"
				end if
				
				'Additional Conditionals
				intDontShow=0
				if intDontShow=0 AND pcv_LogAdminUserID<>"" then
					if (pcv_LogAdminUserID<>pcv_StreamArray(2)) then
						intDontShow=1
					end if
				end if
				
				if intDontShow=0 AND pcv_LogAdminAccessLevel<>"" then
					if (pcv_LogAdminAccessLevel<>pcv_StreamArray(3)) then
						intDontShow=1
					end if
				end if
				
				if intDontShow=0 AND pcv_LogPageAccessed<>"" then
					if instr(Ucase(pcv_StreamArray(9)),pcv_LogPageAccessed) then
					else
						intDontShow=1
					end if
				end if

				if intDontShow=0 AND pcv_LogIP<>"" then
					if (pcv_LogIP<>pcv_StreamArray(4)) then
						intDontShow=1
					end if
				end if
				
				if intDontShow=0 AND pcv_LogSessionID<>"" then
					if (pcv_LogSessionID<>pcv_StreamArray(5)) then
						intDontShow=1
					end if
				end if

				'Check for Time Filters
				if intDontShow=0 then
					if pcv_LogTimeFrom<>"" and pcv_LogTimeTo<>"" then
						iTimeDiffCnt=DateDiff("h", cDate(pcv_LogTimeFrom), cDate(pcv_LogTimeTo))
						if iTimeDiffCnt<0 then
							pcv_LogTimeFrom=""
							pcv_LogTimeTo=""
							pcv_ShowValidationMsg="Your From Date must be ealier than your To Date"
						end if
					end if
							
					if pcv_LogTimeFrom<>"" and pcv_LogTimeTo<>"" then
						iTimeCheckDiffCnt1=DateDiff("h", cDate(pcv_LogTimeFrom), cDate(pcv_StreamArray(7)))
						iTimeCheckDiffCnt2=DateDiff("h", cDate(pcv_LogTimeTo), cDate(pcv_StreamArray(7)))
						IF hour(cDate(pcv_LogTimeTo))<>hour(cDate(pcv_StreamArray(7))) THEN
							if iTimeCheckDiffCnt1=0 OR (iTimeCheckDiffCnt1>0 AND iTimeCheckDiffCnt1<(iTimeDiffCnt+1)) then
								iRecordCnt=iRecordCnt+1
								strDisplayRecords=strDisplayRecords&"<tr>"
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_LoggedIn&"</td>"
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(1)&"</td>"
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(2)&"</td>"
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(3)&"</td>"
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(4)&"</td>"
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(5)&"</td>"
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&ShowDateFrmt(pcv_StreamArray(6))&"</td>"
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(7)&"</td>"
								pcv_RefURL=pcv_StreamArray(8)
								if len(pcv_RefURL)>50 then
									pcv_RefURL="..."&right(pcv_RefURL,50)
								end if
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&"><a href='"&pcv_StreamArray(8)&"' title='"&pcv_StreamArray(8)&"'>"&pcv_RefURL&"</a></td>"
								pcv_PageAccessed=pcv_StreamArray(9)
								if len(pcv_PageAccessed)>50 then
									pcv_PageAccessed="..."&right(pcv_PageAccessed,50)
								end if
								strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&"><a href='"&pcv_StreamArray(9)&"' title='"&pcv_StreamArray(9)&"'>"&pcv_PageAccessed&"</a></td>"
								strDisplayRecords=strDisplayRecords&"</tr>"
							end if
						end if
					else
						iRecordCnt=iRecordCnt+1
						strDisplayRecords=strDisplayRecords&"<tr>"
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_LoggedIn&"</td>"
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(1)&"</td>"
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(2)&"</td>"
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(3)&"</td>"
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(4)&"</td>"
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(5)&"</td>"
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&ShowDateFrmt(pcv_StreamArray(6))&"</td>"
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&">"&pcv_StreamArray(7)&"</td>"
						pcv_RefURL=pcv_StreamArray(8)
						if len(pcv_RefURL)>50 then
							pcv_RefURL="..."&right(pcv_RefURL,50)
						end if
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&"><a href='"&pcv_StreamArray(8)&"' title='"&pcv_StreamArray(8)&"'>"&pcv_RefURL&"</a></td>"
						pcv_PageAccessed=pcv_StreamArray(9)
						if len(pcv_PageAccessed)>50 then
							pcv_PageAccessed="..."&right(pcv_PageAccessed,50)
						end if
						strDisplayRecords=strDisplayRecords&"<td nowrap "&pcv_CP_TRClass&"><a href='"&pcv_StreamArray(9)&"' title='"&pcv_StreamArray(9)&"'>"&pcv_PageAccessed&"</a></td>"
						strDisplayRecords=strDisplayRecords&"</tr>"
					end if
					%>
					<%
				end if
			end if
		Loop 
	Next
	
	if iRecordCnt<>0 then %>
    	<style>
		table.pcCPcontent td{
			font-size: 10px;
		}
		table.pcCPcontent th {
			font-size: 11px;
		}
		</style>
    	<table class="pcCPcontent">
            <tr> 
                <th>Logged In</th>
                <th>System ID</th>
                <th>User ID</th>
                <th>Access Level</th>
                <th>IP Address</th>
                <th>Session ID</th>
                <th>Date</th>
                <th>Time</th>
                <th>Referral URL</th>
                <th>Page Accessed</th>
            </tr>
            <tr>
                <td colspan="10" class="pcCPspacer"></td>
            </tr>
            <%=strDisplayRecords%>
         </table>
    <% else %>
    <div class="pcCPmessage">
        No Records Found. <br><br><a href="javascript:history.back()">Back</a>
    </div>
    <% end if
else

Sub GenDropDown(ddtype)

	pcv_DirPath = Server.Mappath ("CPLogs") 'Physical Path
	pcv_DdType = ddtype
	pcv_ExcludeFile = "_"
	
	Dim rsFSO, objFSO, objFolder, File
	Const adInteger = 3
	Const adDate = 7
	Const adVarChar = 200
  
	'create an ADODB.Recordset and call it rsFSO
	Set rsFSO = Server.CreateObject("ADODB.Recordset")
	
	'Open the FSO object
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	'go get the folder to output it's contents
	Set objFolder = objFSO.GetFolder(pcv_DirPath)
	
	'Now get rid of the objFSO since we're done with it.
	Set objFSO = Nothing
  
	'create the various rows of the recordset
	With rsFSO.Fields
		.Append "Name", adVarChar, 200
		.Append "Type", adVarChar, 200
		.Append "DateCreated", adDate
		.Append "DateLastAccessed", adDate
		.Append "DateLastModified", adDate
		.Append "Size", adInteger
		.Append "TotalFileCount", adInteger
	End With
	rsFSO.Open()
	
	'Now let's find all the files in the folder
	For Each File In objFolder.Files
	
		If (Left(File.Name, 1)) <> pcv_ExcludeFile Then 
			rsFSO.AddNew
			rsFSO("Name") = File.Name
			rsFSO("Type") = File.Type
			rsFSO("DateCreated") = File.DateCreated
			rsFSO("DateLastAccessed") = File.DateLastAccessed
			rsFSO("DateLastModified") = File.DateLastModified
			rsFSO("Size") = File.Size
			rsFSO.Update
		End If

	Next
	
	rsFSO.Sort = "Type DESC, DateCreated ASC "

	Set objFolder = Nothing

	rsFSO.MoveFirst()

	While Not rsFSO.EOF
		'Display changes
		pcv_DateValue=FormatDateTime(rsFSO("DateCreated").Value,2)
		if scDateFrmt = "DD/MM/YY" 	then
			pcv_FrmtDateValue=Day(pcv_DateValue)&"/"&Month(pcv_DateValue)&"/"&Year(pcv_DateValue)
		else
			pcv_FrmtDateValue=Month(pcv_DateValue)&"/"&Day(pcv_DateValue)&"/"&Year(pcv_DateValue)
		end if
		if ddtype="2" then
  	  		Response.Write "<option value='"&pcv_DateValue&"'>"&pcv_FrmtDateValue&"</option>"
		else	
  	  		Response.Write "<option value='"&rsFSO("Name").Value&"'>"&pcv_FrmtDateValue&"</option>"
		end if
		rsFso.MoveNext()
	Wend
  
	rsFSO.close()
	Set rsFSO = Nothing
	
End Sub
%>
<form name="form1" method="post" action="viewCPLogs.asp" class="pcForms">
	<table class="pcCPcontent">
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
      <th colspan="2">Choose the log files</th>
    </tr>
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
      <td colspan="2">Choose the log files that contain the data that you wish to review:</td>
    </tr>
    <tr>
      <td nowrap="nowrap" width="20%" align="right">Pick a <strong>specific date</strong>:</td>
			<td width="80%">
        <label>
        <select name="LogName" id="LogName">
				<% GenDropDown("1") %>
        </select>
      	</label>
			</td>
    </tr>
    <tr>
      <td nowrap="nowrap" align="right">... or a <strong>date range</strong>:</td>
			<td>
        From: 
				<select name="LogDateFrom" id="LogDateFrom">
        	<option value="">Select From Date</option>
        	<% GenDropDown("2") %>
        </select>
        To: 
        <select name="LogDateTo" id="LogDateTo">
        	<option value="">Select To Date</option>
			<% GenDropDown("2") %>
        </select>
			</td>
    </tr>
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
      <th colspan="2">Filter the data</th>
    </tr>
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>

    <tr>
      <td nowrap="nowrap">Time Range:</td>
			<td>From 
        <select name="LogTimeFrom" id="LogTimeFrom">
        	<option value="">Select Time From</option>
            <option value="00:00">12:00 a.m.</option>
            <option value="01:00">1:00 a.m.</option>
            <option value="02:00">2:00 a.m.</option>
            <option value="03:00">3:00 a.m.</option>
            <option value="04:00">4:00 a.m.</option>
            <option value="05:00">5:00 a.m.</option>
            <option value="06:00">6:00 a.m.</option>
            <option value="07:00">7:00 a.m.</option>
            <option value="08:00">8:00 a.m.</option>
            <option value="09:00">9:00 a.m.</option>
            <option value="10:00">10:00 a.m.</option>
            <option value="11:00">11:00 a.m.</option>
            <option value="12:00">12:00 p.m.</option>
            <option value="13:00">1:00 p.m.</option>
            <option value="14:00">2:00 p.m.</option>
            <option value="15:00">3:00 p.m.</option>
            <option value="16:00">4:00 p.m.</option>
            <option value="17:00">5:00 p.m.</option>
            <option value="18:00">6:00 p.m.</option>
            <option value="19:00">7:00 p.m.</option>
            <option value="20:00">8:00 p.m.</option>
            <option value="21:00">9:00 p.m.</option>
            <option value="22:00">10:00 p.m.</option>
            <option value="23:00">11:00 p.m.</option>
        </select>
        To 
        <select name="LogTimeTo" id="LogTimeTo">
        	<option value="">Select Time To</option>
            <option value="00:00">12:00 a.m.</option>
            <option value="01:00">1:00 a.m.</option>
            <option value="02:00">2:00 a.m.</option>
            <option value="03:00">3:00 a.m.</option>
            <option value="04:00">4:00 a.m.</option>
            <option value="05:00">5:00 a.m.</option>
            <option value="06:00">6:00 a.m.</option>
            <option value="07:00">7:00 a.m.</option>
            <option value="08:00">8:00 a.m.</option>
            <option value="09:00">9:00 a.m.</option>
            <option value="10:00">10:00 a.m.</option>
            <option value="11:00">11:00 a.m.</option>
            <option value="12:00">12:00 p.m.</option>
            <option value="13:00">1:00 p.m.</option>
            <option value="14:00">2:00 p.m.</option>
            <option value="15:00">3:00 p.m.</option>
            <option value="16:00">4:00 p.m.</option>
            <option value="17:00">5:00 p.m.</option>
            <option value="18:00">6:00 p.m.</option>
            <option value="19:00">7:00 p.m.</option>
            <option value="20:00">8:00 p.m.</option>
            <option value="21:00">9:00 p.m.</option>
            <option value="22:00">10:00 p.m.</option>
            <option value="23:00">11:00 p.m.</option>
        </select>
				</td>
    </tr>

    <tr>
      <td nowrap="nowrap">User ID:</td>
      <td>
				<label>
        <input type="text" name="LogAdminUserID" id="LogAdminUserID">
     		</label>
			</td>  
			</td>
    </tr>
		
    <tr>
      <td nowrap="nowrap">Access Level:</td>
      <td>
			  <label>
        <input type="text" name="LogAdminAccessLevel" id="LogAdminAccessLevel">
        </label>
			</td>
    </tr>
		
    <tr>
      <td nowrap="nowrap">Page Accessed:</td>
      <td>
        <label>
        <input type="text" name="LogPageAccessed" id="LogPageAccessed">
        </label>
			</td>
    </tr>
		
    <tr>
      <td nowrap="nowrap">IP Address:</td>
      <td>
        <label>
        <input type="text" name="LogIP" id="LogIP">
        </label></td>
    </tr>
		
    <tr>
      <td nowrap="nowrap">Session ID:</td>
      <td>
        <label>
        <input type="text" name="LogSessionID" id="LogSessionID">
        </label></td>
    </tr>
		
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
      <th colspan="2">Formatting</th>
    </tr>
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
        <td colspan="2">You can highlight the results to better understand where a user's activity starts (successful login) and when it ends (user logged out). Please note that ProductCart does not log the "session timed out" or "user closed browser" events.</td>
    </tr>
    <tr>
      <td colspan="2"><input type="checkbox" name="HLLogout" id="HLLogout">
      	Highlight 'Logout Requests' with this color: 
      	<select name="HLLogoutcolor">
      	<option style="background-color:#99CCFF;" value="#99CCFF" selected>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>  
        <option style="background-color:#ffccff;" value="#ffccff"> </option>   
        <option style="background-color:#ccffcc;" value="#ccffcc"> </option>     
        <option style="background-color:#ffffcc;" value="#ffffcc"> </option>    
      </select></td>
    </tr>
    <tr>
        <td colspan="2"><input type="checkbox" name="HLALLLogin" id="HLALLLogin">
        Highlight 'Login Attempts' with this color: 
        <select name="HLALLLoginColor">
        <option style="background-color:#99CCFF;" value="#99CCFF" selected>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
        <option style="background-color:#ffccff;" value="#ffccff"> </option>
        <option style="background-color:#ccffcc;" value="#ccffcc"> </option>
        <option style="background-color:#ffffcc;" value="#ffffcc"> </option>
        </select></td>
    </tr>
    <tr>
        <td colspan="2"><input type="checkbox" name="HLSuccessLogin" id="HLSuccessLogin"> 
        Highlight 'Successful Logins' with this color: 
        <select name="HLSuccessLoginColor">
        <option style="background-color:#99CCFF;" value="#99CCFF" selected>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
        <option style="background-color:#ffccff;" value="#ffccff"> </option>
        <option style="background-color:#ccffcc;" value="#ccffcc"> </option>
        <option style="background-color:#ffffcc;" value="#ffffcc"> </option>
        </select></td>
    </tr>
    <tr>
        <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td>
        <input type="submit" name="submit" id="submit" value="Submit" class="submit2">
      </td>
    </tr>
    </table>
</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->