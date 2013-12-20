<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=0%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/MailUpFunctions.asp"-->
<!--#include file="../includes/rc4.asp"-->
<% pageTitle="MailUp - Last Import Status and History" %>
<% section="mngAcc"
Server.ScriptTimeout = 5400
Dim connTemp,query,rs
Dim tmp_setup
Dim tmpCanNotStart
tmpCanNotStart=0

	tmp_setup=0
	tmp_bulk=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	call opendb()
	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess,pcMailUpSett_BulkRegister FROM pcMailUpSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcMailUpSett_APIUser=rs("pcMailUpSett_APIUser")
		session("CP_MU_APIUser")=pcMailUpSett_APIUser
		pcMailUpSett_APIPassword=enDeCrypt(rs("pcMailUpSett_APIPassword"), scCrypPass)
		session("CP_MU_APIPassword")=pcMailUpSett_APIPassword
		pcMailUpSett_URL=rs("pcMailUpSett_URL")
		session("CP_MU_URL")=pcMailUpSett_URL
		tmp_Auto=rs("pcMailUpSett_AutoReg")
		if IsNull(tmp_Auto) or tmp_Auto="" then
			tmp_Auto=0
		end if
		tmp_setup=rs("pcMailUpSett_RegSuccess")
		if IsNull(tmp_setup) or tmp_setup="" then
			tmp_setup=0
		end if
		tmp_bulk=rs("pcMailUpSett_BulkRegister")
		if IsNull(tmp_bulk) or tmp_bulk="" then
			tmp_bulk=0
		end if
	end if
	set rs=nothing
call closedb()

if tmp_setup=0 then
	response.redirect "mu_manageNewsWiz.asp"
end if

msg=""

'Post Back
IF request("action")="restart" THEN
	call opendb()
	query="SELECT pcMailUpSett_LastIDList,pcMailUpSett_LastIDProcess FROM pcMailUpSettings WHERE pcMailUpSettings.pcMailUpSett_LastIDProcess<>'';"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpIDLists=""
		tmpIDProList=""
		tmp1=split(rs("pcMailUpSett_LastIDList"),"||")
		tmp2=split(rs("pcMailUpSett_LastIDProcess"),"||")
		set rs=nothing
		tmpDidNotStart=0
		For i=lbound(tmp1) to ubound(tmp1)
			if tmp1(i)<>"" then
				tmpIDLists=tmpIDLists & tmp1(i) & "||"
				tmp3=split(tmp2(i),"*")
				if tmp3(1)="0" then
					query="SELECT pcMailUpLists.pcMailUpLists_ListID,pcMailUpLists.pcMailUpLists_ListGuid FROM pcMailUpLists WHERE pcMailUpLists_ListID=" & tmp1(i) & ";"
					set rs=connTemp.execute(query)
					tmpListID=rs("pcMailUpLists_ListID")
					tmpListGuid=rs("pcMailUpLists_ListGuid")
					tmpProcessID=tmp3(0)
					tmpMUResult1=MUStartProcess(session("CP_MU_APIUser"),session("CP_MU_APIPassword"),session("CP_MU_URL"),tmpListID,tmpListGuid,tmpProcessID)
					tmp2(i)=tmp3(0) & "*" & tmpMUResult1
				end if
				tmpIDProList=tmpIDProList & tmp2(i) & "||"
			end if
		Next
		query="UPDATE pcMailUpSettings SET pcMailUpSett_LastIDList='" & tmpIDLists & "',pcMailUpSett_LastIDProcess='" & tmpIDProList & "';"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	end if
	set rs=nothing
	call closedb()
END IF
IF request("action")="uncheck" THEN
	call opendb()
	query="UPDATE pcMailUpSettings SET pcMailUpSett_LastIDList='',pcMailUpSett_LastIDProcess='';"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	call closedb()
	response.redirect "mu_regsyn.asp"
END IF
'End of Post Back
%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%tmpNeedSync=0%>
<table class="pcCPcontent">
	<tr>
		<th>List ID</th>
		<th>List Name</th>
		<th>ID Process</th>
		<th>Import Status</th>
		<th>Confirmation E-mails</th>
	</tr>
	<tr>
		<td colspan="5" class="pcCPspacer"></td>
	</tr>
	<%
	call opendb()
	query="SELECT pcMailUpSett_LastIDList,pcMailUpSett_LastIDProcess FROM pcMailUpSettings WHERE pcMailUpSettings.pcMailUpSett_LastIDProcess<>'';"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmp1=split(rs("pcMailUpSett_LastIDList"),"||")
		tmp2=split(rs("pcMailUpSett_LastIDProcess"),"||")
		set rs=nothing%>
		<%
		tmpDidNotStart=0
		For i=lbound(tmp1) to ubound(tmp1)
			if tmp1(i)<>"" then
				tmp3=split(tmp2(i),"*")
				query="SELECT pcMailUpLists.pcMailUpLists_ListID,pcMailUpLists.pcMailUpLists_ListGuid,pcMailUpLists.pcMailUpLists_ListName FROM pcMailUpLists WHERE pcMailUpLists_ListID=" & tmp1(i) & ";"
				set rs=connTemp.execute(query)
				tmpListID=rs("pcMailUpLists_ListID")
				tmpListGuid=rs("pcMailUpLists_ListGuid")
				tmpListName=rs("pcMailUpLists_ListName")
				tmpProcessID=tmp3(0)
				set rs=nothing
				tmpMUResult=""
				if clng(tmpProcessID)>0 then
					tmpMUResult=MUGetIMStatus(session("CP_MU_APIUser"),session("CP_MU_APIPassword"),session("CP_MU_URL"),tmpListID,tmpListGuid,tmpProcessID)
				else
					Select Case tmp3(1)
						Case "-450": tmpMUResult="Error: listIDs, listGUIDs and groupsIDs don't have the same number of elements. The number must match up."
						Case "-410": tmpMUResult="Error: could not create subscription confirmation email"  
						Case "-400": tmpMUResult="Error: unrecognized error"
						Case "-401": tmpMUResult="Error: xmlDoc is empty"
						Case "-402": tmpMUResult="Error: conversion of xml to csv failed"
						Case "-403": tmpMUResult="Error: creation of new import process failed"
						Case "-100": tmpMUResult="Error: unrecognized error"
						Case "-101": tmpMUResult="Error: verification failed"
						Case "-102": tmpMUResult="Error: List Guid format is not valid"
						Case Else: tmpMUResult="Error Code: " & tmpCode
					End Select
					tmpMUResult="<td colspan=""2""><b>The import process could not be completed successfully. " & tmpMUResult & "</b></td>"
				end if
				%>
				<tr>
					<td><%=tmpListID%></td>
					<td><%=tmpListName%></td>
					<td><%=tmpProcessID%></td>
					<%=tmpMUResult%>
				</tr>
			<%end if
		Next%>
		<tr>
			<td colspan="5" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="5">
				<form class="pcForms">
					<input type="button" name="refreshBtn" value=" Refresh " onclick="javascript:pcf_Open_MailUp(); location='mu_history.asp';" class="submit2">
				&nbsp;<input type="button" name="removeBtn" value=" Remove history " onclick="location='mu_history.asp?action=uncheck';">
				&nbsp;<input type="button" name="Back" value="Back" onClick="location='mu_regsyn.asp';">
				</form>
			</td>
		</tr>
	<%else%>
		<tr>
			<td colspan="5">
				<div class="pcCPmessage">
					No records found!
				</div>
				<br /><br />
				<a href="mu_regsyn.asp">Back</a>
			</td>
		</tr>
	<%end if
	set rs=nothing
	call closedb()%>
</table>
<%Response.write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_MailUp_SynNote2"),"MailUp", 300))%>
<!--#include file="AdminFooter.asp"-->