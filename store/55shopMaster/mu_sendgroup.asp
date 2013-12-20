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
<%
Server.ScriptTimeout = 5400
dim rstemp, conntemp, query
Dim tmp_setup

tmp_setup=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	call opendb()
	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess FROM pcMailUpSettings;"
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
	end if
	set rs=nothing
call closedb()

if tmp_setup=0 OR session("CP_NW_ListID")="" OR IsNull(session("AddrList")) then
	response.redirect "mu_newsWizStep1.asp"
end if

msg=""

IF Request("action")="run" THEN

	tmpXMLDoc=""
	conFirmEmail=request("confirm1")
	if conFirmEmail="" then
		conFirmEmail=0
	end if
	AddrList=session("AddrList")
	call opendb()
	For k=lbound(AddrList) to ubound(AddrList)
		if trim(AddrList(k))<>"" then
		'***************
		IF session("CP_NW_PageType")="0" THEN
			query="SELECT pcSupplier_FirstName,pcSupplier_Lastname,pcSupplier_Company FROM pcSuppliers WHERE pcSupplier_Email like '" & trim(AddrList(k)) & "';"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & AddrList(k) & """ Prefix="""" Number="""">"
					if (rsQ("pcSupplier_FirstName")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo1>" & Server.HTMLEncode(rsQ("pcSupplier_FirstName")) & "</campo1>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo1></campo1>"
					end if
					if (rsQ("pcSupplier_Lastname")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo2>" & Server.HTMLEncode(rsQ("pcSupplier_Lastname")) & "</campo2>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo2></campo2>"
					end if
					if (rsQ("pcSupplier_Company")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo3>" & Server.HTMLEncode(rsQ("pcSupplier_Company")) & "</campo3>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo3></campo3>"
					end if
					tmpXMLDoc=tmpXMLDoc & "</subscriber>"
			else
				tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & AddrList(k) & """ />"
			end if
			set rsQ=nothing
		ELSE
		IF session("CP_NW_PageType")="1" THEN
			query="SELECT pcDropShipper_FirstName,pcDropShipper_Lastname,pcDropShipper_Company FROM pcDropShippers WHERE pcDropShipper_Email like '" & trim(AddrList(k)) & "';"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & AddrList(k) & """ Prefix="""" Number="""">"
					if (rsQ("pcDropShipper_FirstName")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo1>" & Server.HTMLEncode(rsQ("pcDropShipper_FirstName")) & "</campo1>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo1></campo1>"
					end if
					if (rsQ("pcDropShipper_Lastname")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo2>" & Server.HTMLEncode(rsQ("pcDropShipper_Lastname")) & "</campo2>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo2></campo2>"
					end if
					if (rsQ("pcDropShipper_Company")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo3>" & Server.HTMLEncode(rsQ("pcDropShipper_Company")) & "</campo3>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo3></campo3>"
					end if
					tmpXMLDoc=tmpXMLDoc & "</subscriber>"
			else
			query="SELECT pcSupplier_FirstName,pcSupplier_Lastname,pcSupplier_Company FROM pcSuppliers WHERE pcSupplier_Email like '" & trim(AddrList(k)) & "';"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & AddrList(k) & """ Prefix="""" Number="""">"
					if (rsQ("pcSupplier_FirstName")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo1>" & Server.HTMLEncode(rsQ("pcSupplier_FirstName")) & "</campo1>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo1></campo1>"
					end if
					if (rsQ("pcSupplier_Lastname")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo2>" & Server.HTMLEncode(rsQ("pcSupplier_Lastname")) & "</campo2>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo2></campo2>"
					end if
					if (rsQ("pcSupplier_Company")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo3>" & Server.HTMLEncode(rsQ("pcSupplier_Company")) & "</campo3>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo3></campo3>"
					end if
					tmpXMLDoc=tmpXMLDoc & "</subscriber>"
			else
				tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & AddrList(k) & """ />"
			end if
			end if
			set rsQ=nothing
		ELSE
			query="SELECT [name],lastName,customerCompany FROM customers WHERE email like '" & trim(AddrList(k)) & "';"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & AddrList(k) & """ Prefix="""" Number="""">"
					if (rsQ("name")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo1>" & Server.HTMLEncode(rsQ("name")) & "</campo1>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo1></campo1>"
					end if
					if (rsQ("lastname")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo2>" & Server.HTMLEncode(rsQ("lastname")) & "</campo2>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo2></campo2>"
					end if
					if (rsQ("customerCompany")<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo3>" & Server.HTMLEncode(rsQ("customerCompany")) & "</campo3>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo3></campo3>"
					end if
					tmpXMLDoc=tmpXMLDoc & "</subscriber>"
			else
				tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & AddrList(k) & """ />"
			end if
			set rsQ=nothing
		END IF
		END IF
		'***************
		end if
	Next
	call closedb()
	tmpXMLDoc="<subscribers>" & tmpXMLDoc & "</subscribers>"
	call opendb()
	query="SELECT pcMailUpLists_ID,pcMailUpLists_ListID,pcMailUpLists_ListGuid FROM pcMailUpLists WHERE pcMailUpLists_ID=" & session("CP_NW_ListID") & ";"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		listID=rs("pcMailUpLists_ListID")
		listGuid=rs("pcMailUpLists_ListGuid")
	end if
	set rs=nothing
	call closedb()
	
	GroupID=getUserInput(request("EGroupID"),0)
	
	if GroupID="" then
		GroupName=getUserInput(request("NewGroupName"),0)
		tmpResult=MUCreateGroup(session("CP_MU_APIUser"),session("CP_MU_APIPassword"),session("CP_MU_URL"),listID,listGuid,GroupName)
		if tmpResult>"0" then
			GroupID=tmpResult
			call opendb()
			query="INSERT INTO pcMailUpGroups (pcMailUpLists_ID,pcMailUpGroups_GroupID,pcMailUpGroups_GroupName) VALUES (" & session("CP_NW_ListID") & "," & GroupID & ",'" & replace(GroupName,"'","''") & "');"
			set rs=connTemp.execute(query)
			set rs=nothing
			call closedb()
		else
			msg="2"
		end if
	end if
	tmpCanNotStart=0
	if msg<>"2" and GroupID<>"" then
		tmpResult=MUImport(session("CP_MU_APIUser"),session("CP_MU_APIPassword"),session("CP_MU_URL"),listID,listGuid,tmpXMLDoc,GroupID,0,0,conFirmEmail)
		if tmpResult="0" then
			msg="3"
		else
			tmpReturnIDs=session("CP_MU_ReturnIDs")
			tmpReturnList=session("CP_MU_ReturnList")
			tmpReturnCode=session("CP_MU_ReturnCode")
			if tmpReturnIDs="0" then
				msg="3"
			else
				msg="1"
			end if
			call opendb()
			query="UPDATE pcMailUpSettings SET pcMailUpSett_LastIDList='" & listID & "||',pcMailUpSett_LastIDProcess='" & tmpReturnIDs & "*" & tmpReturnCode & "||';"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
			call closedb()
		end if
	end if

END IF
%>
<% pageTitle="Newsletter Wizard - STEP 3: Export Recipients to MailUp" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<form name="form1" method="post" action="mu_sendgroup.asp?action=run" class="pcForms" onsubmit="javascript:pcf_Open_MailUp();">
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<%if msg<>"" then%>
<tr>
	<td colspan="2">
		
			<%Select Case msg
			Case "1":%>
				<div class="pcCPmessageSuccess">E-mail addresses successfully exported to your MailUp console.
				<%if tmpCanNotStart>"0" then%>
				<br />However, the import process <u>has not started yet</u>. <a href="mu_regsyn.asp?action=restart">Try to re-start it now</a>.
				<%end if%>
                </div>
			<%Case "2":%>
            	<div class="pcCPmessage">Cannot create new group. <%if MU_ErrMsg<>"" then%>&nbsp;Error Message: <%=MU_ErrMsg%><%end if%></div>
			<%Case "3":%>
				<div class="pcCPmessage">Cannot export e-mail addresses to MailUp.<%if MU_ErrMsg<>"" then%>&nbsp;Error Message: <%=MU_ErrMsg%><%end if%></div>
			<%End Select%>
		</div>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<%end if%>
<%IF msg<>"1" THEN%>
	<tr>
		<td colspan="2">E-mail addresses are exported to your MailUp console as a new Group or added to an existing Group in the List you selected. When you log into your MailUp console to send your message, you will be able to easily filter recipients by Group.</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<%
	call opendb()
	query="SELECT pcMailUpGroups.pcMailUpGroups_GroupID,pcMailUpGroups.pcMailUpGroups_GroupName FROM pcMailUpGroups INNER JOIN pcMailUpLists ON pcMailUpGroups.pcMailUpLists_ID=pcMailUpLists.pcMailUpLists_ID WHERE pcMailUpLists.pcMailUpLists_ID=" & session("CP_NW_ListID") & " AND (pcMailUpGroups_GroupName<>'TEST' AND pcMailUpGroups_GroupName<>'BOUNCE') ORDER BY pcMailUpGroups.pcMailUpGroups_GroupName;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpArr=rs.getRows()
		intCount=ubound(tmpArr,2)
	%>
	<tr>
		<th colspan="2">Export to an Existing Group</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td valign="top">Select an existing group:</td>
		<td valign="top">
			<select name="EGroupID">
				<option value="" selected></option>
				<%For i=0 to intCount%>
					<option value="<%=tmpArr(0,i)%>"><%=tmpArr(1,i)%></option>
				<%Next%>
			</select>
			<br /><br />
		</td>
	</tr>
<%end if
set rs=nothing%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Export as a New Group</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td valign="top">New Group Name:</td>
		<td valign="top">
			<input type="text" name="NewGroupName" size="60" value="">
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><input type="hidden" name="confirm1" value="0"></td>
	</tr>
	<tr>
		<td colspan="2">
			<input type="submit" name="submit1" value="Export to MailUp" class="submit2">&nbsp;&nbsp;<input type="button" name="back" value="Back"  onclick="location='mu_newsWizStep2.asp';">
		</td>
	</tr>
<%ELSE%>
	<tr>
		<td colspan="2">
			<input type="button" name="back" value="Back to Newsletter Wizard" onclick="location='mu_newsWizStep1.asp';">
		</td>
	</tr>
<%END IF%>
</table>
</form>
<%Response.write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_MailUp_SynNote2"),"MailUp", 300))%>
<%call closedb()%><!--#include file="AdminFooter.asp"-->