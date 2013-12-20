<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=0%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/MailUpFunctions.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp" -->
<% pageTitle="MailUp Integration Settings" %>
<% section="mngAcc"
Dim connTemp,query,rs
Dim tmp_setup

msg=0

if request("action")="upd" then
call opendb()
	tmpAuto=request("auto1")
	if tmpAuto="" then
		tmpAuto=0
	end if
	query="UPDATE pcMailUpSettings SET pcMailUpSett_AutoReg=" & tmpAuto & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
	if request("submit1")<>"" then
		intCount=request("count")
		if intCount<>"" then
			intCount=intCount+1
			For j=1 to intCount
				if request("list"&j)<>"" then
					tmpIDList=request("list"&j)
					tmpListName=request("listname"&j)
					if tmpListName<>"" then
						tmpListName=replace(tmpListName,"'","''")
					end if
					tmpListDesc=request("listdesc"&j)
					if tmpListDesc<>"" then
						tmpListDesc=replace(tmpListDesc,"'","''")
					end if
					tmpActive=getUserInput(request("active"&j),0)
					if tmpActive="" then
						tmpActive=0
					end if
					query="UPDATE pcMailUpLists SET pcMailUpLists_ListName='" & tmpListName & "',pcMailUpLists_ListDesc='" & tmpListDesc & "',pcMailUpLists_Active=" & tmpActive & " WHERE pcMailUpLists_ID=" & tmpIDList & ";"
					set rs=connTemp.execute(query)
					set rs=nothing
				end if
			Next
			msg=1
		end if
	end if
	if request("submit2")<>"" then
		intCount=request("rcount")
		if intCount<>"" then
			intCount=intCount+1
			For j=1 to intCount
				if request("rlist"&j)<>"" then
					tmpIDList=request("rlist"&j)
					query="DELETE FROM pcMailUpSubs WHERE pcMailUpLists_ID=" & tmpIDList & ";"
					set rs=connTemp.execute(query)
					set rs=nothing
					query="DELETE FROM pcMailUpLists WHERE pcMailUpLists_ID=" & tmpIDList & ";"
					set rs=connTemp.execute(query)
					set rs=nothing
				end if
			Next
			msg=2
		end if
	end if
call closedb()
end if

	tmp_setup=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	call opendb()
	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess FROM pcMailUpSettings;"
	set rs=connTemp.execute(query)
		if err.number<>0 then
			set rs = nothing
			call closedb()
			response.Redirect("upddb_MailUp.asp")
		end if
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
	
	if tmp_setup="0" then
		response.redirect "mu_manageNewsWiz.asp"
	end if
	
%>
<!--#include file="AdminHeader.asp"-->
<%
	if request("action")<>"upd" then
		%>
		<%
		call opendb()
		tmpGetList=GetMUList(session("CP_MU_APIUser"),session("CP_MU_APIPassword"),session("CP_MU_URL"))
		if tmpGetList="0" then
			msg=3
		end if
		call closedb()%>
		<%
	end if
%>
<form name="form1" action="mu_settings.asp?action=upd" method="post" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<%if msg<>0 then%>
	<tr>
		<td colspan="2">
				<%Select Case msg
				Case 1:%>
				<div class="pcCPmessageSuccess">MailUp List information was updated successfully!</div>
				<%Case 2:%>
				<div class="pcCPmessageSuccess">MailUp Lists were removed successfully!</div>
				<%Case 3:%>
				<div class="pcCPmessage">Can not get MailUp lists from server.<br>We have loaded the lists that were last saved in the database.<br>Server Error Message: <b><%=MU_ErrMsg%></div>
				<%End Select%>
		</td>
	</tr>
	<%end if%>
	<tr>
		<th colspan="2">MailUp Integration Settings</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="auto1" value="1" class="clearBorder" <%if tmp_Auto="1" then%>checked<%end if%>>
		&nbsp;Automatic Customer Registration
		<div style="margin-top: 6px">When you enable this setting, ProductCart will contact your MailUp console to update a customer's e-mail preferences when the customer registers with the store or updates his/her profile.</div>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<%call opendb()
	query="SELECT pcMailUpLists_ID,pcMailUpLists_ListID,pcMailUpLists_ListGuid,pcMailUpLists_ListName,pcMailUpLists_ListDesc,pcMailUpLists_Active,pcMailUpLists_Removed FROM pcMailUpLists ORDER BY pcMailUpLists_Active DESC, pcMailUpLists_ListID ASC;"
	set rs=connTemp.execute(query)
	intCount=-1
	if not rs.eof then
		pcArr=rs.getRows()
		intCount=ubound(pcArr,2)
	end if
	set rs=nothing
	pcv_HaveLists=0
	pcv_HaveRemoved=0
	call closedb()
	%>
	<tr>
		<th colspan="2">Manage Lists</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<%if intCount>=0 then%>
	<tr>
		<td colspan="2">
			<table class="pcCPcontent">
				<tr style="border-bottom: 1px solid #CCCCCC;">
					<td><strong>ID</strong></td>
					<td><strong>DETAILS</strong></td>
					<td><input type="checkbox" name="C1" value="1" onclick="javascript:checkAll(this.checked);" class="clearBorder"><b>ACTIVE</b></td>
				</tr>
				<%For i=0 to intCount
					if pcArr(6,i)<>"1" then
					If strCol <> "#FFFFFF" Then
						strCol = "#FFFFFF"
					Else 
						strCol = "#E1E1E1"
					End If
					pcv_HaveLists=pcv_HaveLists+1%>
					<tr><td colspan="3">&nbsp;</td></tr>
					<tr style="background-color: <%=strCol%>; border: 1px dashed #CCCCCC;">
						<td nowrap style="padding: 5px; vertical-align: top">
							<input type="hidden" name="list<%=pcv_HaveLists%>" value="<%=pcArr(0,i)%>">
							<%=pcArr(1,i)%>
						</td>
						<td nowrap="nowrap" style="vertical-align: top; padding: 5px;">
							<div>
								<strong>Name</strong>: <input type="text" name="listname<%=pcv_HaveLists%>" value="<%=pcArr(3,i)%>" size="60">
							</div>
							<div style="padding-top: 6px; vertical-align: top;">
								<strong>Description</strong>:<br />
								<textarea name="listdesc<%=pcv_HaveLists%>" cols="80" rows="3"><%=pcArr(4,i)%></textarea>
							</div>
							<div style="padding-top: 6px; padding-bottom: 10px;">
							  List GUID: <%=pcArr(2,i)%>
							</div>
						</td>
						<td style="vertical-align: top; padding: 5px;"><input type="checkbox" name="active<%=pcv_HaveLists%>" value="1" <%if pcArr(5,i)="1" then%>checked<%end if%> class="clearBorder"></td>
					</tr>
					<%else
						pcv_HaveRemoved=pcv_HaveRemoved+1
					end if%>
				<%Next%>
				<%if pcv_HaveLists=0 then%>
				<tr>
					<td colspan="3"><div class="pcCPmessage">No Lists found</div></td>
				</tr>
				<%end if%>
			</table>
		</td>
	<tr>
	<%if pcv_HaveLists>0 then%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colpan="2">
			<script language="JavaScript">
				function checkAll(tmpvalue) {
				for (var j = 1; j <= <%=pcv_HaveLists%>; j++) {
				box = eval("document.form1.active" + j); 
				box.checked = tmpvalue;
				   }
				}
			</script>
			<input type="hidden" name="count" value="<%=intCount%>">
			<input type="submit" name="submit1" class="submit2" value="Update">
		</td>
	</tr>
	<%end if%>
	<%else%>
	<tr>
		<td colspan="2"><div class="pcCPmessage">No Lists found</div></td>
	</tr>
	<%end if%>
	<%if pcv_HaveRemoved>0 then%>
	<tr>
		<th colspan="2">Removed Lists</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">
			<p>When synchronizing with your MailUp console, we found that the lists below no longer exist in your MailUp account.</p>
			<br><br>
			<table class="pcCPcontent">
				<tr style="border-bottom: 1px solid #CCCCCC;">
					<td><strong>ID</strong></td>
					<td><strong>DETAILS</strong></td>
					<td><b>ACTIVE</b></td>
				</tr>
				<%pcv_HaveRemoved=0
				strCol = "#E1E1E1"
				For i=0 to intCount
					if pcArr(6,i)="1" then
					If strCol <> "#FFFFFF" Then
						strCol = "#FFFFFF"
					Else 
						strCol = "#E1E1E1"
					End If
					pcv_HaveRemoved=pcv_HaveRemoved+1%>			
					
					<tr><td colspan="3">&nbsp;</td></tr>
					<tr style="background-color: <%=strCol%>; border: 1px dashed #CCCCCC;">
						<td nowrap style="padding: 5px; vertical-align: top">
							<input type="hidden" name="rlist<%=pcv_HaveRemoved%>" value="<%=pcArr(0,i)%>">
							<%=pcArr(1,i)%>
						</td>
						<td nowrap="nowrap" style="vertical-align: top; padding: 5px;">
							<div>
								<strong>Name</strong>: <input type="text" name="rlistname<%=pcv_HaveRemoved%>" value="<%=pcArr(3,i)%>" size="60">
							</div>
							<div style="padding-top: 6px; vertical-align: top;">
								<strong>Description</strong>:<br />
								<textarea name="rlistdesc<%=pcv_HaveRemoved%>" cols="80" rows="3"><%=pcArr(4,i)%></textarea>
							</div>
							<div style="padding-top: 6px; padding-bottom: 10px;">
							  List GUID: <%=pcArr(2,i)%>
							</div>
						</td>
						<td style="vertical-align: top; padding: 5px;"><input type="checkbox" name="ractive<%=pcv_HaveRemoved%>" value="1" <%if pcArr(5,i)="1" then%>checked<%end if%> class="clearBorder"></td>
					</tr>
					<%end if%>
				<%Next%>
			</table>
		</td>
	<tr>
		<td colpan="2">
			<input type="submit" name="submit2" class="submit2" value="Remove" onclick="javascript:if (confirm('You are about to remove the MailUp lists from your database. It also remove all opted-in customers from these lists. Are you sure you want to complete this action?')) {return(true);} else {return(false);}">
			<input type="hidden" name="rcount" value="<%=intCount%>">
		</td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2">
			<input type="button" name="Back" value="Back to MailUp Management" onClick="location='mu_manageNewsWiz.asp';" class="ibtnGrey">
		</td>
	</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->