<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin="7*9*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!-- #Include file="../pc/checkdate.asp" -->
<% Dim pageTitle, Section
pageTitle="Edit message"
Section="orders" %>
<!-- #Include File="Adminheader.asp" -->
<%
'Display Settings

FFont=FFType
FSize=2
LColor=Link
AFont=FFont
ASize=FSize
SColor=Mtype
AllowUpload="1"

'on error resume next
Dim rs, connTemp, query
call openDB()

intIdOrder=getUserInput(request("IDOrder"),0)
IDFeedback=getUserInput(request("IDFeedback"),0)

query="SELECT pcComm_IDOrder,pcComm_IDFeedback FROM pcComments WHERE pcComm_IDOrder=" & intIdOrder & " and pcComm_IDFeedback=" & IDFeedback
set rs=connTemp.execute(query)
 
if rs.eof then
 response.redirect "adminviewfeedback.asp?IDOrder=" & intIdOrder & "&IDFeedback=" & IDFeedback & "&r=1&msg=This feedback was not found or you don't have permission to modify it."
end if

'Update feedback
if (request("action")="update") and (request("rewrite")="0") then
	intIdOrder=getUserInput(request("IDOrder"),0)
	strFDesc=getUserInput(request("Description"),0)
	strFDetails=getUserInput(request("Details"),0)
	intFStatus=getUserInput(request("FStatus"),0)
	intFType=getUserInput(request("FType"),0)
	intPriority=getUserInput(request("Priority"),0)
	dtComDate=CheckDateSQL(now())
	
	if scDB="SQL" then
		query="UPDATE pcComments SET pcComm_IDOrder=" & intIdOrder & ",pcComm_EditedDate='" & dtComDate & "',pcComm_FType=" & intFType & ",pcComm_FStatus=" & intFStatus & ",pcComm_Priority=" & intPriority & ",pcComm_Description='" & strFDesc & "',pcComm_Details='" & strFDetails & "' WHERE pcComm_IDOrder=" & intIdOrder & " and pcComm_IDFeedback=" & IDFeedback
	else
		query="UPDATE pcComments SET pcComm_IDOrder=" & intIdOrder & ",pcComm_EditedDate=#" & dtComDate & "#,pcComm_FType=" & intFType & ",pcComm_FStatus=" & intFStatus & ",pcComm_Priority=" & intPriority & ",pcComm_Description='" & strFDesc & "',pcComm_Details='" & strFDetails & "' WHERE pcComm_IDOrder=" & intIdOrder & " and pcComm_IDFeedback=" & IDFeedback
	end if
	set rs=connTemp.execute(query)
	
	query="UPDATE pcComments SET pcComm_IDOrder=" & intIdOrder & " WHERE pcComm_IDOrder=" & intIdOrder & " and pcComm_IDParent=" & IDFeedback
	set rs=connTemp.execute(query)
	
	if AllowUpload="1" then
		ACount=getUserInput(request("ACount"),0)
		if ACount<>"" then
			ACount1=clng(ACount)
			For k=1 to ACount1
				if request("AC" & k)="1" then
					query="UPDATE pcUploadFiles SET pcUpld_IDFeedback=" & IDFeedback & " WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
					set rs=connTemp.execute(query)
				else
					query="SELECT * FROM pcUploadFiles WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0) & " and pcUpld_IDFeedback=" & IDFeedback
					set rs=connTemp.execute(query)
					if not rs.eof then
	 					strFilename=rs("pcUpld_Filename")
	 					if strFilename<>"" then
							QfilePath="../pc/Library/" & strFilename
	   					findit = Server.MapPath(QfilePath)
							Set fso = server.CreateObject("Scripting.FileSystemObject")
							Set f = fso.GetFile(findit)
							f.Delete
							Set fso = nothing
							Set f = nothing
							Err.number=0
							Err.Description=""
	 					end if
   				end if

					query="DELETE FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & IDFeedback & " and pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
					set rs=connTemp.execute(query)
	
				end if
			next
		end if
	end if
	%>
	<center>
	<table width="60%" border="0" cellspacing="0" cellpadding="2" height="8" bgcolor="#FFFFFF" style="border: 2 solid #FF0000">
		<tr> 
			<td width="25" valign="top"> 
			<img src="images/successful.gif"> 
			</td>
			<td width="100%"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=SColor%>">Your feedback was
				updated successfully!</font></td>
		</tr>
	</table>
	</center>
	<br>                  
<%end if%>
<script language="JavaScript">
<!--
	
function Form1_Validator(theForm)
{

<%if session("UserType")=3 then%>
		if (theForm.FType.value == "")
 	{
		    alert("Please select one feedback type.");
		    theForm.FType.focus();
		    return (false);
	}
			if (theForm.Priority.value == "")
 	{
		    alert("Please select one priority.");
		    theForm.Priority.focus();
		    return (false);
	}
				if (theForm.FStatus.value == "")
 	{
		    alert("Please select one feedback status.");
		    theForm.FStatus.focus();
		    return (false);
	}
<%end if%>

			if (theForm.Description.value == "")
 	{
		    alert("Please enter a value for Short Description.");
		    theForm.Description.focus();
		    return (false);
	}
	
			if (theForm.Details.value == "")
 	{
		    alert("Please enter a value for Long Description.");
		    theForm.Details.focus();
		    return (false);
	}
  
return (true);
}
//-->
</script>

<%
query="SELECT pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details FROM pcComments WHERE pcComm_IDFeedback=" & IDFeedback & " and pcComm_IDParent=0 and pcComm_IDOrder=" & intIdOrder
set rs=connTemp.execute(query)
intFType=rs("pcComm_FType")
intFStatus=rs("pcComm_FStatus")
intPriority=rs("pcComm_Priority")
strDesc=rs("pcComm_Description")
strDetails=rs("pcComm_Details")
%>
 
<form name="hForm" method="post" action="admineditFeedback.asp?action=update" onSubmit="return Form1_Validator(this)">
<script language="JavaScript"><!--
function newWindow(file,window) {
		msgWindow=open(file,window,'resizable=no,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
}
//--></script>
<input type=hidden name=IDOrder value="<%=intIdOrder%>">
<input type=hidden name=IDFeedback value="<%=IDFeedback%>">
<div align="center">
<table width="600" bgcolor="#666666" cellpadding="1" cellspacing="0">
	<tr>
		<td>
			<table width="100%" bgcolor="#FFFFFF" cellpadding="5" cellspacing="0">
				<tr bgcolor="#e5e5e5"><td colspan="2" align="left"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>"><b>Edit Feedback</b></font></td>
				</tr>
				<tr><td colspan="2" align="left">&nbsp;</td>
				</tr>
				<tr>
					<td width="25%" align="right"> <font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>"> Order #:</font></td>
					<td width="75%">
						<font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">
						<b><%=clng(scpre)+clng(intIdOrder)%></b>
						</font>
					</td>
				</tr>
				<tr>
					<td align="right"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">Feedback Type:</font></td>
					<td><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">
					<select name="FType">
          	<option value=""></option>
						<%query="SELECT pcFType_idtype,pcFType_idtype,pcFType_name FROM pcFTypes"
						set rs=connTemp.execute(query)
						do while not rs.eof%>
   						<option value="<%=rs("pcFType_idtype")%>" <%if request("FType")<>"" then%><%if clng(request("FType"))=clng(rs("pcFType_idtype")) then%>selected<%end if%><%else%><% if rs("pcFType_idtype")=intFType then%>selected<%end if%><%end if%> ><%=rs("pcFType_name")%></option>
   						<%rs.MoveNext
   					Loop%>
					</select></font></td>
  			</tr>
				<tr>
					<td width="25%" align="right"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">Priority:</font></td>
					<td width="75%">
					<select name="Priority">
          	<option value=""></option>
    				<% query="SELECT pcPri_idPri,pcPri_name FROM pcPriority"
						set rs=connTemp.execute(query)
						do while not rs.eof %>
   						<option value="<%=rs("pcPri_idPri")%>" <%if request("Priority")<>"" then%><%if clng(request("Priority"))=clng(rs("pcPri_idPri")) then%>selected<%end if%><%else%><% if rs("pcPri_idPri")=intPriority then%>selected<%end if%><%end if%>><%=rs("pcPri_name")%></option>
   						<%rs.MoveNext
   					Loop%>
						</select> 
						</td>
					</tr>
				<tr>
   				<td width="25%" align="right"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">Status:</font></td>
    			<td width="75%">
					<select name="FStatus">
          	<option value=""></option>
    				<% query="SELECT pcFStat_idStatus,pcFStat_name FROM pcFStatus"
  					set rs=connTemp.execute(query)
   					do while not rs.eof %>
   						<option value="<%=rs("pcFStat_idStatus")%>" <%if request("FStatus")<>"" then%><%if clng(request("FStatus"))=clng(rs("pcFStat_idStatus")) then%>selected<%end if%><%else%><% if rs("pcFStat_idStatus")=intFStatus then%>selected<%end if%><%end if%>><%=rs("pcFStat_name")%></option>
   						<%rs.MoveNext
   					Loop%>
        	</select> 
     	 		</td>
  			</tr>
				<tr>
					<td width="25%" align="right"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">Short Description:</font></td>
					<td width="75%"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>"><input name="Description" type="text" value="<%if request("Description")<>"" then%><%=request("Description")%><%else%><%=strDesc%><%end if%>" size="25" maxlength="100"> 
						</font>
						</td>
				</tr>
				<tr>
					<td width="25%" align="right" valign="top"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">Long Description:</font><br><br>
					<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=Details','window2')" style="font-family: <%=FFont%>; font-size: 8pt; color: #000000; border: 1px solid gray"></td>
					<td width="75%"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>"><textarea name="Details" cols="40" rows="7" id="bugLongDsc"><%if request("Details")<>"" then%><%=request("Details")%><%else%><%=strDetails%><%end if%></textarea></font><br>
					</td>
				</tr>
  			<%if AllowUpload="1" then%>
					<tr><td nowrap width="25%" valign="top">
    				<p align="right"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">Attachment(s):</font></p>
  					</td>
						<td width="75%" valign="top">
						<%query="SELECT pcUpld_IDFile,pcUpld_FileName FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & IDFeedback
						set rs=connTemp.execute(query)
						if rs.eof then%>
							<font face="<%=FFont%>" size="<%=FSize%>" color="<%=SColor%>">No attached files.</font><br>
						<%else
							ACount=0
							do while not rs.eof
								ACount=ACount+1%>
								<input type=hidden name="AID<%=ACount%>" value="<%=rs("pcUpld_IDFile")%>"><input type=checkbox name="AC<%=ACount%>" value="1" checked>&nbsp;<font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>"><%
								strFilename= rs("pcUpld_FileName")
								strFilename = mid(strFilename,instr(strFilename,"_")+1,len(strFilename))%>
								<%=strFilename%></font><br>
								<%rs.MoveNext
							loop%>
							<input type=hidden name=ACount value="<%=ACount%>">
						<%end if%>
						<script language="JavaScript"><!--
							function newWindow1(file,window) {
							catWindow=open(file,window,'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360');
							if (catWindow.opener == null) catWindow.opener = self;
							}
						//--></script>
						<br><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">To upload file(s) <a href="#" onclick="javascript:newWindow1('adminfileuploada_popup.asp?IDFeedback=<%=IDFeedback%>&ReLink=<%=Server.URLencode("admineditfeedback.asp?IDOrder=" & intIdOrder & "&IDFeedback=" & IDFeedback)%>','window2')"><font face="<%=FFont%>" size="1" color="<%=LColor%>">click here</font></a></font></td>
					</tr>
				<%end if%>
				<tr>
					<td width="25%" align="right"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">&nbsp;</font></td>
					<td width="75%"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">&nbsp;</font></td>
				</tr>
				<tr>
					<td width="25%" align="right"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>">&nbsp;</font></td>
					<td width="75%"><font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>"><input type="submit" name="Submit" value=" Update " class="submit2" onclick="document.hForm.rewrite.value='0';">&nbsp;<input type="button" value="Back" onClick="location='adminviewfeedback.asp?IDOrder=<%=intIdOrder%>&IDFeedback=<%=IDFeedback%>'" class="ibtnGrey">&nbsp;
						<input type="button" name="back" value=" View all Postings " onClick="location='adminviewallmsgs.asp';" class="ibtnGrey">
					 <%if session("admin_IDOrder")>0 then%><input type="button" name="go" value=" View Postings " onClick="location='adminviewallmsgs.asp?IDOrder=<%=session("admin_IDOrder")%>';" class="ibtnGrey"><%end if%></font>
					 <input type="hidden" name="uploaded" value="">
						<input type="hidden" name="rewrite" value="1">
					 </td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</div>
</form>
<%call closedb()%><!-- #Include File="adminfooter.asp" -->