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
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!-- #Include file="../pc/checkdate.asp" -->
<% Dim pageTitle, Section
pageTitle="Edit comment"
Section="orders" %>
<!-- #Include File="Adminheader.asp" -->
<%

'Display Settings

FFont="Arial"
FSize=2
LColor=Link
AFont=FFont
ASize=FSize
SColor=Mtype
AllowUpload="1"

'on error resume next
Dim rs, connTemp, query
call openDB()

IDOrder=getUserInput(request("IDOrder"),0)
IDFeedback=getUserInput(request("IDFeedback"),0)
IDComment=getUserInput(request("IDComment"),0)

query="SELECT * FROM pcComments WHERE pcComm_IDFeedback=" & IDComment & " and pcComm_IDParent=" & IDFeedback & " and pcComm_IDOrder=" & IDOrder 
set rs=connTemp.execute(query)
if rs.eof then
	response.redirect "adminviewfeedback.asp?IDOrder=" & IDOrder & "&IDFeedback=" & IDFeedback & "&r=1&msg=This comment was not found or you don't have permission to modify it."
end if

'Create new feedback
if (request("action")="update") and (request("rewrite")="0") then
	strFDetails=getUserInput(request("Details"),0)
	dtComDate=CheckDateSQL(now())
	
	if scDB="SQL" then
		query="UPDATE pcComments SET pcComm_EditedDate='" & dtComDate & "', pcComm_Details='" & strFDetails & "' WHERE pcComm_IDFeedback=" & IDComment & ";"
	else
		query="UPDATE pcComments SET pcComm_EditedDate=#" & dtComDate & "#, pcComm_Details='" & strFDetails & "' WHERE pcComm_IDFeedback=" & IDComment & ";"
	end if
	
	set rs=connTemp.execute(query)
	
	if AllowUpload="1" then
		ACount=getUserInput(request("ACount"),0)
		if ACount<>"" then
			ACount1=clng(ACount)
			For k=1 to ACount1
				if request("AC" & k)="1" then
					query="UPDATE pcUploadFiles set pcUpld_IDFeedback=" & IDComment & " WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
					set rs=connTemp.execute(query)
				else
					query="SELECT pcUpld_Filename FROM pcUploadFiles WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0)&";"
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

					query="DELETE FROM pcUploadFiles WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
					set rs=connTemp.execute(query)
				end if
			next
		end if
	end if
	
	msg="Comment updated successfully."
	msgtype=1
	%>
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>

<%end if%>
<script language="JavaScript">
<!--
	
function Form1_Validator(theForm)
{

			if (theForm.Details.value == "")
 	{
		    alert("Please enter a value for Comment Details.");
		    theForm.Details.focus();
		    return (false);
	}
  
return (true);
}
//-->
</script>
<% query="SELECT pcComm_Details FROM pcComments WHERE pcComm_IDFeedback=" & IDComment & ";"
set rs=connTemp.execute(query)
strDetails=rs("pcComm_Details")
%>
<form name="hForm" method="post" action="admineditComment.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
<script language="JavaScript"><!--
function newWindow(file,window) {
		msgWindow=open(file,window,'resizable=no,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
}
//--></script>
<input type=hidden name=Priority value="<%=Priority%>">
<input type=hidden name=FStatus value="<%=FStatus%>">
<input type=hidden name=FType value="<%=FType%>">
<input type=hidden name=IDOrder value="<%=IDOrder%>">
<input type=hidden name=IDFeedback value="<%=IDFeedback%>">
<input type=hidden name=IDComment value="<%=IDComment%>">
<div align="center">
			<table class="pcCPcontent" style="width: 600px;">
				<tr>
					<td width="25%" align="right">Order #:</td>
					<td width="75%"><b><%=clng(scpre)+clng(IDOrder)%></b></td>
				</tr>
				<tr>
					<td width="25%" align="right">Feedback:</td>
					<td width="75%">
					<% query="SELECT pcComm_Description FROM pcComments WHERE pcComm_IDFeedback=" & IDFeedback &";"
					set rs=connTemp.execute(query)%>
					<%=rs("pcComm_Description")%>
					</td>
				</tr>
				<tr>
					<td width="25%" align="right" valign="top">Comment:<br><br>
					<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=Details','window2')" style="font-family: <%=FFont%>; font-size: 8pt; color: #000000; border: 1px solid gray"></td>
					<td width="75%"><textarea name="Details" cols="40" rows="7" id="bugLongDsc"><%if request("Details")<>"" then%><%=request("Details")%><%else%><%=strDetails%><%end if%></textarea><br>
					</td>
				</tr>
  			<%if AllowUpload="1" then%>
					<tr>
                    	<td nowrap width="25%" valign="top" align="right">Attachment(s):</td>
						<td width="75%" valign="top">
						<%query="SELECT pcUpld_IDFile,pcUpld_FileName FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & IDComment
						set rs=connTemp.execute(query)
						if rs.eof then%>
							No attached files.<br>
						<%else
							ACount=0
							do while not rs.eof
								pc_pcUpld_IDFile=rs("pcUpld_IDFile")
								pc_pcUpld_FileName=rs("pcUpld_FileName")
								
								ACount=ACount+1 %>
								<input type=hidden name="AID<%=ACount%>" value="<%=pc_pcUpld_IDFile%>">
								<input type=checkbox name="AC<%=ACount%>" value="1" checked>&nbsp;<font face="<%=FFont%>" size="<%=FSize%>" color="<%=FColor%>"><%
								strFilename= pc_pcUpld_FileName
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
						<br>To upload file(s) <a href="#" onclick="javascript:newWindow1('adminfileuploada_popup.asp?IDFeedback=<%=IDComment%>&ReLink=<%=Server.URLencode("admineditcomment.asp?IDComment=" & IDComment & "&IDOrder=" & IDOrder & "&IDFeedback=" & IDFeedback)%>','window2')">click here</a></td></tr>
					<%end if%>
					<tr>
						<td colspan="2"><hr></td>
					</tr>
					<tr>
						<td width="25%" align="right"></td>
						<td width="75%"><input type="submit" name="Submit" value=" Update " class="submit2" onclick="document.hForm.rewrite.value='0';">&nbsp;<input type="button" value="Back" onClick="location='adminviewfeedback.asp?IDOrder=<%=IDOrder%>&IDFeedback=<%=IDFeedback%>'">&nbsp;
							<input type="button" name="back" value=" View all Postings " onClick="location='adminviewallmsgs.asp';">
					 	<%if session("admin_IDOrder")>0 then%><input type="button" name="go" value=" View Postings " onClick="location='adminviewallmsgs.asp?IDOrder=<%=session("admin_IDOrder")%>';"><%end if%></font>
					 	<input type="hidden" name="uploaded" value="">
						<input type="hidden" name="rewrite" value="1">
					 </td>
					</tr>
				</table>
</div>
</form>
<%call closedb()%><!-- #Include File="Adminfooter.asp" -->