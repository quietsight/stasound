<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="7*9*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!-- #Include file="../pc/checkdate.asp" -->
<% 
Dim pageTitle, Section
pageTitle="View &amp; Edit Help Desk Message"
Section="orders"

Dim rs, connTemp, query

call openDB() 

intIdOrder=getUserInput(request("IDOrder"),0)
intIdFeedback=getUserInput(request("IDFeedback"),0)

query="Select pcComm_IDOrder,pcComm_IDFeedback,pcComm_IDUser from pcComments where pcComm_IDOrder=" & intIdOrder & " and pcComm_IDFeedback=" & intIdFeedback
set rs=connTemp.execute(query)
if rs.eof then
	set rs = nothing
	call closedb()
	response.redirect "adminviewallmsgs.asp?IDOrder=" & intIdOrder
	else
	pcv_IDUser=rs("pcComm_IDUser")
		if pcv_IDUser = "0" then
			query="SELECT idCustomer FROM orders WHERE idOrder="&intIdOrder
			set rsTemp=connTemp.execute(query)
			pcv_IDUser=rsTemp("idCustomer")
			set rsTemp=nothing
		end if
	set rs = nothing
End if
%>
<!-- #Include File="Adminheader.asp" -->
<%
query="SELECT idcustomer FROM customers WHERE idcustomer=" & pcv_IDUser
set rstemp=connTemp.execute(query)
if rstemp.eof then
%>
	<div class="pcCPmessage">This message thread cannot be updated because the customer account associated with this message no longer exists in the database. <a href="adminviewallmsgs.asp">View other Help Desk messages</a></div>

<%
end if
set rstemp=nothing

ToolTips=0

intIdOrder=getUserInput(request("IDOrder"),0)
intIdFeedback=getUserInput(request("IDFeedback"),0)
session("admin_IDOrder")=intIdOrder

'Change FeedBack Status
if request("action")="changestatus" then
	intNewStatus=getUserInput(request("new"),0)
	'Only Admin can change Feedback Status
	query="Update pcComments set pcComm_FStatus=" & intNewStatus & " where pcComm_IDFeedback=" & intIdFeedback & ";"
	set rs=connTemp.execute(query)
end if

'Create new comment
if (request("action")="add") and (request("rewrite")="0") then
	r=0
	strNewComments=getUserInput(request("comments"),0)
	dtComDate=CheckDateSQL(now())
	if scDB="SQL" then
		query="Insert Into pcComments (pcComm_IDOrder,pcComm_IDParent,pcComm_IDUser,pcComm_CreatedDate,pcComm_EditedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details) values (" & intIdOrder & "," & intIdFeedback & ",0,'" & dtComDate & "','" & dtComDate & "',0,0,0,'','" & strNewComments & "')"
	else
		query="Insert Into pcComments (pcComm_IDOrder,pcComm_IDParent,pcComm_IDUser,pcComm_CreatedDate,pcComm_EditedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details) values (" & intIdOrder & "," & intIdFeedback & ",0,#" & dtComDate & "#,'" & dtComDate & "',0,0,0,'','" & strNewComments & "')"
	end if
	 
	set rs=connTemp.execute(query)
	
	if scDB="Access" then
		query="select pcComm_IDFeedback from pcComments where pcComm_IDParent=" & intIdFeedback & " and pcComm_IDUser=0 and pcComm_IDOrder=" & intIdOrder & " and pcComm_CreatedDate=#" & dtComDate & "# ORDER BY pcComm_IDFeedback DESC"
	else
		query="select pcComm_IDFeedback from pcComments where pcComm_IDParent=" & intIdFeedback & " and pcComm_IDUser=0 and pcComm_IDOrder=" & intIdOrder & " and pcComm_CreatedDate='" & dtComDate & "' ORDER BY pcComm_IDFeedback DESC"
	end if
	set rs=connTemp.execute(query)
	
	if rs.eof then
		strMsg=dictLanguage.Item(Session("language")&"_viewFeedback_12")
	else
		intLastFB=rs("pcComm_IDFeedback")
		set rs=nothing
		IDComment=intLastFB
	
		if scDB="SQL" then
			query="update pcComments set pcComm_EditedDate='" & dtComDate & "' where pcComm_IDFeedback=" & intIdFeedback
		else
			query="update pcComments set pcComm_EditedDate=#" & dtComDate & "# where pcComm_IDFeedback=" & intIdFeedback
		end if
		set rs=connTemp.execute(query)
	
			ACount=getUserInput(request("ACount"),0)
			if ACount<>"" then
				ACount1=clng(ACount)
				For k=1 to ACount1
					if request("AC" & k)="1" then
						query="update pcUploadFiles set pcUpld_IDFeedback=" & IDComment & " where pcUpld_IDFile=" & getUserInput(request("AID" & k),0) & " and pcUpld_IDFeedback=0"
						set rs=connTemp.execute(query)
					end if
				next
				query="delete from pcUploadFiles where pcUpld_IDFeedback=0"
				set rs=connTemp.execute(query)
			end if

		'Generate View Comment Link for Customer
		strPath=Request.ServerVariables("PATH_INFO")
		iCnt=0
		do while iCnt<2
			if mid(strPath,len(strPath),1)="/" then
				iCnt=iCnt+1
			end if
			if iCnt<2 then
				strPath=mid(strPath,1,len(strPath)-1)
			end if
		loop
		strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath
	
		if Right(strPathInfo,1)="/" then
		else
			strPathInfo=strPathInfo & "/"
		end if
	
		dURL=strPathInfo & "pc/Checkout.asp?cmode=1&redirectUrl=" & Server.URLEnCode(strPathInfo & "pc/userviewfeedback.asp?IDOrder=" & scpre+clng(intIdOrder) & "&IDFeedback=" & intIdFeedback)

		'Send mail to USers
		strMsgBody=""
		strMsgBody=dictLanguage.Item(Session("language")&"_addFB_email1") & scpre+clng(intIdOrder) & dictLanguage.Item(Session("language")&"_addFB_email2") & VBCrlf & VBCrlf
	
		query="Select pcComm_Description,pcComm_FStatus, pcComm_IDUser,pcComm_Priority, pcComm_CreatedDate,pcComm_EditedDate from pcComments where pcComm_IDOrder=" & intIdOrder & " and pcComm_IDFeedback=" & intIdFeedback & " and pcComm_IDParent=0"
		set rs=connTemp.execute(query)
		FTitle=rs("pcComm_Description")
		FStatus=rs("pcComm_FStatus")
		UPosted=rs("pcComm_IDUser")
		Priority=rs("pcComm_Priority")
		PostedDate=rs("pcComm_CreatedDate")
		EditedDate=rs("pcComm_EditedDate")
		
		query="Select pcFStat_name from pcFStatus where pcFStat_IDStatus=" & FStatus 
		set rs=connTemp.execute(query)
		FBStatus=rs("pcFStat_name")
	
		if UPosted="0" then
			UserPosted=dictLanguage.Item(Session("language")&"_viewPostings_2")
		else
			query="Select name,lastname from Customers where IDCustomer=" & UPosted 
			set rs=connTemp.execute(query)
			UserPosted=rs("name") & " " & rs("lastname")
		end if
	
		UserEdited=dictLanguage.Item(Session("language")&"_viewPostings_2")

    query="Select pcPri_name from pcPriority where pcPri_IDPri=" & Priority 
		set rs=connTemp.execute(query)
		FPriority=rs("pcPri_name")
	
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email3") & scpre+clng(intIdOrder) & VBCrlf
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email4") & FTitle & VBCrlf
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email5") & UserPosted & VBCrlf
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email14") & CheckDate(PostedDate) & VBCrlf	
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email6") & FPriority & VBCrlf	
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email11") & FBStatus & VBCrlf
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email12") & UserEdited & VBCrlf
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email13") & CheckDate(EditedDate) & VBCrlf	
	
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email8") & dURL & VBCrlf & VBCrlf
		strMsgBody=strMsgBody & scCompanyName
	
		query="select customers.name,customers.lastName,customers.email from Orders,Customers where Orders.IDOrder=" & intIdOrder & " and Customers.IDCustomer=Orders.IDCustomer"
		set rs=connTemp.execute(query)
		pcstrCustomerName = rs("name")
		pcstrCustomerLast = rs("lastName")
		pcstrCustomerEmail = rs("email")
		set rs = nothing

		strMsgBodyMain=pcstrCustomerName & " " & pcstrCustomerLast & ","&VBCrlf&VBCrlf&strMsgBody
		strMsgSubject=scCompanyName & dictLanguage.Item(Session("language")&"_addFB_email9") & scpre+clng(intIdOrder)
		call sendmail(scCompanyName, scEmail, pcstrCustomerEmail, strMsgSubject, strMsgBodyMain)
	
		r=1
		strMsg=dictLanguage.Item(Session("language")&"_viewFeedback_a")
	end if %>
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
<%
end if 'Create new comment

if request("uploaded")<>"" then
	session("uploaded")="1"
else
	session("uploaded")="0"
end if	

'Delete Temponary uploaded files
if session("uploaded")="1" then
	session("uploaded")="0"
else
	query="Select pcUpld_Filename from pcUploadFiles where pcUpld_IDFeedback=0"
	set rs=connTemp.execute(query)
	do while not rs.eof
		strFilename=rs("pcUpld_Filename")
		if strFilename<>"" then
			QfilePath="Library/" & strFilename
			findit = Server.MapPath(QfilePath)
			Set fso = server.CreateObject("Scripting.FileSystemObject")
			Set f = fso.GetFile(findit)
			f.Delete
			Set fso = nothing
			Set f = nothing
			Err.number=0
			Err.Description=""
		end if
		rs.MoveNext
	loop
 	query="delete from pcUploadFiles where pcUpld_IDFeedback=0"
 	set rs=connTemp.execute(query)
 	session("uploaded")="0"
end if 
%>

<% 
if request("msg")<>"" then
	if request("msg")="1" then
		msg="The comment was deleted successfully!"
		msgType=1
	else
		msg=getUserInput(request("msg"),0)
	end if
%>
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
<%
end if
%>

<table class="pcCPcontent">
	<tr>
		<td>
			<div style="float: right; margin-top: -35px;"><a href="adminviewallmsgs.asp?IDOrder=<%=intIdOrder%>">View messages for order <%=scpre+int(intIdOrder)%></a> | <a href="adminviewallmsgs.asp">View all messages</a> | <a href="ordDetails.asp?id=<%=intIdOrder%>">View order details</a></div>
<%
query="Select pcComm_idfeedback,pcComm_iduser,pcComm_createdDate,pcComm_editedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details from pcComments where pcComm_IDFeedback=" & intIdFeedback

set rs=connTemp.execute(query)
intIdFeedback=rs("pcComm_idfeedback")
IDUser=rs("pcComm_iduser")
createdDate=rs("pcComm_createdDate")
editedDate=rs("pcComm_editedDate")
FType=rs("pcComm_FType")
FStatus=rs("pcComm_FStatus")
Priority=rs("pcComm_Priority")
FDesc=rs("pcComm_Description")
FDetails=rs("pcComm_Details")

FDetails = replace(FDetails,"&lt;","<")
FDetails = replace(FDetails,"&gt;",">") 
%>
<h2><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_c")%><%=intIdFeedback%></h2>

<table width="90%" border="0" align="center" cellpadding="6" cellspacing="0">
	<tr align="left" bgcolor="#e5e5e5">
		<td colspan="2" width="100%" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_d")%><strong><%=CheckDate(createdDate)%></strong> by <strong><%
		if (IDUser<>"") and (IDUser<>"0") then
		query="Select name,lastname from Customers where IDCustomer=" & IDUser
		set rs=connTemp.execute(query)
		if not rs.eof then %>
			<%=rs("Name") & " " & rs("LastName") %>
		<% end if
		else%>
			<%response.write dictLanguage.Item(Session("language")&"_viewPostings_2")%>
		<%end if%>
		</strong>
		</td>
		<td nowrap>
		<a href="javascript:if (confirm('You are about to remove this feedback from this order. Are you sure you want to complete this action?')) location='admindelfeedback.asp?IDfeedback=<%=intIdFeedback%>&IDOrder=<%=intIdOrder%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete"></a>&nbsp;<a href="admineditfeedback.asp?IDfeedback=<%=intIdFeedback%>&IDOrder=<%=intIdOrder%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit"></a> 
		</td>
    </tr>
	<td width="20%" align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_h")%></td>
	<td width="80%"><strong>
	<% 
	query="SELECT pcFType_name,pcFType_Img, pcFType_ShowImg FROM pcFTypes WHERE pcFType_IDType=" & FType
    set rs=connTemp.execute(query)
    if not rs.eof then
			PName=rs("pcFtype_Name")
			PImg=rs("pcFtype_Img")
			TypeImage=rs("pcFtype_ShowImg")
			if TypeImage=1 then
				if PImg<>"" then%>
					<img src="../pc/images/<%=PImg%>" alt="<%=PName%>" border="0">
				<%end if
			else%>
				<%=ucase(PName)%>
			<%end if
		end if%>
        </strong>
	</td>
  </tr>
  <tr>
    <td width="20%" align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_i")%></td>
    <td width="80%">
    <% query="Select pcPri_Name,pcPri_Img,pcPri_ShowImg from pcPriority where pcPri_IDPri=" & Priority
    set rs=connTemp.execute(query)
    if not rs.eof then
    	PName=rs("pcPri_Name")
    	PImg=rs("pcPri_Img")
    	PriorityImage=rs("pcPri_ShowImg")
    	if PriorityImage=1 then
    		if PImg<>"" then%>
    			<img src="../pc/images/<%=PImg%>" alt="<%=PName%>" border="0">
    		<%end if
    	else%>
    		<%=PName%>
    	<%end if
    end if%>
    </td>
  </tr>
  <tr>
    <td width="20%" align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_j")%></td>
    <td width="80%"><%=FDesc%></td>
  </tr>
  <tr>
    <td align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_k")%></td>
    <td><%=FDetails%></td>
  </tr>
  <tr>
    <td align="right" valign="middle"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_l")%></td>
    <td>
	<%
    query="Select pcFStat_Name,pcFStat_Img,pcFStat_ShowImg from pcFStatus where pcFStat_IDStatus=" & FStatus
    set rs=connTemp.execute(query)
    if not rs.eof then
    	PName=rs("pcFStat_Name")
    	PImg=rs("pcFStat_Img")
    	StatusImage=rs("pcFStat_ShowImg")
    	if StatusImage=1 then
    		if PImg<>"" then%>
    			<img src="../pc/images/<%=PImg%>" alt="<%=PName%>" border="0">
    		<%end if
    	else%>
    		<%=PName%>
    	<%end if
    end if%>
		<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_m")%>
		<select name="FStatus" id="FStatus" onChange="location='adminviewfeedback.asp?IDOrder=<%=intIdOrder%>&idfeedback=<%=intIdFeedback%>&action=changestatus&new='+document.getElementById('FStatus').value;">
		<%
		query="Select pcFstat_idstatus,pcFstat_name from pcFStatus"
		set rs1=connTemp.execute(query)
		do while not rs1.eof%>
			<option value="<%=rs1("pcFstat_idstatus")%>" <%if rs1("pcFstat_idstatus")=FStatus then%>selected<%end if%>><%=rs1("pcFstat_name")%></option>
			<%rs1.MoveNext
		Loop%>
		</select>
		</td>
  	</tr>
	<%
		query="Select pcUpld_IDFile,pcUpld_FileName from pcUploadFiles where pcUpld_IDFeedback=" & intIdFeedback
		set rs=connTemp.execute(query)
		if not rs.eof then
			%>
			<tr>
				<td valign="top" colspan="2"><span style="background-color: #FFCC00"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_s")%></span><br>
				<%Do while not rs.eof%>
					<a href="admindownload.asp?IDFile=<%=rs("pcUpld_IDFile")%>" target="_blank"><img src="images/DownLoad.gif" border=0 height=19 width=18 alt="Download File"></a>&nbsp;<b><%
					strFilename= rs("pcUpld_FileName")
					strFilename = mid(strFilename,instr(strFilename,"_")+1,len(strFilename))%>
					<%=strFilename%></b><br>
					<%rs.MoveNext
				loop
				%>
      	</td>
		</tr>
		<%
        end if
        %>
</table>
<%
query="Select pcComm_idfeedback,pcComm_iduser,pcComm_createdDate,pcComm_editedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details from pcComments where pcComm_IDParent=" & intIdFeedback & " order by pcComm_IDFeedback"

set rs=connTemp.execute(query)

if not rs.eof then %>
	<h2 style="margin-top: 25px; margin-bottom: 15px;"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_t")%></h2>
<%end if

Do while not rs.eof
	IDComment=rs("pcComm_idfeedback")
	IDUser=rs("pcComm_iduser")
	createdDate=rs("pcComm_createdDate")
	editedDate=rs("pcComm_editedDate")
	FType=rs("pcComm_FType")
	FStatus=rs("pcComm_FStatus")
	Priority=rs("pcComm_Priority")
	FDesc=rs("pcComm_Description")
	FDetails=rs("pcComm_Details")
	
	FDetails	= replace(FDetails,"&lt;","<")
 	FDetails	= replace(FDetails,"&gt;",">")
	%>
	<table width="90%" border="0" align="center" cellpadding="6" cellspacing="0">
		<tr align="left" bgcolor="#e5e5e5">
			<td colspan="2" valign="top" width="100%"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_u")%><%if (createdDate & "") = (editedDate & "") then%><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_v")%><strong><%=CheckDate(createdDate)%></strong><%else%><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_w")%><strong><%=CheckDate(editedDate)%></strong><%end if%><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_x")%><strong><%
			if (IDUser<>"") and (IDUser<>"0") then
    		query="Select name,lastname from Customers where IDCustomer=" & IDUser
    		set rstemp=connTemp.execute(query)
				if not rstemp.eof then%>
					<%=rstemp("Name") & " " & rstemp("LastName")%>
    		<% end if
    	else %>
    		<%response.write dictLanguage.Item(Session("language")&"_viewPostings_2")%>
    	<% end if %></strong></td>
			<td nowrap>
      	<a href="javascript:if (confirm('You are about to remove this comment from order feedback. Are you sure you want to complete this action?')) location='admindelcomment.asp?IDComment=<%=IDComment%>&IDfeedback=<%=intIdFeedback%>&IDOrder=<%=intIdOrder%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete"></a>&nbsp;<a href="admineditcomment.asp?IDComment=<%=IDComment%>&IDfeedback=<%=intIdFeedback%>&IDOrder=<%=intIdOrder%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit"></a>
      </td>
    </tr>
	</table>
	<table width="90%" border="0" align="center" cellpadding="6" cellspacing="0">    
		<tr>
			<td width="20%" align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_z")%></td>
			<td width="80%"><%=FDetails%></td>
		</tr>
		<%
			query="Select pcUpld_IDFile,pcUpld_FileName from pcUploadFiles where pcUpld_IDFeedback=" & IDComment
			set rstemp=connTemp.execute(query)
			if not rstemp.eof then
			%>
  			<tr>
    			<td valign="top" colspan="2"><br><span style="background-color: #FFCC00"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_s")%></span><br>
  				<%Do while not rstemp.eof%>
  					<a href="admindownload.asp?IDFile=<%=rstemp("pcUpld_IDFile")%>" target="_blank"><img src="images/DownLoad.gif" border=0 height=19 width=18 alt="Download File"></a>&nbsp;<b><%
strFilename= rstemp("pcUpld_FileName")
strFilename = mid(strFilename,instr(strFilename,"_")+1,len(strFilename))%>
<%=strFilename%></b><br>
  					<%rstemp.MoveNext
  				loop %>
				</td>
			</tr>
		<%end if%>
	</table>
	<br>
	<%rs.MoveNext
Loop%>
<p>&nbsp;</p>
<p align="center"></p>
<script language="JavaScript">
<!--
function Form1_Validator(theForm)
{

	if (theForm.comments.value == "")
 	{
		    alert("This field cannot be blank. Please add a comment before proceeding.");
		    theForm.comments.focus();
		    return (false);
	}
  
return (true);
}
//-->
</script>

<form name="hForm" method="post" action="adminviewfeedback.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<script language="JavaScript"><!--
function newWindow(file,window) {
		msgWindow=open(file,window,'resizable=no,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
}
//--></script>
<div align="center">
<table class="pcCPcontent" style="border: 1px solid #999; width: 600px;">
	<tr bgcolor="#f5f5f5"><td colspan="2" align="left"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_1")%></td></tr>
	<tr bgcolor="#f5f5f5"><td colspan="2" align="left">
        <ol>
          <li><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_2")%></li>
          <li><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_3")%></li>
        </ol>
    </td>
    </tr>
    <tr><td colspan="2" align="left">&nbsp;</td>
    </tr>
	<tr>
    <td align="right" valign="top" nowrap><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_5")%></td>
	<td valign="top">
	<%query="Select pcUpld_IDFile,pcUpld_FileName from pcUploadFiles where pcUpld_IDFeedback=0"
	set rs=connTemp.execute(query)
	if rs.eof then%>
		<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_6")%><br>
	<%
	else
		ACount=0
		do while not rs.eof
			ACount=ACount+1 %>
			<input type=hidden name="AID<%=ACount%>" value="<%=rs("pcUpld_IDFile")%>">
			<input type=checkbox name="AC<%=ACount%>" value="1" checked>&nbsp;<%
			strFilename= rs("pcUpld_FileName")
			strFilename = mid(strFilename,instr(strFilename,"_")+1,len(strFilename))%>
			<%=strFilename%><br>
			<%rs.MoveNext
		loop%>
		<input type=hidden name=ACount value="<%=ACount%>">
	<%
	end if
	%>
	<script language="JavaScript"><!--
		function newWindow1(file,window) {
		catWindow=open(file,window,'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360');
		if (catWindow.opener == null) catWindow.opener = self;
		}
	//--></script>
	<br><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_7")%><a href="#" onClick="javascript:newWindow1('adminfileuploada_popup.asp?IDFeedback=0&ReLink=<%=Server.URLencode("adminviewfeedback.asp?IDOrder=" & intIdOrder & "&IDFeedback=" & intIdFeedback)%>','window2')"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_8")%></a>.
    </td>
    </tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td valign="top"><p align="right"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_9")%>
		<br><br>
		<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=comments','window2')">
		</p>
  	</td>
		<td valign="top">				
  	<input type=hidden name=IDOrder value="<%=intIdOrder%>">
  	<input type=hidden name=IDFeedback value="<%=intIdFeedback%>">
    <textarea name="comments" cols="50" rows="8"><%=request("comments")%></textarea>
		</td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td colspan="2" align="center">
		<input type="submit" name="Submit" value=" Add Comment " class="submit2" onclick="document.hForm.rewrite.value='0';">
		<input type="hidden" name="uploaded" value="">
		<input type="hidden" name="rewrite" value="1">
		</td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
</table>
</div>
</form>
<iframe name="downloadwindow" src="about:blank"  marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 width=0 height=0 scrolling="no" noresize></iframe>
</td>
</tr>
</table>
<%call closedb()%>
<!-- #Include File="Adminfooter.asp" -->