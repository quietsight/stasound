<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!-- #Include File="checkdate.asp" -->
<%
'Allow upload: change to "0" to disallow
AllowUpload="1"

on error resume next
Dim rs, connTemp, query
Dim LngIdOrder,LngIDFeedback
call openDB() 

LngIdOrder=getUserInput(request("IDOrder"),0)
session("IDOrder")=LngIdOrder

LngIdOrder=Clng(LngIdOrder)-Clng(scpre)

LngIDFeedback=getUserInput(request("IDFeedback"),0)

query="SELECT IDCustomer FROM Orders WHERE IDOrder=" & LngIdOrder & " and IDCustomer=" & session("IDCustomer")

set rs=connTemp.execute(query)

if rs.eof then
	call closedb()
	response.redirect "userviewallposts.asp?IDOrder=" & Clng(scpre)+Clng(LngIdOrder)
end if
%>
<!-- #Include File="header.asp" -->
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<h1><%response.write dictLanguage.Item(Session("language")&"_viewPostings_3")%></h1>
			</td>
		</tr>
		<tr>
			<td>

		<%
		'Change FeedBack Status
		if request("action")="changestatus" then
			NewStatus=getUserInput(request("new"),0)
			'Only Admin can change Feedback Status
			if (session("UserType")=3) then
				query="Update pcComments set pcComm_FStatus=" & NewStatus & " WHERE pcComm_IDFeedback=" & LngIDFeedback & " and pcComm_IDParent=0 and pcComm_IDOrder=" & LngIdOrder
				set rs=connTemp.execute(query)
			end if
		end if

		'Create new comment
		if (request("action")="add") and (request("rewrite")="0") then
			NewComments=getUserInput(request("comments"),0)
			dtComDate=CheckDateSQL(now())
			if scDB="SQL" then
				query="Insert Into pcComments (pcComm_IDOrder,pcComm_IDParent,pcComm_IDUser,pcComm_CreatedDate,pcComm_EditedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details) values (" & LngIdOrder & "," & LngIDFeedback & "," & session("IDCustomer") & ",'" & dtComDate & "','" & dtComDate & "',0,0,0,'','" & NewComments & "')" 
			else
				query="Insert Into pcComments (pcComm_IDOrder,pcComm_IDParent,pcComm_IDUser,pcComm_CreatedDate,pcComm_EditedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details) values (" & LngIdOrder & "," & LngIDFeedback & "," & session("IDCustomer") & ",#" & dtComDate & "#,#" & dtComDate & "#,0,0,0,'','" & NewComments & "')" 
			end if
			
			set rs=connTemp.execute(query)
	
			if scDB="Access" then
				query="SELECT pcComm_IDFeedback FROM pcComments WHERE pcComm_IDParent=" & LngIDFeedback & " and pcComm_IDUser=" & session("IDCustomer") & " and pcComm_IDOrder=" & LngIdOrder & " and pcComm_CreatedDate=#" & dtComDate & "# ORDER BY pcComm_IDFeedback DESC"
			else
				query="SELECT pcComm_IDFeedback FROM pcComments WHERE pcComm_IDParent=" & LngIDFeedback & " and pcComm_IDUser=" & session("IDCustomer") & " and pcComm_IDOrder=" & LngIdOrder & " and pcComm_CreatedDate='" & dtComDate & "' ORDER BY pcComm_IDFeedback DESC"
			end if
			set rs=connTemp.execute(query)
	
			IF rs.eof THEN
				Msg=dictLanguage.Item(Session("language")&"_viewFeedback_12")
			ELSE
				LastFB=rs("pcComm_IDFeedback")
				set rs=nothing
				IDComment=LastFB
	
				query="update pcComments set pcComm_EditedDate='" & dtComDate & "' WHERE pcComm_IDFeedback=" & LngIDFeedback
				set rs=connTemp.execute(query)
	
				if AllowUpload="1" then
					ACount=getUserInput(request("ACount"),0)
					if ACount<>"" then
						ACount1=clng(ACount)
						For k=1 to ACount1
							if request("AC" & k)="1" then
								query="update pcUploadFiles set pcUpld_IDFeedback=" & IDComment & " WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0) & " and pcUpld_IDFeedback=0"
								set rs=connTemp.execute(query)
							end if
						next
						query="delete FROM pcUploadFiles WHERE pcUpld_IDFeedback=0"
						set rs=connTemp.execute(query)
					end if
				end if

				'Generate View Comment Link for Store Owner
				SPath1=Request.ServerVariables("PATH_INFO")
				mycount1=0
				do while mycount1<2
					if mid(SPath1,len(SPath1),1)="/" then
						mycount1=mycount1+1
					end if
					if mycount1<2 then
						SPath1=mid(SPath1,1,len(SPath1)-1)
					end if
				loop
				SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
	
				if Right(SPathInfo,1)="/" then
				else
					SPathInfo=SPathInfo & "/"
				end if
	
				dURL=SPathInfo & scAdminFolderName & "/login_1.asp?redirectUrl=" & Server.URLEnCode(SPathInfo & scAdminFolderName &  "/adminviewfeedback.asp?IDOrder=" & LngIdOrder & "&IDFeedback=" & LngIDFeedback)
	
				'Send mail to Store Owner
				MsgBody=""
				MsgBody=dictLanguage.Item(Session("language")&"_addFB_email1") & Clng(scpre)+Clng(LngIdOrder) & dictLanguage.Item(Session("language")&"_addFB_email2") & VBCrlf & VBCrlf
	
				query="SELECT pcComm_Description,pcComm_FStatus,pcComm_IDUser,pcComm_Priority,pcComm_CreatedDate,pcComm_EditedDate FROM pcComments WHERE pcComm_IDFeedback=" & LngIDFeedback & ";"
				set rs=connTemp.execute(query)
				FTitle=rs("pcComm_Description")
				FStatus=rs("pcComm_FStatus")
				UPosted=rs("pcComm_IDUser")
				Priority=rs("pcComm_Priority")
				PostedDate=rs("pcComm_CreatedDate")
				EditedDate=rs("pcComm_EditedDate")
	
				query="SELECT pcFStat_name FROM pcFStatus WHERE pcFStat_IDStatus=" & FStatus 
				set rs=connTemp.execute(query)
				FBStatus=rs("pcFStat_name")
	
				if UPosted<>"0" then
					query="SELECT name,lastname FROM Customers WHERE IDCustomer=" & UPosted 
					set rs=connTemp.execute(query)
					UserPosted=rs("name") & " " & rs("lastname")
				else
					UserPosted=dictLanguage.Item(Session("language")&"_viewPostings_2")
				end if
	
				query="SELECT name,lastname,email FROM Customers WHERE IDCustomer=" & session("IDCustomer") 
				set rs=connTemp.execute(query)
				UserEdited=rs("name") & " " & rs("lastname")
	
    		query="SELECT pcPri_name FROM pcPriority WHERE pcPri_IDPri=" & Priority 
				set rs=connTemp.execute(query)
				FPriority=rs("pcPri_name")
	
	
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_addFB_email3") & Clng(scpre)+Clng(LngIdOrder) & vbcrlf
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_addFB_email4") & FTitle & vbcrlf
				
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_addFB_email5") & UserPosted & vbcrlf
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_addFB_email14") & CheckDate(PostedDate) & vbcrlf	
				
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_addFB_email6") & FPriority & vbcrlf	
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_addFB_email11") & FBStatus & vbcrlf
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_addFB_email12") & UserEdited & vbcrlf
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_addFB_email13") & CheckDate(EditedDate) & vbcrlf	
				
				MsgBody=MsgBody & dictLanguage.Item(Session("language")&"_addFB_email8") & dURL & VBCrlf&VBCrlf
				MsgBody=MsgBody & scCompanyName
	
				MsgBody1=scCompanyName & ","&VBCrlf&VBCrlf&MsgBody
				
				Dim strCustServEmail
				strCustServEmail=scCustServEmail
				if trim(strCustServEmail)="" then strCustServEmail=scFrmEmail

				call sendmail(scCompanyName,scEmail,strCustServEmail,scCompanyName & dictLanguage.Item(Session("language")&"_addFB_email9") & clng(scpre)+clng(LngIdOrder),MsgBody1)
			
				r=1
				Msg=dictLanguage.Item(Session("language")&"_viewFeedback_a")
			END IF%>
				<div class="pcErrorMessage">
					<%=Msg%>
				</div>
		<%end if

		if request("uploaded")<>"" then
			session("uploaded")="1"
		else
			session("uploaded")="0"
		end if	

		'Delete Temponary uploaded files
		if session("uploaded")="1" then
			session("uploaded")="0"
		else
 			query="SELECT pcUpld_Filename FROM pcUploadFiles WHERE pcUpld_IDFeedback=0"
 			set rs=connTemp.execute(query)
 			do while not rs.eof
 				Filename=rs("pcUpld_Filename")
				if Filename<>"" then
					QfilePath="Library/" & Filename
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
 			query="delete FROM pcUploadFiles WHERE pcUpld_IDFeedback=0"
 			set rs=connTemp.execute(query)
 			session("uploaded")="0"
		end if 
		%>

		<% if request("msg")<>"" then
			if request("msg")="1" then
				msg=dictLanguage.Item(Session("language")&"_editFeedback_i")
			else
				msg=getUserInput(request("msg"),0)
			end if%>
				<div class="pcErrorMessage">
					<%=Msg%>
				</div>
		<%end if%>
		<br>
		<p><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_b")%>
		<a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(LngIdOrder)%>"><strong><%=clng(scpre)+clng(LngIdOrder)%></strong></a>
		</p>
		<% query="SELECT pcComm_idfeedback, pcComm_iduser, pcComm_createdDate,pcComm_editedDate, pcComm_FType, pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details FROM pcComments WHERE pcComm_IDFeedback=" & LngIDFeedback

		set rs=connTemp.execute(query)
		IDfeedback=rs("pcComm_idfeedback")
		intIDUser=rs("pcComm_iduser")
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
		<p><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_c")%><b>#<%=LngIDFeedback%></b></p>
		
		<table class="pcShowContent">
  	<tr align="left" class="pcSectionTitle">
    	<td colspan="2"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_d")%><strong><%=CheckDate(createdDate)%></strong> <%response.write dictLanguage.Item(Session("language")&"_viewFeedback_x")%> <strong><%
    	if (intIDUser<>"") and (intIDUser<>"0") then
    		query="SELECT name,lastname FROM Customers WHERE IDCustomer=" & intIDUser
    		set rs=connTemp.execute(query)
    		if not rs.eof then%>
    			<%=rs("Name") & " " & rs("LastName") %>
    		<%	end if
    	else%>
    		<%response.write dictLanguage.Item(Session("language")&"_viewPostings_2")%>
    	<%end if%>
      </strong>
			</td>
      <td nowrap>
      <%if (session("UserType")=1) or (session("UserType")=2) or (intIDUser=session("IDCustomer")) then%>
      	<a href="javascript:if (confirm('<%response.write dictLanguage.Item(Session("language")&"_editFeedback_j")%>')) location='userdelfeedback.asp?IDfeedback=<%=LngIDFeedback%>&IDOrder=<%=clng(scpre)+clng(LngIdOrder)%>'"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_e")%></a> | <a href="usereditfeedback.asp?IDfeedback=<%=LngIDFeedback%>&IDOrder=<%=clng(scpre)+clng(LngIdOrder)%>"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_f")%></a> 
      <%else%>
      	&nbsp;
      <%end if%>
      </td>
    </tr>
  </table>
  <table class="pcShowContent">  
  <tr class="main">
    <td width="20%" align="right" valign="top">
		<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_h")%>
		</td>
    <td width="80%">
		<strong>
		<%
    query="SELECT pcFtype_Name,pcFtype_img,pcFType_ShowImg FROM pcFTypes WHERE pcFtype_IDType=" & FType
    set rs=connTemp.execute(query)
    if not rs.eof then
			PName=rs("pcFtype_Name")
			PImg=rs("pcFtype_Img")
			TypeImage=rs("pcFtype_ShowImg")
    	if TypeImage=1 then
    		if PImg<>"" then%>
					<img src="images/<%=PImg%>" alt="<%=PName%>" border="0">
    		<%end if
    	else%>
    		<%=ucase(PName)%>
    	<%end if
    end if%>
		</strong>
		</td>
  </tr>
  <tr>
    <td align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_i")%></td>
    <td>
    <%
    query="SELECT pcPri_Name, pcPri_Img, pcPri_ShowImg FROM pcPriority WHERE pcPri_IDPri=" & Priority
    set rs=connTemp.execute(query)
    if not rs.eof then
    	PName=rs("pcPri_Name")
    	PImg=rs("pcPri_Img")
    	PriorityImage=rs("pcPri_ShowImg")
    	if PriorityImage=1 then
    		if PImg<>"" then%>
					<img src="images/<%=PImg%>" alt="<%=PName%>" border="0">
    		<%end if
    	else%>
    		<%=PName%>
    	<%end if
    end if%>
		</td>
  </tr>
  <tr>
    <td align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_j")%></td>
    <td><%=FDesc%></td>
  </tr>
  <tr>
    <td align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_k")%></td>
    <td><%=FDetails%></td>
  </tr>
  <tr class="main">
    <td align="right" valign="middle">
		<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_l")%>
		</td>
    <td>
		<%
    query="SELECT pcFStat_Name,pcFStat_Img,pcFStat_ShowImg FROM pcFStatus WHERE pcFStat_IDStatus=" & FStatus
    set rs=connTemp.execute(query)
    if not rs.eof then
			PName=rs("pcFStat_Name")
			PImg=rs("pcFStat_Img")
			StatusImage=rs("pcFStat_ShowImg")
    	if StatusImage=1 then
    		if PImg<>"" then%>
    			<img src="images/<%=PImg%>" alt="<%=PName%>" border="0">
    		<%end if
    	else%>
    		<%=PName%>
    	<%end if
    end if
    if (session("UserType")=3) then%><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_m")%>
			<select name="FStatus" onChange="location='userviewfeedback.asp?IDOrder=<%=clng(scpre)+clng(LngIdOrder)%>&idfeedback=<%=LngIDFeedback%>&action=changestatus&new='+FStatus.value;">
			<%
   		query="SELECT pcFStat_idstatus,pcFStat_Name FROM pcFStatus"
   		set rs=connTemp.execute(query)
  		do while not rs.eof%>
   			<option value="<%=rs("pcFstat_idstatus")%>" <%if rs("pcFstat_idstatus")=FStatus then%>selected<%end if%>><%=rs("pcFstat_name")%></option>
   			<%rs.MoveNext
   		Loop%>
			</select>
		<%end if%>
		</td>
  </tr>
	<% if AllowUpload="1" then
		query="SELECT pcUpld_IDFile,pcUpld_FileName FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & LngIDFeedback
		set rs=connTemp.execute(query)
		if not rs.eof then
		%>
  	<tr>
    	<td valign="top" colspan="2" class="pcSmallText">
			<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_s")%><br>
  		<% Do until rs.eof %>
  			<a href="userdownload.asp?IDFile=<%=rs("pcUpld_IDFile")%>" target="_blank"><img src="images/DownLoad.gif" border=0 height=19 width=18 alt="Download File"></a>&nbsp;<b><%
				Filename= rs("pcUpld_FileName")
				FileName = mid(FileName,instr(Filename,"_")+1,len(FileName))%>
				<%=FileName%></b><br>
				<%rs.MoveNext
  		loop
 		 	%>
			</td>
		</tr>
	<%end if
end if
%>
</table>
<%
query="SELECT pcComm_idfeedback,pcComm_iduser,pcComm_createdDate,pcComm_editedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details FROM pcComments WHERE pcComm_IDParent=" & LngIDFeedback & " order by pcComm_IDFeedback"

set rs=connTemp.execute(query)

if not rs.eof then
	%>
	<div class="pcSectionTitle" style="margin: 10px 0 10px 0;"><strong><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_t")%></strong></div>
<%end if

Do while not rs.eof
	IDComment=rs("pcComm_idfeedback")
	intIDUser=rs("pcComm_iduser")
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
	<table class="pcShowContent">
		<tr align="left">
			<td colspan="2" valign="top">
			<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_u")%><%if (createdDate & "") = (editedDate & "") then%><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_v")%><strong><%=CheckDate(createdDate)%></strong><%else%><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_w")%><strong><%=CheckDate(editedDate)%></strong><%end if%><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_x")%><strong><%
			if (intIDUser<>"") and (intIDUser<>"0") then
    		query="SELECT name,lastname FROM Customers WHERE IDCustomer=" & intIDUser
    		set rstemp=connTemp.execute(query)
    		if not rstemp.eof then%>
    			<%=rstemp("Name") & " " & rstemp("LastName")%>
    		<%end if
    	else%>
    		<%response.write dictLanguage.Item(Session("language")&"_viewPostings_2")%>
    	<%end if%></strong>
			</td>
			<td nowrap>
      <%if (session("UserType")=3) or (session("UserType")=2) or (intIDUser=session("IDCustomer")) then%>
      	<a href="javascript:if (confirm('<%response.write dictLanguage.Item(Session("language")&"_editFeedback_j")%>')) location='userdelcomment.asp?IDComment=<%=IDComment%>&IDfeedback=<%=LngIDFeedback%>&IDOrder=<%=clng(scpre)+clng(LngIdOrder)%>'"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_e")%></a> | <a href="usereditcomment.asp?IDComment=<%=IDComment%>&IDfeedback=<%=LngIDFeedback%>&IDOrder=<%=clng(scpre)+clng(LngIdOrder)%>"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_f")%></a>
      <%else%>
      	&nbsp;
      <%end if%>
      </td>
    </tr>
	</table>
	<table class="pcShowContent">    
		<tr>
			<td width="20%" align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_z")%></td>
			<td width="80%"><%=FDetails%></td>
		</tr>
		<%if AllowUpload="1" then
			query="SELECT pcUpld_IDFile,pcUpld_FileName FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & IDComment
			set rstemp=connTemp.execute(query)
			if not rstemp.eof then
			%>
			<tr class="main">
				<td valign="top" colspan="2" class="pcSmallText"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_s")%><br>
				<%Do until rstemp.eof%>
  				<a href="userdownload.asp?IDFile=<%=rstemp("pcUpld_IDFile")%>" target="_blank"><img src="images/DownLoad.gif" border=0 height=19 width=18 alt="Download File"></a>&nbsp;<b><%
					Filename= rstemp("pcUpld_FileName")
					FileName = mid(FileName,instr(Filename,"_")+1,len(FileName))%>
					<%=FileName%></b><br>
 				 	<%rstemp.MoveNext
  			loop
  			%>
				</td>
			</tr>
		<%end if
	end if%>
</table>
<hr>
<%rs.MoveNext
Loop%>
		<p>&nbsp;</p>
		<script language="JavaScript">
		<!--
			
		function Form1_Validator(theForm)
		{
		
			if (theForm.comments.value == "")
			{
						alert("<%response.write dictLanguage.Item(Session("language")&"_editFeedback_h")%>");
						theForm.comments.focus();
						return (false);
			}
			
		return (true);
		}
		
		function newWindow(file,window) {
				msgWindow=open(file,window,'resizable=no,width=400,height=500');
				if (msgWindow.opener == null) msgWindow.opener = self;
		}
		//-->
		</script>

			<form name="hForm" method="post" action="userviewfeedback.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
			<table class="pcShowContent">
				<tr class="pcSectionTitle">
					<td colspan="2">
					<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_1")%>
					</td>
				</tr>
				<%if AllowUpload="1" then%>
					<tr>
						<td colspan="2" align="left">
						<ol>
							<li><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_2")%></li>
							<li><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_3")%></li>
						</ol>
						</td>
					</tr>
				<%else%>
				<tr>
					<td colspan="2" align="left">
						<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_4")%>
					</td>
				</tr>
				<%end if%>
				<tr>
					<td colspan="2" class="pcSpacer"></td>
				</tr>
			<%if AllowUpload="1" then%>
				<tr>
					<td align="right" valign="top" nowrap>
					<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_5")%>
					</td>
					<td valign="top">
				<%query="SELECT pcUpld_IDFile,pcUpld_FileName FROM pcUploadFiles WHERE pcUpld_IDFeedback=0"
				set rs=connTemp.execute(query)
				if rs.eof then%>
					<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_6")%><br>
				<%else
				ACount=0
				do while not rs.eof
				ACount=ACount+1
					%>
					<input type="hidden" name="AID<%=ACount%>" value="<%=rs("pcUpld_IDFile")%>">
					<input type="checkbox" name="AC<%=ACount%>" value="1" checked class="clearBorder">
					<%
					Filename= rs("pcUpld_FileName")
					FileName = mid(FileName,instr(Filename,"_")+1,len(FileName))%>
					<%=FileName%></font><br>
					<%rs.MoveNext
				loop%>
				<input type="hidden" name=ACount value="<%=ACount%>">
			<%end if%>
			<script language="JavaScript"><!--
				function newWindow1(file,window) {
				catWindow=open(file,window,'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360');
				if (catWindow.opener == null) catWindow.opener = self;
				}
			//-->
			</script>
			<br><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_7")%><a href="#" onClick="javascript:newWindow1('userfileuploada_popup.asp?IDFeedback=0&ReLink=<%=Server.URLencode("userviewfeedback.asp?IDOrder=" & clng(scpre)+clng(LngIdOrder) & "&IDFeedback=" & LngIDFeedback)%>','window2')"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_8")%></a>.
			</td>
			</tr>
			<tr>
				<td colspan="2" class="pcSpacer"></td>
			</tr>
			<%end if%>
			<tr>
				<td valign="top">
					<p align="right"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_9")%>
				<br><br>
				<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=comments','window2')">
				</p>
				</td>
			<td valign="top">				
				<p align="left">
					<input type="hidden" name="IDOrder" value="<%=clng(scpre)+LngIdOrder%>">
					<input type="hidden" name="IDFeedback" value="<%=LngIDFeedback%>">
					<textarea name="comments" cols="40" rows="8"><%=request("comments")%></textarea>
				</p> 
			</td>
			</tr>
			<tr>
				<td colspan="2" class="pcSpacer"></td>
			</tr>
			<tr>
				<td colspan="2" align="center">
				<input type="submit" name="Submit" value="Add Comment" class="submit2" onclick="document.hForm.rewrite.value='0';">
				<input type="hidden" name="uploaded" value="">
				<input type="hidden" name="rewrite" value="1">
				</td>
			</tr>
			<tr>
				<td colspan="2" class="pcSpacer"></td>
			</tr>
		</table>
	</form>
	<p>&nbsp;</p>
	<p align="center">
  <a href="userviewallposts.asp?IDOrder=<%=clng(scpre)+clng(LngIdOrder)%>"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_10")%></a>
	</p>
	<iframe name="downloadwindow" src="about:blank" marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 width=0 height=0 scrolling="no" noresize></iframe>
</td>
</tr>
</table>
</div>
<%call closedb()%><!-- #Include File="footer.asp" -->