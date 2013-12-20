<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="7*9*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!-- #Include file="../pc/checkdate.asp" -->
<% Dim pageTitle, Section
pageTitle="Post new message and files"
Section="orders" %>
<!-- #Include File="Adminheader.asp" -->
<%

Dim rs, connTemp, query
call openDB()

session("cfrom")=1

dim intIdOrder
intIdOrder=getUserInput(request("IDOrder"),0)
if validNum(intIdOrder) and intIdOrder<>"0" then
	session("admin_IDOrder")=Clng(intIdOrder)
else
	intIdOrder=0
end if

'Create new feedback
if (request("action")="add") and (request("rewrite")="0") then
	Dim strFDesc, strFDetails, intFStatus, intFType, intPriority
	strFDesc=getUserInput(request("Description"),0)
	strFDetails=getUserInput(request("Details"),0)
	intFStatus=getUserInput(request("FStatus"),0)
	FType=getUserInput(request("FType"),0)
	intPriority=getUserInput(request("Priority"),0)
	
	query="Select pcPri_name from pcPriority where pcPri_IDPri=" & intPriority
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	dim strFPriority
	strFPriority=rs("pcPri_name")
	set rs = nothing
	
	dtComDate=CheckDateSQL(now())
	
	if scDB="SQL" then	
		query="Insert Into pcComments (pcComm_IDOrder,pcComm_IDParent,pcComm_IDUser,pcComm_CreatedDate,pcComm_EditedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details) values (" & intIdOrder & ",0,0,'" & dtComDate & "','" & dtComDate & "'," & FType & "," & intFStatus & "," & intPriority & ",'" & strFDesc & "','" & strFDetails & "')"
	else
		query="Insert Into pcComments (pcComm_IDOrder,pcComm_IDParent,pcComm_IDUser,pcComm_CreatedDate,pcComm_EditedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details) values (" & intIdOrder & ",0,0,#" & dtComDate & "#,#" & dtComDate & "#," & FType & "," & intFStatus & "," & intPriority & ",'" & strFDesc & "','" & strFDetails & "')"
	end if
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	
	query="select pcComm_IDFeedback from pcComments where pcComm_IDParent=0 and pcComm_IDOrder=" & intIdOrder & " and pcComm_IDUSer=0 ORDER BY pcComm_IDFeedback DESC;"
	set rs=connTemp.execute(query)
	
	Dim strMsg, r
	r=0
	if rs.eof then
		strMsg=dictLanguage.Item(Session("language")&"_addFB_s")
	else
		Dim intLastFB
		intLastFB=rs("pcComm_IDFeedback")
		set rs=nothing
	
		'Generate View Feedback Link for Customer
		dim strPath, iCnt, strPathInfo
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
		dim strURL
		strURL=strPathInfo & "pc/Checkout.asp?cmode=1&redirectUrl=" & Server.URLEnCode(strPathInfo & "pc/userviewfeedback.asp?IDOrder=" & scpre+clng(intIdOrder) & "&IDFeedback=" & intLastFB)

			ACount=getUserInput(request("ACount"),0)
			if ACount<>"" then
				ACount1=clng(ACount)
				For k=1 to ACount1
					if request("AC" & k)="1" then
						query="update pcUploadFiles set pcUpld_IDFeedback=" & intLastFB & " where pcUpld_IDFile=" & getUserInput(request("AID" & k),0) & " and pcUpld_IDFeedback=0"
						set rs=connTemp.execute(query)
					end if
				next
				query="delete from pcUploadFiles where pcUpld_IDFeedback=0"
				set rs=connTemp.execute(query)
			end if

		'Send mail to customer
		Dim strMsgBody
		strMsgBody=""
		strMsgBody=dictLanguage.Item(Session("language")&"_addFB_email1") & scpre+clng(intIdOrder) & dictLanguage.Item(Session("language")&"_addFB_email2") & vbcrlf & vbcrlf
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email3") & scpre+clng(intIdOrder) & vbcrlf
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email4") & strFDesc & vbcrlf

		AdminName= dictLanguage.Item(Session("language")&"_viewPostings_2")

		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email5") & AdminName & vbcrlf
			
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email6") & strFPriority & vbcrlf
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email7") & vbcrlf
		strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email8") & strURL &VBCrlf&VBCrlf
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

		msgType=1
		msg="The message was posted successfully. <a href=adminviewallmsgs.asp?IDOrder=" & intIdOrder & ">View other messages</a> associated with this order."
	end if %>
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
<%
end if

if request("uploaded")<>"" then
	session("uploaded")="1"
else
	session("uploaded")="0"
end if


'Delete Temponary uploaded files
if request("k")="1" then
else
	if session("uploaded")="1" then
		session("uploaded")="0"
	else
		query="Select pcUpld_Filename from pcUploadFiles where pcUpld_IDFeedback=0"
		set rs=connTemp.execute(query)
		dim strFilename, strQfilePath, findit, fso, f
		do while not rs.eof
			strFilename=rs("pcUpld_Filename")
			if strFilename<>"" then
				strQfilePath="../pc/Library/" & strFilename
				findit = Server.MapPath(strQfilePath)
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
end if 
%>
<script language="JavaScript">
<!--
	
function Form1_Validator(theForm)
{

		if (theForm.FType.value == "")
 	{
		    alert("Please select a feedback type.");
		    theForm.FType.focus();
		    return (false);
	}
	
			if (theForm.Priority.value == "")
 	{
		    alert("Please select a priority.");
		    theForm.Priority.focus();
		    return (false);
	}
	
	
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
IF intIdOrder=0 THEN
%>
<div style="margin: 20px;">
	Please select the order for which you want to open a Help Desk ticket:

   <form name="selectOrder" method="get" action="adminaddfeedback.asp">	
    <%
	query="SELECT IDOrder FROM Orders WHERE orderStatus>1 ORDER by IDOrder DESC"
	set rs=server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if not rs.eof then
	%>
		<select name="IDOrder" onChange="this.form.submit();">
        <option value="0" selected>Select the order number...</option>
		<%
		do while not rs.eof
		%>
		<option value="<%=rs("IDOrder")%>"><%=clng(scpre)+clng(rs("IDOrder"))%></option>
		<%
		rs.MoveNext
		loop
		%>
		</select>
	<%
    else
	%>
    There are no orders yet in this store.
    <%
	end if
	%>
	</form>
</div>
<%
ELSE
'// There is an order ID already
'// Get customer id
pidcustomer=getCustIDfromOrder(intIdOrder)
%>

<form name="hForm" method="post" action="adminaddfeedback.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<input type="hidden" value="<%=intIdOrder%>" name="IDOrder">

<script language="JavaScript"><!--
function newWindow(file,window) {
		msgWindow=open(file,window,'resizable=no,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
}
//-->
</script>

			<table class="pcCPcontent" style="width:auto;">
				<tr>
					<td colspan="2"><h2>You are posting a message for <strong>Order #<%=clng(scpre)+intIdOrder%></strong></h2></td>
                </tr>
                <tr>
                    <td colspan="2">
                    <ol>
                        <li><%response.write dictLanguage.Item(Session("language")&"_addFB_c")%></li>
                        <li><%response.write dictLanguage.Item(Session("language")&"_addFB_d")%></li>
                    </ol>
                    </td>
                </tr>
					<tr>
						<td nowrap width="25%" align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_addFB_f")%></td>
						<td width="75%" valign="top">
						<%
						query="Select pcUpld_IDFile,pcUpld_FileName from pcUploadFiles where pcUpld_IDFeedback=0"
						set rs=connTemp.execute(query)
						if rs.eof then
						%>
						<%response.write dictLanguage.Item(Session("language")&"_addFB_g")%><br>
					<% else
						ACount=0
						do while not rs.eof
							ACount=ACount+1
							pc_pcUpld_IDFile=rs("pcUpld_IDFile")
							pc_pcUpld_FileName= rs("pcUpld_FileName") %>
							<input type="hidden" name="AID<%=ACount%>" value="<%=pc_pcUpld_IDFile%>">
							<input type="checkbox" name="AC<%=ACount%>" value="1" checked class="clearBorder">&nbsp;
							<%
							strFilename= pc_pcUpld_FileName
							strFilename = mid(strFilename,instr(strFilename,"_")+1,len(strFilename))%>
							<%=strFilename%>
							<br />
							<%rs.MoveNext
						loop%>
						<input type="hidden" name="ACount" value="<%=ACount%>">
					<%end if%>
						<script language="JavaScript"><!--
							function newWindow1(file,window) {
							catWindow=open(file,window,'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360');
							if (catWindow.opener == null) catWindow.opener = self;
							}
						//--></script>
						<br>
						<%response.write dictLanguage.Item(Session("language")&"_addFB_h")%>
						<a href="#" onClick="javascript:newWindow1('adminfileuploada_popup.asp?IDFeedback=0&ReLink=<%=Server.URLencode("adminaddfeedback.asp?d=1")%>','window2')"><%response.write dictLanguage.Item(Session("language")&"_addFB_i")%></a>
					</td>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer">
                    <hr>
					<%
					'Default Status = Open
					strTemp=""
					query="Select pcFStat_idstatus,pcFStat_name from pcFStatus"
					set rs=connTemp.execute(query)
						do while not rs.eof
							IDStatus=rs("pcFStat_idstatus")
							SName=ucase(rs("pcFStat_name"))
							if SName="OPEN" then
								strTemp="" & IDStatus
							end if
							rs.MoveNext
						Loop
					if strTemp="" then
						strTemp="1"
					end if%>
					<input type="hidden" name=FStatus value="<%=strTemp%>">
                    
                    </td>
				</tr>
  				<tr>
					<td align="right"><%response.write dictLanguage.Item(Session("language")&"_addFB_k")%></td>
					<td>
					<select name="FType">
						<option value=""></option>
						<% query="Select pcFType_idtype,pcFType_name from pcFTypes"
					 	set rs=connTemp.execute(query)
					 	do while not rs.eof %>
   						<option value="<%=rs("pcFType_idtype")%>" <%if request.form("FType")<>"" then%><%if clng(request("FType"))=clng(rs("pcFType_idtype")) then%>selected<%end if%><%end if%>><%=rs("pcFType_name")%></option>
   						<%rs.MoveNext
   					Loop%>
					</select>
					</td>
                </tr>
                <tr>
					<td align="right"><%response.write dictLanguage.Item(Session("language")&"_addFB_l")%></td>
					<td>
					<select name="Priority">
						<option value=""></option>
						<% query="Select pcPri_idPri,pcPri_name from pcPriority"
						set rs=connTemp.execute(query)
						do while not rs.eof %>
							<option value="<%=rs("pcPri_idPri")%>" <%if request("Priority")<>"" then%><%if clng(request("Priority"))=clng(rs("pcPri_idPri")) then%>selected<%end if%><%end if%>><%=rs("pcPri_name")%></option>
							<%
						rs.MoveNext
						Loop
						set rs=nothing
						%>
					</select> 
					</td>
				</tr>
				<tr>
					<td align="right"><%response.write dictLanguage.Item(Session("language")&"_addFB_q")%></td>
					<td><input name="Description" type="text" id="bugShortDsc" size="25" maxlength="100" value="<%=request("Description")%>"></td>
				</tr>
				<tr>
					<td align="right" valign="top"><%response.write dictLanguage.Item(Session("language")&"_addFB_r")%>
					<br /><br />
					<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=Details','window2')">
					</td>
					<td><textarea name="Details" cols="60" rows="7" id="bugLongDsc"><%=request("Details")%></textarea></td>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<td colspan="2" align="center">
                    <%
					'// Hide ability to post to Help Desk if customer is a Guest
					if pcf_GetCustType(pidcustomer)=0 then
					%>
						<input type="submit" name="Submit" value="Add Feedback" class="submit2" onclick="document.hForm.rewrite.value='0';">
						&nbsp;<input type="button" name="back" value=" View all Postings" onClick="location='adminviewallmsgs.asp';">
					 	<%if session("admin_IDOrder")>0 then%>
					 		<input type="button" name="go" value="View Postings" onClick="location='adminviewallmsgs.asp?IDOrder=<%=session("admin_IDOrder")%>';">
						<%end if%>
						<input type="hidden" name="uploaded" value="">
						<input type="hidden" name="rewrite" value="1">
                    
                    <%
					else
					%>
                    	<div class="pcCPmessage">The Help Desk is disabled for this order: this customer is a &quot;Guest&quot; and would not be able to view/reply to the ticket. <a href="modcusta.asp?idcustomer=<%=pidcustomer%>" target="_blank">View the customer details page</a> to change the customer status.</div>
                    <%
					end if
					%>
					</td>
				</tr>
			</table>
</form>
<%
END IF
call closeDb()
%>
<!-- #Include File="Adminfooter.asp" -->