<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=0%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim rstemp, conntemp, query

'Start SDBA
if request("pagetype")="1" then
	pcv_PageType="1"
	pcv_Title="Drop-Shippers"
else
	if request("pagetype")="0" then
		pcv_PageType="0"
		pcv_Title="Suppliers"
	else
		pcv_PageType=""
		pcv_Title="Customers"
	end if
end if
'End SDBA

if (request("action")<>"add") and ((session("AllowUsing")="1") or (request("using")<>""))  then
	if (session("usingM")<>"") and (session("AllowUsing")="1") then
		mUsing=session("usingM")
		session("AllowUsing")="0"
	end if
	if request("using")<>"" then
		mUsing=request("using")
		session("usingM")=mUsing
	end if

	call opendb()
	query="select * from News where IDNews=" & mUsing
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)
	if not rstemp.eof then
		session("News_FromEmail")=rstemp("FromEmail")
		session("News_FromName")=rstemp("FromName")
		session("News_Title")=rstemp("Title")
		session("News_MsgBody")=rstemp("MsgBody")
		session("News_MsgType")=rstemp("MsgType")
	end if
	set rstemp=nothing
	call closeDb()
end if

if request("action")="add" then
	session("News_FromEmail")=request("FromEmail")
	session("News_FromName")=request("FromName")
	session("News_Title")=request("Title")
	session("News_MsgBody")=request("Details")
	session("News_MsgType")=request("MType")
	response.redirect "newsWizStep4.asp?pagetype=" & pcv_PageType
end if

%>
<% pageTitle="Newsletter Wizard - STEP 3: Enter Message" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--
	
function Form1_Validator(theForm)
{

	if (theForm.fromName.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.fromName.focus();
		    return (false);
	}
	
		if (theForm.fromEmail.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.fromEmail.focus();
		    return (false);
	}

	if (theForm.Title.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.Title.focus();
		    return (false);
	}
	
	if (theForm.details.value == "")
 	{
		    alert("Please enter a value for this field.");
		    theForm.details.focus();
		    return (false);
	}	

return (true);
}
//-->
</script>
<form name="hForm" method="post" action="mu_newsWizStep3.asp?action=add&pagetype=<%=pcv_PageType%>" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2">
		<img src="images/pc2008_MailUp_Wizard.gif" align="Newsletter Wizard - MailUp Integration" style="margin-bottom: 10px;" />
	</td>
</tr>
<tr>
	<td colspan="2">
	<strong>NOTE</strong>: The ProductCart Newsletter Wizard is not intended to handle large email lists. We recommend that you use MailUp or another professional newsletter management tool instead (<a href="mu_newsWizStep2.asp">go back</a>). Messages are sent one by one, to avoid exceeding limitations to the number of concurrent receipients that may be in place on your Web server's mail server. In our tests, sending a message typically took between 1 and 2 seconds each. Therefore, sending a newsletter to 100 recipients should take about 3 minutes.
	</td>
</tr>
<tr>
	<td colspan="2"><strong>NO SPAM</strong>: DO NOT USE this feature to send SPAM email. Regardless of whether spam is considered illegal in your Country (e.g. <a href="http://www.ftc.gov/bcp/conline/pubs/buspubs/canspam.shtm" target="_blank">CAN-SPAM Act in the US</a>), sending unsolicited messages is certainly not what this feature is meant for. It is also a counter-productive marketing practice.</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Enter the message that you want to send to <%=pcv_Title%> in the form below:</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<%
call openDb()
query="select idNews,title from News"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
if not rstemp.eof then
%>
<tr>
	<td>Copy from previously sent messages: </td>
	<td>
	<select size="1" name="SentMList" onchange="if (document.hForm.SentMList.value != '') location='mu_newsWizStep3.asp?pagetype=<%=pcv_PageType%>&using='+document.hForm.SentMList.value;">
		<option value="">Select sent message</option>
		<%do while not rstemp.eof%>
			<option value="<%=rstemp("IDNews")%>"><%=rstemp("Title")%></option>
		<%rstemp.movenext
		loop%>
	</select>
	</td>
</tr>
<%
end if
set rstemp=nothing
call closeDb()
%>
<tr>
	<td>From Name:</td>
	<td><input type="text" name="fromName" size="43" value="<%=session("News_FromName")%>"></td>
</tr>
<tr>
	<td>From Email:</td>
	<td><input type="text" name="fromEmail" size="43" value="<%=session("News_FromEMail")%>"></td>
</tr>
<tr>
	<td>Subject:</td>
	<td><input type="text" name="Title" size="43" value="<%=session("News_Title")%>"></td>
</tr>
<tr>
	<td valign="top"><script language="JavaScript"><!--
				function newWindow(file,window) {
						msgWindow=open(file,window,'resizable=no,width=400,height=500');
						if (msgWindow.opener == null) msgWindow.opener = self;
				}
				//--></script>
		Message:
        <div style="margin-top: 10px;">                           
		<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp','window2')">
        </div>
	</td>
	<td>
		<textarea name="details" cols="80" rows="10"><%=session("News_Msgbody")%></textarea>
	</td>
</tr>
<tr>
	<td>Send as:</td>
	<td>
	<input type="radio" value="0" name="MType" <%if session("News_MsgType")<>"1" then%>checked<%end if%>>
	Plain Text
	<input type="radio" value="1" name="MType" <%if session("News_MsgType")<>"1" then%><%else%>checked<%end if%>>
	HTML</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>NOTE: if you are using CDONTS your store will <u>always send text messages</u>. Check your e-mail settings to see if other components are supported on your server.</td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<tr>
	<td align="center" width="531" colspan="2">
	<input type="submit" name=submit value="Continue" class="submit2">
	&nbsp;
	<input type="button" name="back" value="Back" onClick="location='mu_newsWizStep2.asp?pagetype=<%=pcv_PageType%>'">
</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->