<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=7%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->    
<%

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

if request("action")="test" then

	toEmail=request("toEmail")
	MsgBody=session("News_MsgBody")

	if session("News_MsgType")="1" then
	MsgBody="<html><body>" & MsgBody & "</body></html>"
	end if
	MsgFromName=session("News_FromName")
	MsgFromEmail=session("News_FromEmail")
	MsgTitle=session("News_Title")

	call sendMail(MsgFromName, MsgFromEmail, toEmail,MsgTitle , MsgBody)

	response.redirect "newsWizStep5.asp?from=4&pagetype=" & pcv_PageType

end if

%>
<% pageTitle="Newsletter Wizard: Test Your Message" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--
	
function Form1_Validator(theForm)
{

	if (theForm.toEmail.value == "")
 	{
		    alert("Please enter an e-mail address for your testing.");
		    theForm.toEmail.focus();
		    return (false);
	}
	

return (true);
}
//-->
</script>
<form name="hForm" method="post" action="newsWizStep4.asp?action=test&pagetype=<%=pcv_PageType%>" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2">
		<table width="100%">
		<tr>
			<td width="5%" align="center"><img border="0" src="images/step1.gif"></td>
			<td width="95%"><font color="#A8A8A8">Select <%=pcv_Title%></font></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step2.gif"></td>
			<td><font color="#A8A8A8">Verify <%=pcv_Title%></font></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step3.gif"></td>
			<td><font color="#A8A8A8">Enter message</font></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step4a.gif"></td>
			<td><b>Test message</b></td>
		</tr>
		<tr>
			<td align="center"><img border="0" src="images/step5.gif"></td>
			<td><font color="#A8A8A8">Send message</font></td>
		</tr>
		</table>
	<p>&nbsp;</p>
	</td>
</tr>
<tr>
	<td colspan="2">You can test this message before you send it to your entire <%=pcv_Title%> list. To do so, enter your e-mail address below and click on '<strong>Test Message</strong>' to send it to yourself so that you can review how it looks in your e-mail program(s).</td>
</tr>
<tr>
	<td align="right" nowrap width="10%">E-mail Address:</td>
	<td width="90%">
		<input type="text" name="toEmail" size="50">
	</td>
</tr>
<tr>
	<td align="center" colspan="2">&nbsp;</td>
</tr>
<tr>
	<td align="center" colspan="2">
		<input type="submit" name="submit" value="Test message" class="submit2">&nbsp;
		<input type="button" name="forward" value="Continue Without Testing" onClick="location='newsWizStep5.asp?pagetype=<%=pcv_PageType%>'">&nbsp;
		<input type="button" name="back" value="Back" onClick="location='newsWizStep3.asp?pagetype=<%=pcv_PageType%>'">
	</td>
</tr>
<tr>
	<td align="center" colspan="2">&nbsp;</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->