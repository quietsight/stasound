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
<% pageTitle="MailUp Newsletter Management System" %>
<% section="mngAcc"
Dim connTemp,query,rs
Dim tmp_setup

tmp_setup=0

pcv_pageType=request("pagetype")

call opendb()
if request("a")<>"" then
	if request("a")="on" then
		query="UPDATE pcMailUpSettings SET pcMailUpSett_RegSuccess=1,pcMailUpSett_TurnOff=0;"
	else
		query="UPDATE pcMailUpSettings SET pcMailUpSett_RegSuccess=0,pcMailUpSett_TurnOff=1;"
	end if
	set rs=connTemp.execute(query)
	set rs=nothing
end if

query="SELECT pcMailUpSett_RegSuccess,pcMailUpSett_TurnOff FROM pcMailUpSettings;"
set rs=connTemp.execute(query)
	if err.number<>0 then
		set rs = nothing
		call closedb()
		response.Redirect("upddb_MailUp.asp")
	end if
if not rs.eof then
	tmp_setup=rs("pcMailUpSett_RegSuccess")
	tmp_TurnOff=rs("pcMailUpSett_TurnOff")
end if
set rs=nothing
call closedb()
%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">ProductCart - MailUp Integration: Home</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<%if request("msg")<>"" then%>
	<tr>
		<td>
				<%Select Case request("msg")
				Case "1":%>
					<div class="pcCPmessageSuccess">Your ProductCart-powered store has been successfully registered with your MailUp console. You can now exchange data between the two system.</div>
				<%End Select%>
		</td>
	</tr>
	<%end if%>
	<%if tmp_TurnOff="1" then%>
	<tr>
		<td>
			<p><img src="images/MailUp_200.jpg" alt="MailUp Newsletter Management System" align="right" style="margin: 10px 10px 10px 30px;" />MailUp is a professional e-mail &amp; SMS message management system. It allows you to efficiently broadcast and track messages sent to e-mail or mobile phone recipients. The integration with ProductCart connects your customer database - and your customer newsletter subscription preferences - with your MailUp console.</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<div class="pcCPmessageInfo">MailUp Integration turned OFF. <br><br><a href="mu_manageNewsWiz.asp?a=on">Turn it ON now</a>.</div>
		</td>
	</tr>
	<%else%>
	<tr>
		<td>
		<p><img src="images/MailUp_200.jpg" alt="MailUp Newsletter Management System" align="right" style="margin: 10px 10px 10px 30px;" />MailUp is a professional e-mail &amp; SMS message management system. It allows you to efficiently broadcast and track messages sent to e-mail or mobile phone recipients. The integration with ProductCart connects your customer database - and your customer newsletter subscription preferences - with your MailUp console.</p>
		<p>&nbsp;</p>
		<ul>
		<li><a href="mu_manageAcc.asp">Setup/Manage MailUp Account</a></li>
		<%if tmp_setup=1 then%>
			<li><a href="mu_settings.asp" onclick="javascript:pcf_Open_MailUp();">Retrieve and Manage Lists</a></li>
			<li><a href="mu_regsyn.asp">Register/Synchronize Customers with MailUp</a></li>
			<li><a href="<%if pcv_pageType<>"" then%>mu_sds_newsWizStep1.asp?pagetype=<%=pcv_pageType%><%else%>mu_newsWizStep1.asp<%end if%>">Export Recipient Group to MailUp</a></li>
		<%else%>
			<li><i>Retrieve and Manage Lists</i></li>
			<li><i>Register/Synchronize Customers with MailUp</i></li>
			<li><i>Export Recipient Group to MailUp</i></li>
		<%end if%>
		<%if tmp_setup="1" then%>
			<li style="margin-top: 15px;"><a href="JavaScript:if(confirm('The store will stop communicating with MailUp. Newsletter sign up settings will revert back to the default newsletter signup feature [under Checkout Options]. Are you sure you want to continue?')) location='mu_manageNewsWiz.asp?a=off'">Turn OFF the MailUp Integration</a></li>
		<%end if%>
			<li style="margin-top: 15px;"><a href="http://www.earlyimpact.com/support/userGuides/mailup.asp" target="_blank">User Guide</a></li>
			<li><a href="http://www.earlyimpact.com/support/userGuides/mailupSupport.asp" target="_blank">Technical Support</a></li>
		</ul>
		</td>
	</tr>
	<%if tmp_setup=0 then%>
	<tr>
		<td>
			<div class="pcCPmessage">All features are disabled until your MailUp account has been activated. <a href="mu_manageAcc.asp">Activate now</a>.</div>
		</td>
	</tr>
	<%end if%>
	<%end if%>
</table>
<%Response.write(pcf_ModalWindow("Importing lists from your MailUp console","MailUp", 300))%>
<!--#include file="AdminFooter.asp"-->