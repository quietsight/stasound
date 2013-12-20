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
<!--#include file="../includes/languages.asp"-->
<% pageTitle="Setup/Manage MailUp Account" %>
<% section="mngAcc"
Dim connTemp,query,rs
Dim tmp_setup

if request("action")="reg" then

	tmp_setup=0
	
	pcMailUpSett_APIUser=request("pcMailUpSett_APIUser")
	pcMailUpSett_APIPassword=request("pcMailUpSett_APIPassword")
	pcMailUpSett_URL=request("pcMailUpSett_URL")
		'// Clean up URL
		tempURL=replace(pcMailUpSett_URL,"http://","")
		tempURL=replace(tempURL,"https://","")
		pcMailUpSett_URL = tempURL
	
	msg=0
	
	if pcMailUpSett_APIUser<>"" AND pcMailUpSett_APIPassword<>"" AND pcMailUpSett_URL<>"" then
		tmpReg=RegisterAcc(pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL)
		if tmpReg=1 and MU_ErrMsg="" then
			session("CP_MU_APIUser")=pcMailUpSett_APIUser
			session("CP_MU_APIPassword")=pcMailUpSett_APIPassword
			session("CP_MU_URL")=pcMailUpSett_URL
			pcMailUpSett_APIPassword1=enDeCrypt(pcMailUpSett_APIPassword, scCrypPass)
			call opendb()
			query="SELECT pcMailUpSett_ID FROM pcMailUpSettings;"
			set rs=connTemp.execute(query)
			if rs.eof then
				query="INSERT INTO pcMailUpSettings (pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess) VALUES ('" & pcMailUpSett_APIUser & "','" & pcMailUpSett_APIPassword1 & "','" & pcMailUpSett_URL & "',1,1);"
			else
				query="UPDATE pcMailUpSettings SET pcMailUpSett_APIUser='" & pcMailUpSett_APIUser & "',pcMailUpSett_APIPassword='" & pcMailUpSett_APIPassword1 & "',pcMailUpSett_URL='" & pcMailUpSett_URL & "',pcMailUpSett_RegSuccess=1;"
			end if
			set rs=nothing
			set rs=connTemp.execute(query)
			set rs=nothing
			call closedb()
			call opendb()
			tmpGetList=GetMUList(session("CP_MU_APIUser"),session("CP_MU_APIPassword"),session("CP_MU_URL"))
			call closedb()
			if tmpGetList=1 and MU_ErrMsg="" then
				response.redirect "mu_manageNewsWiz.asp?msg=1"
			else
				call opendb()
				query="UPDATE pcMailUpSettings SET pcMailUpSett_RegSuccess=0;"
				set rs=connTemp.execute(query)
				set rs=nothing
				call closedb()
				msg=3
			end if
		else
			msg=2
		end if
	else
		msg=1
	end if
else
	tmp_setup=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	call opendb()
	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_RegSuccess FROM pcMailUpSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcMailUpSett_APIUser=rs("pcMailUpSett_APIUser")
		pcMailUpSett_APIPassword=enDeCrypt(rs("pcMailUpSett_APIPassword"), scCrypPass)
		pcMailUpSett_URL=rs("pcMailUpSett_URL")
		tmp_setup=rs("pcMailUpSett_RegSuccess")
		if IsNull(tmp_setup) or tmp_setup="" then
			tmp_setup=0
		end if
	end if
	set rs=nothing
	call closedb()
end if
%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<form name="form1" action="mu_manageAcc.asp?action=reg" method="post" class="pcForms" onsubmit="javascript:pcf_Open_MailUp();">
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<%if tmp_setup="1" then%>
	<tr>
		<td colspan="2">
			<div class="pcCPmessage">You already registered a MailUp Account.</div>
		</td>
	</tr>
	<%end if%>
	<%if msg<>0 then%>
	<tr>
		<td colspan="2">
				<%Select Case msg
				Case 1:%>
				<div class="pcCPmessage">You have to enter your full MailUp console credentials</div>
				<%Case 2:%>
				<div class="pcCPmessage">Your MailUp account could not be registered<br>Server Error Message: <b><%=MU_ErrMsg%></b></div>
				<%Case 3:%>
				<div class="pcCPmessage">Your account has been registered successfully.<br>However, ProductCart cannot connect to the MailUp Web Service.<br>Server Error Message: <b><%=MU_ErrMsg%></b></div>
				<%End Select%>
			</div>
		</td>
	</tr>
	<%end if%>
	<tr>
		<th colspan="2">Enter MailUp account information</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>MailUp API Username:</td>
		<td><input type="text" name="pcMailUpSett_APIUser" value="<%=pcMailUpSett_APIUser%>" size="30"></td>
	</tr>
	<tr>
		<td>MailUp API Password:</td>
		<td><input type="password" name="pcMailUpSett_APIPassword" value="<%=pcMailUpSett_APIPassword%>" size="30"></td>
	</tr>
	<tr>
		<td>MailUp Console URL:</td>
		<td>http:// <input type="text" name="pcMailUpSett_URL" value="<%=pcMailUpSett_URL%>" size="40"> <span class="smallText">(e.g. myCompany.mailupnet.it)</span></td>
	</tr>
	<tr>
		<td colspan="2">
			<br>
			<br>
			<input type="submit" name="submit1" value="Register" class="submit2">&nbsp;
			<input type="button" name="Back" value="Back" onClick="location='mu_manageNewsWiz.asp';" class="ibtnGrey">
			</td>
	</tr>
</table>
</form>
<%Response.write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_MailUp_SynNote2"),"MailUp", 300))%><!--#include file="AdminFooter.asp"-->