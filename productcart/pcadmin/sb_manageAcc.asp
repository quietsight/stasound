<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Dim pageTitle, pageName, pageIcon, Section
pageTitle="Activate SubscriptionBridge Integration"
pageName="sb_manageAcc.asp"
pageIcon="pcv4_icon_sb.png"
Section="SB" 
%>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="sb_inc.asp"-->
<% Dim connTemp,query,rs

call opendb()

on error goto 0			

if request("action")="reg" then

	tmp_setup=0
	
	Setting_APIUser=request("Setting_APIUser")
	Setting_APIPassword=request("Setting_APIPassword")
	Setting_APIKey=request("Setting_APIKey")
	
	msg=0
	
	if Setting_APIUser<>"" AND Setting_APIPassword<>"" AND Setting_APIKey<>"" then
		
		'// Register Account
		Dim objSB 
		Set objSB = NEW pcARBClass
		result=objSB.RegisterAcc(Setting_APIUser, Setting_APIPassword, Setting_APIKey)
		
		if result="1" and SB_ErrMsg="" then

			Setting_APIPassword1=enDeCrypt(Setting_APIPassword, scCrypPass)
			Setting_APIKey1=enDeCrypt(Setting_APIKey, scCrypPass)

			
			query="SELECT Setting_ID FROM SB_Settings;"
			set rs=connTemp.execute(query)
			if rs.eof then
				query="INSERT INTO SB_Settings (Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess) VALUES ('" & Setting_APIUser & "','" & Setting_APIPassword1 & "','" & Setting_APIKey1 & "',1);"
			else
				query="UPDATE SB_Settings SET Setting_APIUser='" & Setting_APIUser & "',Setting_APIPassword='" & Setting_APIPassword1 & "',Setting_APIKey='" & Setting_APIKey1 & "',Setting_RegSuccess=1;"
			end if
			set rs=nothing
			set rs=connTemp.execute(query)
			set rs=nothing

			if SB_ErrMsg="" then
				response.redirect "sb_Default.asp?msg=1"
			else

				query="UPDATE SB_Settings SET Setting_RegSuccess=0;"
				set rs=connTemp.execute(query)
				set rs=nothing

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
	Setting_APIUser=""
	Setting_APIPassword=""
	Setting_APIKey=""

	query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		Setting_APIUser=rs("Setting_APIUser")
		Setting_APIPassword=enDeCrypt(rs("Setting_APIPassword"), scCrypPass)
		Setting_APIKey=enDeCrypt(rs("Setting_APIKey"), scCrypPass)
		tmp_setup=rs("Setting_RegSuccess")
		if IsNull(tmp_setup) or tmp_setup="" then
			tmp_setup=0
		end if
	end if
	set rs=nothing

end if
			if err.number <> 0 then
				response.Write("Where:  " & err.description)
				response.End()
			end if
%>
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<form name="form1" action="sb_manageAcc.asp?action=reg" method="post" class="pcForms" onsubmit="javascript:pcf_Open_SubscriptionBridge();">
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<%if tmp_setup="1" then%>
	<tr>
		<td colspan="2">
			<div class="pcCPmessageSuccess">This store is <strong>already communicating</strong> with SubscriptionBridge. It is using the API credentials listed below.</div>
		</td>
	</tr>
	<%end if%>
	<%if msg<>0 then%>
	<tr>
		<td colspan="2">
			<div class="pcCPmessage">
				<%Select Case msg
				Case 1:%>
					You must enter an API username, password, and key.
				<%Case 2:%>
					We cannot activate the link to your SubscriptionBridge account at this time. <br /> Server Error Message: <b><%=SB_ErrMsg%></b>
				<%Case 3:%>
					Your store has been successfully integrated with SubscriptionBridge.<br> However, ProductCart cannot currently connect to the SubscriptionBridge Web Services.<br>Server Error Message: <b><%=SB_ErrMsg%></b>
				<%End Select%>
			</div>
		</td>
	</tr>
	<%end if%>
	<tr>
		<th colspan="2">Enter SubscriptionBridge API Credentials</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>API Username:</td>
		<td><input type="text" name="Setting_APIUser" value="<%=Setting_APIUser%>" size="40"></td>
	</tr>
	<tr>
		<td>API Password:</td>
		<td><input type="text" name="Setting_APIPassword" value="<%=Setting_APIPassword%>" size="40"></td>
	</tr>
	<tr>
		<td>API Key:</td>
		<td><input type="text" name="Setting_APIKey" value="<%=Setting_APIKey%>" size="60"> </td>
	</tr>
	<tr>
		<td colspan="2">
        	<hr />
			<input type="submit" name="submit1" value="Register" class="submit2">&nbsp;
			<input type="button" name="Back" value="Back" onClick="location='sb_Default.asp';" class="ibtnGrey">
		</td>
	</tr>
</table>
</form>
<%
call closedb()
Response.write(pcf_ModalWindow("Please wait... contacting SubscriptionBridge.","SubscriptionBridge", 300))%>
<!--#include file="AdminFooter.asp"-->