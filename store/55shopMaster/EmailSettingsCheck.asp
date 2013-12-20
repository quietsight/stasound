<%@Language="VBScript"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
<TITLE>E-mail Settings Test</TITLE>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
<style>
	.pcCPcontent td {
		font-size: 12px;
	}
</style>
</HEAD>
<body style="background-image: none;">
<form action="EmailSettingsCheck.asp" method="post" class="pcForms">
<input type="hidden" name="pcFormAction" value="send">
<table class="pcCPcontent" style="width: 100%;">
    <tr> 
        <th colspan="2">
        <div style="float: right; margin: 0 10px 0 0;"><a href=# class="pcSmallText" onClick="self.close();">Close</a></div>
        Test Your E-mail Settings
      </th>
    </tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>

<% dim pcv_fromname, pcv_fromemail, pcv_toname, pcv_toemail, pcv_subject, pcv_message, pcv_success

pcv_FormAction=request.Form("pcFormAction")

if pcv_FormAction = "send" then

	pcv_fromname=scCompanyName
	pcv_fromemail=getUserInput(request.form("pcFromEmail"),0)
	pcv_toname="Store Administrator"
	pcv_toemail=getUserInput(request.form("pcAdminEmail"),0)
	pcv_message=getUserInput(request.form("pcEmailTestMessage"),0)
	pcv_subject="ProductCart Email Settings Test Message" & " - " & replace(scCompanyName,"'","") 
	pcv_errMsg=""
	
	call sendmail (pcv_fromname, pcv_fromemail, pcv_toemail, pcv_subject, pcv_message)
	if pcv_errMsg<>"" then
		pcv_err = InStr(1,pcv_errMsg,"Object required",1) %>
			 <tr> 
				<td><img src="images/pcv4_icon_alert.gif"></td>
                <td>
				<% if pcv_err > 0 then %>
					You have selected an email component that is not supported on this server
				<% else
					response.write pcv_errMsg
				end if %>					
				</td>
			</tr>
	<% else %>
             <tr> 
                <td colspan="2">
                The message was sent. Check your email to make sure you have received it successfully.
                <div style="margin-top: 20px;"><a href=# onClick="self.close();">Close</a></div>
                </td>
            </tr>
	<% end if	
else %>

		<tr> 
			<td align="right">Email Component:</td>
			<td><%=scEmailComObj%></td>
		</tr>
		<tr> 
            <td align="right">SMTP Server:</td>
			<td><%=scSMTP%></td>
		</tr>
		<tr> 
            <td align="right">From Email:</td>
			<td><input type="text" value="<%=scEmail%>" name="pcFromEmail"></td>
		</tr>
		<tr> 
            <td align="right">From Name:</td>
			<td><%=scCompanyName%></td>
		</tr>
		<tr> 
			<td align="right">Admin Email:</td>
			<td><input type="text" value="<%=scFrmEmail%>" name="pcAdminEmail"></td>
		</tr>
		<tr> 
			<td align="right">Message:</td>
			<td></td>
		</tr>
		<tr> 
			<td align="center" colspan="2"><textarea name="pcEmailTestMessage" cols="30" rows="5">If you receive this message it means that your store is successfully sending messages using: <%=scEmailComObj%>. The SMTP server is: <%=scSMTP%>.</textarea></td>
		</tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
		<tr> 
			<td colspan="2" align="center">
            <input type="submit" value="Send Test Message" id="searchSubmit">
            </td>
		</tr>
<% end if %>
</table>
</form>
</BODY>
</HTML>