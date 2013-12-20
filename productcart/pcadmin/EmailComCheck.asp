<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%@Language="VBScript"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%><!--#include file="adminv.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This script was originally created by Richard Kinser.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<%
Dim theComponent(7)
Dim theComponentName(7)

' the components
theComponent(0) = "CDONTS.NewMail"
theComponent(1) = "Bamboo.SMTP"
theComponent(2) = "SMTPsvg.Mailer"
theComponent(3) = "JMail.SMTPMail"
theComponent(4) = "JMail.Message"
theComponent(5) = "CDO.Message"
theComponent(6) = "ABMailer.Mailman"
theComponent(7) = "Persits.MailSender"

' the name of the components
theComponentName(0) = "CDONTS"
theComponentName(1) = "Bamboo SMTP"
theComponentName(2) = "ServerObjects ASPMail"
theComponentName(3) = "JMail 3.7"
theComponentName(4) = "JMail 4"
theComponentName(5) = "CDOSYS"
theComponentName(6) = "ABMailer"
theComponentName(7) = "Persits ASPMail"

Function IsObjInstalled(strClassString)
 On Error Resume Next
 ' initialize default values
 IsObjInstalled = False
 Err = 0
 ' testing code
 Dim xTestObj
 Set xTestObj = Server.CreateObject(strClassString)
 If 0 = Err Then IsObjInstalled = True
 ' cleanup
 Set xTestObj = Nothing
 Err = 0
End Function
%>
<HEAD>
<TITLE>E-mail Component Test</TITLE>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</HEAD>
<body style="background-image: none;">
<table class="pcCPcontent" style="width:100%;">
    <tr> 
        <th colspan="2">Email Component Test</th>
    </tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<% Dim i
	For i=0 to UBound(theComponent) %>
		<tr> 
			<td nowrap="nowrap" style="border-bottom: 1px dashed #e1e1e1;"> 
			<%= theComponentName(i)%>:
			</td>
			<td style="border-bottom: 1px dashed #e1e1e1;">
			<% If Not IsObjInstalled(theComponent(i)) Then %>
				Not Installed 
			<% Else %>
				<strong>Installed</strong>
			<% End If %>
			</td>
		</tr>
	<% Next %>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><a href=# onClick="self.close();">Close Window</a></td>
	</tr>
</table>
</BODY>
</HTML>