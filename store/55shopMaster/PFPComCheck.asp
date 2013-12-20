<%@Language="VBScript"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=19%>
<!--#include file="adminv.asp"-->
<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This script was originally created by Richard Kinser.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Option Explicit

Dim theComponent(1)
Dim theComponentName(1)
Dim theComponentLink(1)
Dim theComponentMsg(1)

' the components
theComponent(0) = "PFProCOMControl.PFProCOMControl.1"
theComponent(1) = "PayPal.Payments.Communication.PayflowNETAPI" 

' the name of the components
theComponentName(0) = "Payflow Pro Component"
theComponentName(1) = "Payflow Pro .NET SDK"

theComponentLink(0) = ""
theComponentLink(1) = "https://cms.paypal.com/us/cgi-bin/?cmd=_render-content&content_ID=developer/library_download_sdks"

theComponentMsg(0) = ""
theComponentMsg(1) = "ProductCart will use this recommended COM Object"

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

<HTML>
<HEAD>
<TITLE>Payflow Pro Component Test</TITLE>
</HEAD>
<body bgcolor="#ffffff" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" marginwidth="0" marginheight="0" stylesrc="1bg_info.htm">
<table border="0" width="100%" cellspacing="0" cellpadding="3">
  <tr> 
    <td width="100%" bgcolor="#000080"><font size="2" color="mintcream"><b><font face="Arial, Helvetica, sans-serif">Pay 
      Flow Pro Component Test</font></b></font></td>
  </tr>
</table>

<br>
<br>
<table border=1 bordercolor="#000000" cellspacing=0 cellpadding=4 align="center">
  <% Dim i
           For i=0 to UBound(theComponent) %>
  <tr> 
    <td bgcolor="#FFFFFF" valign="top" align="left" nowrap> 
      <div align="right"><strong> 
        <font face="Arial, Helvetica, sans-serif" size="2"><%= theComponentName(i)%>:</font></strong></div>
    </td>
    <td bgcolor="#FFFFFF" align="" A0B0E0"" center""> <font size="2"> <font face="Arial, Helvetica, sans-serif">
      <% If Not IsObjInstalled(theComponent(i)) Then %>
      </font></font> 
      <div align="left"><font size="2" face="Arial, Helvetica, sans-serif">Not 
        Installed <BR>
		<% if theComponentLink(i)<>"" then %>
			<a href="<%=theComponentLink(i)%>" target="_blank">Download</a>
        <% end if %>
        <% Else %>
        	<strong>Installed</strong>
            <% if i=1 then 
             response.write "<BR><font size='1' color=FF0000>"&theComponentMsg(i)
			end if
       	End If %>
        </font></div>
    </td>
  </tr>
  <% Next %>
</table>
<p align="center"><a href=# onClick="self.close();"><font size=2 face="Verdana,Helvetica,Arial,sans-serif"><b>Close 
  Window</b></font></a> 
 
</BODY>
</HTML>
