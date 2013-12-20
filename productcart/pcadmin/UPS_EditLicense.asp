<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Edit UPS Online&reg; Tools License" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript"><!--
function win(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=640,height=500')
	myFloater.location.href=fileName;
	}
//--></script>

<table class="pcCPcontent">
	<tr>
		<td>
		<% 
		Dim mySQL, connTemp, rs
		call openDb()
		if request.form("submit")<>"" then
			UPSAccessLicense=request.form("UPSAccessLicense")
			UPSID=request.form("UPSID")
			UPSPassword=request.form("UPSPassword")
			if UPSAccessLicense="" or UPSID="" or UPSPassword="" then
				response.redirect "UPS_EditLicense.asp?msg="&Server.URLEncode("All fields are required.")
				response.end
			end if
			'update db
			mySQL="UPDATE ShipmentTypes SET userID='"&UPSID&"',[password]='"&UPSPassword&"', AccessLicense='"&UPSAccessLicense&"' WHERE idShipment=3"
			set rs=server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(mySQL)
			set rs=nothing
			call closedb()
			response.redirect "viewShippingOptions.asp?s=1&msg=" & server.URLEncode("UPS License updated successfully.")
		else 
		mySQL="SELECT userID,[password],AccessLicense FROM ShipmentTypes WHERE idShipment=3"
		set rs=server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(mySQL)
		%>
		<form name="form1" method="post" action="UPS_EditLicense.asp" class="pcForms">
						
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>

<table class="pcCPcontent">
<tr>                          
<td colspan="2"> 
<p>In order to use UPS OnLine&reg; Tools, you need to <a href="javascript:win('../UPSLicense/licenseAgrRequest.asp')">register</a> an account with the company. Registration is free and includes access to the following UPS OnLine&reg; Tools:</p>
<ul>
<li>UPS OnLine&reg; Tools Tracking</li>
<li>UPS OnLine&reg; Tools Rates &amp; Service Selection</li>
</ul>
<p>If you need to register an account <a href="javascript:win('../UPSLicense/licenseAgrRequest.asp')">click here</a>.</p>
<hr>
</td>
</tr>                        
<tr> 
<td width="22%">
<div align="right">Access License:</div></td>
<td width="78%"> 
<input type="text" name="UPSAccessLicense" size="50" value="<%=rs("AccessLicense")%>">
</td>
</tr>
<tr> 
<td>
<div align="right">User ID:</div></td>
<td> 
<input type="text" name="UPSID" size="30" value="<%=rs("userID")%>">
</td>
</tr>
<tr> 
<td>
<div align="right">Password:</div></td>
<td> 
<input type="text" name="UPSpassword" size="30" value="<%=rs("password")%>">
</td>
</tr>                      
<tr> 
<td colspan="2">&nbsp;</td>
</tr>
<tr> 
<td colspan="2" align="center"> 
<input type="submit" name="Submit" value="Continue" class="submit2"></td>
</tr>                    
</table>
</form>
<% end if %>
</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->