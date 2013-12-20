<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Activate United States Postal Service" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<% if request.form("submit")<>"" then
	USPSServer=request.form("USPSServer")
	USPSLabelServer=request.Form("USPSLabelServer")
	Session("ship_USPS_Server")=USPSServer
	Session("ship_USPS_LabelServer")=USPSLabelServer
	USPSID=request.form("USPSID")
	Session("ship_USPS_ID")=USPSID

	response.redirect "2_Step2.asp"
	response.end
else %>

		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>

		<form name="form1" method="post" action="ConfigureOption2.asp" class="pcForms">
			  <table class="pcCPcontent">
				<tr> 
					<td colspan="2"> 
					<p>In order to use USPS, you need to register to obtain your User ID (it's free). Go to: <a href="http://www.usps.com/webtools/" target="_blank">http://www.usps.com/webtools/</a> 
					(XML API Used).</p>
					<p><b>Note</b>: USPS will begin to function only after your account has been set 
					to production status.<br><br></p></td>
				</tr>
                <% if Session("ship_USPS_Server")="" then
					Session("ship_USPS_Server")="http://production.shippingapis.com/ShippingAPI.dll"
				end if
				if Session("ship_USPS_LabelServer")="" then
					Session("ship_USPS_LabelServer")="https://secure.shippingapis.com/ShippingAPI.dll"
				end if %>
				<tr> 
					<td width="19%">
					<div align="right">USPS Server:</div></td>
					<td width="81%"> 
					<input type="text" name="USPSServer" size="50" value="<%=Session("ship_USPS_Server")%>"></td>
				</tr>
				<tr>
                  <td><div align="right">USPS Secured Server:</div></td>
				  <td><input type="text" name="USPSLabelServer" size="50" value="<%=Session("ship_USPS_LabelServer")%>">
                  </td>
			    </tr>
				<tr> 
					<td>
					<div align="right">User ID:</div></td>
					<td> 
					<input type="text" name="USPSID" size="30" value="<%=Session(Ship_USPS_ID)%>"></td>
				</tr>
				<tr> 
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr> 
                	<td></td>
					<td>
					<input type="submit" name="Submit" value="Save" class="submit2">
					&nbsp;
					<input type="button" name="back" value="Back" onClick="javascript:history.back()" class="ibtnGrey"></td>
				</tr>
			</table>
		</form>
	<% end if %>
<!--#include file="AdminFooter.asp"-->