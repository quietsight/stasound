<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="USPS Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr>
		<td>
		<% Dim query, connTemp, rs, pcv_USPSID, pcv_USPSShipServer, pcv_USPSLabelServer
		if request.form("submit")<>"" then
			pcv_USPSShipServer=request.form("USPSShipServer")
			pcv_USPSLabelServer=request.form("USPSLabelServer")
			if pcv_USPSShipServer="" then
				pcv_USPSShipServer="http://production.shippingapis.com/ShippingAPI.dll"
			end if
			if pcv_USPSLabelServer="" then
				pcv_USPSLabelServer="https://secure.shippingapis.com/ShippingAPI.dll"
			end if
			pcv_USPSID=request.form("USPSID")
			'update db
			call openDb()
			query="UPDATE ShipmentTypes SET shipserver='"&pcv_USPSShipServer&"', userID='"&pcv_USPSID&"', AccessLicense='"&pcv_USPSLabelServer&"' WHERE idShipment=4"
			set rs=server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			set rs=nothing
			call closedb()
			response.redirect "viewShippingOptions.asp#USPS"
			response.end
		else 
			call opendb()
			query="SELECT shipserver, userID, AccessLicense FROM ShipmentTypes WHERE idShipment=4"
			set rs=server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			pcv_ShipServer=rs("shipserver")
			pcv_UserID=rs("userID")
			pcv_LabelServer=rs("AccessLicense")
			set rs=nothing
			call closedb() 
		%>
        
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>

			<form name="form1" method="post" action="USPS_EditLicense.asp">
						
                <table class="pcCPcontent">
                	<tr>
                    	<td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    <tr> 
                        <th colspan="2" align="left">Enable USPS - <span class="pcSmallText"><a href="http://www.USPS.com" target="_blank">Web site</a></span></th>
                    </tr>
                	<tr>
                    	<td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> 
                        <p>In order to use USPS, you need to register to obtain your User ID (it's free). Go to: <a href="http://www.usps.com/webtools/" target="_blank">http://www.usps.com/webtools/</a> (XML API Used).</p>
                        <p>&nbsp;</p>
                        <p>Enter the server URLs that were sent to you by USPS Customer Care Center when you registered. You should have received two URLs, one that is secured and one that is not secured. The secured URL will start with &quot;https://&quot;. Make sure to include the entire URLs into the fields below, including the &quot;http://&quot; or &quot;https://&quot;.<br>
                          <br>
                          <b>Note</b>: USPS will begin to function only after your account has been set to production status. </p></td>
                    </tr>
                	<tr>
                    	<td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    <tr>
                        <th colspan="2">Production Server Settings</th>
                    </tr>
                	<tr>
                    	<td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    <tr> 
                        <td width="26%"><div align="right">USPS Server:</div></td>
                      <td width="74%"> 
                        <input type="text" name="USPSShipServer" size="60" value="<%=pcv_ShipServer%>">
                        </td>
                  </tr>
                    <tr>
                        <td><div align="right"> USPS Secured Server:</div></td>
                        <td><input type="text" name="USPSLabelServer" size="60" value="<%=pcv_LabelServer%>">  </td>
                    </tr>
                    <tr> 
                        <td><div align="right">User ID:</div></td>
                        <td> 
                            <input type="text" name="USPSID" size="30" value="<%=pcv_UserID%>">
                        </td>
                    </tr>
                	<tr>
                    	<td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    <tr>
						<td>&nbsp;</td>
                        <td align="left"> 
                        <input type="submit" name="Submit" value="Continue" class="submit2"></td>
                    </tr>
                </table>
            </form>
		<% end if %>
        </td>
    </tr>
</table>
<!--#include file="AdminFooter.asp"-->