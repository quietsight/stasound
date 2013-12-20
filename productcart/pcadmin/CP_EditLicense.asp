<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration - Edit License" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<% 
Dim mySQL, connTemp, rs
call openDb()
if request.form("submit")<>"" then
	CPServer=request.form("CPServer")
	CPID=request.form("CPID")
	if CPServer="" or CPID="" then
		response.redirect "CP_EditLicense.asp?msg="&Server.URLEncode("All fields are required.")
		response.end
	end if
	'update db
	mySQL="UPDATE ShipmentTypes SET shipserver='"&CPServer&"', userID='"&CPID&"' WHERE idShipment=7"
	set rs=server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(mySQL)
	set rs=nothing
	call closeDb()
	response.redirect "viewShippingOptions.asp#CP"
	response.end
else 

	mySQL="SELECT shipserver,userID FROM ShipmentTypes WHERE idShipment=7"
	set rs=server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(mySQL) %>
		<form name="form1" method="post" action="CP_EditLicense.asp" class="pcForms">
			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer">
						<% ' START show message, if any %>
							<!--#include file="pcv4_showMessage.asp"-->
						<% 	' END show message %>
					</td>
				</tr>
				<tr> 
				  <td colspan="2"><h2>Enable Canada Post (<a href="http://www.canadapost.ca/personal/offerings/sell_online_contact/can/tech_questions-e.asp" target="_blank">Web site</a>)</h2>
					In order to use Canada Post, you need to request a shipping profile account from <a href="mailto:eparcel@canadapost.ca ">eparcel@canadapost.ca</a>. ProductCart utilizes Canada Post's Sell Online's XML Direct Connection. The Sell Online Direct Connection to the server can be obtained by sending an email to <a href="mailto:sellonline@canadapost.ca">sellonline@canadapost.ca</a> and by asking for the &quot;Sell Online Direct Connection&quot;.<br /><br />YOU MUST provide the following information: 
					<ul>
						<li>Company name </li>
						<li>Contact name and telephone number</li>
					</ul>
				 </td>
				</tr>
				<tr> 
				<td width="15%" align="right">Server:</td>
				<td width="85%"> <input type="text" name="CPServer" size="50" value="<%=rs("shipserver")%>"></td>
				</tr>
				<tr> 
				<td align="right">User ID:</td>
				<td><input type="text" name="CPID" size="30" value="<%=rs("userID")%>"></td>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr> 
				<td></td>
				<td>
				<input type="submit" name="Submit" value="Continue" class="submit2"></td>
				</tr>
            </table>
    </form>
    <% 
    set rs=nothing
    call closedb()
end if 
%>
<!--#include file="AdminFooter.asp"-->