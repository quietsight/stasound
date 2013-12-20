<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Canada Post Shipping Configuration" %>
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
if request.form("submit")<>"" then
	CPServer=request.form("CPServer")
	Session("ship_CP_Server")=CPServer
	CPID=request.form("CPID")
	Session("ship_CP_ID")=CPID
	if CPServer="" or CPID="" then
		response.redirect "ConfigureOption4.asp?msg="&Server.URLEncode("All fields are required.")
		response.end
	end if
	response.redirect "4_Step2.asp"
	response.end
else %>
<form name="form1" method="post" action="ConfigureOption4.asp" class="pcForms">
      <table class="pcCPcontent">
            <tr>
                <td colspan="3" class="pcCPspacer">
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
                </td>
            </tr>
            <tr> 
            <td colspan="2"><h2>Enable Canada Post - <a href="http://www.canadapost.ca/personal/offerings/sell_online_contact/can/tech_questions-e.asp" target="_blank">Web site</a></h2></td>
            </tr>
            <tr> 
                <td colspan="2">
                In order to use Canada Post, you need to request a shipping profile account from <a href="mailto:eparcel@canadapost.ca ">eparcel@canadapost.ca</a>. ProductCart utilizes Canada Post's Sell Online's XML Direct Connection. The Sell Online Direct Connection to the server can be obtained by sending an email to <a href="mailto:sellonline@canadapost.ca">sellonline@canadapost.ca</a> and by asking for the &quot;Sell Online Direct Connection&quot;. MUST provide the following information: 
                <ul>
                	<li>Company name </li>
                    <li>Contact name and telephone number</li>
                </ul>
             	</td>
            </tr>
            <tr> 
                                          
            <td width="20%" align="right">Server:</td>
            <td width="80%"><input type="text" name="CPServer" size="50" value="<%=Session(Ship_CPServer)%>"></td>
            </tr>
            <tr> 
            <td align="right">User ID:</td>
            <td><input type="text" name="CPID" size="30" value="<%=Session(Ship_CP_ID)%>"></td>
            </tr>
            <tr> 
            <td colspan="2">&nbsp;</td>
            </tr>
            <tr> 
            <td>&nbsp;</td>
            <td> 
            <input type="submit" name="Submit" value="Continue" class="submit2">
            &nbsp;
            <input type="button" name="back" value="Back" onClick="javascript:history.back()">
            </td>
        </tr>
    </table>
</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->