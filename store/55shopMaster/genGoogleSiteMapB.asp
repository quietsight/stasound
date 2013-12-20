<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*3*"%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 
pageTitle="Google Site Map Generation Results" 
pageIcon="pcv4_icon_xml.gif"
section="layout"
%>
<!--#include file="AdminHeader.asp"-->
<%
SiteMapFile=request("fn")

IF SiteMapFile<>"" THEN

	SPathHeader="http://www.google.com/webmasters/sitemaps/ping?sitemap=" & Server.URLEncode(SiteMapFile)

	Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
	srvXmlHttp.open "POST", SPathHeader, False
	srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	srvXmlHttp.send ""
	pcv_strSMHeader = srvXmlHttp.responseText
	myErr=0
		if err.number<>0 then
		err.Description=""
		err.number=0
		myErr=1
	end if
%>

	<table class="pcCPcontent">
    <tr>
        <td class="pcCPspacer"></td>
    </tr>
	<% if myErr=0 then %>
		<tr>
			<td>
            	<div class="pcCPmessageSuccess">Google received the SiteMap notification successfully!</div>
                <div style="padding-top: 10px;">Google returned the following message:</div>
                <div style="padding-top: 10px;"><%=pcv_strSMHeader%></div>
            </td>
        </tr>
	<% else %>
        <tr>
            <td>
                <div class="pcCPmessage">ProductCart was NOT able to send the SiteMap Notification to Google. You can manually submit the sitemap using your <a href="http://www.google.com/webmasters/tools/" target="_blank">Google Webmaster Tools</a> account.</div>
            </td>
        </tr>
	<% end if %>
    <tr>
        <td class="pcCPspacer"></td>
    </tr>
    <tr>
        <td><a href="menu.asp">Return to the Start page</a>.</td>
    </tr>
    <tr>
        <td class="pcCPspacer"></td>
    </tr>
	</table>
<%
END IF
%>
<!--#include file="AdminFooter.asp"-->