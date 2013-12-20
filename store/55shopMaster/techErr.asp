<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "ProductCart Online Help - Technical Error Information" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim ErrVar, MsgDifined
MsgDefined=0
ErrVar=server.HTMLEncode(request.QueryString("error"))
if ErrVar="Permissions Not Set to Log" then
	MsgDefined=1 %>
    <div class="pcCPmessage">An error occurred while accessing your Control Panel. Please make sure that your &quot;<%=scPcFolder%>/<%=scAdminFolderName%>/CPLogs/&quot; folder has &quot;<strong>Modify</strong>&quot; or &quot;<strong>Delete</strong>&quot; permissions.</div>
<% end if
if ErrVar="Permissions Not Set to Modify Tax" then
	MsgDefined=1 %>
    <div class="pcCPmessage">An error occurred while updating your store settings. Please make sure that your &quot;<%=scPcFolder%>/includes/&quot; folder has &quot;<strong>Modify</strong>&quot; or &quot;<strong>Delete</strong>&quot; permissions.</div>
<% end if
if ErrVar="Permissions Not Set to Modify Constants" then
	MsgDefined=1 %>
    <div class="pcCPmessage">An error occurred while activating your store. Please make sure that your &quot;<%=scPcFolder%>/includes/&quot; folder has &quot;<strong>Modify</strong>&quot; or &quot;<strong>Delete</strong>&quot; permissions.</div>
<% end if
if ErrVar="Permissions Not Set to Modify Create" then
	MsgDefined=1 %>
    <div class="pcCPmessage">An error occurred while updating your store settings. Please make sure that your &quot;<%=scPcFolder%>/includes/&quot; folder has &quot;<strong>Modify</strong>&quot; or &quot;<strong>Delete</strong>&quot; permissions.</div>
<% end if
if ErrVar="Permissions Not Set to Modify First" then
	MsgDefined=1 %>
    <div class="pcCPmessage">An error occurred while updating your store settings. Please make sure that your &quot;<%=scPcFolder%>/includes/&quot; folder has &quot;<strong>Modify</strong>&quot; or &quot;<strong>Delete</strong>&quot; permissions.</div>
<% end if 
if ErrVar="Permissions Not Set to Modify Email" then
	MsgDefined=1 %>
    <div class="pcCPmessage">An error occurred while updating your store settings. Please make sure that your &quot;<%=scPcFolder%>/includes/&quot; folder has &quot;<strong>Modify</strong>&quot; or &quot;<strong>Delete</strong>&quot; permissions.</div>
<% end if 
if MsgDefined=0 then %>
    <div class="pcCPmessage"><%response.write request.querystring("error")%></div>
<% end if %>
<!--#include file="AdminFooter.asp"--> 