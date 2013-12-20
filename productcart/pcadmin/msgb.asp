<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
response.Buffer=true
pageTitle="Control Panel - Message" %>
<%PmAdmin=0%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<div class="pcCPmessage">
<% 
msg=request.querystring ("message")
if msg="" or isNull(msg) then
	msg=request.querystring ("msg")
end if
if instr(ucase(msg),"<S") then
    msg=replace(msg,"<","&lt;")
    msg=replace(msg,">","&gt;")
end if 
if instr(ucase(msg),"JAVASCRIPT") then
    if instr(ucase(msg),"HISTORY.GO") then 
    else
        msg=replace(msg,"<","&lt;")
        msg=replace(msg,">","&gt;")
    end if 
end if %>
<%=msg%>
</div>
<!--#include file="AdminFooter.asp"-->