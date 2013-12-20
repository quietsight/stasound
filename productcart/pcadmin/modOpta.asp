<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Rename Option Attribute" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/languages.asp" -->
<% 
dim query, conntemp, rstemp

pidOption=request.Querystring("idOption")
pidOptionGroup=request.Querystring("idOptionGroup")

if not validNum(pidOption) then
   response.redirect "msg.asp?message=21"
end if

call openDb()
query="SELECT * FROM options WHERE idOption=" &pidOption
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error renaming option attribute on modOpta.asp") 
end If

poptionDescrip=replace(rstemp("optionDescrip"),"''","'")
poptionDescrip=replace(rstemp("optionDescrip"),"""","&quot;")
set rstemp=nothing
call closeDb()
%>
<!--#include file="AdminHeader.asp"-->
<form action="modOptb.asp" method="post" name="modOpGr" class="pcForms">
<input type="hidden" name="idOption" size="60" value="<%=pidOption%>">
<input type="hidden" name="idOptionGroup" value="<%=pidOptionGroup%>">
<% if request.querystring("redirectURL")<>"" then %>
<input type="hidden" name="redirectURL" value="<%=request.querystring("redirectURL")%>">
<input type="hidden" name="mode" value="<%=request.querystring("mode")%>">
<% end if %>
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>             
	<tr> 
		<td colspan="2">
		Attribute: <input name="optionDescrip" type="text" value="<%=poptionDescrip%>" size="40" maxlength="250">
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>   
	<tr>
		<td colspan="2">
		<input type="button" name="Button" value="Back" onClick="javascript:history.back()">
		&nbsp;<input type="submit" name="modify" value="Rename" class="submit2">
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr> 
</table>
</form>
<!--#include file="AdminFooter.asp"-->