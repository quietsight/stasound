<% pageTitle = "Reverse Category Import Wizard - Locate categories" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
Dim connTemp,rs,query

pcv_catlist=trim(request("catlist"))
session("cp_revCatImport_fseparator")=request("fseparator")
if session("cp_revCatImport_fseparator")="" then
	session("cp_revCatImport_fseparator")=0
end if
session("cp_ExportSize")=request("ExportSize")
session("cp_revCatImport_cseparator")=request("cseparator")
Select Case session("cp_revCatImport_fseparator")
	Case 0: session("cp_revCatImport_cseparator")=","
	Case 1: session("cp_revCatImport_cseparator")=chr(9)
End Select
if pcv_catlist<>"" then
	session("cp_revCatImport_catlist")=pcv_catlist
	if pcv_catlist="ALL1" then
		session("cp_revCatImport_catlist")="ALL"
		session("cp_revCatImport_pagecurrent")=""
	else
		session("cp_revCatImport_pagecurrent")=request("pagecurrent")
	end if
	response.redirect "ReverseCatImport_step2.asp"
end if
%>

<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Locate the Categories to Export</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
				<%
				src_FormTitle1="Find categories"
				src_FormTitle2="Reverse Category Import Wizard - Locate categories"
				src_FormTips1="Use the following filters to look for categories in your store."
				src_FormTips2="Select one or more categories that you would like to export."
				src_DisplayType=1
				src_ShowLinks=0
				src_ParentOnly=0
				src_FromPage="ReverseCatImport_step1A.asp?fseparator=" & session("cp_revCatImport_fseparator") & "&cseparator=" & session("cp_revCatImport_cseparator") & "&ExportSize=" & session("cp_ExportSize")
				src_ToPage="ReverseCatImport_step1A.asp?fseparator=" & session("cp_revCatImport_fseparator") & "&cseparator=" & session("cp_revCatImport_cseparator") & "&ExportSize=" & session("cp_ExportSize")
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" New Search "
				src_PageSize=15
				UseSpecial=0
				session("srcCat_from")=""
				session("srcCat_where")=""
				%>
				<!--#include file="inc_srcCats.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->