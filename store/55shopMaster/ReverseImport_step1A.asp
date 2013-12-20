<% pageTitle = "Reverse Import Wizard - Locate products to export" %>
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

pcv_prdlist=trim(request("prdlist"))
session("cp_revImport_fseparator")=request("fseparator")
if session("cp_revImport_fseparator")="" then
	session("cp_revImport_fseparator")=0
end if
session("cp_ExportSize")=request("ExportSize")
session("cp_revImport_cseparator")=request("cseparator")
Select Case session("cp_revImport_fseparator")
	Case 0: session("cp_revImport_cseparator")=","
	Case 1: session("cp_revImport_cseparator")=chr(9)
End Select
if pcv_prdlist<>"" then
	session("cp_revImport_prdlist")=pcv_prdlist
	if pcv_prdlist="ALL1" then
		session("cp_revImport_prdlist")="ALL"
		session("cp_revImport_pagecurrent")=""
	else
		session("cp_revImport_pagecurrent")=request("pagecurrent")
	end if
	response.redirect "ReverseImport_step2.asp"
end if
%>

<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td colspan="2">
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
				<%
				src_ShowPrdTypeBtns=1
				src_FormTitle1="Find Products"
				src_FormTitle2="Reverse Import Wizard - Locate products"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select one or more products that you would like to export."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=1
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="ReverseImport_step1A.asp?fseparator=" & session("cp_revImport_fseparator") & "&cseparator=" & session("cp_revImport_cseparator") & "&ExportSize=" & session("cp_ExportSize")
				src_ToPage="ReverseImport_step1A.asp?fseparator=" & session("cp_revImport_fseparator") & "&cseparator=" & session("cp_revImport_cseparator") & "&ExportSize=" & session("cp_ExportSize")
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" New Search "
				src_PageSize=15
				UseSpecial=0
				session("srcprd_from")=""
				session("srcprd_where")=""
				cat_DisplayRoot="1"
				%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->