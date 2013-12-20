<% pageTitle = "Export to NexTag Wizard - Locate products" %>
<% section = "specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
Dim connTemp,rs,query

pcv_prdlist=trim(request("prdlist"))
if pcv_prdlist<>"" then
	session("cp_exportNextTag_prdlist")=pcv_prdlist
	if pcv_prdlist="ALL1" then
		session("cp_exportNextTag_prdlist")="ALL"
		session("cp_exportNextTag_pagecurrent")=""
	else
		session("cp_exportNextTag_pagecurrent")=request("pagecurrent")
	end if
	response.redirect "pcNextTag_step2.asp"
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
				src_FormTitle2="Export to NexTag Wizard - Locate products"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select one or more products that you would like to export."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_ShowPrdTypeBtns=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="pcNextTag_step1A.asp"
				src_ToPage="pcNextTag_step1A.asp"
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" New Search "
				src_DontShowInactive="1"
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