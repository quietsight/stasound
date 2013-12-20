<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Dim pageTitle, pageName, pageIcon, Section
pageTitle="Create Subscription Package Link"
pageName="sb_CreatePackages.asp"
pageIcon="pcv4_icon_sb.png"
Section="SB" 
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="sb_inc.asp"-->
<% 'on error resume next
Dim rsOrd, connTemp, strSQL, pid,rstemp

call openDb()

%>
	<% ' START show message, if any
		pcMessage = getUserInput(Request.QueryString("msg"),0)
		If pcMessage <> "" Then %>
		<div class="pcCPmessage">
			<%=pcMessage%>
		</div>
	<% 	end if
	 ' END show message %>
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
				<%
				src_ShowPrdTypeBtns=0
				src_FormTitle1="Find Products"
				src_FormTitle2="Create Subscription Package Link"
				src_FormTips1="Use the following filters to look for products in your store that you would like to flag as subscription packages. "
				src_FormTips2="Select a product you would like to flag as a subscription package."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=2
				src_ShowLinks=0
				src_FromPage="sb_Default.asp"
				src_ToPage="sb_CreatePackages2.asp?action=setup"
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" Back "
				src_PageSize=25
				UseSpecial=1
				session("srcprd_from")=""
				query=""
				query=" AND (products.idproduct NOT IN (SELECT DISTINCT idProduct FROM SB_Packages))"
				session("srcprd_where")=query
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->