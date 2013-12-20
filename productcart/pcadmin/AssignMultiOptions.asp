<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->

<% 
idOptionGroup=request.querystring("idOptionGroup")
if idOptionGroup="" or idOptionGroup="0" then
	response.redirect "manageOptions.asp"
end if

Dim query, conntemp, rs
call opendb()
 
query="SELECT OptionGroupDesc FROM OptionsGroups WHERE idOptionGroup="&idOptionGroup
set rs=server.createobject("adodb.recordset") 
set rs=conntemp.execute(query) 
OptionGroupDesc=rs("OptionGroupDesc")
set rs=nothing

call closedb()

pageTitle="Assign &quot;" & OptionGroupDesc & "&quot; to Multiple Products"

%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr>
		<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Assign Options to Multiple Products"
				src_FormTips1="Use the following filters to look for products in your store that you would like to assign options."
				src_FormTips2="Select one or more products that you would like to assign options."
				src_IncNormal=1
				src_IncBTO=0
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="AssignMultiOptions.asp?idOptionGroup=" & idOptionGroup
				src_ToPage="modPrdOpta4.asp?action=add&idOptionGroup=" & idOptionGroup
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND (products.idProduct NOT IN (SELECT idProduct FROM options_optionsGroups WHERE idOptionGroup="&idOptionGroup&")) "
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->