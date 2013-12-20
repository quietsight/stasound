<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Remove Option Attributes from Multiple Products" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
idOptionGroup=request.querystring("idOptionGroup")
if not validNum(idOptionGroup) then
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
%>

<table class="pcCPcontent">
<tr> 
	<td>Removing product option group: <b><%=OptionGroupDesc%></b></td>
</tr>
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Remove Option Attributes from Multiple Products"
				src_FormTips1="Use the following filters to look for products in your store that you would like to update."
				src_FormTips2="Select one or more products that you would like to update."
				src_IncNormal=1
				src_IncBTO=0
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="RevMultiOptions.asp?idOptionGroup=" & idOptionGroup
				src_ToPage="RevMultiOptions1.asp?action=upd&idOptionGroup=" & idOptionGroup
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=" ,options_optionsGroups "
				session("srcprd_where")=" AND (options_optionsGroups.idOptionGroup="&idOptionGroup&" AND options_optionsGroups.idProduct=products.idProduct) "
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->