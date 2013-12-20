<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add custom field to products" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
Dim rsOrd, connTemp, strSQL, pid,rstemp
call openDb()
%>
<!--#include file="AdminHeader.asp"-->
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Copy custom fields to other products"
				src_FormTips1="Use the following filters to look for products in your store that you would like to copy this custom field to."
				src_FormTips2="Select which products you would like to copy this custom field to."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="addCFtoPrds1.asp"
				src_ToPage="addCFtoPrds2.asp?action=apply"
				src_Button1=" Search "
				src_Button2=" Apply Custom Field "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				query=""
				if session("admin_customtype")=2 then
					query=" AND products.xfield1<>" & session("admin_idxfield") & " AND products.xfield2<>" & session("admin_idxfield") & " AND products.xfield3<>" & session("admin_idxfield")
				else
					query=" AND (products.idproduct NOT IN (SELECT DISTINCT idProduct FROM pcSearchFields_Products WHERE idSearchData=" & session("admin_skeyword") & "))"
				end if
				session("srcprd_where")=query
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->