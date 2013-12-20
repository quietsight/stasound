<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Quantity Discounts - Locate Product" %>
<% section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%Dim connTemp,rs,query

call opendb()%>

	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Quantity Discounts - Product Search Results"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="NOTE: once you have assigned discounts to a product, you can apply the same discounts to multiple other products at once. Just click on the &quot;Modify&quot; icon, then select &quot;Apply to Other Products&quot;."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=1
				src_DisplayType=0
				src_ShowLinks=0
				src_FromPage="viewDisca.asp"
				src_ToPage=""
				src_Button1=" Search "
				src_Button2=""
				src_Button3=" Back "
				src_PageSize=15
				src_ShowDiscTypes=1
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=""
				session("srcprd_DiscArea")="1"
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->