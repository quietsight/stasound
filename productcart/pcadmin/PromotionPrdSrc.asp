<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Promotions - Locate Product"
pcStrPageName="PromotionPrdSrc.asp"
section="specials" 
%>
<%PmAdmin=3%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim connTemp,rs,query
call opendb()
%>

	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Promotions - Product Search Results"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="If there are promotions applied to a product, you can modify the promotion by clicking on the Modify icon, if no promotion have been set then you can choose to add promotion by clicking on the Add icon."
				src_IncNormal=1
				src_DontShowInactive=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=0
				src_ShowLinks=0
				src_FromPage="PromotionPrdSrc.asp"
				src_ToPage=""
				src_Button1=" Search "
				src_Button2=""
				src_Button3=" Back "
				src_PageSize=15
				src_ShowPromoTypes=1
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