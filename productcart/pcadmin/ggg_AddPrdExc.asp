<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Products not eligible for Gift Wrapping" %>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim connTemp,query,rstemp
call opendb()

if request("action")="add" then

	pcv_prdlist=request("prdlist")
	if pcv_prdlist<>"" then
		pcA=split(pcv_prdlist,",")

		For i=lbound(pcA) to ubound(pcA)
			if trim(pcA(i))<>"" then
				IDPro=pcA(i)
				query="insert into pcProductsExc (pcPE_IDProduct) values (" & IDPro & ")"
				set rstemp=connTemp.execute(query)
				set rstemp=nothing
			end if
		Next
	end if
	call closeDb()
	response.redirect "ggg-GiftWrapOptions.asp"
end if

%>
<table class="pcCPcontent">
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Product exclusions (not eligible for Gift Wrapping)"
				src_FormTips1="Use the following filters to look for products in your store that you would like to add to the Product Exclusions List."
				src_FormTips2="Select one or more products that you would like to add to the Product Exclusions List."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="ggg_AddPrdExc.asp"
				src_ToPage="ggg_AddPrdExc.asp?action=add"
				src_Button1=" Search "
				src_Button2=" Add to Exclusion List "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND (products.idproduct NOT IN (SELECT pcPE_IDProduct FROM pcProductsExc)) "
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->