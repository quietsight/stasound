<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Add New Product(s) to Exclusion List" 
pageIcon="pcv4_icon_reviews.png"
Section="reviews" 
%>
<%PmAdmin=2%>
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
Dim connTemp,query,rs
if request("action")="add" then
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		call openDb()
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			IDPro=prdlist(i)
			If (IDPro<>"0") and (IDPro<>"") then
				query="INSERT INTO pcRevExc (pcRE_IDProduct) values (" & IDPro & ")"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				set rs=nothing
			End if
		Next
		call closedb()
		response.redirect "prv_PrdExc.asp?s=1&msg=" & Server.UrlEncode("Products successfully added to the exclusion list.")
	else
		response.redirect "prv_PrdExc.asp"
	end if
end if
%>
<p class="pcCPsectionTitle">Select available products</p>
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Add New Product(s) to Exclusion List"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select one or more products that you want to add to the list of products for which customers cannot view and/or post reviews"
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="prv_AddPrdExc.asp"
				src_ToPage="prv_AddPrdExc.asp?action=add"
				src_Button1=" Search "
				src_Button2=" Add to the Product Exclusion List "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND (products.idProduct NOT IN (SELECT DISTINCT pcRE_IDProduct FROM pcRevExc)) "
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->