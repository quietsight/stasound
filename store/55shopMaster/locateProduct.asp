<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add product to order" %>
<% section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
idOrder=request.QueryString("ido")
if not validNum(idOrder) then response.redirect "resultsAdvancedAll.asp?B1=View+All&dd=1"
Dim query, conntemp, rs
call openDB()

if request("action")="add" then
	pcv_ArrPrdlist=request("prdlist")
	if trim(pcv_ArrPrdlist)<>"" then
		pcv_ArrPrdlist=split(pcv_ArrPrdlist,",")
		pcv_intIDProduct=pcv_ArrPrdlist(0)
		query="SELECT serviceSpec FROM products WHERE idproduct=" & pcv_intIDProduct
		set rs=connTemp.execute(query)
		pcv_IsBTO=rs("serviceSpec")
		set rs=nothing
		if pcv_IsBTO<>0 then
			query="SELECT configProduct FROM configSpec_products where specProduct=" & pcv_intIDProduct
			set rs=conntemp.execute(query)
	
			if rs.eof then
				pcv_IsBTO=0
			end if
			set rs=nothing
		end if
		
		if pcv_IsBTO<>0 then
			call closedb()
			response.redirect "bto_configurePrd.asp?ido=" & idOrder & "&idproduct=" & pcv_intIDProduct
		else
			call closedb()
			response.redirect "addprdToOrd.asp?ido=" & idOrder & "&idproduct=" &  pcv_intIDProduct & "&adminPreview=1"
		end if
	else
		call closedb()
		response.redirect "menu.asp"
	end if	
end if
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
				<%
				src_FormTitle1="Find a Product"
				src_FormTitle2="Add product to order"
				src_FormTips1="Use one or more of the following search criteria to locate a product in your store catalog. You will then be able to view and edit any of the products returned in the search."
				src_FormTips2="Select the product that you would like to add to order."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=2
				src_ShowLinks=0
				src_FromPage="LocateProduct.asp?ido=" & idOrder
				src_ToPage="LocateProduct.asp?ido=" & idOrder & "&action=add"
				src_Button1=" Search "
				src_Button2=" Add to Order "
				src_Button3=" New Search "
				src_PageSize=15
				UseSpecial=0
				session("srcprd_from")=""
				session("srcprd_where")=""
				%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>      
<!--#include file="AdminFooter.asp"-->