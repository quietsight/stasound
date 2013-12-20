<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Product(s) to the Promotion" %>
<% Section="specials" %>
<%PmAdmin=3%>
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

pidcode=request("idcode")

if pidcode="" then
	pidcode="0"
end if

pIDProduct=request("idproduct")

if pIDProduct="" or pIDProduct="0" then
	response.redirect "menu.asp"
end if

if request("action")="add" then
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
				if pidcode<>"0" then
					query="INSERT INTO pcPPFProducts (pcPrdPro_ID,idproduct) values (" & pidcode & "," & ID & ")"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
				else
					if (Instr(session("admin_PromoFPrds"),ID & ",")=1) or (Instr(session("admin_PromoFPrds"),"," & ID & ",")>1) then
					else
						session("admin_PromoFPrds")=session("admin_PromoFPrds") & ID & ","
					end if
				end if
			end if
		Next
	end if
	
	if pidcode=0 then
		response.redirect "AddPromotionPrd.asp?idproduct=" & pIDProduct
	else
		response.redirect "ModPromotionPrd.asp?idproduct=" & pIDProduct
	end if
end if
%>

<p class="pcCPsectionTitle">Select available products</p>
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Add New Product(s) to the Promotion"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select one or more products that you want to add to the Promotion"
				src_IncNormal=1
				src_DontShowInactive=1
				src_IncBTO=0
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="addPrdsToPR.asp?idcode=" & pidcode & "&idproduct=" & pIDProduct
				src_ToPage="addPrdsToPR.asp?action=add&idcode=" & pidcode & "&idproduct=" & pIDProduct
				src_Button1=" Search "
				src_Button2=" Add to the Promotion "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND (products.idProduct NOT IN (select idproduct FROM pcPPFProducts WHERE pcPrdPro_id=" & pidcode & ")) "
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->