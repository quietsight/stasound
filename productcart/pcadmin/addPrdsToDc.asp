<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Product(s) to the Discount by Code" %>
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

pidDiscount=request("idcode")
if not validNum(pidDiscount) then pidDiscount="0"

if request("action")="add" then
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		call openDb()
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
				if pidDiscount<>"0" then
					query="insert into pcDFProds (pcFPro_IDDiscount,pcFPro_IDProduct) values (" & pidDiscount & "," & ID & ")"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=connTemp.execute(query)
					set rstemp=nothing
				else
					if (Instr(session("admin_DiscFPrds"),ID & ",")=1) or (Instr(session("admin_DiscFPrds"),"," & ID & ",")>1) then
					else
						session("admin_DiscFPrds")=session("admin_DiscFPrds") & ID & ","
					end if
				end if
			end if
		Next
		call closeDb()
	end if
	
	if pidDiscount=0 then
		response.redirect "AddDiscounts.asp"
	else
		response.redirect "modDiscounts.asp?mode=Edit&iddiscount=" & pidDiscount
	end if
end if
%>

<p class="pcCPsectionTitle">Select available products</p>
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Add New Product(s) to the Discount by Code"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select one or more products that you want to add to the Discount by Code"
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="addPrdsToDc.asp?idcode=" & pidDiscount
				src_ToPage="addPrdsToDc.asp?action=add&idcode=" & pidDiscount
				src_Button1=" Search "
				src_Button2=" Add to the Discount by Code "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND (products.idProduct NOT IN (select DISTINCT pcFPro_IDProduct from pcDFProds where pcFPro_IDDiscount=" & pidDiscount & ")) "
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->