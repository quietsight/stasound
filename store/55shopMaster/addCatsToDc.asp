<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Categories to the Discount by Code" %>
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
	if (request("catlist")<>"") and (request("catlist")<>",") then
		call openDb()
		catlist=split(request("catlist"),",")
		For i=lbound(catlist) to ubound(catlist)
			IDPro=catlist(i)
			IDSub="1"
			If (IDPro<>"0") and (IDPro<>"") then
				if pidDiscount<>"0" then
					query="INSERT INTO pcDFCats (pcFCat_IDDiscount,pcFCat_IDCategory,pcFCat_SubCats) VALUES (" & pidDiscount & "," & IDPro & "," & IDSub & ");"
					set rstemp=connTemp.execute(query)
				else
					session("admin_DiscFCATs")=session("admin_DiscFCATs") & IDPro & "-" & IDSub & ","
				end if
			End if
		Next
		set rstemp=nothing
		call closeDb()
	end if
	
	if pidDiscount=0 then
		response.redirect "AddDiscounts.asp"
	else
		response.redirect "modDiscounts.asp?mode=Edit&iddiscount=" & pidDiscount
	end if
end if
%>
<p class="pcCPsectionTitle">Select available categories</p>
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Categories"
				src_FormTitle2="Add New Categories to the Discount by Code"
				src_FormTips1="Use the following filters to look for categories in your store."
				src_FormTips2="Select one or more categories that you want to add to the Discount by Code"
				src_DisplayType=1
				src_IncNotDisplay=2
				src_ParentOnly=0
				src_ShowLinks=0
				src_FromPage="addCatsToDc.asp?idcode=" & pidDiscount
				src_ToPage="addCatsToDc.asp?action=add&idcode=" & pidDiscount
				src_Button1=" Search "
				src_Button2=" Add to the Discount by Code "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcCat_From")=""
				session("srcCat_Where")=" AND (categories.idCategory NOT IN (SELECT DISTINCT pcFCat_idcategory FROM pcDFCats WHERE pcFCat_IDDiscount=" & pidDiscount & ")) "
			%>
				<!--#include file="inc_srcCATs.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->