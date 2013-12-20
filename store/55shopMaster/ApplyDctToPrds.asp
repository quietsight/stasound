<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Assign quantity discounts to multiple products" %>
<% section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<% on error resume next
Dim rsOrd, connTemp, strSQL, pid,rstemp

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

'sorting order
Dim strORD

strORD=request("order")
if strORD="" then
	strORD="description"
End If

strSort=request("sort")
if strSort="" Then
	strSort="ASC"
End If 
	
call openDb()

idproduct=request("idproduct")

if idproduct<>"" then
session("ADQidproduct")=idproduct
else
idproduct=session("ADQidproduct")
end if

if request("action")="apply" then
	if idproduct<>"" then
		query="SELECT * FROM discountsPerQuantity WHERE idproduct="&idproduct&" AND discountdesc='PD' ORDER BY num"
		set rstemp=conntemp.execute(query)
		if not rstemp.eof then
			if (request("prdlist")<>"") and (request("prdlist")<>",") then
				prdlist=split(request("prdlist"),",")
				For i=lbound(prdlist) to ubound(prdlist)
					id=prdlist(i)
					If (id<>"0") and (id<>"") then
						query="delete from discountsPerQuantity where idproduct=" & id
						set rstemp1=connTemp.execute(query)
					end if
				Next
			 	
				do while not rstemp.eof
					For i=lbound(prdlist) to ubound(prdlist)
						id=prdlist(i)
						If (id<>"0") and (id<>"") then
							query=""
							query="" & id & ","
							query=query & rstemp("idcategory") & ","
							query=query &"'" &rstemp("discountDesc")&"'" & ","
							query=query &rstemp("quantityFrom")& ","
							query=query &rstemp("quantityUntil")& ","
							query=query &rstemp("discountPerUnit")& ","
							query=query &rstemp("num")& ","
							query=query &rstemp("percentage")& ","
							query=query &rstemp("discountPerWUnit")& ","
							query=query &rstemp("baseproductonly")
			
							query="insert into discountsPerQuantity (idproduct,idcategory,discountDesc,quantityFrom,quantityUntil,discountPerUnit,num,percentage,discountPerWUnit,baseproductonly) values (" & query & ")"
							set rs=conntemp.execute(query)
							set rs=nothing
						end if
					Next
					rstemp.movenext
				loop
			end if
			set rstemp=nothing
			call closedb()
			response.redirect "modDctQtyPrd.asp?idproduct=" & idproduct & "&s=1&msg=" & "Assigned quantity discounts to selected products successfully!" 
		else
			set rstemp=nothing
			call closedb()
			response.redirect "modDctQtyPrd.asp?idproduct=" & idproduct & "&r=1&msg=" & "The product does not have any quantity discounts assigned to it."
		end if
	else
		response.redirect "modDctQtyPrd.asp?idproduct=" & idproduct & "&r=1&msg=" & "Please select source product before assigning quantity discounts"
	end if
end if


%>
<!--#include file="AdminHeader.asp"-->

	<% ' START show message, if any
		pcMessage = getUserInput(Request.QueryString("msg"),0)
		If pcMessage <> "" Then %>
		<div class="pcCPmessage">
			<%=pcMessage%>
		</div>
		<% 	end if
	 ' END show message %>

	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Assign quantity discounts to multiple products"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select which products you would like to apply quantity discounts to."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=1
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="ApplyDctToPrds.asp"
				src_ToPage="ApplyDctToPrds.asp?action=apply"
				src_Button1=" Search "
				src_Button2=" Add Discounts to Selected Products "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND (products.idproduct<>" & idproduct & ") AND (products.idproduct NOT IN (SELECT DISTINCT pcPrdPromotions.idproduct FROM pcPrdPromotions))"
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->