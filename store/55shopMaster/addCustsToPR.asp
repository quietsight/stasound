<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Customer(s) to the Promotion" %>
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
<%Dim connTemp,mySQL,rstemp

pIDPromo=request("idcode")

if pIDPromo="" then
	pIDPromo="0"
end if

pIDProduct=request("idproduct")

if pIDProduct="" or pIDProduct="0" then
	response.redirect "menu.asp"
end if

if request("action")="add" then
	if (request("custlist")<>"") and (request("custlist")<>",") then
		custlist=split(request("custlist"),",")
		For i=lbound(custlist) to ubound(custlist)
			id=custlist(i)
			If (id<>"0") and (id<>"") then
				if pIDPromo<>"0" then
					call opendb()
					query="insert into pcPPFCusts (pcPrdPro_id,IDCustomer) values (" & pIDPromo & "," & ID & ")"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					set rs=nothing
					call closedb()
				else
					if (Instr(session("admin_PRFCusts"),ID & ",")=1) or (Instr(session("admin_PRFCusts"),"," & ID & ",")>1) then
					else
						session("admin_PRFCusts")=session("admin_PRFCusts") & ID & ","
					end if
				end if
			End if
		Next
	End if
	if pIDPromo=0 then
		response.redirect "AddPromotionPrd.asp?idproduct=" & pIDProduct
	else
		response.redirect "ModPromotionPrd.asp?idproduct=" & pIDProduct
	end if
end if
%>

<p class="pcCPsectionTitle">Select available customers</p>
	<table id="FindCustomers" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Customers"
				src_FormTitle2="Add New Customer(s) to the Promotion"
				src_FormTips1="Use the following filters to look for customers in your store."
				src_FormTips2="Select one or more customers that you want to add to the Promotion"
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="addCustsToPR.asp?idcode=" & pIDPromo & "&idproduct=" & pIDProduct
				src_ToPage="addCustsToPR.asp?action=add&idcode=" & pIDPromo & "&idproduct=" & pIDProduct
				src_Button1=" Search "
				src_Button2=" Add to the Promotion "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcCust_from")=""
				session("srcCust_where")=" (customers.idcustomer NOT IN (select DISTINCT IDCustomer from pcPPFCusts WHERE pcPrdPro_id=" & pIDPromo & ")) "
			%>
				<!--#include file="inc_srcCusts.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->