<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Cross Selling - Add new relationship: select related products" %>
<% Section="products" %>
<%PmAdmin="2*3*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<% dim conntemp, rs, idmain, query
idmain=request.QueryString("idmain") 
if request("action")="source" then
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
				idmain=id
				exit for
			end if
		Next
	end if
end if
if idmain<>"" then
	session("cross_idmain")=idmain
else
	idmain=session("cross_idmain")
end if
if idmain="" then
	response.redirect "crossSellAdd.asp"
	response.end
end if
if request("action")="apply" then
	call openDb()
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
				query="INSERT INTO cs_relationships (idproduct, idrelation, num) VALUES ("&idmain&","&id&",0);"
				set rs=Server.CreateObject("ADODB.Recordset") 
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		next
	end if
	call closedb()
	response.redirect "crossSellAddc.asp?idmain="&idmain
	response.End
end if %>
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Cross Selling - Add new relationship"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select one or more related products."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="crossSellAddb.asp"
				src_ToPage="crossSellAddb.asp?action=apply"
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND products.idproduct <> " & idmain & " AND (products.idProduct NOT IN (SELECT DISTINCT idrelation FROM cs_relationships WHERE idproduct=" & idmain & ")) "
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->