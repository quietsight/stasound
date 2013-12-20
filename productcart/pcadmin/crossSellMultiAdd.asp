<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Cross Selling - Add the relationship to multiple products" %>
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
<% 
dim query, conntemp, rs, idmain
call openDb()
idmain=request("idmain")
if idmain<>"" then
    query="SELECT idproduct,idrelation,num,cs_type,discount,isPercent,isRequired FROM cs_relationships WHERE idproduct=" & idmain
	set rs=connTemp.execute(query)
	if not rs.eof then
	    pcArray=rs.getRows()
	    pcv_intCount=ubound(pcArray,2)
	end if
	set rs=nothing
else
	call closeDb()
	response.redirect "crossSellView.asp"
end if
if request("action")="apply" then
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
				query="DELETE FROM cs_relationships WHERE idproduct="&id&";"
				set rs=Server.CreateObject("ADODB.Recordset") 
				set rs=conntemp.execute(query)
				set rs=nothing
				For j=0 to pcv_intCount
					query="INSERT INTO cs_relationships (idproduct, idrelation, num, cs_type, discount, isPercent, isRequired) "
					query = query+ "VALUES ("&id&"," & pcArray(1,j) & "," & pcArray(2,j) & ",'" & pcArray(3,j) & "'," & pcArray(4,j) & "," & pcArray(5,j) & "," & pcArray(6,j) & ");"
					set rs=Server.CreateObject("ADODB.Recordset") 
					set rs=conntemp.execute(query)
					set rs=nothing
				Next
			End if
		next
	end if
	call closedb()
	response.redirect "crossSellView.asp"
end if %>

	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Cross Selling - Add this relationship to multiple products"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select one or more products to which you would like to apply this relationship."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="crossSellMultiAdd.asp?idmain=" & idmain
				src_ToPage="crossSellMultiAdd.asp?action=apply&idmain=" & idmain
				src_Button1=" Search "
				src_Button2=" Apply the relationship "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND products.idProduct <> " & idmain & " AND (products.idProduct NOT IN (SELECT DISTINCT idProduct FROM cs_relationships)) "
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->