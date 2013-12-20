<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<%	
pcStrPageName="BrandsAssign.asp"
Dim connTemp, query, rs, rstemp, pcIntBrandID

pcIntBrandID = request.QueryString("idbrand")
	if not validNum(pcIntBrandID) then
		response.Redirect("BrandsManage.asp")
	end if
	
call opendb()
	query="SELECT BrandName FROM Brands WHERE IDBrand="&pcIntBrandID
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)
	BrandName=rstemp("BrandName")
	set rstemp=nothing
	pageTitle="Manage Brands - Add New Products to " & BrandName
%>
<!--#include file="AdminHeader.asp"-->
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Manage Brands - Add new products to " & BrandName
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select the products that you would like to assign to " & BrandName
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="BrandsAssign.asp"
				src_ToPage="BrandsAssign2.asp?idbrand=" & pcIntBrandID
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND products.IDBrand <> " & pcIntBrandID
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->