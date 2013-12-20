<% 'BTO ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<% pageTitle = "Copy Options Groups to other products" %>
<% section = "services" %>
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/currencyformatinc.asp" -->

<% dim f, query, conntemp, rstemp, pidProduct
If request("action")<>"add" then
	response.redirect "menu.asp"
End if
call openDB() 
%>
<!--#include file="Adminheader.asp"-->
<%
pIdproduct=request("prdlist")
pcArr=split(pIdproduct,",")
pIdproduct=pcArr(0)
session("pcAdminProductID")=pIdproduct 
%>
<table class="pcCPcontent">
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
				<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Select Products"
				src_FormTips1="Locate the products to which the options will be assigned."
				src_FormTips2="Select the products to which the options will be assigned."
				src_IncNormal=1
				src_IncBTO=0
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="ApplyOptionsMulti2.asp?action=add&&prdlist=" & request("prdlist")
				src_ToPage="dupMultiOptions.asp?action=add"
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" Back "
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
<!--#include file="Adminfooter.asp"-->