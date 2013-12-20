<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
if request("pagetype")="1" then
	pcv_PageType=1
	pcv_Title="Drop-Shipper"
	pcv_Table="pcDropShipper"
else
	pcv_PageType=0
	pcv_Title="Supplier"
	pcv_Table="pcSupplier"
end if

pageTitle=pcv_Title & " Management" %>
<% Section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->

	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find " & pcv_Title & "s"
				src_FormTitle2=pcv_Title & " Management"
				src_FormTips1="Use the following filters to look for " & pcv_Title & "s in your store."
				src_FormTips2="List of " & pcv_Title & "s in your store:"
				src_DisplayType=0
				src_ShowLinks=1
				src_FromPage="sds_manage.asp?pagetype=" & pcv_PageType
				src_ToPage=""
				src_Button1=" Search "
				src_Button2=""
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=0
				session("srcSDS_from")=""
				session("srcSDS_where")=""
				src_PageType=pcv_PageType
			%>
				<!--#include file="inc_srcSDSs.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->