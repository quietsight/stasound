<% 'BTO ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<% 
pageTitle="View and Update Inventory Levels for Multiple Products"
pageIcon="pcv4_icon_inventoryAdded.gif"
section="products" 
%>
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/currencyformatinc.asp" -->

<% dim f, query, conntemp, rstemp, pidProduct
 
call openDB() 

	' START: Check to see if there are products in the store
	' If not, redirect to appropriate message
	query="SELECT idProduct FROM products"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if rs.EOF then
		set rs=nothing
		call closeDb()
		response.Redirect("msg.asp?message=5")
	end if
	set rs=nothing
	' END: Check to see if there are products in the store

%>
<!--#include file="Adminheader.asp"-->
<%src_checkPrdType="0"%>
<table class="pcCPcontent">
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
				<%
				src_ShowPrdTypeBtns=1
				src_FormTitle1="Find Products"
				src_FormTitle2="Product Search Results"
				src_FormTips1="Use the following filters to look for products in your store to update inventory."
				src_FormTips2=""
				src_IncNormal=0
				src_IncBTO=0
				src_IncItem=0
				src_DisplayType=0
				
				src_StockChoices=1
				
				src_ShowLinks=1
				src_FromPage="viewStock.asp"
				src_ToPage=""
				src_Button1=" Search "
				src_Button2=" Continue "
				src_Button3=" New Search "
				src_PageSize=25
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND products.noStock=0 "
				src_ShowStockLevel=1
				%>
				<!--#include file="inc_srcPrds.asp"-->
                <script>
					document.ajaxSearch.action="viewStocka.asp?action=newsrc";
				</script>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="Adminfooter.asp"-->