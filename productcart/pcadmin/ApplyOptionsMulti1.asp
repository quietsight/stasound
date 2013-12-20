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

<% 
dim f, query, conntemp, rstemp, pidProduct
 
call openDB()

session("pcAdminRedirectFlag")=""
session("pcAdminProductID")=""
%>


<!--#include file="Adminheader.asp"-->
<table class="pcCPcontent">
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
        <tr>
            <td class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr>
			<td>
				<%
					src_FormTitle1="Find Products"
					src_FormTips1="Use the following filters to look for a product whose options you want to copy to other products."
					src_FormTitle2="Select a Product"
					src_FormTips2="Select the product whose option groups and attributes you wish to copy to other products."
					src_DisplayType=2
					src_ShowLinks=0
					src_FromPage="ApplyOptionsMulti1.asp"
					src_ToPage="ApplyOptionsMulti2.asp?action=add"
					src_Button1=" Search "
					src_Button2="Continue"
					src_Button3="Back"
					src_ParentOnly=0
					UseSpecial=1
					session("srcprd_from")=""
					session("srcprd_where")=" AND (products.idProduct IN (SELECT idProduct FROM options_optionsGroups)) "					
				%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="Adminfooter.asp"-->