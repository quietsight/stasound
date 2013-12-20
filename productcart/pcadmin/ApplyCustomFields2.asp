<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Copy custom fields to other products" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
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

if request("action")="next" then
	CustFieldCopy=request("CustField")
	idproduct=request("idproduct")
	if CustFieldCopy<>"" then
		session("CustFieldCopy")=CustFieldCopy
	else
		response.redirect "ApplyCustomFields1.asp?idproduct=" & idproduct & "&msg=" & "Please choose a custom field before continuing."
	end if
end if

if idproduct<>"" then
	session("ACidproduct")=idproduct
else
	idproduct=session("ACidproduct")
end if

%>
<!--#include file="AdminHeader.asp"-->

	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>

	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Copy custom fields to other products"
				src_FormTips1="Use the following filters to look for products in your store that you would like to copy this custom field to."
				src_FormTips2="Select which products you would like to copy this custom field to."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="ApplyCustomFields1.asp?idproduct=" & idproduct
				src_ToPage="ApplyCustomFields3.asp?action=apply"
				src_Button1=" Search "
				src_Button2=" Apply Custom Field "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND products.idproduct <> " & idproduct
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->