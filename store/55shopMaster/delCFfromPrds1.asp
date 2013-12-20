<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Remove custom field from products" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
dim query, conntemp, rs

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

if (request("action")="delfield") and (request("idcustom")<>"") then
else
response.redirect "ManageCFields.asp"
response.end
end if

idcustom=mid(request("idcustom"),2,len(request("idcustom")))
idcustomType=Left(request("idcustom"),1)

%>
<div style="display:none">
<%
	src_FormTitle1=""
	src_FormTitle2="Remove custom field from selected products"
	src_FormTips1=""
	src_FormTips2="Select which products you would like to remove this custom field."
	src_IncNormal=1
	src_IncBTO=1
	src_IncItem=0
	src_DisplayType=1
	src_ShowLinks=0
	src_FromPage="ManageCFields.asp"
	src_ToPage="delCFfromPrds2.asp?action=apply&idcustom=" & request("idcustom")
	src_Button1=" Search "
	src_Button2=" Remove from Selected "
	src_Button3=" Back "
	src_PageSize=15
	UseSpecial=1
	session("srcprd_from")=""
	query=""
	if idcustomType="C" then
		query=" AND (products.xfield1=" & idcustom & " OR products.xfield2=" & idcustom & " OR products.xfield3=" & idcustom & ")"
	else
		query=" AND (products.idproduct IN (SELECT DISTINCT pcSearchFields_Products.idProduct FROM pcSearchFields_Products INNER JOIN pcSearchData ON pcSearchFields_Products.idSearchData=pcSearchData.idSearchData WHERE pcSearchData.idSearchField=" & idcustom & "))"
	end if
	session("srcprd_where")=query
%>
<!--#include file="inc_srcPrds.asp"-->
</div>
<script>
	document.ajaxSearch.submit();
</script>