<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Add New Categories to the Promotion" %>
<% Section="specials" %>
<%PmAdmin=3%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<%Dim connTemp,query,rstemp

call opendb()

pidcode=request("idcode")
if pidcode="" then
	pidcode="0"
end if

pIDProduct=request("idproduct")

if pIDProduct="" or pIDProduct="0" then
	response.redirect "menu.asp"
end if

if request("action")="add" then

	Count=request("Count")
	if Count="" then
	Count="0"
	end if
	
	For i=1 to Count
		if request("Pro" & i)="1" then
		IDPro=request("IDPro" & i)
		IDSub=request("IDSub" & i)
		if IDSub="" then
		IDSub="0"
		end if
		if pidcode<>"0" then
			query="INSERT INTO pcPPFCategories (pcPrdPro_ID,idcategory,pcPPFCats_IncSubCats) values (" & pidcode & "," & IDPro & "," & IDSub & ")"
			set rs=connTemp.execute(query)
			set rs=nothing
		else
			session("admin_PromoFCATs")=session("admin_PromoFCATs") & IDPro & "-" & IDSub & ","
		end if
		end if
	Next
	if pidcode=0 then
		response.redirect "AddPromotionPrd.asp?idproduct=" & pIDProduct
	else
		response.redirect "ModPromotionPrd.asp?idproduct=" & pIDProduct
	end if
end if

%>

<table width="94%" border="0" cellspacing="0" cellpadding="4" align="center">
<form name="hForm" method="post" action="addCatsToPR.asp?action=add">
<tr class="normal">
<td valign="top" width="98%" colspan="3"><strong>List of available categories:</strong></td>
</tr>
<tr class="normal"><td></td><td></td><td>
  <p align="center">Include subcategories</td></tr>
<%
Dim Count,pcArr,intCount
Count=0

Sub TestAvailableCats(idparent,pathname)
Dim query,rs,i,pIDCat,pName,myTest,tmpPath
	For i=0 to intCount
	if clng(pcArr(2,i))=clng(idparent) then
		pIDCat=pcArr(0,i)
		pName=pcArr(1,i)
		query="select idcategory,pcPPFCats_IncSubCats from pcPPFCategories where pcPrdPro_ID=" & pidcode & " and idcategory=" & pIDCat
		set rs=connTemp.execute(query)
		myTest=0
		if rs.eof then
			myTest=1
		else
			if rs("pcPPFCats_IncSubCats")="0" then
				myTest=1
			end if
		end if
		if rs.eof then
			Count=Count+1%>
			<tr class="normal">
			<td width="4%"><input type="checkbox" name="Pro<%=Count%>" value="1"><input type=hidden name="IDPro<%=Count%>" value="<%=pIDCat%>"></td>
			<td nowrap width="94%"><%=pName%> <%if pathname<>"" then%>[<%=pathname%>]<%end if%></td><td><input type="checkbox" name="IDSub<%=Count%>" value="1"></td>
			</tr>
			<%
		end if
		set rs=nothing

		IF myTest=1 then
			tmpPath=pathname
			if tmpPath<>"" then
				tmpPath=tmpPath & "/"
			end if
			tmpPath=tmpPath & pName
			call TestAvailableCats(pIDCat,tmpPath)
		END IF
	end if
	Next
End Sub

query="SELECT categories.idcategory,categories.categoryDesc,idParentCategory FROM categories WHERE categories.idcategory<>1 AND categories.iBTOhide=0 ORDER BY categories.categoryDesc"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArr=rs.getRows()
	set rs=nothing
	intCount=ubound(pcArr,2)
	call TestAvailableCats(1,"")
end if
set rs=nothing

if Count=0 then%>
<tr class="normal">
<td colspan="3"><br><font color="#FF0000"><b>No Items to display.</b></font><br><br></td>
</tr>
<%end if%>
<tr class="normal">
<td colspan="3"><%if Count>0 then%><a href="javascript:checkAll();"><b>Check All</b></a><b>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a></b>
<script language="JavaScript">
<!--
function checkAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.hForm.Pro" + j); 
if (box.checked == false) box.checked = true;
   }
}

function uncheckAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.hForm.Pro" + j); 
if (box.checked == true) box.checked = false;
   }
}

//-->
</script>
<%end if%>
</td>
</tr>
<tr>
<td colspan="3">
<input name="submit" type="submit" class="ibtnGrey" value=" Add to the Promotion ">
&nbsp;
<input name="back" type="button" class="ibtnGrey" onClick="javascript:history.back();" value="Back">
<input type=hidden name="Count" value="<%=Count%>">
<input type=hidden name="idcode" value="<%=pidcode%>">
<input type=hidden name="idproduct" value="<%=pIDProduct%>">
</td>
</tr>
</form>
</table>
<!--#include file="AdminFooter.asp"-->