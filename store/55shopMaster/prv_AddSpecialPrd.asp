<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Select Products to Apply Special Settings" 
pageIcon="pcv4_icon_reviews.png"
Section="reviews" 
%>
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
<% 
Dim connTemp,query,rs
pidDiscount=request("idcode")
if request("action")="add" then
	Count=request("Count")
	if Count="" then
		Count="0"
	end if
	pcv_strPrd=""
	For i=1 to Count
		if request("Pro" & i)="1" then
			IDPro=request("IDPro" & i)
			pcv_strPrd=pcv_strPrd & IDPro & ","
		end if
	Next
	if pcv_strPrd<>"" then
		response.redirect "prv_CustomizePrd.asp?PrdList=" & pcv_strPrd
	end if
end if
%>
<form name="hForm" method="post" action="prv_AddSpecialPrd.asp?action=add" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2"><h2>List of eligible products</h2></td>
		</tr>
		<%  
		call opendb()
		query="SELECT products.idproduct, products.description FROM products WHERE (products.idproduct NOT IN (SELECT DISTINCT pcRS_IDProduct FROM pcReviewSpecials)) AND products.removed=0 AND products.active=-1 AND products.ConfigOnly=0 ORDER BY products.Description"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if rs.eof then
			set rs=nothing
			call closeDb()
		%>
			<tr>
				<td colspan="2">No Items to display. All eligible products have already been edited. <a href="JavaScript:;" onClick="JavaScript:history.back();">Back</a></td>
				</tr>
        <%
		else
			Dim prdCount, count
			count=0
			prdArray = rs.getRows()
			prdCount = ubound(prdArray,2)
			set rs = nothing
			call closedb()

			for i=0 to prdCount
				pIDProduct=prdArray(0,i)
				pName=prdArray(1,i)
				count=count+1
				%>
				<tr>
					<td width="5%">
                    	<input type="checkbox" name="Pro<%=Count%>" value="1">
                        <input type="hidden" name="IDPro<%=Count%>" value="<%=pIDProduct%>">
                     </td>
					<td width="95%"><a href="FindProductType.asp?id=<%=pIDProduct%>" target="_blank"><%=pName%></a></td>
				</tr>
			<%
			next
		
			if Count>0 then
			%>
            <tr>
            	<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
			<td colspan="2" class="cpLinksList">
				<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
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
			</td>
		</tr>
		<tr>
            	<td colspan="2" class="pcCPspacer"></td>
            </tr>
            <%
            end if
            %>
		<tr>
			<td colspan="2">
			<input name="submit" type="submit" class="submit2" value=" Add to Product List ">
			&nbsp;
			<input name="back" type="button" onClick="javascript:history.back();" value="Back">
			<input type=hidden name="Count" value="<%=Count%>">
			</td>
		</tr>
        <%
		end if
		%>
    </table>
	</form>
<!--#include file="AdminFooter.asp"-->