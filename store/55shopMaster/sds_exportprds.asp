<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
if (request("pagetype")="1") or (request("src_PageType")="1") then
	pcv_PageType=1
	pcv_Title="Drop-Shipper"
	pcv_Table="pcDropShipper"
else
	pcv_PageType=0
	pcv_Title="Supplier"
	pcv_Table="pcSupplier"
end if

pageTitle="Export Products Assigned to the Selected " & pcv_Title%>
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
<%Dim connTemp,rs1,query

'--- Get Search Form parameters ---

src_FormTitle1=getUserInput(request("src_FormTitle1"),0)
src_FormTitle2=getUserInput(request("src_FormTitle2"),0)
src_FormTips1=getUserInput(request("src_FormTips1"),0)
src_FormTips2=getUserInput(request("src_FormTips2"),0)
src_IncNormal=getUserInput(request("src_IncNormal"),0)
src_IncBTO=getUserInput(request("src_IncBTO"),0)
src_IncItem=getUserInput(request("src_IncItem"),0)
src_Special=getUserInput(request("src_Special"),0)
src_Featured=getUserInput(request("src_Featured"),0)
src_DisplayType=getUserInput(request("src_DisplayType"),0)
src_ShowLinks=getUserInput(request("src_ShowLinks"),0)
src_FromPage=getUserInput(request("src_FromPage"),0)
src_ToPage=getUserInput(request("src_ToPage"),0)
src_Button2=getUserInput(request("src_Button2"),0)
src_Button3=getUserInput(request("src_Button3"),0)

'Start SDBA
src_PageType=getUserInput(request("src_PageType"),0)
src_IDSDS=getUserInput(request("src_IDSDS"),0)
src_IsDropShipper=getUserInput(request("src_IsDropShipper"),0)
src_sdsAssign=getUserInput(request("src_sdsAssign"),0)
src_sdsStockAlarm=getUserInput(request("src_sdsStockAlarm"),0)
'End SDBA

form_idcategory=getUserInput(request("idcategory"),0)
form_customfield=getUserInput(request("customfield"),0)
form_SearchValues=getUserInput(request("SearchValues"),0)
form_priceFrom=getUserInput(request("priceFrom"),0)
form_priceUntil=getUserInput(request("priceUntil"),0)
form_withstock=getUserInput(request("withstock"),0)
form_sku=getUserInput(request("sku"),0)
form_IDBrand=getUserInput(request("IDBrand"),0)
form_keyWord=getUserInput(request("keyWord"),0)
form_exact=getUserInput(request("exact"),0)
form_pinactive=getUserInput(request("pinactive"),0)
form_resultCnt=getUserInput(request("resultCnt"),0)
form_order=getUserInput(request("order"),0)

'--- End of Search Form parameters ---
%>

<%if (src_IDSDS<>"") and (src_IDSDS<>"0") then
	call opendb()%>
	<h2 style="padding-left: 10px;">
		<%if src_PageType="0" then%>
			<%="Supplier: "%>
			<%query="SELECT pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName FROM pcSuppliers WHERE pcSupplier_ID=" & src_IDSDS
			set rs1=ConnTemp.execute(query)%>
			<strong><%=rs1("pcSupplier_Company") & " (" & rs1("pcSupplier_FirstName") & " " & rs1("pcSupplier_LastName") & ")"%></strong>
			<%set rs1=nothing
		else%>
			<%="Drop-Shipper: "%>
			<%if src_IsDropShipper="1" then
				query="SELECT pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName FROM pcSuppliers WHERE pcSupplier_ID=" & src_IDSDS
				set rs1=ConnTemp.execute(query)%>
				<strong><%=rs1("pcSupplier_Company") & " (" & rs1("pcSupplier_FirstName") & " " & rs1("pcSupplier_LastName") & ")"%></strong>
				<%set rs1=nothing
			else
				query="SELECT pcDropShipper_Company,pcDropShipper_FirstName,pcDropShipper_LastName FROM pcDropShippers WHERE pcDropShipper_ID=" & src_IDSDS
				set rs1=ConnTemp.execute(query)%>
				<strong><%=rs1("pcDropShipper_Company") & " (" & rs1("pcDropShipper_FirstName") & " " & rs1("pcDropShipper_LastName") & ")"%></strong>
				<%set rs1=nothing
			end if%>
		<%end if%>
	</h2>
	<%call closedb()
end if%>
	
	<form name="ajaxSearch" method="post" action="sds_exportprdsA.asp" class="pcForms">
	<input type="hidden" name="src_IncNormal" value="<%=src_IncNormal%>">
	<input type="hidden" name="src_IncBTO" value="<%=src_IncBTO%>">
	<input type="hidden" name="src_IncItem" value="<%=src_IncItem%>">
	<input type="hidden" name="src_Special" value="<%=src_Special%>">
	<input type="hidden" name="src_Featured" value="<%=src_Featured%>">
	<input type="hidden" name="src_DisplayType" value="0">
	<input type="hidden" name="src_ShowLinks" value="0">
	<input type="hidden" name="src_FromPage" value="<%=src_FromPage%>">
	<input type="hidden" name="src_ToPage" value="<%=src_ToPage%>">
	<input type="hidden" name="src_Button2" value="<%=src_Button2%>">
	<input type="hidden" name="src_Button3" value="<%=src_Button3%>">
	
	<%'Start SDBA%>
	<input type="hidden" name="src_PageType" value="<%=src_PageType%>">
	<input type="hidden" name="src_IDSDS" value="<%=src_IDSDS%>">
	<input type="hidden" name="src_IsDropShipper" value="<%=src_IsDropShipper%>">
	<input type="hidden" name="src_sdsAssign" value="0">
	<%'End SDBA%>

	<input type="hidden" name="idcategory" value="<%=form_idcategory%>">
	<input type="hidden" name="customfield" value="<%=form_customfield%>">
	<input type="hidden" name="SearchValues" value="<%=form_SearchValues%>">
	<input type="hidden" name="priceFrom" value="<%=form_priceFrom%>">
	<input type="hidden" name="priceUntil" value="<%=form_priceUntil%>">
	<input type="hidden" name="withstock" value="<%=form_withstock%>">
	<input type="hidden" name="sku" value="<%=form_sku%>">
	<input type="hidden" name="IDBrand" value="<%=form_IDBrand%>">
	<input type="hidden" name="keyWord" value="<%=form_keyWord%>">
	<input type="hidden" name="exact" value="<%=form_exact%>">
	<input type="hidden" name="pinactive" value="<%=form_pinactive%>">
	<input type="hidden" name="resultCnt" value="<%=form_resultCnt%>">
	<input type="hidden" name="order" value="<%=form_order%>">
	<input type="hidden" name="iPageCurrent" value="1">
	<input type="hidden" name="prdlist" value="">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2">
				Set the following export options:
			</td>
		</tr>
        <tr>
        	<td colspan="2" class="pcCPspacer"></td>
        </tr>
		<tr>
			<td>Type:</td>
			<td>
				<select name="src_sdsStockAlarm">
					<option value="0" <%if src_sdsStockAlarm<>"1" then%>selected<%end if%>>All products that belong to this <%=pcv_Title%></option>
					<option value="1" <%if src_sdsStockAlarm="1" then%>selected<%end if%>>All products whose inventory count is lower than the reorder level</option>
				</select>
			</td>
		</tr>
		<tr>
			<td>Format:</td>
			<td>
				<select name="src_expFormat">
					<option value="0" <%if src_expFormat<>"1" then%>selected<%end if%>>HTML Table</option>
					<option value="1" <%if src_expFormat="1" then%>selected<%end if%>>CSV File</option>
				</select>
			</td>
		</tr>
        <tr>
        	<td colspan="2" class="pcCPspacer"></td>
        </tr>
		<tr>
			<td>&nbsp;</td>
			<td><input type="submit" name="Submit" value=" Export products " class="submit2"></td>
			</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->