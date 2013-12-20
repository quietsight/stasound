<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="specials" %>
<%PmAdmin=3%>
<%
pcSaleID=session("sm_pcSaleID")
if (pcSaleID="") AND (request("a")<>"rev") then
pageTitle="Sales Manager - Create New Sale - Step 1: Choose Products"
else
if (pcSaleID="") then
	pageTitle="Sales Manager -  Create New Sale - Add More Products"
else
	pageTitle="Sales Manager - Edit Sale - Step 1: Add More Products"
end if
end if%>
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
Dim connTemp,query,rstemp

pcSaleID=session("sm_pcSaleID")
if pcSaleID="" then
	pcSaleID="0"
end if
if not validNum(pcSaleID) then pcSaleID="0"

if request("action")="add" then
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		call openDb()
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
					query="SELECT idProduct FROM pcSales_Pending WHERE (idProduct=" & ID & ") AND (pcSales_ID=" & pcSaleID & ");" 
					set rstemp=connTemp.execute(query)
					if rstemp.eof then
						query="INSERT INTO pcSales_Pending (idProduct,pcSales_ID) values (" & ID & "," & pcSaleID & ")"
						set rstemp=connTemp.execute(query)
						set rstemp=nothing
					end if
					set rstemp=nothing
			end if
		Next
		call closeDb()
	end if
	
	Set cn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	cn.Open scDSN
	Set cmd.ActiveConnection = cn
	cmd.CommandText = "uspGetPrdCount"
	cmd.CommandType = adCmdStoredProc

	cmd.Parameters.Refresh
	cmd.Parameters("@Param1") = "Products INNER JOIN pcSales_Pending ON Products.idProduct=pcSales_Pending.idProduct"
	cmd.Parameters("@Param2") = "pcSales_Pending.pcSales_ID=" & pcSaleID
	cmd.Execute
	
	tmpPrdCount=cmd.Parameters("@SMCOUNT")
	session("sm_PrdCount")=tmpPrdCount
	Set cmd=nothing
	set cn=nothing
	
	if pcSaleID=0 then
		session("sm_Param1")="Products INNER JOIN pcSales_Pending ON Products.idProduct=pcSales_Pending.idProduct"
		session("sm_Param2")="pcSales_Pending.pcSales_ID=" & pcSaleID
		if request("a")="rev" then
			response.redirect "sm_addedit_S1.asp?a=rev" & "&b=" & request("b")
		else
			response.redirect "sm_addedit_S3.asp"
		end if
	else
		response.redirect "sm_addedit_S1.asp?a=" & request("a") & "&b=" & request("b")
	end if
end if
%>

<p class="pcCPsectionTitle">Select available products</p>
	<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%	session("sm_ShowNext")=0
				src_ShowPrdTypeBtns=1
				src_checkPrdType=0
				src_FormTitle1="Find Products"
				src_FormTitle2="Add New Product(s) to the Sale"
				src_FormTips1="Use the following filters to look for products in your store."
				src_FormTips2="Select one or more products that you want to add to the Sale"
				src_SM=1
				src_IncNormal=0
				src_IncBTO=0
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="sm_addprds.asp?a=" & request("a") & "&b=" & request("b")
				src_ToPage="sm_addprds.asp?action=add&a=" & request("a") & "&b=" & request("b")
				src_Button1=" Search "
				src_Button2=" Add to the Sale "
				src_Button3=" Back "
				src_PageSize=25
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND (products.idProduct NOT IN (select idProduct FROM pcSales_Pending WHERE pcSales_ID=" & pcSaleID & ")) "
				session("sm_selectall")="1"
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		<%if request("a")="rev" then%>
		<tr>
		<td>
			&nbsp;&nbsp;&nbsp;<input type="button" name="Go" value="View/Edit Products Included in this Sale" onClick="location='sm_addedit_S1.asp?a=rev&b=<%=request("b")%>';" class="ibtnGrey">
		</td>
		</tr>
		<%end if%>
	</table>
<!--#include file="AdminFooter.asp"-->