<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials" 
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="sm_check.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<% 
Dim query, conntemp, rstemp, rstemp4, rstemp5

call openDB()

if (request("a")="new") OR (request("c")="new") then
	session("sm_pcSaleID")=""
	session("sm_Param1")=""
	session("sm_Param2")=""
	session("sm_UP1")=""
	session("sm_UP2")=""
	session("sm_UP3")=""
	session("sm_UP4")=""
	session("sm_UP5")=""
	session("sm_ChangeTxt")=""
	session("sm_ChangeName")=""
	session("sm_TechDetails")=""
	session("sm_TechDetails1")=""
	session("sm_SaleName")=""
	session("sm_SaleDesc")=""
	session("sm_SaleIcon")=""
	session("sm_loadedSale")=""
	session("sm_PrdCount")=0
	session("sm_ShowNext")=0
	query="DELETE FROM pcSales_Pending WHERE pcSales_ID=0;"
	set rs=connTemp.execute(query)
	set rs=nothing
end if

pcSaleID=request("id")

if pcSaleID<>"" then
	query="SELECT pcSales_ID FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		session("sm_pcSaleID")=pcSaleID
	else
		pcSaleID=""
		session("sm_pcSaleID")=pcSaleID
	end if
	set rs=nothing
else
	pcSaleID=session("sm_pcSaleID")
end if
pageIcon="pcv4_icon_salesManager.png"

if (pcSaleID="") AND (request("a")<>"rev") then
	call closedb()
	response.redirect "sm_addprds.asp"
end if

if request("action")="rmv" then
	if (request("prdlist")<>"") and (request("prdlist")<>",") then
		call openDb()
		prdlist=split(request("prdlist"),",")
		For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
					if IsNull(pcSaleID) OR pcSaleID="" then
						pcSaleID=0
					end if
					query="DELETE FROM pcSales_Pending WHERE idProduct=" & ID & " AND pcSales_ID=" & pcSaleID & ";"
					set rstemp=connTemp.execute(query)
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
end if

	if request("a")="rev" then
		if pcSaleID="" then
			pcSaleID=0
		end if
	end if%>
	<div style="display:none">
	<%
	if (request("b")<>"") then
		session("sm_ShowNext")=request("b")
	else
		session("sm_ShowNext")=1
	end if
	src_ShowPrdTypeBtns=1
	src_FormTitle1=""
	if ((request("b")="5") OR (request("b")="6") OR (request("b")="7") OR (request("b")="8") OR (request("b")="9")) then
		src_FormTitle2="View Products Included in this Sale"
	else
		if pcSaleID=0 then
			src_FormTitle2="Sales Manager - Create New Sale - Step 1: View/Edit Products Included in this Sale"
		else
			src_FormTitle2="Sales Manager - Edit Sale - Step 1: View/Edit Products Included in this Sale"
		end if
	end if
	src_FormTips1=""
	src_FormTips2=""
	src_IncNormal=0
	src_IncBTO=0
	src_IncItem=0
	src_ShowLinks=0
	if ((request("b")="5") OR (request("b")="6") OR (request("b")="7")  OR (request("b")="8") OR (request("b")="9")) then
		src_DisplayType=0
		Select Case request("b")
			Case "5":
				src_FromPage="sm_saledetails.asp?id=" & request("scid") & "&e=1"
			Case "6":
				src_FromPage="sm_sales.asp"
			Case "7":
				src_FromPage="sm_saledetails.asp?id=" & request("scid")
			Case "8":
				src_FromPage="sm_stop.asp?a=stop&id=" & request("scid")
			Case "9":
				src_FromPage="sm_showdetails.asp?id=" & request("scid")
		End Select
		
		src_ToPage=""
		src_Button1=""
		src_Button2=""
		src_Button3=" Back "
	else
		src_DisplayType=1
		if request("a")="rev" then
			src_FromPage="sm_addprds.asp?a=rev" & "&b=" & request("b")
		else
			src_FromPage="sm_addprds.asp"
		end if
		src_ToPage="sm_addedit_S1.asp?action=rmv&a=" & request("a") & "&b=" & request("b")
		src_Button1=""
		src_Button2=" Remove Selected Products "
		src_Button3=" Add More Products to the Sale "
	end if
	src_PageSize=25
	UseSpecial=1
	session("srcprd_from")=""
	session("srcprd_where")=" AND (products.idProduct IN (select idProduct FROM pcSales_Pending WHERE pcSales_ID=" & pcSaleID & ")) "
	session("sm_selectall")="0"
	if session("sm_PrdCount")>"0" then
		session("sm_TechDetails1")="Products:<div style=""padding-top: 5px; font-weight:bold;"">The sale will affect " & session("sm_PrdCount") & " product(s) in your store.</div><div style=""padding-top: 5px;"" class=""pcSmallText"">If you just updated the product included in the sale, this number may be incorrect. it will be updated when you save the sale.</div>"
	end if
	%>
	<!--#include file="inc_srcPrds.asp"-->
	</div>
	<script>
		document.ajaxSearch.submit();
	</script>
<%call closedb()%>
