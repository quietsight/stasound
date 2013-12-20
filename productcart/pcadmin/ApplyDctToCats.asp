<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Assign quantity discounts to multiple categories" %>
<% section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
Dim rsOrd, connTemp, strSQL, pid,rstemp

call openDb()

idcategory=request("idcategory")
if idcategory<>"" then
	session("ADQidcategory")=idcategory
else
	idcategory=session("ADQidcategory")
end if


if request("action")="apply" then
	if validNum(idcategory) then
		query="SELECT * FROM pcCatDiscounts WHERE pcCD_idcategory="&idcategory&" ORDER BY pcCD_num"
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		if not rstemp.eof then
			if (request("catlist")<>"") and (request("catlist")<>",") then
				catlist=split(request("catlist"),",")
				For i=lbound(catlist) to ubound(catlist)
					id=catlist(i)
					If (id<>"0") and (id<>"") then
						query="delete from pcCatDiscounts where pcCD_idcategory=" & id
						set rs=connTemp.execute(query)
						set rs=nothing
					end if
				Next

				do while not rstemp.eof
					For i=lbound(catlist) to ubound(catlist)
						id=catlist(i)
						If (id<>"0") and (id<>"") then
							query=""
							query="" & id & ","
							query=query &rstemp("pcCD_quantityFrom")& ","
							query=query &rstemp("pcCD_quantityUntil")& ","
							query=query &rstemp("pcCD_discountPerUnit")& ","
							query=query &rstemp("pcCD_num")& ","
							query=query &rstemp("pcCD_percentage")& ","
							query=query &rstemp("pcCD_discountPerWUnit")& ","
							query=query &rstemp("pcCD_baseproductonly")

							query="insert into pcCatDiscounts (pcCD_idcategory,pcCD_quantityFrom,pcCD_quantityUntil,pcCD_discountPerUnit,pcCD_num,pcCD_percentage,pcCD_discountPerWUnit,pcCD_baseproductonly) values (" & query & ")"
							set rs=conntemp.execute(query)
							set rs=nothing
						end if
					Next
					rstemp.movenext
				loop
			end if
			set rstemp=nothing
			call closedb()
			response.redirect "modDctQtyCat.asp?idcategory=" & idcategory & "&s=1&msg=" & "Quantity discounts were successfully assigned to the selected categories!" 
		else
			set rstemp=nothing
			call closedb()
			response.redirect "modDctQtyCat.asp?idcategory=" & idcategory & "&r=1&msg=" & "The category you selected does not contain any quantity discounts"
		end if
	else
		call closeDb()
		response.redirect "modDctQtyCat.asp?idcategory=" & idcategory & "&r=1&msg=" & "Please select a category before assigning quantity discounts"
	end if
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
				src_FormTitle1="Find Categories"
				src_FormTitle2="Assign quantity discounts to multiple categories"
				src_FormTips1="Use the following filters to look for categories in your store."
				src_FormTips2="Select which categories you would like to apply quantity discounts to."
				src_DisplayType=1
				src_IncNotDisplay=2
				src_ParentOnly=2
				src_ShowLinks=0
				src_FromPage="ApplyDctToCats.asp"
				src_ToPage="ApplyDctToCats.asp?action=apply"
				src_Button1=" Search "
				src_Button2=" Add Discounts to Selected Categories "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcCat_from")=""
				session("srcCat_where")=" AND categories.idcategory<>" & idcategory & " "
			%>
				<!--#include file="inc_srcCATs.asp"-->
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->