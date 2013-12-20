<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Product Options" %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<%
Dim rs, rstemp, connTemp, query, pid

optionGroupID=getUserInput(request.querystring("idOptionGroup"),4)

if not validNum(optionGroupID) then
	msg="Not a valid option group ID"
	response.Redirect "manageOptions.asp?msg="&msg
	response.End()
end if

	call openDb()

	' get products that use option group
	query="SELECT DISTINCT pcProductsOptions.idProduct, pcProductsOptions.idOptionGroup, products.description, products.idproduct, products.removed FROM pcProductsOptions INNER JOIN products ON pcProductsOptions.idProduct = products.idProduct WHERE pcProductsOptions.idOptionGroup = " & optionGroupID & " AND products.removed = 0 ORDER BY products.description ASC"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	
	if rs.EOF then
		set rs=nothing
		call closeDb()
			msg="There are no products using this option group"
			response.Redirect "manageOptions.asp?msg="&msg
			response.End()
	else
		productsArray=rs.getRows()
		intCount=ubound(productsArray,2)
		set rs = nothing
			' Get option group name
			query="SELECT OptionGroupDesc FROM optionsGroups WHERE idOptionGroup = " & optionGroupID
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			pcv_OptionGroupName = rs("OptionGroupDesc")
			set rs=nothing
			call closeDb()
		%>
    	<tr>
      	<td class="pcCPspacer" colspan="2"></td>
      </tr>
    	<tr>
      	<th colspan="2">The following products use the Option Group &quot;<strong><%=pcv_OptionGroupName%></strong>&quot;</th>
      </tr>
    	<tr>
      	<td class="pcCPspacer" colspan="2"></td>
      </tr>
    <%
		For i=0 to intCount
	%>         
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td width="80%"><a href="findproducttype.asp?idproduct=<%=productsArray(3,i)%>"><%=productsArray(2,i)%></a></td>
          		<td width="20%" align="right" nowrap class="cpLinksList"><a href="findproducttype.asp?idproduct=<%=productsArray(3,i)%>">Edit Product</a> - <a href="modPrdOpta.asp?idproduct=<%=productsArray(3,i)%>">Manage Options</a></td>
        </tr>
<% 
		intCount=intCount+1
		next
  End If
%>
	<tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
    	<td colspan="2"><a href="editGrpOptions.asp?idOptionGroup=<%=optionGroupID%>">Remove one or more products from this Option Group &gt;&gt;</a></td>
    </tr>
</table>
<!--#include file="AdminFooter.asp"-->