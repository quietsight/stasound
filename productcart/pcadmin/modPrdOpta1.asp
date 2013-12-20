<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/currencyformatinc.asp" --> 
<% dim conntemp, query, pIdProduct, pIdoptionGroup, NewGroup

If request.Form("Submit2")<>"" then	
	pIdProduct=request.Form("idProduct")
	pIdoptionGroup=request.Form("idOptionGroup")
	if trim(pIdoptionGroup)="0" then
	   '// nothing selected
	   response.redirect "modPrdOpta1.asp?idProduct="&pIdProduct&"&msg=You must select an Option Group"
	else
		response.redirect "modPrdOpta2.asp?idProduct="&pIdProduct&"&idOptionGroup="&pIdoptionGroup	
	end if
End If

' form parameter 
pIdProduct=request("idProduct")
if trim(pidproduct)="" then
   response.redirect "msg.asp?message=2"
end if

' get item details from db
call openDB()
query="SELECT idProduct, description FROM products WHERE products.idProduct=" &pIdProduct
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
    response.redirect "techErr.asp?error="& Server.Urlencode("Error in modPrdOpta1: "&Err.Description) 
end if

' charge rscordset data into local variables
pIdProduct=rstemp("idProduct")
pDescription=rstemp("description")
pageTitle="Adding Product Options to: <strong>"&pDescription&"</strong>"
%>
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript" type="text/javascript" src="../includes/spry/xpath.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryData.js"></script>
<script type="text/javascript">
	var dscategories = new Spry.Data.XMLDataSet("pcSpryCategoriesXML.asp?CP=3&idRootCategory=<%=pcv_IdRootCategory%>", "categories/category",{sortOnLoad:"categoryDesc",sortOrderOnLoad:"ascending",useCache:false});
	dscategories.setColumnType("idCategory", "number");
	dscategories.setColumnType("idParentCategory", "number");
	var dsProducts = new Spry.Data.XMLDataSet("categories_productsxml.asp?mode=3&x_idCategory={dscategories::idCategory}", "/root/row");
</script>
<form method="post" name="modifyProduct" action="modPrdOpta1.asp" class="pcForms">
<input type="hidden" name="idproduct" value="<%=pIdProduct%>">
<table class="pcCPcontent">
<tr>
	<td>
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
    <h2><strong>ADD</strong> a new option group to the product</h2>
    </td>
</tr>                 
<tr> 
	<td> 
		Available option groups:&nbsp;
		<% 
		query="SELECT idOptionGroup,optionGroupDesc FROM optionsGroups WHERE idoptiongroup>1 ORDER BY optionGroupDesc"
		set rstemp=conntemp.execute(query)
		If rstemp.eof then
			set rstemp=nothing
			call closeDb()
			response.redirect "msg.asp?message=23"
		end if 
		%>								
		<select name="idOptionGroup">
			<option value="0" selected>Select One</option>
			<%
			do until rstemp.eof 
			pIdOptionGroup = rstemp("idOptionGroup")			
				if len(pIdProduct)>0 then
					strSQL="SELECT idOptionGroup, idproduct FROM pcProductsOptions WHERE idproduct="& pIdProduct &" AND idOptionGroup="& pIdOptionGroup &" "
					set rsOptionCheck=conntemp.execute(strSQL)	
					if rsOptionCheck.eof then				
					%>
					<option value="<%=rstemp("idOptionGroup")%>"><%=rstemp("optionGroupDesc")%></option>
					<%
					end if
					set rsOptionCheck = nothing
				end if			
			rstemp.moveNext
			loop
			set rstemp=nothing 
			%>                            
		</select>
	</td>
</tr>  
<tr> 
    <td> 
			<input type="submit" name="Submit2" value="Continue" class="submit2">
			&nbsp;<input TYPE="button" VALUE="Create New Group" onClick="location.href='instOptGrpa.asp?prdFrom=<%=pIdProduct%>'">
	</td>
</tr>
</table>
</form>

<% 
query="SELECT DISTINCT options_optionsGroups.idProduct, products.description FROM options_optionsGroups INNER JOIN products ON options_optionsGroups.idProduct=products.idProduct ORDER BY products.description;"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
If NOT rstemp.eof then 
%>

<div align="center"><img src="images/pc_admin.gif" width="85" height="19" vspace="10"></div>

<form name="formPrdOpt" action="dupOptions.asp" method="post" class="pcForms">
<input type="hidden" name="idproduct" value="<%response.write pIdProduct%>">
<table class="pcCPcontent"> 
	<tr>
		<td><h2><strong>COPY ALL</strong> option groups from another product</h2></td>
	</tr>
	<tr> 
		<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="2">
		  <tr>
		    <td width="13%" nowrap>Select a Product:&nbsp;</td>
		    <td width="87%">
            <div spry:region="dscategories" id="categorySelector">
                <div spry:state="loading"><img src="images/pc_AjaxLoader.gif"/>&nbsp;Loading Categories...</div>
                <div spry:state="ready">
    			<select spry:repeatchildren="dscategories" name="categorySelect" onchange="document.formPrdOpt.iddup.disabled = true; dscategories.setCurrentRowNumber(this.selectedIndex);">
                    <option spry:if="{ds_RowNumber} == {ds_CurrentRowNumber}" value="{idCategory}" selected="selected">{categoryDesc} {pcCats_BreadCrumbs}</option>
                    <option spry:if="{ds_RowNumber} != {ds_CurrentRowNumber}" value="{idCategory}">{categoryDesc} {pcCats_BreadCrumbs}</option>
                </select>
                </div>
            </div>
            </td>
		    </tr>
		  <tr>
		    <td>&nbsp;</td>
		    <td>
            <div spry:region="dsProducts" id="productSelector">
                <div spry:state="loading"><img src="images/pc_AjaxLoader.gif"/>&nbsp;Loading Products...</div>
                <div spry:state="ready">
                <select spry:repeatchildren="dsProducts" id="productSelect" name="iddup">
                    <option spry:if="{ds_RowNumber} == {ds_CurrentRowNumber}" value="{idProduct}" selected="selected">{description}</option>
                    <option spry:if="{ds_RowNumber} != {ds_CurrentRowNumber}" value="{idProduct}">{description}</option>
                </select>
                </div>
            </div>        
            </td>
		    </tr>
		  </table>
		</td>
	</tr>
	<tr>                         
		<td>
			Copy attribute pricing? 
			<input type="radio" name="pricingdup" value="1" class="clearBorder">Yes 
			<input type="radio" name="pricingdup" value="0" checked class="clearBorder">No
		</td>
	</tr>
	<tr>                         
		<td>
			Overwrite existing attributes (if they are the same)?
			<input type="radio" name="updatedup" value="1" class="clearBorder">Yes
			<input type="radio" name="updatedup" value="0" checked class="clearBorder">No, skip the attribute
		</td>
	</tr>
	<tr>                         
		<td><input type="submit" name="Submit1" value="Continue" class="submit2"></td>
	</tr>           
</table>
</form>
<% end if %>

<% 
query="SELECT DISTINCT options_optionsGroups.idOptionGroup, products.description, products.removed, options_optionsGroups.idProduct, optionsGroups.OptionGroupDesc FROM options_optionsGroups , optionsGroups, products WHERE options_optionsGroups.idProduct=products.idProduct and optionsGroups.idOptionGroup = options_optionsGroups.idOptionGroup ORDER BY products.description;"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
If NOT rstemp.eof then 
%>

<div align="center"><img src="images/pc_admin.gif" width="85" height="19" vspace="10"></div>

<form action="dupOptions.asp" method="post" class="pcForms">
<input type="hidden" name="idproduct" value="<%response.write pIdProduct%>">            
<table class="pcCPcontent"> 
<tr>
	<td><h2><strong>COPY</strong> one option group from another product</h2></td>
</tr>                    
<tr> 
	<td>
		Select a Product:&nbsp;
		<select name="iddup">
			<% 
			do until rstemp.eof
			if rstemp("removed")=0 then 
			%>
			<option value="<%=rstemp("idProduct")%>|<%=rstemp("idOptionGroup")%>"><%=rstemp("description")%> - <%=rstemp("OptionGroupDesc")%></option>
			<% 
			end if
			rstemp.moveNext
			loop
			set rstemp=nothing
			call closeDB() 
			%>                          
		</select>
	</td>
</tr>
                      
<tr>                        
	<td>
		Copy attribute pricing? 
		<input type="radio" name="pricingdup" value="1" class="clearBorder"> Yes 
		<input type="radio" name="pricingdup" value="0" checked class="clearBorder"> No
	</td>
</tr>

<tr>                         
	<td>
			Overwrite existing attributes (if they are the same)?
			<input type="radio" name="updatedup" value="1" class="clearBorder"> Yes
		<input type="radio" name="updatedup" value="0" checked class="clearBorder">	No, skip the attribute
	</td>
</tr>
<tr>                         
	<td><input type="submit" name="Submit2" value="Continue" class="submit2"></td>
</tr>
<tr>                         
	<td>&nbsp;</td>
</tr>
<tr>                         
	<td align="center">
		<input type="button" name="back" value="Back" onClick="javascript:history.back()">&nbsp;
		<input type="button" value="Manage Product Options" onClick="location.href='manageOptions.asp'">
	</td>
</tr>
</table>
</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->